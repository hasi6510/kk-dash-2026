/**
 * 家計ダッシュボード v2 - データ更新スクリプト
 * 使い方: node update.js
 *
 * MoneyForward ME の CSVを mf_csv/ フォルダに入れてから実行してください。
 * ※ CSVはExcelで「CSV UTF-8 (コンマ区切り)」で保存し直してください。
 *
 * 【高配当株シート】stocks.xlsx を同フォルダに置くと配当データを自動読み込みします。
 *   Googleスプレッドシート → ファイル → ダウンロード → Microsoft Excel (.xlsx)
 *   → ファイル名を stocks.xlsx に変更して保存
 *
 * 【サブスクシート】subscriptions.xlsx を同フォルダに置くと最新の「verX」タブを自動読み込みします。
 *   タブ名例: "2026年改善ver2", "2026年改善ver3" → 最も大きいverを使用
 *   ファイル → ダウンロード → Microsoft Excel (.xlsx) → subscriptions.xlsx に名前変更して保存
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const os   = require('os');

const DIR    = __dirname;
const MF_DIR = path.join(DIR, 'mf_csv');

// exceljs パス（既存インストール場所）
const EXCELJS_PATH = path.join(os.tmpdir(), 'xlsx_work', 'node_modules', 'exceljs');

// ===== 1. Load budget config =====
const budgetPath = path.join(DIR, 'budget.json');
if (!fs.existsSync(budgetPath)) {
  console.error('❌ budget.json が見つかりません: ' + budgetPath);
  process.exit(1);
}
const budget = JSON.parse(fs.readFileSync(budgetPath, 'utf8'));
console.log('✅ budget.json 読み込み完了 (対象月: ' + budget.month + ')');

// ===== Main async function =====
async function main() {

  // ===== 2a. Try loading subscriptions.xlsx for subscription detail =====
  let subscriptionsOverride = null;
  const subsXlsx = path.join(DIR, 'subscriptions.xlsx');

  if (fs.existsSync(subsXlsx) && fs.existsSync(EXCELJS_PATH)) {
    try {
      const ExcelJS = require(EXCELJS_PATH);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(subsXlsx);

      // Find the latest "verX" sheet
      const sheet = findLatestVerSheet(wb);
      console.log('📺 subscriptions.xlsx 読み込み: シート「' + sheet.name + '」を使用');

      subscriptionsOverride = extractSubscriptionData(sheet);
      if (subscriptionsOverride) {
        const total = subscriptionsOverride.reduce((s, x) => s + x.amount, 0);
        console.log('  アクティブなサブスク: ' + subscriptionsOverride.length + '件 合計 ' + formatYen(total) + '/月');
      }
    } catch (e) {
      console.log('⚠ subscriptions.xlsx の読み込みをスキップしました: ' + e.message);
    }
  } else if (fs.existsSync(subsXlsx)) {
    console.log('⚠ subscriptions.xlsx が見つかりましたが exceljs がありません。budget.jsonの値を使用します。');
  }

  // ===== 2b. Try loading stocks.xlsx for dividend data =====
  let dividendOverride = null;
  const stocksXlsx = path.join(DIR, 'stocks.xlsx');

  if (fs.existsSync(stocksXlsx) && fs.existsSync(EXCELJS_PATH)) {
    try {
      const ExcelJS = require(EXCELJS_PATH);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(stocksXlsx);

      // Find the latest-dated sheet
      const sheet = findLatestSheet(wb);
      console.log('📈 stocks.xlsx 読み込み: シート「' + sheet.name + '」を使用');

      dividendOverride = extractDividendData(sheet);
      if (dividendOverride) {
        console.log('  年間配当: ' + formatYen(dividendOverride.annual_total));
        console.log('  投資元本: ' + formatYen(dividendOverride.principal));
      }
    } catch (e) {
      console.log('⚠ stocks.xlsx の読み込みをスキップしました: ' + e.message);
    }
  } else if (fs.existsSync(stocksXlsx)) {
    console.log('⚠ stocks.xlsx が見つかりましたが exceljs がありません。budget.jsonの値を使用します。');
  }

  // ===== 2c. Check for mf_dashboard_input.json (Claude自動生成JSON) =====
  const jsonInputPath = path.join(DIR, 'mf_dashboard_input.json');
  let jsonInput = null;
  if (fs.existsSync(jsonInputPath)) {
    try {
      const parsed = JSON.parse(fs.readFileSync(jsonInputPath, 'utf8'));
      if (parsed.month !== budget.month) {
        console.log('⚠ mf_dashboard_input.json の月(' + parsed.month + ')が budget.json の月(' + budget.month + ')と一致しません。無視します。');
      } else {
        jsonInput = parsed;
        console.log('✅ mf_dashboard_input.json を使用します (月: ' + jsonInput.month + ')');
      }
    } catch (e) {
      console.log('⚠ mf_dashboard_input.json の読み込みに失敗しました: ' + e.message);
    }
  }

  // ===== 3. Find newest CSV in mf_csv/ =====
  if (!fs.existsSync(MF_DIR)) {
    fs.mkdirSync(MF_DIR, { recursive: true });
    console.log('📁 mf_csv フォルダを作成しました');
  }

  const csvFiles = fs.readdirSync(MF_DIR)
    .filter(f => f.toLowerCase().endsWith('.csv'))
    .map(f => ({ name: f, mtime: fs.statSync(path.join(MF_DIR, f)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);

  // ===== 4. Parse CSV or use JSON input =====
  let transactions = [];
  let csvSource    = '（CSVなし）';
  let datePeriod   = { from: budget.month + '-01', to: new Date().toISOString().slice(0,10) };
  const unmappedSet = new Set();

  // ===== 5. Aggregate spending by category =====
  const spendingMap = {};
  budget.variable_categories.forEach(c => { spendingMap[c.name] = 0; });

  if (jsonInput) {
    // --- JSON入力モード（Claude自動生成） ---
    csvSource = 'mf_dashboard_input.json (Claude自動取得)';
    if (jsonInput.data_period) datePeriod = jsonInput.data_period;
    jsonInput.variable_spending.forEach(item => {
      if (spendingMap[item.name] !== undefined) spendingMap[item.name] = item.actual || 0;
    });
    console.log('📊 JSONデータを使用: ' + jsonInput.variable_spending.length + ' カテゴリ');
  } else if (csvFiles.length === 0) {
    console.log('⚠ mf_csv フォルダにCSVファイルがありません。ゼロ実績でdata.jsonを生成します。');
    console.log('  → MoneyForwardからCSVをダウンロードして mf_csv/ フォルダに入れてください。');
  } else {
    // --- CSVモード（従来通り） ---
    if (csvFiles.length > 1) {
      console.log('ℹ 複数のCSVが見つかりました。最新のファイルを使用します:');
      csvFiles.forEach((f, i) => console.log('   ' + (i === 0 ? '▶' : ' ') + ' ' + f.name));
    }
    const csvPath = path.join(MF_DIR, csvFiles[0].name);
    csvSource = csvFiles[0].name;
    console.log('📂 CSVファイル: ' + csvSource);

    let raw;
    try {
      raw = fs.readFileSync(csvPath, 'utf8').replace(/^﻿/, '');
    } catch (e) {
      console.error('❌ CSVの読み込みに失敗しました: ' + e.message);
      console.log('  → ファイルがUTF-8で保存されているか確認してください。');
      process.exit(1);
    }

    const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    const targetMonth = budget.month.replace('-', '/');
    let dates = [];

    for (const line of lines) {
      const cols = parseCSVLine(line);
      if (cols.length < 8) continue;

      const [date, desc, amountRaw, institution, major, minor, memo, transfer] = cols;

      if (!/^\d{4}\/\d{2}\/\d{2}$/.test(date)) continue;
      if (!date.startsWith(targetMonth)) continue;
      if (transfer === '1' || transfer === 'true' || transfer === '振替') continue;

      const amount = parseFloat(amountRaw.replace(/,/g, '').replace(/"/g, ''));
      if (isNaN(amount) || amount >= 0) continue;

      const mfCat = minor || major || '';
      const appCat = mapCategory(mfCat, major, budget.mf_category_map, unmappedSet);

      if (appCat === '固定費') continue;

      transactions.push({ date, description: desc, amount: Math.abs(amount), mf_category: mfCat, category: appCat });
      dates.push(date);
    }

    console.log('📊 取引件数: ' + transactions.length + ' 件（対象月: ' + targetMonth + '）');

    if (unmappedSet.size > 0) {
      console.log('⚠ 未マップカテゴリ（「その他」に計上）:');
      unmappedSet.forEach(c => console.log('   - ' + c));
    }

    if (dates.length > 0) {
      dates.sort();
      datePeriod = { from: dates[0].replace(/\//g, '-'), to: dates[dates.length - 1].replace(/\//g, '-') };
    }

    transactions.forEach(t => {
      if (spendingMap[t.category] !== undefined) spendingMap[t.category] += t.amount;
      else spendingMap['その他'] = (spendingMap['その他'] || 0) + t.amount;
    });
  }

  // ===== 6. Build variable_spending with traffic lights =====

  const variable_spending = budget.variable_categories.map(c => {
    const actual = Math.round(spendingMap[c.name] || 0);
    const pct    = c.budget > 0 ? (actual / c.budget) * 100 : 0;
    const light  = pct < 70 ? 'green' : pct < 100 ? 'yellow' : 'red';
    return { name: c.name, budget: c.budget, actual, pct: Math.round(pct * 10) / 10, traffic_light: light };
  });

  // ===== 7. Totals =====
  const fixed_total           = budget.fixed_costs.reduce((s, fc) => s + fc.amount, 0);
  const variable_budget_total = budget.variable_categories.reduce((s, c) => s + c.budget, 0);
  const variable_actual_total = variable_spending.reduce((s, c) => s + c.actual, 0);
  const grand_budget          = fixed_total + variable_budget_total;
  const grand_actual          = fixed_total + variable_actual_total;

  // ===== 8. Dividend data (from xlsx override or budget.json) =====
  const annual_div  = dividendOverride ? dividendOverride.annual_total  : budget.dividends.annual_total;
  const monthly_div = dividendOverride ? Math.round(annual_div / 12)    : budget.dividends.monthly_total;
  const principal   = dividendOverride ? dividendOverride.principal      : (budget.investment_principal || 0);

  // ===== 9. Dividend coverage (smallest fixed cost first) =====
  const fixedSorted = [...budget.fixed_costs].sort((a, b) => a.amount - b.amount);
  let cumulative = 0;
  const covers = fixedSorted.map(fc => {
    cumulative += fc.amount;
    return { name: fc.name, amount: fc.amount, covered: monthly_div >= cumulative, cumulative };
  });
  const coverage_pct = fixed_total > 0 ? Math.round((monthly_div / fixed_total) * 1000) / 10 : 0;

  // ===== 10. r > g =====
  const income_monthly = budget.income_monthly  || 0;
  const inv_monthly    = budget.investment_monthly || 0;
  const g_pct          = budget.salary_growth_rate || 3.0;
  const r_pct          = principal > 0 ? Math.round((annual_div / principal) * 10000) / 100 : 0;
  const r_gt_g         = r_pct > g_pct;

  // ===== 11. Investment panel =====
  const investment_ratio_pct = income_monthly > 0
    ? Math.round((inv_monthly / income_monthly) * 1000) / 10
    : 0;
  const investment_surplus = income_monthly - fixed_total - variable_actual_total - inv_monthly;

  // ===== 12. Highlights（今月良かったこと） =====
  const highlights = [];
  variable_spending
    .filter(c => c.actual > 0 && c.pct < 50)
    .forEach(c => highlights.push(c.name + 'を予算の' + c.pct + '%で抑えられました！'));

  if (variable_spending.every(c => c.traffic_light === 'green') && variable_actual_total > 0) {
    highlights.push('全カテゴリが予算内です！素晴らしい家計管理です🎉');
  }
  if (r_gt_g) {
    highlights.push('r(' + r_pct.toFixed(2) + '%) > g(' + g_pct + '%) 達成中！資産所得が賃金上昇率を上回っています📈');
  }
  if (coverage_pct >= 20) {
    highlights.push('配当が固定費の' + coverage_pct + '%をカバーしています。着実に資産が育っています！');
  }

  // ===== 13. Improvements（来月の改善案） =====
  const improvements = [];
  variable_spending
    .filter(c => c.traffic_light === 'red')
    .forEach(c => {
      const over = formatYen(c.actual - c.budget);
      improvements.push({ icon: '🔴', text: c.name + 'が予算超過(' + c.pct + '%)です。' + over + '円オーバー。来月は予算を増やすか支出を見直しましょう。' });
    });
  variable_spending
    .filter(c => c.pct >= 80 && c.traffic_light === 'yellow')
    .forEach(c => {
      improvements.push({ icon: '⚠️', text: c.name + 'が' + c.pct + '%使用中です。月末まで注意して使いましょう。' });
    });
  if (investment_ratio_pct < 10 && income_monthly > 0) {
    improvements.push({ icon: '💡', text: '投資比率が' + investment_ratio_pct + '%です。余力があれば積み立てを増やしましょう（目標：収入の10〜20%）。' });
  }
  if (coverage_pct < 30) {
    improvements.push({ icon: '📈', text: '配当で固定費の' + coverage_pct + '%をカバー中。高配当株の買い増しでカバー率を上げていきましょう！' });
  }
  if (investment_surplus < 0) {
    improvements.push({ icon: '🔴', text: '今月は投資余力がマイナス(' + formatYen(investment_surplus) + ')です。支出を見直してみましょう。' });
  }

  // ===== 14. Build final data object =====
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const generated_at = jst.toISOString().replace('Z', '+09:00');

  const data = {
    generated_at,
    month: budget.month,
    data_period: datePeriod,
    csv_source: csvSource,
    variable_spending,
    fixed_costs: budget.fixed_costs,
    fixed_total,
    variable_budget_total,
    variable_actual_total,
    grand_budget,
    grand_actual,
    dividends: {
      monthly_total: monthly_div,
      annual_total:  annual_div,
      coverage_pct,
      covers
    },
    investment: {
      monthly_amount:    inv_monthly,
      income_monthly:    income_monthly,
      ratio_pct:         investment_ratio_pct,
      surplus:           investment_surplus
    },
    r_vs_g: {
      r_pct,
      g_pct,
      r_gt_g,
      annual_dividend: annual_div,
      principal
    },
    highlights,
    improvements,
    subscriptions_detail: subscriptionsOverride || budget.subscriptions_detail || [],
    subscriptions_source: subscriptionsOverride ? 'subscriptions.xlsx' : 'budget.json',
    unmapped_categories: Array.from(unmappedSet)
  };

  // ===== 15. Write output files =====
  fs.writeFileSync(path.join(DIR, 'data.json'), JSON.stringify(data, null, 2), 'utf8');
  console.log('✅ data.json を書き出しました');

  const inlineContent = [
    '// Auto-generated by update.js — do not edit manually',
    '// 更新日時: ' + generated_at,
    'const BUDGET_DATA = ' + JSON.stringify(data, null, 2) + ';'
  ].join('\n') + '\n';
  fs.writeFileSync(path.join(DIR, 'data_inline.js'), inlineContent, 'utf8');
  console.log('✅ data_inline.js を書き出しました');

  // ===== 16. Console summary =====
  console.log('');
  console.log('========================================');
  console.log('  更新完了: ' + new Date().toLocaleString('ja-JP'));
  console.log('========================================');
  console.log('  変動費実績: ' + formatYen(variable_actual_total) + ' / ' + formatYen(variable_budget_total));
  console.log('  固定費合計: ' + formatYen(fixed_total));
  console.log('  月間合計:   ' + formatYen(grand_actual) + ' / ' + formatYen(grand_budget));
  console.log('  月配当:     ' + formatYen(monthly_div) + ' （固定費の' + coverage_pct + '%カバー）');
  console.log('  r vs g:     r=' + r_pct.toFixed(2) + '% ' + (r_gt_g ? '>' : '<') + ' g=' + g_pct + '% → ' + (r_gt_g ? '✅ r>g達成！' : '❌ まだg>r'));
  console.log('  投資余力:   ' + formatYen(investment_surplus));
  console.log('');
  if (variable_spending.some(c => c.traffic_light === 'red')) {
    console.log('  🔴 予算超過: ' + variable_spending.filter(c => c.traffic_light === 'red').map(c => c.name).join(', '));
  } else if (variable_spending.some(c => c.traffic_light === 'yellow')) {
    console.log('  🟡 要注意:   ' + variable_spending.filter(c => c.traffic_light === 'yellow').map(c => c.name).join(', '));
  } else {
    console.log('  🟢 全カテゴリ予算内です！');
  }
  console.log('');
  console.log('  index.html をブラウザで開いて確認してください。');
}

// ===== HELPERS =====

/**
 * サブスクシートの最新「verX」タブを見つける
 * 例: "2026年改善ver2", "2026年改善ver3" → 数字が最大のものを返す
 */
function findLatestVerSheet(workbook) {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) throw new Error('シートが見つかりません');
  if (sheets.length === 1) return sheets[0];

  // ver番号でスコアリング
  const scored = sheets.map(ws => {
    const m = ws.name.match(/ver\s*(\d+)/i);
    return { ws, ver: m ? parseInt(m[1]) : 0 };
  });

  // 最大ver番号のシートを返す（同じなら後ろのシートを優先）
  scored.sort((a, b) => b.ver !== a.ver ? b.ver - a.ver : b.ws.id - a.ws.id);
  return scored[0].ws;
}

/**
 * サブスクシートからアクティブなサービス一覧を抽出する
 * - 終了日が過去 → 除外（解約済み）
 * - 年払い → 月換算（÷12、または"3年"を含む場合は÷36）
 * - ジムなど固定費扱いのものは note に "固定費計上済み" を付ける
 */
function extractSubscriptionData(sheet) {
  const today = new Date();
  const results = [];

  // ヘッダー行を探す（「サービス名」を含む行）
  let headerRow = null;
  let headerRowNum = 0;
  let colMap = {}; // { name, type, amount, startDate, endDate, memo }

  sheet.eachRow((row, rowNum) => {
    if (headerRow) return; // 既に見つかったらスキップ
    const vals = [];
    row.eachCell({ includeEmpty: true }, cell => vals.push(String(cell.value || '')));
    // "サービス名" か "サービス" が含まれる行をヘッダーとして認識
    if (vals.some(v => v.includes('サービス') || v.toLowerCase().includes('service'))) {
      headerRow = vals;
      headerRowNum = rowNum;
      vals.forEach((v, i) => {
        const lv = v.trim();
        if (/サービス名|サービス/.test(lv))   colMap.name      = i;
        if (/支払い方式|支払方式|頻度/.test(lv)) colMap.type      = i;
        if (/金額|料金|費用/.test(lv))          colMap.amount    = i;
        if (/開始/.test(lv))                    colMap.startDate = i;
        if (/終了/.test(lv))                    colMap.endDate   = i;
        if (/メモ|備考|note/i.test(lv))         colMap.memo      = i;
      });
    }
  });

  if (!headerRow) {
    console.log('  ⚠ subscriptions.xlsx のヘッダー行が見つかりませんでした');
    return null;
  }

  // データ行を処理
  sheet.eachRow((row, rowNum) => {
    if (rowNum <= headerRowNum) return; // ヘッダー行以前はスキップ

    const vals = [];
    row.eachCell({ includeEmpty: true }, cell => {
      // 日付型セルの場合は文字列に変換
      let v = cell.value;
      if (v instanceof Date) v = v.toISOString().slice(0, 7); // YYYY-MM
      vals.push(v === null || v === undefined ? '' : String(v));
    });

    const name   = colMap.name      !== undefined ? vals[colMap.name]?.trim()  : '';
    const type   = colMap.type      !== undefined ? vals[colMap.type]?.trim()  : '月払い';
    const rawAmt = colMap.amount    !== undefined ? vals[colMap.amount]        : '';
    const endRaw = colMap.endDate   !== undefined ? vals[colMap.endDate]?.trim() : '';
    const memo   = colMap.memo      !== undefined ? vals[colMap.memo]?.trim()  : '';

    if (!name || name === '') return; // 空行スキップ
    if (/^[\d]+$/.test(name)) return; // 番号だけの行はスキップ

    // 終了日チェック（過去なら除外）
    if (endRaw && endRaw !== '' && endRaw !== 'null') {
      const endDate = parseJapaneseDate(endRaw);
      if (endDate && endDate < today) {
        console.log('  ✂ 除外（解約済み）: ' + name + ' （終了: ' + endRaw + '）');
        return;
      }
    }

    // 金額パース
    const amountRaw = parseFloat(String(rawAmt).replace(/[¥,円\s]/g, ''));
    if (isNaN(amountRaw) || amountRaw <= 0) return;

    // 月換算
    let monthlyAmount = amountRaw;
    const is3Year = /3年/.test(name) || /3年/.test(memo);
    if (/年払い|年額|年間/.test(type)) {
      monthlyAmount = is3Year ? Math.round(amountRaw / 36) : Math.round(amountRaw / 12);
    }

    // ジムなど固定費扱いのものを識別（メモやサービス名に「固定費」を含む場合）
    const isFixed = /ジム|固定費/.test(name) || /ジム|固定費/.test(memo);

    results.push({
      name,
      amount: monthlyAmount,
      note: (isFixed ? '固定費計上済み・' : '') +
            (/年払い|年額|年間/.test(type) ? (is3Year ? '3年払い月換算' : '年払い月換算') : '') +
            (memo && memo !== 'null' ? (memo.length > 0 ? memo : '') : '')
    });
  });

  return results.length > 0 ? results : null;
}

/**
 * 日本語日付文字列をDateオブジェクトに変換
 * 対応フォーマット: "2025/01", "2025年1月", "2025-01"
 */
function parseJapaneseDate(str) {
  if (!str) return null;
  // YYYY/MM or YYYY-MM
  let m = str.match(/(\d{4})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, 1);
  // YYYY年MM月
  m = str.match(/(\d{4})年(\d{1,2})月/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, 1);
  return null;
}

function findLatestSheet(workbook) {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) throw new Error('シートが見つかりません');
  if (sheets.length === 1) return sheets[0];

  // Try to find the sheet with the latest date in its name
  const dated = sheets.map(ws => {
    const name = ws.name;
    // Try various date patterns: YYYY-MM, YYYY/MM, YYYY年MM月, YY.MM, YYYYMMDD
    const patterns = [
      /(\d{4})[-\/](\d{1,2})/,
      /(\d{4})年(\d{1,2})月/,
      /(\d{2})\.(\d{2})/,
      /(\d{4})(\d{2})\d{2}/
    ];
    for (const pat of patterns) {
      const m = name.match(pat);
      if (m) {
        const year  = parseInt(m[1]) < 100 ? 2000 + parseInt(m[1]) : parseInt(m[1]);
        const month = parseInt(m[2]);
        return { ws, date: year * 100 + month };
      }
    }
    return { ws, date: 0 };
  });

  const hasDated = dated.filter(d => d.date > 0);
  if (hasDated.length > 0) {
    hasDated.sort((a, b) => b.date - a.date);
    return hasDated[0].ws;
  }

  // No dates found — use the last sheet
  return sheets[sheets.length - 1];
}

function extractDividendData(sheet) {
  let totalDividend = 0;
  let totalPrincipal = 0;
  let foundDividend  = false;
  let foundPrincipal = false;

  // Scan all rows looking for summary/total rows
  sheet.eachRow((row, rowNum) => {
    const cells = [];
    row.eachCell({ includeEmpty: true }, (cell) => {
      cells.push(cell.value);
    });

    const rowText = cells.map(c => c === null || c === undefined ? '' : String(c)).join('|');

    // Look for annual dividend total row
    // Common patterns: 配当合計, 年間配当, 合計 in column A with a number in another column
    if (/配当(合計|総計|計|金額)|年間配当/.test(rowText)) {
      const nums = cells.filter(c => typeof c === 'number' && c > 1000);
      if (nums.length > 0) {
        totalDividend = Math.max(...nums);
        foundDividend = true;
      }
    }

    // Look for acquisition/principal total row
    if (/取得(金額|価額|合計|総計)|投資(元本|合計)|買付(合計|金額)/.test(rowText)) {
      const nums = cells.filter(c => typeof c === 'number' && c > 10000);
      if (nums.length > 0) {
        totalPrincipal = Math.max(...nums);
        foundPrincipal = true;
      }
    }

    // Also check for rows labeled 合計 or 総計 with large numbers (potential totals)
    if (/^(合計|総計|TOTAL|total)$/.test(String(cells[0] || '').trim())) {
      const nums = cells.filter(c => typeof c === 'number');
      if (nums.length >= 2) {
        // Heuristic: if we haven't found yet, largest number might be principal
        if (!foundPrincipal && Math.max(...nums) > 100000) {
          totalPrincipal = Math.max(...nums);
          foundPrincipal = true;
        }
      }
    }
  });

  if (foundDividend || foundPrincipal) {
    return {
      annual_total: foundDividend  ? totalDividend  : null,
      principal:    foundPrincipal ? totalPrincipal : null
    };
  }

  console.log('  ⚠ stocks.xlsx から配当・元本の合計行が見つかりませんでした。budget.jsonの値を使用します。');
  console.log('  ヒント: シートに「配当合計」「取得金額」などの行があることを確認してください。');
  return null;
}

function mapCategory(minor, major, categoryMap, unmappedSet) {
  if (minor && categoryMap[minor]) return categoryMap[minor];
  if (major && categoryMap[major]) return categoryMap[major];
  const candidates = Object.keys(categoryMap);
  for (const key of candidates) {
    if (minor && minor.includes(key)) return categoryMap[key];
    if (major && major.includes(key)) return categoryMap[key];
  }
  const label = minor || major || '（不明）';
  if (label && label !== '（不明）') unmappedSet.add(label);
  return 'その他';
}

function parseCSVLine(line) {
  const result = [];
  let current  = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result;
}

function formatYen(n) {
  return '¥' + Math.round(n).toLocaleString('ja-JP');
}

// Run
main().catch(e => {
  console.error('❌ エラー: ' + e.message);
  process.exit(1);
});
