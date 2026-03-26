// ============================================================
// test.js - コアロジックのユニットテスト（Node.js実行用）
// GASのシート操作を除いた純粋なロジック部分を検証する
// ============================================================

let passed = 0;
let failed = 0;

function assert(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) {
    console.log(`  ✅ ${label}`);
    passed++;
  } else {
    console.log(`  ❌ ${label}`);
    console.log(`     期待値: ${JSON.stringify(expected)}`);
    console.log(`     実際値: ${JSON.stringify(actual)}`);
    failed++;
  }
}

// ============================================================
// テスト対象のロジックをインライン定義（GASから抽出）
// ============================================================

function formatNumber_(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function calculateTaxBreakdown_(items) {
  let subtotal10 = 0, subtotal8 = 0, subtotalNon = 0;
  for (const item of items) {
    if (item.taxCategory === '10%')      subtotal10 += item.amount;
    else if (item.taxCategory === '8%')  subtotal8  += item.amount;
    else                                 subtotalNon += item.amount;
  }
  const tax10 = Math.floor(subtotal10 * 0.10);
  const tax8  = Math.floor(subtotal8  * 0.08);
  const subtotal = subtotal10 + subtotal8 + subtotalNon;
  const grandTotal = subtotal + tax10 + tax8;
  return { subtotal, tax10, tax8, grandTotal };
}

function validateFormData_(formData) {
  if (!formData.docType)       return '書類種別を選択してください。';
  if (!formData.customerName)  return '顧客名を選択してください。';
  if (!formData.issueDate)     return '発行日を入力してください。';
  if (formData.docType === '請求書' && !formData.paymentDeadline) {
    return '請求書の場合、支払期限を入力してください。';
  }
  return null;
}

function getNextDocNumber_(docType, historyData, todayStr) {
  const prefix = (docType === '請求書') ? 'INV' : 'EST';
  const searchPrefix = prefix + '-' + todayStr + '-';
  let maxNum = 0;
  for (let i = 1; i < historyData.length; i++) {
    const existing = historyData[i][0].toString();
    if (existing.indexOf(searchPrefix) === 0) {
      const numPart = parseInt(existing.substring(searchPrefix.length), 10);
      if (numPart > maxNum) maxNum = numPart;
    }
  }
  const nextNum = maxNum + 1;
  const numStr = ('000' + nextNum).slice(-3);
  return prefix + '-' + todayStr + '-' + numStr;
}

function getItemsFromForm_(formItems, itemMasterData) {
  const items = [];
  for (const item of formItems) {
    if (item.name && item.quantity > 0) {
      // 品目マスタ検索
      let masterItem = null;
      for (let i = 1; i < itemMasterData.length; i++) {
        if (itemMasterData[i][1] === item.name) {
          const taxCategory = itemMasterData[i][4];
          if (taxCategory !== '10%' && taxCategory !== '8%' && taxCategory !== '非課税') {
            throw new Error(`品目「${item.name}」の税区分が不正です: "${taxCategory}"`);
          }
          masterItem = { unitPrice: itemMasterData[i][2], unit: itemMasterData[i][3], taxCategory };
          break;
        }
      }
      if (masterItem) {
        items.push({
          name: item.name,
          unitPrice: masterItem.unitPrice,
          unit: masterItem.unit,
          taxCategory: masterItem.taxCategory,
          quantity: item.quantity,
          amount: masterItem.unitPrice * item.quantity,
        });
      }
    }
  }
  return items;
}

// ============================================================
// テスト実行
// ============================================================

// --- [1] formatNumber_ ---
console.log('\n[1] 数値フォーマット');
assert('1,000',          formatNumber_(1000),       '1,000');
assert('1,000,000',      formatNumber_(1000000),    '1,000,000');
assert('0',              formatNumber_(0),           '0');
assert('300,000',        formatNumber_(300000),     '300,000');
assert('1,234,567',      formatNumber_(1234567),    '1,234,567');

// --- [2] calculateTaxBreakdown_ ---
console.log('\n[2] 税計算');

// パターン1: 10%のみ
assert('10%のみ: 100,000円 → 税10,000円 合計110,000円',
  calculateTaxBreakdown_([{ amount: 100000, taxCategory: '10%' }]),
  { subtotal: 100000, tax10: 10000, tax8: 0, grandTotal: 110000 }
);

// パターン2: 8%のみ
assert('8%のみ: 100,000円 → 税8,000円 合計108,000円',
  calculateTaxBreakdown_([{ amount: 100000, taxCategory: '8%' }]),
  { subtotal: 100000, tax10: 0, tax8: 8000, grandTotal: 108000 }
);

// パターン3: 非課税のみ
assert('非課税のみ: 50,000円 → 税0円 合計50,000円',
  calculateTaxBreakdown_([{ amount: 50000, taxCategory: '非課税' }]),
  { subtotal: 50000, tax10: 0, tax8: 0, grandTotal: 50000 }
);

// パターン4: 10% + 8% + 非課税 混在
assert('混在: 10%(50,000)+8%(20,000)+非課税(10,000) → 税6,600 合計86,600',
  calculateTaxBreakdown_([
    { amount: 50000, taxCategory: '10%' },
    { amount: 20000, taxCategory: '8%' },
    { amount: 10000, taxCategory: '非課税' },
  ]),
  { subtotal: 80000, tax10: 5000, tax8: 1600, grandTotal: 86600 }
);

// パターン5: 端数切り捨て（10%対象33,333円 → 税3,333（切捨）→ 合計36,666）
assert('端数切り捨て: 33,333円 × 10% → 税3,333 合計36,666',
  calculateTaxBreakdown_([{ amount: 33333, taxCategory: '10%' }]),
  { subtotal: 33333, tax10: 3333, tax8: 0, grandTotal: 36666 }
);

// パターン6: 複数品目（同税率）
assert('10%品目3件の合算',
  calculateTaxBreakdown_([
    { amount: 300000, taxCategory: '10%' },
    { amount: 80000,  taxCategory: '10%' },
    { amount: 15000,  taxCategory: '10%' },
  ]),
  { subtotal: 395000, tax10: 39500, tax8: 0, grandTotal: 434500 }
);

// --- [3] validateFormData_ ---
console.log('\n[3] バリデーション');

assert('書類種別なし → エラー',
  validateFormData_({ docType: '', customerName: '株式会社A', issueDate: '2026/03/26', paymentDeadline: '' }),
  '書類種別を選択してください。'
);
assert('顧客名なし → エラー',
  validateFormData_({ docType: '請求書', customerName: '', issueDate: '2026/03/26', paymentDeadline: '2026/04/25' }),
  '顧客名を選択してください。'
);
assert('発行日なし → エラー',
  validateFormData_({ docType: '請求書', customerName: '株式会社A', issueDate: '', paymentDeadline: '2026/04/25' }),
  '発行日を入力してください。'
);
assert('請求書で支払期限なし → エラー',
  validateFormData_({ docType: '請求書', customerName: '株式会社A', issueDate: '2026/03/26', paymentDeadline: '' }),
  '請求書の場合、支払期限を入力してください。'
);
assert('見積書で支払期限なし → OK（エラーなし）',
  validateFormData_({ docType: '見積書', customerName: '株式会社A', issueDate: '2026/03/26', paymentDeadline: '' }),
  null
);
assert('全項目入力済み（請求書） → OK',
  validateFormData_({ docType: '請求書', customerName: '株式会社A', issueDate: '2026/03/26', paymentDeadline: '2026/04/25' }),
  null
);

// --- [4] getNextDocNumber_ ---
console.log('\n[4] 発行番号採番');
const TODAY = '20260326';

assert('履歴なし → INV-20260326-001',
  getNextDocNumber_('請求書', [['']], TODAY),
  'INV-20260326-001'
);
assert('同日に1件あり → INV-20260326-002',
  getNextDocNumber_('請求書', [[''], ['INV-20260326-001']], TODAY),
  'INV-20260326-002'
);
assert('同日に9件あり → INV-20260326-010',
  getNextDocNumber_('請求書', [
    [''],
    ['INV-20260326-001'],['INV-20260326-002'],['INV-20260326-003'],
    ['INV-20260326-004'],['INV-20260326-005'],['INV-20260326-006'],
    ['INV-20260326-007'],['INV-20260326-008'],['INV-20260326-009'],
  ], TODAY),
  'INV-20260326-010'
);
assert('別日の履歴のみ → INV-20260326-001（別日は無視）',
  getNextDocNumber_('請求書', [[''], ['INV-20260325-001'], ['INV-20260324-005']], TODAY),
  'INV-20260326-001'
);
assert('見積書の採番 → EST-20260326-001',
  getNextDocNumber_('見積書', [['']], TODAY),
  'EST-20260326-001'
);
assert('請求書・見積書混在履歴 → それぞれ独立',
  getNextDocNumber_('見積書', [[''], ['INV-20260326-001'], ['EST-20260326-001']], TODAY),
  'EST-20260326-002'
);

// --- [5] getItemsFromForm_ ---
console.log('\n[5] 品目フォームデータ処理');

const MASTER = [
  ['品目ID', '品名', '単価', '単位', '税区分'],
  ['I001', 'Webサイト制作', 300000, '式', '10%'],
  ['I002', 'コンサルティング', 10000, '時間', '10%'],
  ['I010', '食品加工費', 20000, '式', '8%'],
  ['I011', '印紙代', 400, '枚', '非課税'],
];

assert('1品目（数量2）→ 金額600,000円',
  getItemsFromForm_([{ name: 'Webサイト制作', quantity: 2 }], MASTER).map(i => i.amount),
  [600000]
);
assert('数量0の品目は除外される',
  getItemsFromForm_([
    { name: 'Webサイト制作', quantity: 0 },
    { name: 'コンサルティング', quantity: 3 },
  ], MASTER).map(i => i.name),
  ['コンサルティング']
);
assert('品名なしの品目は除外される',
  getItemsFromForm_([
    { name: '', quantity: 5 },
    { name: 'Webサイト制作', quantity: 1 },
  ], MASTER).map(i => i.name),
  ['Webサイト制作']
);
assert('全品目空欄 → 空配列',
  getItemsFromForm_([
    { name: '', quantity: 0 },
    { name: '', quantity: 0 },
  ], MASTER),
  []
);
assert('8%・非課税品目も取得できる',
  getItemsFromForm_([
    { name: '食品加工費', quantity: 1 },
    { name: '印紙代', quantity: 3 },
  ], MASTER).map(i => ({ name: i.name, taxCategory: i.taxCategory, amount: i.amount })),
  [
    { name: '食品加工費', taxCategory: '8%', amount: 20000 },
    { name: '印紙代', taxCategory: '非課税', amount: 1200 },
  ]
);

// 不正な税区分
assert('不正な税区分 → エラーをthrow',
  (() => {
    try {
      getItemsFromForm_([{ name: '食品加工費', quantity: 1 }], [
        ['品目ID', '品名', '単価', '単位', '税区分'],
        ['I010', '食品加工費', 20000, '式', '5%'],  // 不正
      ]);
      return 'エラーなし';
    } catch (e) {
      return e.message.includes('税区分が不正') ? 'エラーthrow成功' : e.message;
    }
  })(),
  'エラーthrow成功'
);

// --- [6] 統合シナリオ ---
console.log('\n[6] 統合シナリオ');

// 請求書：Webサイト制作(10%) + 食品加工費(8%) + 印紙代(非課税)
const invoiceItems = getItemsFromForm_([
  { name: 'Webサイト制作', quantity: 1 },
  { name: '食品加工費', quantity: 2 },
  { name: '印紙代', quantity: 4 },
], MASTER);
const invoiceTotals = calculateTaxBreakdown_(invoiceItems);
assert('請求書・混在品目の合計金額',
  invoiceTotals,
  {
    subtotal: 300000 + 40000 + 1600,  // 341600
    tax10: 30000,
    tax8:  3200,
    grandTotal: 341600 + 30000 + 3200, // 374800
  }
);

// 見積書：単一品目（コンサルティング10時間）
const quoteItems = getItemsFromForm_([
  { name: 'コンサルティング', quantity: 10 },
], MASTER);
const quoteTotals = calculateTaxBreakdown_(quoteItems);
assert('見積書・単一品目',
  quoteTotals,
  { subtotal: 100000, tax10: 10000, tax8: 0, grandTotal: 110000 }
);

// ============================================================
// 結果サマリー
// ============================================================
console.log('\n' + '='.repeat(50));
console.log(`テスト結果: ${passed + failed}件中 ${passed}件成功 / ${failed}件失敗`);
if (failed === 0) {
  console.log('🎉 全テスト通過！');
} else {
  console.log(`⚠️  ${failed}件のテストが失敗しました。`);
  process.exit(1);
}
