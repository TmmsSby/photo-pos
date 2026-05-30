/**
 * 写真展 POS - Google Apps Script バックエンド
 * 
 * 【シート構成】
 *   売上明細     - 精算データ（自動生成）
 *   商品マスター  - 商品一覧（自動生成）
 *   担当者マスター - 担当者一覧（自動生成）
 *
 * 【デプロイ設定】
 *   種類: Webアプリ
 *   次のユーザーとして実行: 自分
 *   アクセスできるユーザー: 全員
 */

// ============================================================
// HTMLを配信（GETリクエスト → POSアプリ本体を返す）
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // 接続テスト用
  if (action === 'ping') {
    return ContentService
      .createTextOutput(JSON.stringify({status:'ok', message:'POS GAS接続確認OK'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // POSアプリ本体HTMLを配信
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('写真展 POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// APIエンドポイント（POSTリクエスト）
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    switch (action) {
      case 'getMasters':  return jsonResponse(getMasters());
      case 'saveMaster':  return jsonResponse(saveMaster(payload.type, payload.data));
      case 'addSales':    return jsonResponse(addSales(payload.rows));
      default:
        return jsonResponse({status:'error', message:'Unknown action: ' + action});
    }
  } catch(err) {
    return jsonResponse({status:'error', message: err.toString()});
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// マスターデータ取得
// ============================================================
function getMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 商品マスター
  const products = [];
  const pSheet = getOrCreateSheet(ss, '商品マスター');
  if (pSheet.getLastRow() <= 1) {
    initProductSheet(pSheet);
  }
  const pData = pSheet.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) {
    const row = pData[i];
    if (!row[0] && !row[1]) continue; // 空行スキップ
    products.push({
      code:  String(row[0] || ''),
      name:  String(row[1] || ''),
      price: Number(row[2] || 0),
      cat:   String(row[3] || 'goods'),
      sub:   String(row[4] || '')
    });
  }

  // 担当者マスター
  const staff = [];
  const sSheet = getOrCreateSheet(ss, '担当者マスター');
  if (sSheet.getLastRow() <= 1) {
    initStaffSheet(sSheet);
  }
  const sData = sSheet.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    const name = String(sData[i][0] || '').trim();
    if (name) staff.push(name);
  }

  return {status:'ok', products, staff};
}

// ============================================================
// マスターデータ保存
// ============================================================
function saveMaster(type, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === 'products') {
    const sh = getOrCreateSheet(ss, '商品マスター');
    sh.clearContents();
    sh.appendRow(['コード','商品名','単価','種別','サブカテゴリ']);
    data.forEach(p => sh.appendRow([p.code, p.name, p.price, p.cat, p.sub || '']));
    formatHeader(sh);
  }

  if (type === 'staff') {
    const sh = getOrCreateSheet(ss, '担当者マスター');
    sh.clearContents();
    sh.appendRow(['担当者名']);
    data.forEach(name => sh.appendRow([name]));
    formatHeader(sh);
  }

  return {status:'ok', message:'保存しました'};
}

// ============================================================
// 売上明細追記
// ============================================================
function addSales(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet(ss, '売上明細');

  // ヘッダー行がなければ作成
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      '日時','取引番号','担当者','商品名','カテゴリ',
      'サブカテゴリ','数量','単価','小計',
      '合計','支払方法','お預かり','おつり'
    ]);
    formatHeader(sh);
  }

  rows.forEach(row => sh.appendRow(row));

  return {status:'ok', message: rows.length + '件追加しました'};
}

// ============================================================
// ユーティリティ
// ============================================================
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function formatHeader(sh) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn());
  header.setBackground('#1c1b18')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
  sh.setFrozenRows(1);
}

function initProductSheet(sh) {
  sh.appendRow(['コード','商品名','単価','種別','サブカテゴリ']);
  const defaults = [
    ['TKT-FREE', '入場無料チケット', 0,    'ticket', ''],
    ['TKT-GUIDE','ガイド付チケット', 1500, 'ticket', ''],
    ['NT-001',   'ノート（A5）',    800,  'goods',  'ノート'],
    ['NT-002',   'ノート（B6）',    600,  'goods',  'ノート'],
    ['FL-001',   'クリアファイルA', 400,  'goods',  'ファイル'],
    ['PC-001',   'ポストカード①',  200,  'goods',  'ポストカード'],
    ['MG-001',   'マグネット①',    300,  'goods',  'マグネット'],
    ['PT-001',   'ポストイット①',  350,  'goods',  'ポストイット'],
  ];
  defaults.forEach(row => sh.appendRow(row));
  formatHeader(sh);
}

function initStaffSheet(sh) {
  sh.appendRow(['担当者名']);
  ['田中','佐藤','鈴木','伊藤','山田','中村'].forEach(n => sh.appendRow([n]));
  formatHeader(sh);
}
