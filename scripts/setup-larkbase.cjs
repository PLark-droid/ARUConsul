/**
 * LarkBase セットアップスクリプト
 * 就労継続支援B型開業プロジェクト管理用のテーブルを自動構築
 */

const lark = require('@larksuiteoapi/node-sdk');
const XLSX = require('xlsx');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

// Lark クライアント初期化
const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: process.env.LARK_DOMAIN === 'larksuite' ? lark.Domain.Lark : lark.Domain.Feishu,
});

// Excelシリアル値を日付文字列に変換
function excelSerialToDate(serial) {
  if (!serial || isNaN(serial)) return null;
  const date = new Date((serial - 25569) * 86400 * 1000);
  return date.toISOString().split('T')[0];
}

// テーブル定義
const TABLE_DEFINITIONS = {
  // 1. プロジェクト情報
  project_info: {
    name: 'プロジェクト情報',
    fields: [
      { field_name: '顧客名', type: 1 }, // テキスト
      { field_name: 'コンサル会社', type: 1 },
      { field_name: '対象エリア', type: 1 },
      { field_name: '開業予定日', type: 5 }, // 日付
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '進行中', color: 0 },
        { name: '完了', color: 1 },
        { name: '保留', color: 2 }
      ]}} // 単一選択
    ]
  },

  // 2. WBSカテゴリ
  wbs_category: {
    name: 'WBSカテゴリ',
    fields: [
      { field_name: 'WBS番号', type: 1 },
      { field_name: 'カテゴリ名', type: 1 },
      { field_name: '表示順', type: 2 } // 数値
    ]
  },

  // 3. タスク
  tasks: {
    name: 'タスク',
    fields: [
      { field_name: 'WBS番号', type: 1 },
      { field_name: 'タスク名', type: 1 },
      { field_name: 'カテゴリ', type: 1 },
      { field_name: '担当者', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: '貴社', color: 1 }
      ]}},
      { field_name: '開始日', type: 5 },
      { field_name: '期限', type: 5 },
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '未着手', color: 2 },
        { name: '進行中', color: 0 },
        { name: '完了', color: 1 },
        { name: '保留', color: 3 },
        { name: 'ブロック中', color: 4 }
      ]}},
      { field_name: '完了率', type: 2 },
      { field_name: '備考', type: 1 }
    ]
  },

  // 4. 申請書類
  application_docs: {
    name: '申請書類',
    fields: [
      { field_name: 'No.', type: 2 },
      { field_name: '書類名', type: 1 },
      { field_name: '作成担当', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: 'LIVENOW', color: 1 }
      ]}},
      { field_name: '確認担当', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: 'LIVENOW', color: 1 }
      ]}},
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '未', color: 2 },
        { name: '作成中', color: 0 },
        { name: '確認中', color: 3 },
        { name: '完了', color: 1 }
      ]}},
      { field_name: '備考', type: 1 }
    ]
  },

  // 5. 提供書類
  provided_docs: {
    name: '提供書類',
    fields: [
      { field_name: 'No.', type: 2 },
      { field_name: '書類名', type: 1 },
      { field_name: '提供予定日', type: 5 },
      { field_name: '提供元', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: 'LIVENOW', color: 1 }
      ]}},
      { field_name: '提供先', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: 'LIVENOW', color: 1 }
      ]}},
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '未', color: 2 },
        { name: '準備中', color: 0 },
        { name: '提供済', color: 1 }
      ]}},
      { field_name: '備考', type: 1 }
    ]
  },

  // 6. 面接者
  interviewees: {
    name: '面接者',
    fields: [
      { field_name: '氏名', type: 1 },
      { field_name: '面接日時', type: 5 },
      { field_name: '面接種別', type: 3, property: { options: [
        { name: '1次面接', color: 0 },
        { name: '2次面接', color: 1 },
        { name: '最終面接', color: 2 }
      ]}},
      { field_name: '事前情報URL', type: 15 }, // URL
      { field_name: '面接リンク', type: 15 },
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '予定', color: 0 },
        { name: '完了', color: 1 },
        { name: '採用', color: 3 },
        { name: '不採用', color: 4 }
      ]}},
      { field_name: '評価', type: 3, property: { options: [
        { name: 'A', color: 1 },
        { name: 'B', color: 0 },
        { name: 'C', color: 3 },
        { name: 'D', color: 4 }
      ]}},
      { field_name: 'メモ', type: 1 }
    ]
  },

  // 7. 物件
  properties: {
    name: '物件',
    fields: [
      { field_name: 'NO', type: 2 },
      { field_name: '提供日', type: 5 },
      { field_name: '物件名', type: 1 },
      { field_name: 'URL', type: 15 },
      { field_name: '媒体', type: 3, property: { options: [
        { name: 'アットホーム', color: 0 },
        { name: 'スーモ', color: 1 },
        { name: 'スマイティ', color: 2 },
        { name: 'ヤフー不動産', color: 3 },
        { name: '店舗ネットワーク', color: 4 },
        { name: 'その他', color: 5 }
      ]}},
      { field_name: '住所', type: 1 },
      { field_name: '賃料', type: 1 },
      { field_name: '管理料', type: 1 },
      { field_name: '専有面積', type: 1 },
      { field_name: '築年数', type: 1 },
      { field_name: '間取り', type: 1 },
      { field_name: '駐車場', type: 3, property: { options: [
        { name: '有', color: 1 },
        { name: '無', color: 4 }
      ]}},
      { field_name: 'ステータス', type: 3, property: { options: [
        { name: '候補', color: 2 },
        { name: '就労B使用可能', color: 1 },
        { name: '就労B使用不可', color: 4 },
        { name: '確認中', color: 3 },
        { name: '内見済', color: 0 },
        { name: '決定', color: 5 }
      ]}},
      { field_name: 'コメント', type: 1 },
      { field_name: '内見日', type: 5 },
      { field_name: '不動産会社', type: 1 }
    ]
  },

  // 8. 不動産会社
  real_estate: {
    name: '不動産会社',
    fields: [
      { field_name: 'NO', type: 2 },
      { field_name: '会社名', type: 1 },
      { field_name: '担当者名', type: 1 },
      { field_name: 'URL', type: 15 },
      { field_name: '電話番号', type: 13 }, // 電話
      { field_name: '架電担当', type: 3, property: { options: [
        { name: 'ARU', color: 0 },
        { name: 'LIVENOW', color: 1 }
      ]}},
      { field_name: '架電日', type: 5 },
      { field_name: '架電メモ', type: 1 },
      { field_name: '次回ToDo', type: 1 }
    ]
  },

  // 9. 議事録
  minutes: {
    name: '議事録',
    fields: [
      { field_name: '日時', type: 5 },
      { field_name: '場所', type: 1 },
      { field_name: '出席者', type: 1 },
      { field_name: '議事内容', type: 1 },
      { field_name: '決定事項', type: 1 },
      { field_name: 'ToDo', type: 1 }
    ]
  }
};

// WBSカテゴリ初期データ
const WBS_CATEGORIES = [
  { wbs: '1', name: '法人関連', order: 1 },
  { wbs: '2', name: '融資', order: 2 },
  { wbs: '3', name: '物件', order: 3 },
  { wbs: '4', name: '採用', order: 4 },
  { wbs: '5', name: '指定申請', order: 5 },
  { wbs: '6', name: '利用者獲得', order: 6 },
  { wbs: '7', name: '業務獲得', order: 7 },
  { wbs: '8', name: 'その他運営関連', order: 8 }
];

// メイン処理
async function main() {
  console.log('=== LarkBase セットアップ開始 ===\n');

  // 認証情報チェック
  if (!process.env.LARK_APP_ID || process.env.LARK_APP_ID === 'your_app_id_here') {
    console.error('エラー: .envファイルにLARK_APP_IDを設定してください');
    console.log('\n設定手順:');
    console.log('1. Lark Open Platform (https://open.larksuite.com/) でアプリを作成');
    console.log('2. .envファイルのLARK_APP_IDとLARK_APP_SECRETを更新');
    console.log('3. アプリに「多維表格」の権限を追加');
    process.exit(1);
  }

  try {
    // ステップ1: Baseを作成
    console.log('ステップ1: 新規Baseを作成中...');

    let appToken = process.env.LARK_BASE_APP_TOKEN;

    if (!appToken) {
      // 新規Base作成
      const createBaseRes = await client.bitable.app.create({
        data: {
          name: 'LIVENOW様 就労継続支援B型開業プロジェクト',
          folder_token: '' // ルートフォルダに作成
        }
      });

      if (createBaseRes.code !== 0) {
        throw new Error(`Base作成エラー: ${createBaseRes.msg}`);
      }

      appToken = createBaseRes.data.app.app_token;
      console.log(`✓ Base作成完了: ${appToken}`);
      console.log(`  URL: https://www.feishu.cn/base/${appToken}`);
    } else {
      console.log(`✓ 既存のBaseを使用: ${appToken}`);
    }

    // ステップ2: テーブルを作成
    console.log('\nステップ2: テーブルを作成中...');
    const tableIds = {};

    for (const [key, def] of Object.entries(TABLE_DEFINITIONS)) {
      console.log(`  - ${def.name}を作成中...`);

      const createTableRes = await client.bitable.appTable.create({
        path: { app_token: appToken },
        data: {
          table: {
            name: def.name,
            default_view_name: '全件表示',
            fields: def.fields
          }
        }
      });

      if (createTableRes.code !== 0) {
        console.error(`    エラー: ${createTableRes.msg}`);
        continue;
      }

      tableIds[key] = createTableRes.data.table_id;
      console.log(`    ✓ 完了 (ID: ${tableIds[key]})`);

      // レート制限対策
      await new Promise(r => setTimeout(r, 500));
    }

    // ステップ3: データをインポート
    console.log('\nステップ3: Excelデータをインポート中...');

    const excelPath = path.join(__dirname, '../../docs/合同会社LIVENOW様　就労継続支援B型開業スケジュール.xlsx');
    const workbook = XLSX.readFile(excelPath);

    // プロジェクト情報
    console.log('  - プロジェクト情報...');
    await client.bitable.appTableRecord.create({
      path: { app_token: appToken, table_id: tableIds.project_info },
      data: {
        fields: {
          '顧客名': '合同会社LIVENOW',
          'コンサル会社': '株式会社ARU',
          '対象エリア': '東京都（大田区メイン）',
          '開業予定日': 1738368000000, // 2026/02/01
          'ステータス': '進行中'
        }
      }
    });
    console.log('    ✓ 完了');

    // WBSカテゴリ
    console.log('  - WBSカテゴリ...');
    for (const cat of WBS_CATEGORIES) {
      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.wbs_category },
        data: {
          fields: {
            'WBS番号': cat.wbs,
            'カテゴリ名': cat.name,
            '表示順': cat.order
          }
        }
      });
      await new Promise(r => setTimeout(r, 200));
    }
    console.log('    ✓ 完了');

    // タスクデータ
    console.log('  - タスク...');
    const scheduleSheet = workbook.Sheets['開業スケジュール'];
    const scheduleData = XLSX.utils.sheet_to_json(scheduleSheet, { header: 1, defval: '' });

    let taskCount = 0;
    for (let i = 5; i < scheduleData.length; i++) {
      const row = scheduleData[i];
      const wbs = String(row[0] || '').trim();
      const title = String(row[1] || '').trim();

      // タスク行のみ処理（WBS番号が x.x 形式）
      if (!wbs.includes('.') || !title) continue;

      const category = WBS_CATEGORIES.find(c => wbs.startsWith(c.wbs + '.'));
      const owner = String(row[2] || '').trim();
      const startDate = row[3] ? excelSerialToDate(row[3]) : null;
      const endDate = row[4] ? excelSerialToDate(row[4]) : null;
      const progress = row[6] === 1 ? '完了' : '未着手';
      const note = String(row[7] || '').trim();

      const fields = {
        'WBS番号': wbs,
        'タスク名': title,
        'カテゴリ': category ? category.name : '',
        'ステータス': progress,
        '備考': note
      };

      if (owner) fields['担当者'] = owner;
      if (startDate) fields['開始日'] = new Date(startDate).getTime();
      if (endDate) fields['期限'] = new Date(endDate).getTime();
      if (row[6] !== undefined) fields['完了率'] = row[6] === 1 ? 100 : 0;

      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.tasks },
        data: { fields }
      });

      taskCount++;
      await new Promise(r => setTimeout(r, 200));
    }
    console.log(`    ✓ ${taskCount}件のタスクをインポート`);

    // 物件データ
    console.log('  - 物件...');
    const propertySheet = workbook.Sheets['物件リスト'];
    const propertyData = XLSX.utils.sheet_to_json(propertySheet, { header: 1, defval: '' });

    let propCount = 0;
    for (let i = 1; i < propertyData.length; i++) {
      const row = propertyData[i];
      if (!row[0] || !row[4]) continue; // NO, 名称がない行はスキップ

      const statusMap = {
        '': '候補',
        '就労B使用可能': '就労B使用可能',
        '就労B使用不可': '就労B使用不可',
        '確認中': '確認中'
      };

      const fields = {
        'NO': parseInt(row[0]) || 0,
        '物件名': String(row[4] || '').substring(0, 100),
        '住所': String(row[8] || ''),
        '賃料': String(row[9] || ''),
        '管理料': String(row[10] || ''),
        '専有面積': String(row[13] || row[12] || ''),
        '築年数': String(row[14] || ''),
        '間取り': String(row[15] || ''),
        'ステータス': statusMap[String(row[3] || '')] || '候補',
        'コメント': String(row[5] || '').substring(0, 500),
        '不動産会社': String(row[16] || '')
      };

      if (row[1]) fields['提供日'] = new Date(excelSerialToDate(row[1])).getTime();
      if (row[2]) fields['URL'] = { link: String(row[2]).substring(0, 1000), text: 'リンク' };
      if (row[7]) fields['媒体'] = String(row[7]);
      if (row[11]) fields['駐車場'] = row[11] === '有' ? '有' : '無';

      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.properties },
        data: { fields }
      });

      propCount++;
      await new Promise(r => setTimeout(r, 200));
    }
    console.log(`    ✓ ${propCount}件の物件をインポート`);

    // 不動産会社データ
    console.log('  - 不動産会社...');
    const realEstateSheet = workbook.Sheets['不動産リスト'];
    const realEstateData = XLSX.utils.sheet_to_json(realEstateSheet, { header: 1, defval: '' });

    let reCount = 0;
    for (let i = 1; i < realEstateData.length; i++) {
      const row = realEstateData[i];
      if (!row[0] || !row[1]) continue;

      const fields = {
        'NO': parseInt(row[0]) || 0,
        '会社名': String(row[1] || ''),
        '担当者名': String(row[2] || ''),
        '電話番号': String(row[4] || '')
      };

      if (row[3]) fields['URL'] = { link: String(row[3]).substring(0, 1000), text: 'HP' };

      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.real_estate },
        data: { fields }
      });

      reCount++;
      await new Promise(r => setTimeout(r, 200));
    }
    console.log(`    ✓ ${reCount}件の不動産会社をインポート`);

    // 面接者データ
    console.log('  - 面接者...');
    const interviewSheet = workbook.Sheets['面接者情報'];
    const interviewData = XLSX.utils.sheet_to_json(interviewSheet, { header: 1, defval: '' });

    let intCount = 0;
    for (let i = 1; i < interviewData.length; i++) {
      const row = interviewData[i];
      if (!row[1]) continue; // 氏名がない行はスキップ

      const fields = {
        '氏名': String(row[1] || '').replace('さん', '').trim(),
        '面接種別': '1次面接',
        'ステータス': '予定'
      };

      if (row[2]) {
        const infoText = String(row[2]);
        if (infoText.includes('http')) {
          const urlMatch = infoText.match(/(https?:\/\/[^\s]+)/);
          if (urlMatch) fields['事前情報URL'] = { link: urlMatch[1], text: '履歴書' };
        }
      }

      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.interviewees },
        data: { fields }
      });

      intCount++;
      await new Promise(r => setTimeout(r, 200));
    }
    console.log(`    ✓ ${intCount}件の面接者をインポート`);

    // 議事録データ
    console.log('  - 議事録...');
    const minuteSheets = ['251002打ち合わせ議事録', '251013打ち合わせ議事録', '251105打ち合わせ議事録'];

    for (const sheetName of minuteSheets) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      const fields = {
        '場所': 'オンライン'
      };

      for (const row of data) {
        const label = String(row[0] || '').trim();
        const value = String(row[1] || '').trim();

        if (label === '日 時' && value) {
          fields['日時'] = value;
        } else if (label === '場 所') {
          fields['場所'] = value || 'オンライン';
        } else if (label === '出席者') {
          fields['出席者'] = value;
        } else if (label.includes('議事内容')) {
          fields['議事内容'] = value;
        } else if (label === '決定事項') {
          fields['決定事項'] = value;
        } else if (label === 'To Do') {
          fields['ToDo'] = value;
        }
      }

      await client.bitable.appTableRecord.create({
        path: { app_token: appToken, table_id: tableIds.minutes },
        data: { fields }
      });

      await new Promise(r => setTimeout(r, 200));
    }
    console.log('    ✓ 3件の議事録をインポート');

    console.log('\n=== セットアップ完了 ===');
    console.log(`\nLarkBase URL: https://www.feishu.cn/base/${appToken}`);
    console.log('\n作成されたテーブル:');
    for (const [key, id] of Object.entries(tableIds)) {
      console.log(`  - ${TABLE_DEFINITIONS[key].name}: ${id}`);
    }

  } catch (error) {
    console.error('\nエラーが発生しました:', error.message);
    if (error.response) {
      console.error('レスポンス:', JSON.stringify(error.response.data, null, 2));
    }
    process.exit(1);
  }
}

main();
