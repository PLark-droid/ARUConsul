/**
 * マルチテナント対応セットアップスクリプト
 * Phase 1: 基盤構築
 * - 顧客マスタテーブル作成
 * - タスクテンプレートテーブル作成
 * - タスクテンプレート52件の初期データ投入
 * - 既存タスクテーブルへの顧客リンク追加
 */

const lark = require('@larksuiteoapi/node-sdk');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

// Lark クライアント初期化
const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: process.env.LARK_DOMAIN === 'larksuite' ? lark.Domain.Lark : lark.Domain.Feishu,
});

let APP_TOKEN = process.env.LARK_BASE_APP_TOKEN;

// ========================================
// テーブル定義
// ========================================

// 顧客マスタテーブル定義
const CUSTOMER_TABLE = {
  name: '顧客マスタ',
  fields: [
    { field_name: '会社名', type: 1 }, // テキスト
    { field_name: '代表者名', type: 1 },
    { field_name: '担当コンサル', type: 3, property: { options: [
      { name: '金沢', color: 0 },
      { name: '山田', color: 1 },
      { name: '鈴木', color: 2 },
      { name: '田中', color: 3 }
    ]}},
    { field_name: '対象エリア', type: 1 },
    { field_name: '開業予定日', type: 5 }, // 日付
    { field_name: '契約開始日', type: 5 },
    { field_name: 'ステータス', type: 3, property: { options: [
      { name: '準備中', color: 2 },
      { name: '進行中', color: 0 },
      { name: '開業済', color: 1 },
      { name: '保留', color: 3 },
      { name: '解約', color: 4 }
    ]}},
    { field_name: '連絡先メール', type: 1 },
    { field_name: '連絡先電話', type: 13 }, // 電話
    { field_name: '備考', type: 1 }
  ]
};

// タスクテンプレートテーブル定義
const TEMPLATE_TABLE = {
  name: 'タスクテンプレート',
  fields: [
    { field_name: 'WBS番号', type: 1 },
    { field_name: 'タスク名', type: 1 },
    { field_name: 'カテゴリ', type: 3, property: { options: [
      { name: '法人関連', color: 0 },
      { name: '融資', color: 1 },
      { name: '物件', color: 2 },
      { name: '採用', color: 3 },
      { name: '指定申請', color: 4 },
      { name: '利用者獲得', color: 5 },
      { name: '業務獲得', color: 6 },
      { name: 'その他運営関連', color: 7 }
    ]}},
    { field_name: '担当区分', type: 3, property: { options: [
      { name: 'ARU', color: 0 },
      { name: '顧客', color: 1 }
    ]}},
    { field_name: '標準所要日数', type: 2 }, // 数値
    { field_name: '開業日オフセット', type: 2 }, // 数値（開業日から何日前に期限か）
    { field_name: '表示順', type: 2 },
    { field_name: '必須フラグ', type: 7 }, // チェックボックス
    { field_name: '備考テンプレート', type: 1 }
  ]
};

// ========================================
// タスクテンプレート初期データ（52タスク）
// ========================================
const TASK_TEMPLATES = [
  // 1. 法人関連（3タスク）
  { wbs: '1.1', name: '定款文言提供', category: '法人関連', owner: 'ARU', days: 36, offset: -90, order: 1 },
  { wbs: '1.2', name: '定款作成（変更）', category: '法人関連', owner: '顧客', days: 36, offset: -90, order: 2 },
  { wbs: '1.3', name: '登記簿登録', category: '法人関連', owner: '顧客', days: 14, offset: -60, order: 3 },

  // 2. 融資（5タスク）
  { wbs: '2.1', name: '融資先検討', category: '融資', owner: '顧客', days: 30, offset: -150, order: 4 },
  { wbs: '2.2', name: '事業計画書作成', category: '融資', owner: 'ARU', days: 30, offset: -120, order: 5 },
  { wbs: '2.3', name: '各種見積もり作成', category: '融資', owner: '顧客', days: 14, offset: -100, order: 6 },
  { wbs: '2.4', name: '融資面談', category: '融資', owner: '顧客', days: 7, offset: -90, order: 7 },
  { wbs: '2.5', name: '融資実行', category: '融資', owner: '顧客', days: 30, offset: -60, order: 8 },

  // 3. 物件（10タスク）
  { wbs: '3.1', name: 'エリア決定', category: '物件', owner: '顧客', days: 26, offset: -150, order: 9 },
  { wbs: '3.2', name: '物件リスト作成', category: '物件', owner: 'ARU', days: 9, offset: -120, order: 10 },
  { wbs: '3.3', name: '内見及び物件検討', category: '物件', owner: '顧客', days: 30, offset: -90, order: 11 },
  { wbs: '3.4', name: 'レイアウト図作成', category: '物件', owner: 'ARU', days: 29, offset: -60, order: 12 },
  { wbs: '3.5', name: '消防確認', category: '物件', owner: 'ARU', days: 29, offset: -60, order: 13 },
  { wbs: '3.6', name: '建築確認', category: '物件', owner: 'ARU', days: 29, offset: -60, order: 14 },
  { wbs: '3.7', name: '都市計画確認', category: '物件', owner: 'ARU', days: 29, offset: -60, order: 15 },
  { wbs: '3.8', name: 'まちづくり条例確認', category: '物件', owner: 'ARU', days: 29, offset: -60, order: 16 },
  { wbs: '3.9', name: '物件決定（契約）', category: '物件', owner: '顧客', days: 30, offset: -30, order: 17 },
  { wbs: '3.10', name: '保険内容決定', category: '物件', owner: '顧客', days: 30, offset: -30, order: 18 },

  // 4. 採用（5タスク）
  { wbs: '4.1', name: 'ジョブメドレー、リタリコ、ウェルミージョブ連絡', category: '採用', owner: 'ARU', days: 9, offset: -120, order: 19 },
  { wbs: '4.2', name: 'ジョブメドレー、リタリコ、ウェルミージョブ契約', category: '採用', owner: '顧客', days: 9, offset: -120, order: 20 },
  { wbs: '4.3', name: 'ハローワーク、Indeed掲載', category: '採用', owner: '顧客', days: 7, offset: -100, order: 21 },
  { wbs: '4.4', name: 'サービス管理責任者決定', category: '採用', owner: '顧客', days: 29, offset: -60, order: 22 },
  { wbs: '4.5', name: 'その他職員決定', category: '採用', owner: '顧客', days: 60, offset: 0, order: 23 },

  // 5. 指定申請（11タスク）
  { wbs: '5.1', name: '都道府県及び市町村へフロー確認', category: '指定申請', owner: 'ARU', days: 9, offset: -120, order: 24 },
  { wbs: '5.2', name: '在宅就労申請フロー確認', category: '指定申請', owner: 'ARU', days: 9, offset: -120, order: 25 },
  { wbs: '5.3', name: '屋号・作業内容・事業所方針検討', category: '指定申請', owner: 'ARU', days: 30, offset: -75, order: 26 },
  { wbs: '5.4', name: '事前協議書作成', category: '指定申請', owner: 'ARU', days: 14, offset: -60, order: 27 },
  { wbs: '5.5', name: '市町村と事前協議', category: '指定申請', owner: '顧客', days: 18, offset: -40, order: 28 },
  { wbs: '5.6', name: '協力医療機関確保', category: '指定申請', owner: '顧客', days: 35, offset: -40, order: 29 },
  { wbs: '5.7', name: '指定申請書類作成', category: '指定申請', owner: 'ARU', days: 9, offset: -60, order: 30 },
  { wbs: '5.8', name: '指定申請書類確認', category: '指定申請', owner: '顧客', days: 18, offset: -40, order: 31 },
  { wbs: '5.9', name: '指定申請書類提出', category: '指定申請', owner: '顧客', days: 11, offset: -30, order: 32 },
  { wbs: '5.10', name: '最終版指定申請書類共有', category: '指定申請', owner: '顧客', days: 15, offset: -15, order: 33 },
  { wbs: '5.11', name: '現地確認', category: '指定申請', owner: '顧客', days: 30, offset: 0, order: 34 },

  // 6. 利用者獲得（9タスク）
  { wbs: '6.1', name: 'チラシ・三つ折りパンフレット作成', category: '利用者獲得', owner: 'ARU', days: 16, offset: -30, order: 35 },
  { wbs: '6.2', name: 'ポスティング会社リストアップ', category: '利用者獲得', owner: 'ARU', days: 18, offset: -40, order: 36 },
  { wbs: '6.3', name: '公営住宅等配布先リストアップ', category: '利用者獲得', owner: 'ARU', days: 18, offset: -40, order: 37 },
  { wbs: '6.4', name: '配布エリア提案、ポスト方法レクチャー', category: '利用者獲得', owner: 'ARU', days: 12, offset: -30, order: 38 },
  { wbs: '6.5', name: 'ポスティング会社依頼（自社配布も可）', category: '利用者獲得', owner: '顧客', days: 12, offset: -30, order: 39 },
  { wbs: '6.6', name: '印刷業者へ依頼、印刷', category: '利用者獲得', owner: '顧客', days: 5, offset: -30, order: 40 },
  { wbs: '6.7', name: '相談支援事業所等営業先リストアップ', category: '利用者獲得', owner: 'ARU', days: 18, offset: -40, order: 41 },
  { wbs: '6.8', name: '相談支援事業所等営業レクチャー', category: '利用者獲得', owner: 'ARU', days: 12, offset: -30, order: 42 },
  { wbs: '6.9', name: '相談支援事業所等営業', category: '利用者獲得', owner: '顧客', days: 30, offset: 0, order: 43 },

  // 7. 業務獲得（4タスク）
  { wbs: '7.1', name: 'PC作業トライアル(職員向け)', category: '業務獲得', owner: '顧客', days: 30, offset: 0, order: 44 },
  { wbs: '7.2', name: '業務受託先リストアップ', category: '業務獲得', owner: 'ARU', days: 18, offset: -40, order: 45 },
  { wbs: '7.3', name: '業務受託先営業レクチャー（テレアポ）', category: '業務獲得', owner: 'ARU', days: 12, offset: -30, order: 46 },
  { wbs: '7.4', name: '業務受託先営業', category: '業務獲得', owner: '顧客', days: 30, offset: 0, order: 47 },

  // 8. その他運営関連（5タスク）
  { wbs: '8.1', name: '利用者獲得、業務獲得', category: 'その他運営関連', owner: 'ARU', days: 30, offset: -30, order: 48 },
  { wbs: '8.2', name: '人員配置、職員の職務内容', category: 'その他運営関連', owner: 'ARU', days: 30, offset: 0, order: 49 },
  { wbs: '8.3', name: '運営マニュアル説明', category: 'その他運営関連', owner: 'ARU', days: 30, offset: 0, order: 50 },
  { wbs: '8.4', name: '法定研修', category: 'その他運営関連', owner: 'ARU', days: 15, offset: 0, order: 51 },
  { wbs: '8.5', name: '営業進捗共有／確認', category: 'その他運営関連', owner: '顧客', days: 15, offset: 0, order: 52 }
];

// ========================================
// メイン処理
// ========================================
async function main() {
  console.log('=== マルチテナント対応 Phase 1: 基盤構築 ===\n');

  // 認証情報チェック
  if (!process.env.LARK_APP_ID || process.env.LARK_APP_ID === 'your_app_id_here') {
    console.error('エラー: .envファイルにLARK_APP_IDを設定してください');
    process.exit(1);
  }

  try {
    // ========================================
    // ステップ0: 新規Base作成（APP_TOKENがない場合）
    // ========================================
    if (!APP_TOKEN) {
      console.log('ステップ0: 新規LarkBaseを作成中...');

      const createBaseRes = await client.bitable.app.create({
        data: {
          name: 'ARUフランチャイズ プロジェクト管理',
          folder_token: '' // ルートフォルダに作成
        }
      });

      if (createBaseRes.code !== 0) {
        throw new Error(`Base作成エラー: ${createBaseRes.msg}`);
      }

      APP_TOKEN = createBaseRes.data.app.app_token;
      const baseUrl = process.env.LARK_DOMAIN === 'larksuite'
        ? `https://www.larksuite.com/base/${APP_TOKEN}`
        : `https://www.feishu.cn/base/${APP_TOKEN}`;
      console.log(`✓ 新規Base作成完了`);
      console.log(`  APP_TOKEN: ${APP_TOKEN}`);
      console.log(`  URL: ${baseUrl}`);
      console.log(`\n  ※ .envに以下を追加してください:`);
      console.log(`  LARK_BASE_APP_TOKEN=${APP_TOKEN}\n`);

      await sleep(1000);
    } else {
      console.log(`既存のBaseを使用: ${APP_TOKEN}`);
    }

    // ========================================
    // ステップ1: 顧客マスタテーブル作成
    // ========================================
    console.log('\nステップ1: 顧客マスタテーブルを作成中...');

    const customerTableRes = await client.bitable.appTable.create({
      path: { app_token: APP_TOKEN },
      data: {
        table: {
          name: CUSTOMER_TABLE.name,
          default_view_name: '顧客一覧',
          fields: CUSTOMER_TABLE.fields
        }
      }
    });

    if (customerTableRes.code !== 0) {
      throw new Error(`顧客マスタ作成エラー: ${customerTableRes.msg}`);
    }

    const customerTableId = customerTableRes.data.table_id;
    console.log(`✓ 顧客マスタ作成完了 (ID: ${customerTableId})`);

    await sleep(500);

    // ========================================
    // ステップ2: タスクテンプレートテーブル作成
    // ========================================
    console.log('\nステップ2: タスクテンプレートテーブルを作成中...');

    const templateTableRes = await client.bitable.appTable.create({
      path: { app_token: APP_TOKEN },
      data: {
        table: {
          name: TEMPLATE_TABLE.name,
          default_view_name: 'テンプレート一覧',
          fields: TEMPLATE_TABLE.fields
        }
      }
    });

    if (templateTableRes.code !== 0) {
      throw new Error(`タスクテンプレート作成エラー: ${templateTableRes.msg}`);
    }

    const templateTableId = templateTableRes.data.table_id;
    console.log(`✓ タスクテンプレート作成完了 (ID: ${templateTableId})`);

    await sleep(500);

    // ========================================
    // ステップ3: タスクテンプレート52件の初期データ投入
    // ========================================
    console.log('\nステップ3: タスクテンプレートデータを投入中...');

    let successCount = 0;
    let errorCount = 0;

    for (const template of TASK_TEMPLATES) {
      const fields = {
        'WBS番号': template.wbs,
        'タスク名': template.name,
        'カテゴリ': template.category,
        '担当区分': template.owner,
        '標準所要日数': template.days,
        '開業日オフセット': template.offset,
        '表示順': template.order,
        '必須フラグ': true
      };

      try {
        const res = await client.bitable.appTableRecord.create({
          path: { app_token: APP_TOKEN, table_id: templateTableId },
          data: { fields }
        });

        if (res.code === 0) {
          successCount++;
          process.stdout.write(`\r  進捗: ${successCount}/${TASK_TEMPLATES.length} 件`);
        } else {
          errorCount++;
          console.error(`\n  エラー (${template.wbs}): ${res.msg}`);
        }
      } catch (err) {
        errorCount++;
        console.error(`\n  エラー (${template.wbs}): ${err.message}`);
      }

      await sleep(200);
    }

    console.log(`\n✓ タスクテンプレート投入完了: ${successCount}件成功, ${errorCount}件エラー`);

    // ========================================
    // ステップ4: タスクテーブル作成（顧客・テンプレートリンク付き）
    // ========================================
    console.log('\nステップ4: タスクテーブルを作成中...');

    const taskTableRes = await client.bitable.appTable.create({
      path: { app_token: APP_TOKEN },
      data: {
        table: {
          name: 'タスク',
          default_view_name: '全タスク一覧',
          fields: [
            { field_name: 'WBS番号', type: 1 },
            { field_name: 'タスク名', type: 1 },
            { field_name: 'カテゴリ', type: 3, property: { options: [
              { name: '法人関連', color: 0 },
              { name: '融資', color: 1 },
              { name: '物件', color: 2 },
              { name: '採用', color: 3 },
              { name: '指定申請', color: 4 },
              { name: '利用者獲得', color: 5 },
              { name: '業務獲得', color: 6 },
              { name: 'その他運営関連', color: 7 }
            ]}},
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
        }
      }
    });

    if (taskTableRes.code !== 0) {
      throw new Error(`タスクテーブル作成エラー: ${taskTableRes.msg}`);
    }

    const taskTableId = taskTableRes.data.table_id;
    console.log(`✓ タスクテーブル作成完了 (ID: ${taskTableId})`);

    await sleep(500);

    // 顧客リンクフィールドを追加（単向関連: type 18）
    console.log('  - 顧客リンクフィールドを追加中...');
    try {
      const linkRes1 = await client.bitable.appTableField.create({
        path: { app_token: APP_TOKEN, table_id: taskTableId },
        data: {
          field_name: '顧客',
          type: 18, // 単向関連
          property: {
            table_id: customerTableId
          }
        }
      });
      if (linkRes1.code === 0) {
        console.log('    ✓ 顧客リンク追加完了');
      } else {
        console.log(`    ⚠ スキップ: ${linkRes1.msg}`);
      }
    } catch (err) {
      console.log(`    ⚠ スキップ: ${err.message}`);
    }

    await sleep(300);

    // テンプレートリンクフィールドを追加（単向関連: type 18）
    console.log('  - テンプレートリンクフィールドを追加中...');
    try {
      const linkRes2 = await client.bitable.appTableField.create({
        path: { app_token: APP_TOKEN, table_id: taskTableId },
        data: {
          field_name: 'テンプレート',
          type: 18, // 単向関連
          property: {
            table_id: templateTableId
          }
        }
      });
      if (linkRes2.code === 0) {
        console.log('    ✓ テンプレートリンク追加完了');
      } else {
        console.log(`    ⚠ スキップ: ${linkRes2.msg}`);
      }
    } catch (err) {
      console.log(`    ⚠ スキップ: ${err.message}`);
    }

    // ========================================
    // ステップ5: LIVENOW様を顧客マスタに登録
    // ========================================
    console.log('\nステップ5: 既存顧客（LIVENOW様）を顧客マスタに登録中...');

    const customerRes = await client.bitable.appTableRecord.create({
      path: { app_token: APP_TOKEN, table_id: customerTableId },
      data: {
        fields: {
          '会社名': '合同会社LIVENOW',
          '代表者名': '柴田様',
          '担当コンサル': '金沢',
          '対象エリア': '東京都（大田区メイン）',
          '開業予定日': new Date('2026-02-01').getTime(),
          '契約開始日': new Date('2025-10-01').getTime(),
          'ステータス': '進行中',
          '備考': '既存顧客。Phase1移行時に登録。'
        }
      }
    });

    if (customerRes.code === 0) {
      console.log('✓ 顧客マスタにLIVENOW様を登録しました');
    } else {
      console.log(`  ⚠ 顧客登録エラー: ${customerRes.msg}`);
    }

    // ========================================
    // 完了
    // ========================================
    const finalBaseUrl = process.env.LARK_DOMAIN === 'larksuite'
      ? `https://www.larksuite.com/base/${APP_TOKEN}`
      : `https://www.feishu.cn/base/${APP_TOKEN}`;
    console.log('\n=== Phase 1 セットアップ完了 ===');
    console.log(`\nLarkBase URL: ${finalBaseUrl}`);
    console.log('\n作成されたテーブル:');
    console.log(`  - 顧客マスタ: ${customerTableId}`);
    console.log(`  - タスクテンプレート: ${templateTableId}`);
    console.log(`  - タスク: ${taskTableId}`);
    console.log(`\nタスクテンプレート: ${successCount}件登録`);
    console.log('\n.envに追加してください:');
    console.log(`  LARK_BASE_APP_TOKEN=${APP_TOKEN}`);
    console.log(`  CUSTOMER_TABLE_ID=${customerTableId}`);
    console.log(`  TEMPLATE_TABLE_ID=${templateTableId}`);
    console.log(`  TASK_TABLE_ID=${taskTableId}`);
    console.log('\n次のステップ:');
    console.log('  1. LarkBaseでビュー（ガントチャート、カンバン）を設定');
    console.log('  2. LIVENOW様のタスク生成: npm run generate-tasks <顧客レコードID>');
    console.log('  3. Phase 2: Larkオートメーション設定');

  } catch (error) {
    console.error('\nエラーが発生しました:', error.message);
    if (error.response) {
      console.error('レスポンス:', JSON.stringify(error.response.data, null, 2));
    }
    process.exit(1);
  }
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

main();
