/**
 * 顧客タスク自動生成スクリプト
 * Phase 2: 自動化
 *
 * 使用方法:
 *   node scripts/generate-customer-tasks.cjs <顧客レコードID>
 *
 * 例:
 *   node scripts/generate-customer-tasks.cjs rec_xxxxxxxx
 *
 * 機能:
 *   - 指定した顧客の開業予定日を取得
 *   - タスクテンプレートから全タスクを生成
 *   - 開業日から逆算して期限を自動設定
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

const APP_TOKEN = process.env.LARK_BASE_APP_TOKEN;

// テーブルID（実行後に.envに設定するか、ここに直接記載）
const CUSTOMER_TABLE_ID = process.env.CUSTOMER_TABLE_ID || '';
const TEMPLATE_TABLE_ID = process.env.TEMPLATE_TABLE_ID || '';
const TASK_TABLE_ID = process.env.TASK_TABLE_ID || '';

async function main() {
  const customerId = process.argv[2];

  if (!customerId) {
    console.log('使用方法: node scripts/generate-customer-tasks.cjs <顧客レコードID>');
    console.log('例: node scripts/generate-customer-tasks.cjs rec_xxxxxxxx');
    process.exit(1);
  }

  console.log(`=== 顧客タスク自動生成 ===\n`);
  console.log(`顧客ID: ${customerId}`);

  try {
    // テーブルIDが未設定の場合、テーブル一覧から取得
    let customerTableId = CUSTOMER_TABLE_ID;
    let templateTableId = TEMPLATE_TABLE_ID;
    let taskTableId = TASK_TABLE_ID;

    if (!customerTableId || !templateTableId || !taskTableId) {
      console.log('\nテーブルIDを取得中...');
      const tablesRes = await client.bitable.appTable.list({
        path: { app_token: APP_TOKEN }
      });

      if (tablesRes.code !== 0) {
        throw new Error(`テーブル一覧取得エラー: ${tablesRes.msg}`);
      }

      for (const table of tablesRes.data.items) {
        if (table.name === '顧客マスタ') customerTableId = table.table_id;
        if (table.name === 'タスクテンプレート') templateTableId = table.table_id;
        if (table.name === 'タスク') taskTableId = table.table_id;
      }

      if (!customerTableId || !templateTableId || !taskTableId) {
        console.error('必要なテーブルが見つかりません');
        console.log('  顧客マスタ:', customerTableId || '未検出');
        console.log('  タスクテンプレート:', templateTableId || '未検出');
        console.log('  タスク:', taskTableId || '未検出');
        process.exit(1);
      }
    }

    // ========================================
    // ステップ1: 顧客情報を取得
    // ========================================
    console.log('\nステップ1: 顧客情報を取得中...');

    const customerRes = await client.bitable.appTableRecord.get({
      path: { app_token: APP_TOKEN, table_id: customerTableId, record_id: customerId }
    });

    if (customerRes.code !== 0) {
      throw new Error(`顧客情報取得エラー: ${customerRes.msg}`);
    }

    const customer = customerRes.data.record.fields;
    const companyName = customer['会社名'];
    const openingDate = customer['開業予定日']; // ミリ秒タイムスタンプ

    if (!openingDate) {
      throw new Error('顧客の開業予定日が設定されていません');
    }

    const openingDateObj = new Date(openingDate);
    console.log(`✓ 顧客: ${companyName}`);
    console.log(`  開業予定日: ${openingDateObj.toISOString().split('T')[0]}`);

    // ========================================
    // ステップ2: タスクテンプレートを取得
    // ========================================
    console.log('\nステップ2: タスクテンプレートを取得中...');

    const templatesRes = await client.bitable.appTableRecord.list({
      path: { app_token: APP_TOKEN, table_id: templateTableId },
      params: { page_size: 100 }
    });

    if (templatesRes.code !== 0) {
      throw new Error(`テンプレート取得エラー: ${templatesRes.msg}`);
    }

    const templates = templatesRes.data.items;
    console.log(`✓ テンプレート: ${templates.length}件取得`);

    // ========================================
    // ステップ3: タスクを生成
    // ========================================
    console.log('\nステップ3: タスクを生成中...');

    let successCount = 0;
    let errorCount = 0;

    for (const template of templates) {
      const t = template.fields;

      // 期限を計算（開業日 + オフセット）
      const offset = t['開業日オフセット'] || 0;
      const dueDate = new Date(openingDate);
      dueDate.setDate(dueDate.getDate() + offset);

      // 開始日を計算（期限 - 所要日数）
      const duration = t['標準所要日数'] || 14;
      const startDate = new Date(dueDate);
      startDate.setDate(startDate.getDate() - duration);

      const fields = {
        'WBS番号': t['WBS番号'],
        'タスク名': t['タスク名'],
        'カテゴリ': t['カテゴリ'],
        '担当者': t['担当区分'] === 'ARU' ? 'ARU' : '貴社',
        '開始日': startDate.getTime(),
        '期限': dueDate.getTime(),
        'ステータス': '未着手',
        '完了率': 0,
        '顧客': [customerId], // リンクフィールド: record_idの配列
        'テンプレート': [template.record_id], // リンクフィールド: record_idの配列
        '備考': t['備考テンプレート'] || ''
      };

      try {
        const res = await client.bitable.appTableRecord.create({
          path: { app_token: APP_TOKEN, table_id: taskTableId },
          data: { fields }
        });

        if (res.code === 0) {
          successCount++;
          process.stdout.write(`\r  進捗: ${successCount}/${templates.length} 件`);
        } else {
          errorCount++;
          console.error(`\n  エラー (${t['WBS番号']}): ${res.msg}`);
        }
      } catch (err) {
        errorCount++;
        console.error(`\n  エラー (${t['WBS番号']}): ${err.message}`);
      }

      await sleep(200);
    }

    console.log(`\n\n=== タスク生成完了 ===`);
    console.log(`顧客: ${companyName}`);
    console.log(`成功: ${successCount}件`);
    console.log(`エラー: ${errorCount}件`);
    console.log(`\n開業予定日(${openingDateObj.toISOString().split('T')[0]})を基準にタスクの期限が設定されました。`);

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
