/**
 * 既存タスクテーブルにリンクフィールドを追加するスクリプト
 */

const lark = require('@larksuiteoapi/node-sdk');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: process.env.LARK_DOMAIN === 'larksuite' ? lark.Domain.Lark : lark.Domain.Feishu,
});

const APP_TOKEN = process.env.LARK_BASE_APP_TOKEN;
const TASK_TABLE_ID = process.env.TASK_TABLE_ID;
const CUSTOMER_TABLE_ID = process.env.CUSTOMER_TABLE_ID;
const TEMPLATE_TABLE_ID = process.env.TEMPLATE_TABLE_ID;

async function main() {
  console.log('=== リンクフィールド追加 ===\n');

  if (!APP_TOKEN || !TASK_TABLE_ID || !CUSTOMER_TABLE_ID || !TEMPLATE_TABLE_ID) {
    console.error('エラー: .envに必要なIDが設定されていません');
    console.log('必要な設定:');
    console.log('  LARK_BASE_APP_TOKEN');
    console.log('  TASK_TABLE_ID');
    console.log('  CUSTOMER_TABLE_ID');
    console.log('  TEMPLATE_TABLE_ID');
    process.exit(1);
  }

  console.log(`APP_TOKEN: ${APP_TOKEN}`);
  console.log(`TASK_TABLE_ID: ${TASK_TABLE_ID}`);
  console.log(`CUSTOMER_TABLE_ID: ${CUSTOMER_TABLE_ID}`);
  console.log(`TEMPLATE_TABLE_ID: ${TEMPLATE_TABLE_ID}`);

  try {
    // 顧客リンクフィールドを追加
    console.log('\n1. 顧客リンクフィールドを追加中...');
    const linkRes1 = await client.bitable.appTableField.create({
      path: { app_token: APP_TOKEN, table_id: TASK_TABLE_ID },
      data: {
        field_name: '顧客',
        type: 18, // 単向関連
        property: {
          table_id: CUSTOMER_TABLE_ID
        }
      }
    });

    if (linkRes1.code === 0) {
      console.log('   ✓ 顧客リンク追加完了');
    } else {
      console.log(`   ⚠ エラー: ${linkRes1.msg} (code: ${linkRes1.code})`);
    }

    await new Promise(r => setTimeout(r, 500));

    // テンプレートリンクフィールドを追加
    console.log('\n2. テンプレートリンクフィールドを追加中...');
    const linkRes2 = await client.bitable.appTableField.create({
      path: { app_token: APP_TOKEN, table_id: TASK_TABLE_ID },
      data: {
        field_name: 'テンプレート',
        type: 18, // 単向関連
        property: {
          table_id: TEMPLATE_TABLE_ID
        }
      }
    });

    if (linkRes2.code === 0) {
      console.log('   ✓ テンプレートリンク追加完了');
    } else {
      console.log(`   ⚠ エラー: ${linkRes2.msg} (code: ${linkRes2.code})`);
    }

    console.log('\n=== 完了 ===');

  } catch (error) {
    console.error('\nエラー:', error.message);
    process.exit(1);
  }
}

main();
