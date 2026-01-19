/**
 * ロールアップと数式フィールドを追加
 */

const lark = require('@larksuiteoapi/node-sdk');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: lark.Domain.Feishu,
});

const appToken = process.env.LARK_BASE_APP_TOKEN;
const categoryTableId = 'tbleaLCFajhx7KcN';
const taskTableId = 'tblHiAuZsUyWxlAY';

async function main() {
  console.log('=== ロールアップ・数式フィールド追加 ===\n');

  try {
    // WBSカテゴリのフィールドを取得
    const catFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields = catFieldsRes.data.items || [];

    console.log('WBSカテゴリのフィールド:');
    catFields.forEach(f => console.log(`  - ${f.field_name} (type: ${f.type}, id: ${f.field_id})`));

    // タスクテーブルのフィールドを取得
    const taskFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const taskFields = taskFieldsRes.data.items || [];

    console.log('\nタスクのフィールド:');
    taskFields.forEach(f => console.log(`  - ${f.field_name} (type: ${f.type}, id: ${f.field_id})`));

    // 双方向リンクフィールドを探す（type 18 = リンク）
    const backLinkField = catFields.find(f => f.type === 18);
    const statusField = taskFields.find(f => f.field_name === 'ステータス');

    if (!backLinkField) {
      console.log('\nエラー: 双方向リンクフィールドが見つかりません');
      return;
    }

    console.log(`\n双方向リンクフィールド: ${backLinkField.field_name} (${backLinkField.field_id})`);
    console.log(`ステータスフィールド: ${statusField ? statusField.field_id : 'not found'}`);

    // タスク数のロールアップを追加
    console.log('\nタスク数（自動）ロールアップを追加中...');
    const existingTaskCount = catFields.find(f => f.field_name === 'タスク数（自動）');

    if (!existingTaskCount) {
      const rollupRes = await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
        data: {
          field_name: 'タスク数（自動）',
          type: 20,
          property: {
            link_field_id: backLinkField.field_id,
            rollup_type: 'COUNTA'
          }
        }
      });
      console.log('応答:', JSON.stringify(rollupRes, null, 2));
    } else {
      console.log('既存');
    }

    await new Promise(r => setTimeout(r, 500));

    // 完了数のロールアップを追加（ステータスが完了のものをカウント）
    console.log('\n完了数（自動）ロールアップを追加中...');
    const existingCompleted = catFields.find(f => f.field_name === '完了数（自動）');

    if (!existingCompleted && statusField) {
      // 条件付きロールアップが難しい場合、数式で対応
      // まずは単純なロールアップを試す
      try {
        const rollupRes2 = await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
          data: {
            field_name: '完了数（自動）',
            type: 20,
            property: {
              link_field_id: backLinkField.field_id,
              target_field_id: statusField.field_id,
              rollup_type: 'COUNTALL',
              filter_info: {
                conjunction: 'and',
                conditions: [{
                  field_id: statusField.field_id,
                  operator: 'is',
                  value: ['完了']
                }]
              }
            }
          }
        });
        console.log('応答:', JSON.stringify(rollupRes2, null, 2));
      } catch (e) {
        console.log('条件付きロールアップエラー:', e.message);
      }
    } else {
      console.log('既存');
    }

    await new Promise(r => setTimeout(r, 500));

    // 完了率の数式フィールドを追加
    console.log('\n完了率（数式）フィールドを追加中...');
    const existingRate = catFields.find(f => f.field_name === '完了率');

    // 最新のフィールド一覧を再取得
    const catFieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields2 = catFieldsRes2.data.items || [];

    const taskCountField = catFields2.find(f => f.field_name === 'タスク数（自動）');
    const completedField = catFields2.find(f => f.field_name === '完了数（自動）');

    console.log('タスク数フィールドID:', taskCountField ? taskCountField.field_id : 'なし');
    console.log('完了数フィールドID:', completedField ? completedField.field_id : 'なし');

    if (!existingRate && taskCountField) {
      // 数式フィールドを追加
      try {
        const formulaRes = await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
          data: {
            field_name: '完了率',
            type: 19,  // 数式
            property: {
              formula_expression: completedField
                ? `IF([タスク数（自動）]>0,ROUND([完了数（自動）]/[タスク数（自動）]*100,0),0)`
                : `0`
            }
          }
        });
        console.log('数式フィールド応答:', JSON.stringify(formulaRes, null, 2));
      } catch (e) {
        console.log('数式フィールドエラー:', e.message);
        if (e.response && e.response.data) {
          console.log('詳細:', JSON.stringify(e.response.data, null, 2));
        }
      }
    }

    console.log('\n=== 完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
