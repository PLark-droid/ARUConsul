/**
 * タスクテーブルの完了率を数式フィールドに変更するスクリプト
 * - WBSカテゴリテーブルから手動追加フィールドを削除
 * - タスク→WBSカテゴリのリンクを作成
 * - WBSカテゴリにロールアップで集計
 * - タスクにルックアップで完了率を表示
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
const categoryTableId = 'tbleaLCFajhx7KcN';  // WBSカテゴリ
const taskTableId = 'tblHiAuZsUyWxlAY';      // タスク

async function main() {
  console.log('=== 完了率フィールド再構成 ===\n');

  try {
    // ステップ1: WBSカテゴリテーブルから手動追加フィールドを削除
    console.log('ステップ1: WBSカテゴリの手動フィールドを削除中...');

    const catFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });

    const catFields = catFieldsRes.data.items || [];
    const fieldsToDelete = ['タスク数', '完了タスク数', '完了率'];

    for (const field of catFields) {
      if (fieldsToDelete.includes(field.field_name)) {
        await client.request({
          method: 'DELETE',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields/${field.field_id}`
        });
        console.log(`  ✓ ${field.field_name} を削除`);
        await new Promise(r => setTimeout(r, 300));
      }
    }

    // ステップ2: タスクテーブルの完了率フィールドを削除（数値→数式に変更のため）
    console.log('\nステップ2: タスクテーブルの既存完了率フィールドを確認...');

    const taskFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });

    const taskFields = taskFieldsRes.data.items || [];
    const completionField = taskFields.find(f => f.field_name === '完了率');

    if (completionField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${completionField.field_id}`
      });
      console.log('  ✓ 既存の完了率フィールドを削除');
      await new Promise(r => setTimeout(r, 300));
    }

    // ステップ3: タスク→WBSカテゴリのリンクフィールドを作成
    console.log('\nステップ3: リンクフィールドを作成中...');

    const existingLink = taskFields.find(f => f.field_name === 'WBSカテゴリリンク');

    let linkFieldId;
    if (!existingLink) {
      const linkRes = await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: 'WBSカテゴリリンク',
          type: 18,  // リンク
          property: {
            table_id: categoryTableId,
            multiple: false
          }
        }
      });
      linkFieldId = linkRes.data.field.field_id;
      console.log('  ✓ WBSカテゴリリンクフィールド作成');
    } else {
      linkFieldId = existingLink.field_id;
      console.log('  - WBSカテゴリリンクフィールドは既存');
    }

    await new Promise(r => setTimeout(r, 500));

    // ステップ4: WBSカテゴリにロールアップフィールドを作成
    console.log('\nステップ4: WBSカテゴリにロールアップフィールドを作成中...');

    // 再取得
    const catFieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields2 = catFieldsRes2.data.items || [];

    // 双方向リンクフィールドを探す
    const backLinkField = catFields2.find(f => f.type === 18);

    if (backLinkField) {
      console.log(`  - 双方向リンクフィールド: ${backLinkField.field_name}`);

      // タスク数のロールアップ
      const existingTaskCount = catFields2.find(f => f.field_name === 'タスク数（自動）');
      if (!existingTaskCount) {
        await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
          data: {
            field_name: 'タスク数（自動）',
            type: 20,  // ロールアップ
            property: {
              link_field_id: backLinkField.field_id,
              rollup_type: 'COUNTA'
            }
          }
        });
        console.log('  ✓ タスク数（自動）ロールアップ作成');
      }

      await new Promise(r => setTimeout(r, 500));
    }

    // ステップ5: タスクのカテゴリリンクを自動設定
    console.log('\nステップ5: タスクのカテゴリリンクを設定中...');

    // カテゴリレコードを取得
    const categoriesRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/records`,
      params: { page_size: 100 }
    });
    const categories = categoriesRes.data.items || [];

    // WBS番号→record_idのマップを作成
    const categoryMap = {};
    for (const cat of categories) {
      categoryMap[cat.fields['WBS番号']] = cat.record_id;
    }

    // 全タスクを取得してリンクを設定
    let allTasks = [];
    let pageToken = null;

    do {
      const params = { page_size: 500 };
      if (pageToken) params.page_token = pageToken;

      const tasksRes = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
        params
      });

      if (tasksRes.data.items) {
        allTasks = allTasks.concat(tasksRes.data.items);
      }
      pageToken = tasksRes.data.page_token;
    } while (pageToken);

    let linkedCount = 0;
    for (const task of allTasks) {
      const wbs = task.fields['WBS番号'];
      if (!wbs || !wbs.includes('.')) continue;

      const categoryNum = wbs.split('.')[0];
      const categoryRecordId = categoryMap[categoryNum];

      if (categoryRecordId) {
        await client.request({
          method: 'PUT',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records/${task.record_id}`,
          data: {
            fields: {
              'WBSカテゴリリンク': [{ record_id: categoryRecordId }]
            }
          }
        });
        linkedCount++;
        await new Promise(r => setTimeout(r, 100));
      }
    }
    console.log(`  ✓ ${linkedCount}件のタスクをカテゴリにリンク`);

    // ステップ6: 完了率の数式フィールドを作成
    console.log('\nステップ6: 完了率の数式フィールドを作成中...');

    // WBSカテゴリに完了率数式フィールドを追加
    const catFieldsRes3 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields3 = catFieldsRes3.data.items || [];

    const existingCatRate = catFields3.find(f => f.field_name === '完了率');
    if (!existingCatRate) {
      // ステータスフィールドを探す
      const taskFieldsRes2 = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
      });
      const taskFields2 = taskFieldsRes2.data.items || [];
      const statusField = taskFields2.find(f => f.field_name === 'ステータス');
      const backLink = catFields3.find(f => f.type === 18);

      if (backLink && statusField) {
        // 完了タスク数のロールアップ（条件付き）
        await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
          data: {
            field_name: '完了数（自動）',
            type: 20,  // ロールアップ
            property: {
              link_field_id: backLink.field_id,
              target_field_id: statusField.field_id,
              rollup_type: 'COUNTA',
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
        console.log('  ✓ 完了数（自動）ロールアップ作成');
      }
    }

    console.log('\n=== 完了 ===');
    console.log('\n次の手順:');
    console.log('1. LarkBaseを開く: https://www.feishu.cn/base/' + appToken);
    console.log('2. WBSカテゴリテーブルで「フィールド追加」→「数式」を選択');
    console.log('3. 数式: IF(タスク数（自動）>0, 完了数（自動）/タスク数（自動）*100, 0)');
    console.log('4. フィールド名: 完了率');
    console.log('\nまたはタスクテーブルで:');
    console.log('1. 「フィールド追加」→「ルックアップ」を選択');
    console.log('2. リンク先: WBSカテゴリリンク');
    console.log('3. 表示フィールド: 完了率');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
