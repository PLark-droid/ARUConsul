/**
 * テーブル名を確認し、正しい数式を設定
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

async function main() {
  console.log('=== テーブル情報確認 ===\n');

  try {
    // 全テーブル一覧を取得
    const tablesRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables`
    });
    const tables = tablesRes.data.items || [];

    console.log('テーブル一覧:');
    tables.forEach(t => {
      console.log(`  - ${t.name} (ID: ${t.table_id})`);
    });

    // タスクテーブルの正式名称を取得
    const taskTable = tables.find(t => t.table_id === 'tblHiAuZsUyWxlAY');
    if (taskTable) {
      console.log(`\nタスクテーブルの正式名称: "${taskTable.name}"`);

      // 現在の数式を確認
      const fieldsRes = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`
      });
      const fields = fieldsRes.data.items || [];

      const totalField = fields.find(f => f.field_name === 'カテゴリ総数');
      const completedField = fields.find(f => f.field_name === 'カテゴリ完了数');

      if (totalField) {
        console.log(`\n現在のカテゴリ総数の数式:`);
        console.log(`  ${totalField.property?.formula_expression}`);

        // テーブル名が違う場合、修正が必要
        if (taskTable.name !== 'タスク') {
          console.log(`\n⚠️ テーブル名が「タスク」ではなく「${taskTable.name}」です`);
          console.log(`数式を修正する必要があります。`);

          // 数式を修正
          console.log('\n数式を修正中...');

          // カテゴリ総数を削除して再作成
          await client.request({
            method: 'DELETE',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields/${totalField.field_id}`
          });
          console.log('  カテゴリ総数を削除');
          await new Promise(r => setTimeout(r, 500));

          // 正しいテーブル名で再作成
          const correctFormula = `[${taskTable.name}].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ])`;
          console.log(`  新しい数式: ${correctFormula}`);

          await client.request({
            method: 'POST',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`,
            data: {
              field_name: 'カテゴリ総数',
              type: 20,
              ui_type: 'Formula',
              property: {
                formula_expression: correctFormula
              }
            }
          });
          console.log('  ✓ カテゴリ総数を再作成');

          await new Promise(r => setTimeout(r, 500));

          // カテゴリ完了数も修正
          if (completedField) {
            await client.request({
              method: 'DELETE',
              url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields/${completedField.field_id}`
            });
            console.log('  カテゴリ完了数を削除');
            await new Promise(r => setTimeout(r, 500));
          }

          const correctCompletedFormula = `[${taskTable.name}].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ]&&CurrentValue.[ステータス]=="完了")`;
          console.log(`  新しい数式: ${correctCompletedFormula}`);

          await client.request({
            method: 'POST',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`,
            data: {
              field_name: 'カテゴリ完了数',
              type: 20,
              ui_type: 'Formula',
              property: {
                formula_expression: correctCompletedFormula
              }
            }
          });
          console.log('  ✓ カテゴリ完了数を再作成');

          await new Promise(r => setTimeout(r, 500));

          // 完了率も再作成
          const rateField = fields.find(f => f.field_name === '完了率');
          if (rateField) {
            await client.request({
              method: 'DELETE',
              url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields/${rateField.field_id}`
            });
            console.log('  完了率を削除');
            await new Promise(r => setTimeout(r, 500));
          }

          // 最新のフィールドIDを取得
          const newFieldsRes = await client.request({
            method: 'GET',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`
          });
          const newFields = newFieldsRes.data.items || [];
          const newTotalField = newFields.find(f => f.field_name === 'カテゴリ総数');
          const newCompletedField = newFields.find(f => f.field_name === 'カテゴリ完了数');

          if (newTotalField && newCompletedField) {
            const rateFormula = `IF($field[${newTotalField.field_id}]>0,$field[${newCompletedField.field_id}]/$field[${newTotalField.field_id}],0)`;
            console.log(`  新しい数式: ${rateFormula}`);

            await client.request({
              method: 'POST',
              url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`,
              data: {
                field_name: '完了率',
                type: 20,
                ui_type: 'Formula',
                property: {
                  formula_expression: rateFormula
                }
              }
            });
            console.log('  ✓ 完了率を再作成');
          }
        }
      }

      // 最終確認
      console.log('\n=== 最終フィールド状態 ===');
      const finalFieldsRes = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTable.table_id}/fields`
      });
      const finalFields = finalFieldsRes.data.items || [];

      console.log('\n計算用フィールド:');
      ['カテゴリ総数', 'カテゴリ完了数', '完了率'].forEach(name => {
        const f = finalFields.find(field => field.field_name === name);
        if (f) {
          console.log(`\n  ${f.field_name}:`);
          console.log(`    数式: ${f.property?.formula_expression || 'なし'}`);
        }
      });
    }

    console.log('\n=== 完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
