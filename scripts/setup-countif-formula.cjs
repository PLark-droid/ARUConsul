/**
 * COUNTIF + CurrentValue で完全自動計算を実現
 *
 * カテゴリ総数 = [タスク].COUNTIF(CurrentValue.[カテゴリ]=[カテゴリ])
 * カテゴリ完了数 = [タスク].COUNTIF(CurrentValue.[カテゴリ]=[カテゴリ]&&CurrentValue.[ステータス]="完了")
 * 完了率 = カテゴリ完了数 / カテゴリ総数
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
const taskTableId = 'tblHiAuZsUyWxlAY';

async function main() {
  console.log('=== COUNTIF数式で完全自動計算セットアップ ===\n');

  try {
    // ステップ1: 既存フィールドを削除
    console.log('ステップ1: 既存フィールドを削除...');

    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    console.log('現在のフィールド:');
    fields.forEach(f => console.log(`  - ${f.field_name} (type: ${f.type})`));

    const fieldsToDelete = ['完了率', 'カテゴリ総数', 'カテゴリ完了数', 'WBSカテゴリリンク'];
    for (const fieldName of fieldsToDelete) {
      const field = fields.find(f => f.field_name === fieldName);
      if (field) {
        await client.request({
          method: 'DELETE',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${field.field_id}`
        });
        console.log(`  ✓ ${fieldName} を削除`);
        await new Promise(r => setTimeout(r, 300));
      }
    }

    // フィールドIDを取得
    const fieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields2 = fieldsRes2.data.items || [];

    const categoryField = fields2.find(f => f.field_name === 'カテゴリ');
    const statusField = fields2.find(f => f.field_name === 'ステータス');

    if (!categoryField || !statusField) {
      console.log('エラー: カテゴリまたはステータスフィールドが見つかりません');
      return;
    }

    console.log(`\nカテゴリフィールドID: ${categoryField.field_id}`);
    console.log(`ステータスフィールドID: ${statusField.field_id}`);

    // ステップ2: カテゴリ総数（COUNTIF数式）
    console.log('\nステップ2: カテゴリ総数（COUNTIF）を作成...');

    // COUNTIF構文: bitable::$table[tableId].COUNTIF(CurrentValue.$field[fieldId]==$field[fieldId])
    const totalFormula = `bitable::$table[${taskTableId}].COUNTIF(CurrentValue.$field[${categoryField.field_id}]==$field[${categoryField.field_id}])`;

    console.log('数式:', totalFormula);

    const totalRes = await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'カテゴリ総数',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: totalFormula
        }
      }
    });
    console.log('✓ カテゴリ総数 作成成功');
    const totalFieldId = totalRes.data.field.field_id;

    await new Promise(r => setTimeout(r, 500));

    // ステップ3: カテゴリ完了数（COUNTIF数式 - 複数条件）
    console.log('\nステップ3: カテゴリ完了数（COUNTIF）を作成...');

    // 複数条件: CurrentValue.[カテゴリ]=[カテゴリ] && CurrentValue.[ステータス]="完了"
    const completedFormula = `bitable::$table[${taskTableId}].COUNTIF(CurrentValue.$field[${categoryField.field_id}]==$field[${categoryField.field_id}]&&CurrentValue.$field[${statusField.field_id}]=="完了")`;

    console.log('数式:', completedFormula);

    const completedRes = await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'カテゴリ完了数',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: completedFormula
        }
      }
    });
    console.log('✓ カテゴリ完了数 作成成功');
    const completedFieldId = completedRes.data.field.field_id;

    await new Promise(r => setTimeout(r, 500));

    // ステップ4: 完了率（数式）
    console.log('\nステップ4: 完了率を作成...');

    const rateFormula = `IF($field[${totalFieldId}]>0,$field[${completedFieldId}]/$field[${totalFieldId}],0)`;

    console.log('数式:', rateFormula);

    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: '完了率',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: rateFormula
        }
      }
    });
    console.log('✓ 完了率 作成成功');

    console.log('\n=== セットアップ完了 ===');
    console.log('\nLarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n【完全自動計算】');
    console.log('- カテゴリ総数: COUNTIF(カテゴリが同じレコード)');
    console.log('- カテゴリ完了数: COUNTIF(カテゴリが同じ && ステータス=完了)');
    console.log('- 完了率: カテゴリ完了数 / カテゴリ総数');
    console.log('\nステータスを変更すると自動で再計算されます！');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
