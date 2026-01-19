/**
 * LarkBase 管理者権限追加スクリプト
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

async function addAdmin() {
  const appToken = process.env.LARK_BASE_APP_TOKEN;
  const email = 'hiroki.matsui@sei-san-sei.com';

  console.log('=== 管理者権限付与 ===\n');
  console.log('対象Base:', appToken);
  console.log('対象メール:', email);

  try {
    // 権限を付与（REST APIを直接呼び出し）
    console.log('\n権限を付与中...');

    const response = await client.request({
      method: 'POST',
      url: '/open-apis/drive/v1/permissions/' + appToken + '/members',
      params: {
        type: 'bitable',
        need_notification: true
      },
      data: {
        member_type: 'email',
        member_id: email,
        perm: 'full_access'
      }
    });

    console.log('API応答:', JSON.stringify(response, null, 2));

    if (response.code === 0) {
      console.log('\n✓ 管理者権限付与完了！');
      console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);
    } else {
      console.error('\n権限付与エラー:', response.msg);
    }

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

addAdmin();
