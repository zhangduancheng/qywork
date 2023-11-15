import json
import requests
import time
import pandas as pd

sources = {
    0: "未知来源",
    1: "扫描二维码",
    2: "搜索手机号",
    3: "名片分享",
    4: "群聊",
    5: "手机通讯录",
    6: "微信联系人",
    8: "安装第三方应用时自动添加的客服人员",
    9: "搜索邮箱",
    10: "视频号添加",
    11: "通过日程参与人添加",
    12: "通过会议参与人添加",
    13: "添加微信好友对应的企业微信",
    14: "通过智慧硬件专属客服添加",
    15: "通过上门服务客服添加",
    16: "通过获客链接添加",
    17: "通过定制开发添加",
    18: "通过需求回复添加",
    201: "内部成员共享",
    202: "管理员/负责人分配"
}

# 获取企业微信调用接口的token
def get_token(corpid, corpsecret):

    url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid="+ corpid +"&corpsecret=" + corpsecret

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)
    json_data = json.loads(response.text)
    if json_data['errcode'] == 0:
        return json_data['access_token']
    else:
        print(f"Error: {json_data['errmsg']}")

def exec_request(token, payload):
    # 批量获取客户详情 企业/第三方可通过此接口获取指定成员添加的客户信息列表。
    url = "https://qyapi.weixin.qq.com/cgi-bin/externalcontact/batch/get_by_user?access_token=" + token

    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    return response.text
def process_contacts(data, contacts):
    # 迭代处理每个联系人信息
    for contact in data['external_contact_list']:
        follow_info = contact['follow_info']
        external_contact = contact['external_contact']
        # 外部联系人的user_id
        external_userid = external_contact['external_userid']
        # 外部联系人的名称
        name = external_contact['name']
        # 外部联系人的头像
        avatar = external_contact['avatar']
        # 外部联系人的类型，如果为1表示该外部联系人是微信用户，如果为2表示该外部联系人是企业微信用户
        type = "微信用户" if external_contact['type'] == 1 else "企业微信用户"
        # 外部联系人的性别，如果external_contact['gender']为1表示该外部联系人是男性，如果为2表示该外部联系人是女性，0表示未知
        gender = "男" if external_contact['gender'] == 1 else "女" if external_contact['gender'] == 2 else "未知"
        # 添加人的user_id
        add_userid = follow_info['userid']
        # 备注
        remark = follow_info['remark']
        # 描述
        description = follow_info['description']
        # 添加时间，为unix时间戳，单位为秒，转换成格式为2021-01-01 00:00:00的字符串
        createtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(follow_info['createtime']))
        # 企业标签id
        tag_id = follow_info['tag_id']
        # 备注手机号码
        remark_mobiles = follow_info['remark_mobiles']
        # 添加来源
        add_way = follow_info['add_way']
        source = sources.get(add_way, "未知来源")

        # 发起添加的userid
        oper_userid = follow_info['oper_userid']
        # 将每个联系人的信息添加到列表中
        contacts.append({
            "外部联系人的user_id": external_userid,
            "客户名称": name,
            "客户头像": avatar,
            "客户类型": type,
            "性别": gender,
            "添加人的user_id": add_userid,
            "备注": remark,
            "描述": description,
            "添加时间": createtime,
            "企业标签id": tag_id,
            "备注手机号码": remark_mobiles,
            "来源": source,
            "添加人": oper_userid
        })

def export_to_excel(contacts):
    df = pd.DataFrame(contacts)
    df.to_excel("exported_contacts.xlsx", index=False)

def export_qywework_data(token):
    contacts = []
    cursor = ""

    while True:
        payload = json.dumps({
            # 企业成员的userid列表，字符串类型，最多支持100个
            "userid_list": [
                "zhangsan",
                "lisi"
            ],
            "cursor": cursor,
            "limit": 100
        })

        try:
            result = exec_request(token, payload)
            data = json.loads(result)
            if data['errcode'] == 0 and data['errmsg'] == 'ok':
                process_contacts(data, contacts)
                cursor = data.get('cursor', "")
                if not cursor:
                    break
            else:
                print(f"Error: {data['errmsg']}")
                break
        except requests.RequestException as e:
            print(f"Request failed: {e}")
            break

    export_to_excel(contacts)