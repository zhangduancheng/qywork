import export_qywework_data

if __name__ == '__main__':
    corpid = "" # 企业微信的corpid
    corpsecret = "" # 企业微信的corpsecret
    access_token = export_qywework_data.get_token(corpid, corpsecret)
    if access_token:
        export_qywework_data.export_qywework_data(access_token)
    else:
        print("获取token失败")