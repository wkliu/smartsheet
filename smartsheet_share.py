import smartsheet
from myconfig import *
proxies = {'http': 'http://proxy.esl.cisco.com:80/', 'https':'http://proxy.esl.cisco.com:80/'}
#ss = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)
ss = smartsheet.Smartsheet(access_token)
SEs=[davidtai, andrewyang, stanhuang, jimcheng, karlhsieh, vincenthsu, angelalin, barryhuang, vanhsieh, rickywang, tonyhsieh, willyhuang, jerrylin, allentseng, vinceliu]

for se in SEs:
    for item in se:
        if item != "email":
            print(se[item], se["email"])
            newShareObj = ss.models.Share()
            newShareObj.email = se["email"]
            newShareObj.access_level = "EDITOR"
            sheet = ss.Sheets.get_sheet(se[item])
            sheet.share(newShareObj, send_email=False)
        #print(item)
#newShareObj = ss.models.Share()
#newShareObj.email = "angelin@cisco.com"
#newShareObj.access_level = "EDITOR"
#sheet = ss.Sheets.get_sheet(angelalin["KSO"])
#sheet.share(newShareObj)