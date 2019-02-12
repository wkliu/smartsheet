import smartsheet
from myconfig import *
proxies = {'http': 'http://proxy.esl.cisco.com:80/', 'https':'http://proxy.esl.cisco.com:80/'}
#ss = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)
ss = smartsheet.Smartsheet(access_token)
SEs=[davidtai, andrewyang, stanhuang, jimcheng, karlhsieh, vincenthsu, angelalin, barryhuang, vanhsieh, rickywang, tonyhsieh, willyhuang, jerrylin, allentseng, vinceliu]
Share_Sheets = [destination_sheetIds["POST_SALE"]]
for se in SEs:
    for item in Share_Sheets:
        #print(se[item], se["email"])
        newShareObj = ss.models.Share()
        newShareObj.email = se["email"]
        #newShareObj.access_level = "EDITOR"
        newShareObj.access_level = "VIEWER"
        print(newShareObj, item)
        #sheet = ss.Sheets.get_sheet(item)
        #sheet.share(newShareObj, send_email=False)
        sight = ss.Sights.share_sight(item, newShareObj, False)
        #print(item)
#newShareObj = ss.models.Share()
#newShareObj.email = "angelin@cisco.com"
#newShareObj.access_level = "EDITOR"
#sheet = ss.Sheets.get_sheet(angelalin["KSO"])
#sheet.share(newShareObj)
