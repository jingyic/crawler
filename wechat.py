import requests
import time
import json
from pymongo import MongoClient
 
# 西安
sheetName = "插座学院"
 
# 目标url
url = "https://mp.weixin.qq.com/cgi-bin/appmsg"
 
Cookie = "noticeLoginFlag=1; remember_acct=35220942%40qq.com; pac_uid=0_98b541c92afb4; pgv_pvid=7220521180; ua_id=8bBH7Ctfs7SFA8qCAAAAAM3dsw97MZiN4pEok0SzXf4=; mm_lang=zh_CN; pgv_pvi=5941006336; noticeLoginFlag=1; xid=17f954c2de446b1fe63e03159f42c82c; tvfe_boss_uuid=c1444dcb17cbdb80; pgv_info=ssid=s3786794640; pgv_si=s2849615872; uuid=96ad88c3cac18737968d6e4351d7c454; cert=QYiYgN52qAAChPHWdRI6h1RO9Ivky3Ig; sig=h01b0b35e48c6882efb241a48c3fa6a1f14ec2a55ef1f6c850db2ff01c992a0c63b65a6174b5f01489a; _qpsvr_localtk=0.5105186424172053; ticket=4873ef501e878b1724589ecda398d95e7f68c86a; ticket_id=gh_79db24f9e420; rewardsn=; wxtokenkey=777; wxuin=1105735130; devicetype=Windows10; version=62060619; lang=zh_CN; pass_ticket=ZhiSXRSHeWqzO9QoF/ibhBEg4iTTvh04R0+hu3YymMvdWfvSQqte6zPX0+PuZuI0; wap_sid2=CNrboI8EElxNd1NfMm4ydUtUTzdNbm05RzVNYlFkMThCR1loNnZhMWtFUWd0bXl5cUFEVURVZ3Y4X0JOSXZ0QWh5c3VRd3NJY2hlbC1GV19iWXZkQkhhX3FCdTIzT29EQUFBfjD9uYzlBTgNQAE=; bizuin=3016251305; ptui_loginuin=544353686; uin=o0544353686; skey=@2sFgylEgC; RK=iA5srUmJVq; ptcz=210d45f5de26f6de7b89ee1cdd0bae5ce9e7687245084ab745bfd9a36bf6d6a7; device=; remember_acct=35220942%40qq.com; data_bizuin=3000250874; data_ticket=mYRcnkJD6QDLhlPXHdRWEEGGRVi19eVWphnF/Zpxw1d91/2iec7WjuyUvJdCgkHq; slave_sid=YjBXdEhLR3hzdF9Qb1JhazRidUZCZmw5RV9CM3hzMzdCUHZTQXB5dlkxU0hKNjZnSVM0djI0cF9RX0VEa1h1WUFIbVhWXzdabTc1VEZRMXNKRWVucUFMbEhlUUNaWHcxOUpGY1M3bG12V2NuYmpnaWlTaTBwbm1ScUUyWnFRNFFtM0lyOGhoUHBnMzFvblFR; slave_user=gh_79db24f9e420; openid2ticket_o-jHUssqdGVoz73unFoETvJjgy1Q=BkJTJsqJoy/zz9Puzfyz+A5Njz6fJMOTkBaEo5XNjF4="
# 使用Cookie，跳过登陆操作
headers = {
  "Cookie": Cookie,
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0",
}
 
"""
需要提交的data
以下个别字段是否一定需要还未验证。
注意修改yourtoken,number
number表示从第number页开始爬取，为5的倍数，从0开始。如0、5、10……
token可以使用Chrome自带的工具进行获取
fakeid是公众号独一无二的一个id，等同于后面的__biz
"""
token = "85841133"
fakeid ="MzAwNzEzNzU0Ng=="
# type在网页中会是10，但是无法取到对应的消息link地址，改为9就可以了
type = '9'
data1 = {
    "token": token,
    "lang": "zh_CN",
    "f": "json",
    "ajax": "1",
    "action": "list_ex",
    "begin": "365",
    "count": "5",
    "query": "",
    "fakeid": fakeid,
    "type": type,
}
 
 
# 毫秒数转日期
def getDate(times):
    # print(times)
    timearr = time.localtime(times)
    date = time.strftime("%Y-%m-%d %H:%M:%S", timearr)
    return date
 
 
# 获取阅读数和点赞数
def getMoreInfo(link):
    # 获得mid,_biz,idx,sn 这几个在link中的信息
    mid = link.split("&")[1].split("=")[1]
    idx = link.split("&")[2].split("=")[1]
    sn = link.split("&")[3].split("=")[1]
    _biz = link.split("&")[0].split("_biz=")[1]
 
    # fillder 中取得一些不变得信息
    #req_id = "0614ymV0y86FlTVXB02AXd8p"
    pass_ticket = "qqzhgtGjvPLJTLPz6c2H6nyeqTbKiUZee1TXxSvhx%252BfUgvZ%252B3LfccFqRrw%252BIlzMt"
    appmsg_token = "1005_5sl%2BFPZZw7oQdMsc49_tj8leiJI9pRVcKjNNzUjdoew1_UdYTc-Tn-5bRUhuNCzfbVYGyLLfDb32KP3q"
 
    # 目标url
    url = "http://mp.weixin.qq.com/mp/getappmsgext"
    # 添加Cookie避免登陆操作，这里的"User-Agent"最好为手机浏览器的标识
    phoneCookie = "devicetype=iOS12.1.4; lang=zh_CN; pass_ticket=qqzhgtGjvPLJTLPz6c2H6nyeqTbKiUZee1TXxSvhx+fUgvZ+3LfccFqRrw+IlzMt; version=1700032a; wap_sid2=CNrboI8EElx4a25kSHdxZG94cnhfUnctdVBURE9EbS1iUHNUaWJNQ1ZraUwwTU4xa0N0STBwallCNEVhUGZZdzg0MUdTVEF1ZDdNdFY1X20zeUstQV9GTTA3T3ZCTzBEQUFBfjCzmOflBTgNQJVO; wxuin=1105735130; rewardsn=; wxtokenkey=777; pgv_pvid=3789619064; ua_id=sVBCYAfushR89SHFAAAAAJRRE8zdjh_i2T7MznvqO4c=; mm_lang=zh_CN; RK=DI4kvUmZUI; ptcz=3c3d1040277093af5f8592d8836b14e8efc3e09237af35ba1baf6b112e013352; _scan_has_moon=1; ts_uid=4988706788; pac_uid=0_d1e42d80be652; qb_qua=; Q-H5-GUID=81e2a12990574b0a83ad4ceee1258d66; qb_guid=81e2a12990574b0a83ad4ceee1258d66; tvfe_boss_uuid=6163169dd7302e22; pgv_pvid_new=085e9858ee811fddd8c30738b@wx.tenpay.com_cabe8da970; pgv_pvi=3707342848"
    headers = {
        "Cookie": phoneCookie,
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/16D57 MicroMessenger/7.0.3(0x17000321) NetType/WIFI Language/zh_CN"
    }
    # 添加data，`req_id`、`pass_ticket`分别对应文章的信息，从fiddler复制即可。
    data = {
        "is_only_read": "1",
        "is_temp_url": "0",
        "appmsg_type": "9",
        'reward_uin_count':'0'
    }
    """
    添加请求参数
    __biz对应公众号的信息，唯一
    mid、sn、idx分别对应每篇文章的url的信息，需要从url中进行提取
    key、appmsg_token从fiddler上复制即可
    pass_ticket对应的文章的信息，也可以直接从fiddler复制
    """
    params = {
        "__biz": _biz,
        "mid": mid,
        "sn": sn,
        "idx": idx,
        "key": "777",
        "pass_ticket": pass_ticket,
        "appmsg_token": appmsg_token,
        "uin": "777",
        "wxtoken": "777",
    }
 
    # 使用post方法进行提交
    content = requests.post(url, headers=headers, data=data, params=params).json()
    # 提取其中的阅读数和点赞数
    print(content["appmsgstat"]["read_num"], content["appmsgstat"]["like_num"])
    try:
        readNum = content["appmsgstat"]["read_num"]
        # print(readNum)
    except:
        readNum=0
    try:
        likeNum = content["appmsgstat"]["like_num"]
        # print(likeNum)
    except:
        likeNum=0
    try:
        comment_count = content['comment_count']
        print("true:" + str(comment_count))
    except:
        comment_count=0
        print("error:"+str(comment_count))
 
 
    # 歇3s，防止被封
    time.sleep(3)
    return readNum, likeNum,comment_count
 
 
# 最大值365，所以range中就应该是73,15表示前3页
def getAllInfo(url, begin):
    # 拿一页，存一页
    messageAllInfo = []
    # begin 从0开始，365结束
    data1["begin"] = begin
    # 使用get方法进行提交
    content_json = requests.get(url, headers=headers, params=data1, verify=False).json()
    print(content_json)
    time.sleep(10)
    # 返回了一个json，里面是每一页的数据
    if "app_msg_list" in content_json:
        for item in content_json["app_msg_list"]:
            # 提取每页文章的标题及对应的url
            url = item['link']
            #print(url)
            readNum, likeNum ,comment_count= getMoreInfo(url)
            #print(readNum)
            info = {
                "title": item['title'],
                "readNum": readNum,
                "likeNum": likeNum,
                'comment_count':comment_count,
                "digest": item['digest'],
                "date": getDate(item['update_time']),
                "url": item['link']
            }
            messageAllInfo.append(info)
        # print(messageAllInfo)
        return messageAllInfo
    else:
        print("no app_msg_list")
 
 
# 写入数据库
def putIntoMogo(urlList):
    host = "127.0.0.1"
    port = 27017
 
    # 连接数据库
    client = MongoClient(host, port)
    # 建库
    chazuoxueyuan = client['chazuoxueyuan']
    # 建表
    wx_message_sheet = chazuoxueyuan[sheetName]
 
    # 存
    for message in urlList:
        wx_message_sheet.insert_one(message)
    print("成功！")
 
def main():
    # messageAllInfo = []
    # 爬10页成功，从11页开始
    for i in range(1,41):
        begin = i * 5
        messageAllInfo = getAllInfo(url, str(begin))
        print("第%s页" % i)
        putIntoMogo(messageAllInfo)
 
 
if __name__ == '__main__':
    main()
