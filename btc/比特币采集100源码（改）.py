import os
import xlwt
import time
import random
import requests
from pathlib import Path
from xlutils.copy import copy
from xlrd import open_workbook

user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.57.2 (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2 ",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.101 Safari/537.36",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36 OPR/26.0.1656.60",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.71 Safari/537.1 LBBROWSER",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.3.4000 Chrome/30.0.1599.101 Safari/537.36",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.122 UBrowser/4.0.3214.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36",
    "Opera/12.80 (Windows NT 5.1; U; en) Presto/2.10.289 Version/12.02",
]

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    # 'Cache-Control': 'no-cache',
    # 'Connection': 'keep-alive',
    # 'referer': 'https://www.google.com/',
    # 'Upgrade-Insecure-Requests': '1',
    'User-Agent': random.choice(user_agents)
}

def mkdir(path):
    path = path.strip()
    path = path.rstrip("\\")
    isExists = os.path.exists(path)
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print(path + ' 创建成功')
        return True


def get_data(page, pro=False):
    # if pro:
    #     pro_list = rds.lrange("ProxyL", 0, -1)
    #     p = random.choice(pro_list)
    #     proxies = {"http": p, "https": p}
    # else:
    #     proxies = {}
    # proxies = {}
    url = "https://btc.tokenview.com/api/blocks/btc/{}/100".format(page) #api地址根据网页F12-“网络”确定
    # url = "https://explorer-web.api.btc.com/v1/eth/blocks?page={}&size=150".format(page)
    try:
        req = requests.get(url, headers=headers, timeout=20)
        if req.status_code != 200:
            raise
        else:
            return req.json()
    except Exception as e:
        print("req err:", e)
        return get_data(page, pro=True)


def save_item(time_stamp, item_list):
    # 创建文件夹
    year_dir_path = "{}\{}".format(os.getcwd(), time_stamp.tm_year)
    month_dir_path = "{}\{}".format(year_dir_path, '%02d' % time_stamp.tm_mon)
    table_name = "{}.xlsx".format(time.strftime('%Y-%m-%d', time_stamp))

    mkdir(year_dir_path)
    mkdir(month_dir_path)

    path = "{}\{}".format(month_dir_path, table_name)
    print("写入{}".format(path), len(item_list))
    my_file = Path(path)
    if not my_file.is_file():
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('1')
        worksheet.write(0, 0, label='高度')
        worksheet.write(0, 1, label='时间')
        worksheet.write(0, 2, label='播报方')
        worksheet.write(0, 3, label='大小(B)')
        worksheet.write(0, 4, label='奖励(BTC)')
        worksheet.write(0, 5, label='交易数')
        worksheet.write(0, 6, label='手续费(BTC)')
        worksheet.write(0, 7, label='交易总额(BTC)')
        workbook.save(path)
    r_xls = open_workbook(path)  # 读取excel文件
    sheet = r_xls.sheet_by_name("1")
    row = sheet.nrows  # 获取已有的行数
    excel = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    table = excel.get_sheet("1")
    for i in range(8):
        table.col(i).width = 256 * 15
    for one_item in item_list:
        # 获取数据
        block_no = one_item.get("block_no", "")  # 高度
        created_ts = one_item.get("time", "")  # 时间
        created_time_stamp = time.localtime(created_ts)
        created_time = time.strftime('%Y-%m-%d %H:%M:%S', created_time_stamp)
        minerAlias = one_item.get("minerAlias") if one_item.get("minerAlias") else one_item.get("miner", "")  # 播报方
        size = one_item.get("size", "")  # 区块大小(B)
        reward = round(float(one_item.get("reward", "0")), 5)  # 奖励(BTC)
        txCnt = one_item.get("txCnt", "")  # 交易数
        fee = round(float(one_item.get("fee", "0")), 5)  # 手续费(BTC)
        sentValue = round(float(one_item.get("sentValue", "0")), 5)  # 交易总额(BTC)

        table.write(row, 0, block_no)
        table.write(row, 1, created_time)
        table.write(row, 2, minerAlias)
        table.write(row, 3, size)
        table.write(row, 4, reward)
        table.write(row, 5, txCnt)
        table.write(row, 6, fee)
        table.write(row, 7, sentValue)
        row+=1
    excel.save(path)


def start(end_id):
    page = 1
    # page = 24487
    # end_id = 0
    break_swith = False
    save_block_height = 99999999999999
    # save_block_height = 930
    item_list = []
    p = time.localtime(0)
    while True:
        if break_swith:
            print("结束")
            break
        result = get_data(page)  #完整的一页数据
        result_list = result.get("data", {})  #需要具体对照
        # result_list = result.get("data", {}).get("list", [])
        for one_item in result_list:
            # 获取数据
            block_height = one_item.get("block_no", "")  # 高度
            if save_block_height <= block_height:
                print("已保存, 跳过", save_block_height, block_height)
                continue
            save_block_height = block_height
            created_ts = one_item.get("time", "")  # 时间
            created_time_stamp = time.localtime(created_ts)
            if created_time_stamp.tm_mday == p.tm_mday:
                item_list.append(one_item)
            else:
                if item_list:
                    save_item(p, item_list)
                    item_list = [] #清空列表？
                    item_list.append(one_item)
            p = created_time_stamp
            print(page, block_height, time.strftime('%Y-%m-%d %H:%M:%S', created_time_stamp), len(item_list))
            # save_item(month_dir_path, table_name, one_item)
            if block_height <= int(end_id):
                save_item(p, item_list)
                break_swith = True
                break
        page += 1
        # break

def main():
    # if time.time() >= 1618848000:
    #     print("试用结束, 请联系作者")
    #     time.sleep(60)
    end_id = input("请输入停止高度:")
    try:
        int(end_id)
    except:
        time.sleep(1)
        print("输入错误, 请重新输入")
        return main()
    start(end_id)


if __name__ == '__main__':
    # start()
    main()