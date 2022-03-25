import sys
import threading

import requests
import xlrd
import xlwt
import os
from concurrent.futures import ThreadPoolExecutor
import argparse

status = [200, 302, 301, 403]

rsp_success_list = []
rsp_fail_list = []


# lock = threading.RLock

def rsp_scan(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/99.0.4844.51 Safari/537.36 "
    }
    rsp_success_result = {}
    rsp_fail_result = {}
    try:
        rsp = requests.get(url=url, timeout=3, verify=True, headers=headers)
    except Exception as e:
        print(f"{url}请求失败 \nError:{e.args}")
        return False
    print(f"[{rsp.status_code}]:{url}")
    if rsp.status_code in status:
        rsp_success_result['url'] = url
        rsp_success_result['status_code'] = rsp.status_code
        rsp_success_list.append(rsp_success_result)
    else:
        rsp_fail_result['url'] = url
        rsp_fail_result['status_code'] = rsp.status_code
        rsp_fail_list.append(rsp_fail_result)
    return True


def read_xls(xls_path):
    web_url_list = []
    try:
        ws = xlrd.open_workbook(xls_path)
    except Exception as e:
        print(f"表格{xls_path}打开失败！\n{e}")
        exit(0)
    sheet = ws.sheets()[0]
    rows = sheet.nrows
    cols = sheet.ncols
    for row in range(1, rows):
        web_url_list.append(sheet.cell(row, 0).value)
    return web_url_list


def write_xls(xls_success_list, xls_fail_list):
    ws = xlwt.Workbook()
    sheet = ws.add_sheet('success')
    sheet1 = ws.add_sheet('fail')
    sheet.write(0, 0, 'url')
    sheet.write(0, 1, 'status_code')
    sheet1.write(0, 0, 'url')
    sheet1.write(0, 1, 'status_code')
    for i in range(len(xls_success_list)):
        sheet.write(i + 1, 0, xls_success_list[i]['url'])
        sheet.write(i + 1, 1, xls_success_list[i]['status_code'])
    for i in range(len(xls_fail_list)):
        sheet.write(i + 1, 0, xls_fail_list[i]['url'])
        sheet.write(i + 1, 1, xls_fail_list[i]['status_code'])
    print(f'save success, save path {os.getcwd()}/result.xls')
    ws.save("result.xls")


def main(args):
    xls_path = args.file
    if xls_path == None:
        sys.exit()
    # "/Users/ideal/Downloads/test_url.xls"
    url_list = read_xls(xls_path)
    url_list = tuple(url_list)
    with ThreadPoolExecutor(max_workers=20) as pool:
        pool.map(rsp_scan, url_list)
    # for url in url_list:
    #     rsp_scan(url)
    write_xls(xls_success_list=rsp_success_list, xls_fail_list=rsp_fail_list)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="python rsp_scan.py -f [filepath]")
    parser.add_argument("-f", help="输入的xls文件路径！", dest="file")
    parser.print_help()
    args = parser.parse_args()

    main(args)