#!/usr/bin/python
#-*- coding: UTF-8 -*-

import os, sys, xlsxwriter, traceback, requests, json, urllib, re, datetime, time


BASE_URL = "https://cps.qixin18.com/product/tryTrial?"

PARAMS = {
    "callback": "jQuery18200722824527054019_1576236583334",
    "uid": "1036313",
    "prodId": "104004",
    "planId": "128447",
    "restrictGeneParams": {
        "productId": "104004",
        "productPlanId": "128447",
        "genes":[
            {"sort":5,"protectItemId":8490,"key":"","value":"50万元"},
            {"sort":1,"protectItemId":"","key":"insurantDate","value":"2001-10-24"},
            {"sort":2,"protectItemId":"","key":"sex","value":"男"},
            {"sort":3,"protectItemId":"","key":"insurantJob","value":"1-6类"},
            {"sort":4,"protectItemId":"","key":"city","value":"310100"},
            {"protectItemId":"","key":"province","value":"310000"},
            {"sort":6,"protectItemId":"","key":"insurantDateLimit","value":"至60岁"},
            {"sort":7,"protectItemId":"","key":"paymentType","value":"年交"},
            {"sort":8,"protectItemId":"","key":"insureAgeLimit","value":"20年"},
            {"sort":9,"protectItemId":"","key":"insureForSelf","value":"是"},
            {"sort":10,"protectItemId":"","key":"vesterAge","value":"2001-10-31"},
            {"sort":11,"protectItemId":"","key":"vesterSex","value":"男"}
        ],
        "optGeneOldValue":{"sort":1,"protectItemId":"","key":"insurantDate","value":"2001-12-14"}
    },
    "_": "1576236596620"
}

COOKIES = {
    "hz_guest_key": "31Rq3EhNDHZ16JhQ9maj_1576217231024_2_1036313_0",
    "hz_view_key": "31Rq3EhNDHZ16JhQ9maj3CM8TNbSLHZ30825P7ln_1576236583458_https%253A%252F%252Fcps.qixin18.com%252Flxc1036313%252Fproduct%252Fdetail-104004-128447.html%253Ffrom%253Dspks",
    "hz_visit_key": "1EFKEgpkBHZ2QnvwsCNH_1576236566492_4_1576236566492",
    "cuidlxc1036313": "c9cea411-88f8-4025-9865-cfe0b133f944",
    "nodejs_sid": "s%3APk9dNuFjAPdfUb6cu4IbNzF94175ntu6.eAWiaYSe38IcNZQsqwnE52mY7iFVr5%2BGp0hSK6LLaCU",
    "acw_tc": "2f6a1fb315762172304101926e042efd471e74f28b5b546cfeec172b0b9bb0"
}

HEADERS = {
    "Accept" : "text/javascript, application/javascript, application/ecmascript, application/x-ecmascript, */*; q=0.01",
    "Content-Type": "application/json",
    "Host": "cps.qixin18.com",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.2 Safari/605.1.15",
    "Referer": "https://cps.qixin18.com/lxc1036313/product/detail-104004-128447.html?from=spks",
    "X-Requested-With": "XMLHttpRequest",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive"
}

EXCEL_COL_WIDTH = 50
EXCEL_ROW_HEIGHT = 30
EXCEL_COL_NAMES = ["年龄", "性别", "地区", "是否本人投保", "职业", "基本保额", "保障期限", "缴费类型", "缴费年限", "保费", "总保费"]


#Change age to date format
now_date = datetime.datetime.now()
tomorrow_date = now_date + datetime.timedelta(days = 1)
year = int(tomorrow_date.strftime("%Y"))
month_day_str = tomorrow_date.strftime("-%m-%d")

REAL_AGE_RANGE = list(range(17, 61, 1))
AGE_RANGE = ["{}{}".format((year - 18), now_date.strftime("-%m-%d"))] + list(map(lambda x: "{}{}".format((year - x), month_day_str), range(19, 62, 1)))
SEX_RANGE = ["男", "女"]
INSURANCE_AMOUNT = list(map(lambda x: str(x) + "万元", range(50, 160, 10)))
INSURANCE_DURATION = ["至60岁", "至65岁", "至70岁", "至80岁", "20年", "30年"]
PAY_TYPE = ["一次性", "年交"]
PAY_ALL = "趸交"
PAY_DURATION = ["5年", "10年", "20年", "30年"]

failed_case = []

#Make a HTTP request and return a parsed result
#return: (preminum, total prenium)
def requestWebData(url_params, session):
    try:
        #Replace single quote with double quote, remove redundant "+" to match server data format
        request_url = BASE_URL + urllib.parse.urlencode(url_params).replace("%27", "%22").replace("+", "")

        r2 = requests.get(request_url, headers = HEADERS, cookies = COOKIES).text

        matchSearch = re.search(r'.*?\(({.*?})\);', r2)

        if matchSearch:
            jsonResult = json.loads(matchSearch.group(1))

            return (jsonResult['data']['preminum'], jsonResult['data']['trialPrice']['totalPreminum'])
        else:
            print(r2)
    except Exception as e:
        traceback.print_exc()

    return (-1, -1)

#Initialize an Excel file
#return: A new sheet handler and Excel file handler
def initExcel(file_name):
    if os.path.exists(file_name):
        print("File " + file_name + " exists, delete it...")
        os.remove(file_name)

    excel_handler = xlsxwriter.Workbook(file_name)
    new_sheet = excel_handler.add_worksheet()

    col_num = len(EXCEL_COL_NAMES)
    new_sheet.set_column(col_num, col_num, EXCEL_COL_WIDTH)

    #insert a header
    for col_idx in range(0, col_num):
        new_sheet.write(0, col_idx, EXCEL_COL_NAMES[col_idx])

    return (new_sheet, excel_handler)

#Close Excel file
def closeExcel(file_handler):
    file_handler.close()

#Append a row to Excel file
def appendExcel(sheet_handler, row_num, row_data):
    for col_idx in range(0, len(row_data)):
        #Skip header row
        sheet_handler.write(row_num + 1, col_idx, row_data[col_idx])

def buildParam(new_data):
    #insurance amount
    PARAMS["restrictGeneParams"]['genes'][0]["value"] = new_data["amount"]

    #birth date
    PARAMS["restrictGeneParams"]['genes'][1]["value"] = new_data["insuranceDate"]
    PARAMS["restrictGeneParams"]['genes'][-2]["value"] = new_data["insuranceDate"]
    PARAMS["restrictGeneParams"]['optGeneOldValue']["value"] = new_data["insuranceDate"]

    #sex
    PARAMS["restrictGeneParams"]['genes'][2]["value"] = new_data["sex"]
    PARAMS["restrictGeneParams"]['genes'][-1]["value"] = new_data["sex"]

    #occupation
    PARAMS["restrictGeneParams"]['genes'][3]["value"] = new_data["insurantJob"]

    #city & province
    PARAMS["restrictGeneParams"]['genes'][4]["value"] = new_data["city"]
    PARAMS["restrictGeneParams"]['genes'][5]["value"] = new_data["province"]

    #insurance for self
    PARAMS["restrictGeneParams"]['genes'][-3]["value"] = new_data["insureForSelf"]

    #duration
    PARAMS["restrictGeneParams"]['genes'][6]["value"] = new_data["duration"]

    #pay_duration
    PARAMS["restrictGeneParams"]['genes'][-4]["value"] = new_data["pay_duration"]

    #pay_type
    PARAMS["restrictGeneParams"]['genes'][-5]["value"] = new_data["pay_type"]


def main(file_name):
    (sheet_handler, excel_handler) = initExcel(file_name)

    try:
        row_num = 0

        param_data = {
            "insureForSelf": "是",
            "province": "310000",
            "city": "310100",
            "insurantJob": "1-6类",
        }
        
        for age_idx in range(0, len(AGE_RANGE)):
            age = AGE_RANGE[age_idx]
            
            param_data['insuranceDate'] = age

            for sex in SEX_RANGE:
                param_data['sex'] = sex

                for amount in INSURANCE_AMOUNT:
                    param_data['amount'] = amount

                    for duration in INSURANCE_DURATION:
                        param_data['duration'] = duration

                        for pay_type in PAY_TYPE:
                            param_data['pay_type'] = pay_type

                            if pay_type == PAY_TYPE[0]:
                                pay_duration = PAY_ALL
                                param_data['pay_duration'] = PAY_ALL

                                buildParam(param_data)

                                (preminum, total_prenium) = requestWebData(PARAMS, None)

                                #record failed case
                                if preminum == -1:
                                    failed_case.append(dict(PARAMS))

                                #age, sex, district, self_insurance, occupation, amount, duration, pay_type, pay_duration, premium, total_premium
                                excel_data = [REAL_AGE_RANGE[age_idx], sex, "SHANGHAI", "是", "1-6类", amount, duration, pay_type, pay_duration, preminum, total_prenium]

                                appendExcel(sheet_handler, row_num, excel_data)

                                row_num += 1

                                print(row_num)
                            else:
                                for pay_duration in PAY_DURATION:
                                    param_data['pay_duration'] = pay_duration

                                    buildParam(param_data)

                                    (preminum, total_prenium) = requestWebData(PARAMS, None)

                                    #record failed case
                                    if preminum == -1:
                                        failed_case.append(dict(PARAMS))

                                    #age, sex, district, self_insurance, occupation, amount, duration, pay_type, pay_duration, premium, total_premium
                                    excel_data = [REAL_AGE_RANGE[age_idx], sex, "SHANGHAI", "是", "1-6类", amount, duration, pay_type, pay_duration, preminum, total_prenium]

                                    appendExcel(sheet_handler, row_num, excel_data)

                                    row_num += 1

                                    print(row_num)
                            
                            time.sleep(0.2)

    except Exception as e:
        traceback.print_exc()
    
    closeExcel(excel_handler)

    print("----Failed case----")
    for case in failed_case:
        print(case)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print "python grab_premium.py [output Excel file name]"
        sys.exit(1)
    
    main(sys.argv[1])