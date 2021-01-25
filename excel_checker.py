# -*- coding: UTF-8 -*-

import argparse
import re
import sys
import win32com.client

from xlrd import *

VERSION="1.1"

class Data():
    def __init__(self, d02, d06, soft_ver, ck, mat_ver, norm_n2):
        self.d02 = d02 or ""
        self.d06 = d06 or ""
        self.soft_ver = soft_ver or ""
        self.ck = ck or ""
        self.mat_ver = mat_ver or ""
        self.norm_n2 = norm_n2 or ""
    def toString(self):
        return "{}, {}, {}, {}, {}".format(self.d02, self.d06, self.soft_ver, self.ck, self.mat_ver, self.norm_n2)

def get_args():
    parser = argparse.ArgumentParser(description='Tool Version: {}'.format(VERSION))
    parser.add_argument('-r', '--request-excel', required=True, help='Input request excel path')
    parser.add_argument('-rs', '--request-excel-sheet', default=1, help='Input request excel sheet index (begin from 1)')
    parser.add_argument('-rp', '--request-excel-password', default=123, help='Input request excel password')
    parser.add_argument('-d', '--data-excel', required=True, help='Input data excel path')
    parser.add_argument('-ds', '--data-excel-sheet', default=1, help='Input data excel sheet index (begin from 1)')
    parser.add_argument('-dp', '--data-excel-password', default=123, help='Input data excel password')
    parser.add_argument('-l', '--line-number', required=True, help='Check request line number')
    return parser.parse_args()

def get_excel(filename, sheet_index, password=""):
    xlapp = win32com.client.Dispatch("Excel.Application")
    xlwb = xlapp.Workbooks.Open(filename, False, True, None, Password=password)
    return xlwb ,xlwb.Sheets(sheet_index)

def main():
    args = get_args()
    request_xlwb, request_excel = get_excel(args.request_excel, args.request_excel_sheet, args.request_excel_password)
    data_xlwb, data_excel = get_excel(args.data_excel, args.data_excel_sheet, args.request_excel_password)
    request_row = list(request_excel.UsedRange.Rows(args.line_number).value[0])
    check_dict = {}
    strip_f = lambda x: x.strip()
    # 異動內容
    for item in request_row[14].splitlines():
        new_value = ""
        if re.search(r'>', item):
            new_value = item.split(">")[-1].strip()
        elif not re.findall(r'[\u4e00-\u9fff]+', item):
            new_value = item
        else:
            continue
        
        item_dict = {"mat_ver": "", "soft_ver": "", "norm_n2": []}
        if re.search(r"\.", new_value):
            param, mat_ver = new_value.split(".")
            item_dict["mat_ver"] = mat_ver
        elif re.search(r"\(|\)", new_value):
            param, soft_ver = map(strip_f, list(filter(None, re.split(r"\(|\)", new_value))))
            item_dict["soft_ver"] = soft_ver
        else:
            param = new_value
        check_dict[param] = check_dict.get(param, {})
        check_dict[param].update(item_dict)
    # 新參數
    param_ver = "{}-"
    for item in request_row[13].splitlines():
        if re.search(r"^n2:", item, re.I):
            param_ver = "-{}"
            continue
        elif re.search(r":", item, re.I):
            param_ver = "{}-"
            continue
        elif re.findall(r'[\u4e00-\u9fff]+', item):
            continue
        norm_n2, soft_ver = "", ""
        if re.search(r"_", item):
            param, norm_n2 = item.split("_")
        elif re.search(r"\(|\)", item):
            param, soft_ver = map(strip_f, list(filter(None, re.split(r"\(|\)", item))))
        else:
            param = item
        check_dict[param] = check_dict.get(param, {})
        if not check_dict[param]:
            check_dict[param].update({"mat_ver": "", "soft_ver": "", "norm_n2": []})
        if norm_n2:
            check_dict[param]["norm_n2"].append(param_ver.format(norm_n2))
        if soft_ver:
            check_dict[param]["soft_ver"] = soft_ver

    d02_data = [[] for _ in range(10)]
    d06_data = [[] for _ in range(10)]
    ck_data = [[] for _ in range(100)]
    for i, row in enumerate(data_excel.UsedRange.Rows):
        if i < 3:
            continue
        row = list(row.value[0])
        data = Data(row[6], row[7], row[8], row[11], row[12], row[14])
        if re.search(r"^d02", data.d02, re.I):
            num = int(data.d02.split("-")[1][1])
            d02_data[num].append(data)
        if re.search(r"^d06",data.d06, re.I):
            num = int(data.d06.split("-")[1][1])
            d06_data[num].append(data)
        if re.search(r"^cn",data.ck, re.I):
            num = int(data.ck[3:5])
            ck_data[num].append(data)

    pass_str = "[PASS] Check {}"
    fail_str = "[FAIL] Can not find matched {} {}"
    print(check_dict)
    for key, val_dict in check_dict.items():
        pass_mat_ver_check, had_pass_mat_ver_check = False, False
        pass_soft_ver_check, had_pass_soft_ver_check = False, False
        pass_norm_n2_check, had_pass_norm_n2_check = False, False
        if re.search(r"^cn", key, re.I):
            data_list = ck_data[int(key[3:5])]
        elif re.search(r"^d02", key, re.I):
            data_list = d02_data[int(key.split("-")[1][1])]
        else:
            data_list = d06_data[int(key.split("-")[1][1])]
        for data in data_list:
            pass_mat_ver_check, pass_norm_n2_check = False, False
            check_count = 0
            if key in [data.d02, data.d06, data.ck]:
                if val_dict.get("soft_ver", "") in (data.soft_ver, ""):
                    pass_soft_ver_check = had_pass_soft_ver_check = True
                if val_dict.get("mat_ver", "") in (data.mat_ver, ""):
                    pass_mat_ver_check = had_pass_mat_ver_check = True
                for norm_val in val_dict["norm_n2"]:
                    if norm_val in data.norm_n2:
                        check_count += 1
                if check_count >= len(val_dict["norm_n2"]):
                    pass_norm_n2_check = had_pass_norm_n2_check = True
                if pass_mat_ver_check and pass_norm_n2_check and pass_soft_ver_check:
                    print(pass_str.format(key))
                    break
        else:
            if not (pass_soft_ver_check or had_pass_soft_ver_check):
                print(fail_str.format("軟件版本", key))
            if not (pass_mat_ver_check or had_pass_mat_ver_check):
                print(fail_str.format("料件版本", key))
            if not (pass_norm_n2_check or had_pass_norm_n2_check):
                print(fail_str.format("一般成測-N2", key))
    #request_xlwb.Close()
    #data_xlwb.Close()

if __name__ == '__main__':
    main()