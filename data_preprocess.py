# -*- utf-8 -*-

import pandas as pd
from openpyxl import load_workbook
from xlwt import Workbook


def create_sheet(sheet_name, savename, tran_dict, ori_data):
    wb = Workbook()
    data_sheet = wb.add_sheet(sheet_name)
    data_sheet.write(0, 0, "subject")
    data_sheet.write(0, 1, "relation")
    data_sheet.write(0, 2, "object")
    for i in range(len(ori_data[ori_data.columns[0]])):
        data_sheet.write(i + 1, 0, tran_dict[ori_data[ori_data.columns[0]][i]])
        data_sheet.write(i + 1, 1, ori_data[ori_data.columns[1]][i])
        data_sheet.write(i + 1, 2, tran_dict[ori_data[ori_data.columns[2]][i]])
    wb.save(savename)
    print("sheet", sheet_name, "created")


def read_data(filename, sheet_name="node"):
    print("read data ->", filename, sheet_name)
    return pd.read_excel(filename, sheet_name=sheet_name)


def build_dictionary(data, cn_label, en_label):
    cn_names, en_names = data[cn_label], data[en_label]
    print("length of names ->", len(cn_names), len(en_names))
    dictionary = {}
    for i in range(len(cn_names)):
        dictionary[cn_names[i]] = en_names[i]
    return dictionary


if __name__ == "__main__":
    filename = "./data/exemplar_knowledge_graph_v1.xlsx"
    data = read_data(filename)
    cn_label = "ChineseName"
    en_label = "EnglishName"
    tran_dict = build_dictionary(data, cn_label, en_label)
    savename = "./data/en_relation.xls"
    ori_data = read_data(filename, sheet_name="relation")
    create_sheet("en_relation", savename, tran_dict, ori_data)
