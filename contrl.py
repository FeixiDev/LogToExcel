import os
import sys
import action
import re


def fetch_file(suffix, files_path=r'\root\elastic', excel_path='.'):
    file_name_list = []
    file_path_list = []
    sheet_name_list = []
    for root, dirs, files in os.walk(files_path, topdown=False):
        for name in files:
            if os.path.splitext(name)[1] == suffix and '.mapping' not in os.path.splitext(name)[0]:
                file_name_list.append(name)
                file_path_list.append(os.path.join(root, name))
    file_path_list.sort()
    file_name_list.sort()
    for i in range(len(file_name_list)):
        sheet_name = re.findall('2.*?(?=.json)', (file_name_list[i]))
        sheet_name_list.append(sheet_name)
    excel_dict_list = []
    for file_path in file_path_list:
        excel_dict = action.read_json(file_path)
        excel_dict_list.append(excel_dict)
    excel_name = action.creat_excel(excel_path)
    for i in range(len(excel_dict_list)):
        action.to_excel(excel_dic=excel_dict_list[i],
                        excel_name=excel_name,
                        sheet_name=str(sheet_name_list[i][0]))


