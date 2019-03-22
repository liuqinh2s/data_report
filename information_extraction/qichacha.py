# coding=utf-8

# 此程序输出来源于企查查的信息，包括：身份信息，股东信息，变更记录信息
# 以json格式输出

import docx
import re
import yaml
import os

from word_manipulation import docx_enhanced

current_path = os.path.dirname(os.path.realpath(__file__))
f = open(current_path+"\\qichacha_config.yml", encoding="utf-8")
config = yaml.load(f, Loader=yaml.FullLoader)
f.close()


def dict2list(dict):
    result = []
    for i in dict:
        result.append(dict[i])
    return result


def recurse_dict2list(ll):
    if isinstance(ll, list):
        for i in range(len(ll)):
            ll[i] = recurse_dict2list(ll[i])
    elif isinstance(ll, dict):
        ll = recurse_dict2list(dict2list(ll))
    return ll


# title表示表头内容，content表示需要删选的列
def get_table_content(docx_file, title, content):
    result = []
    docx_list = docx_enhanced.docx_to_list(docx_file)
    for i in range(1, len(docx_list)):
        if isinstance(docx_list[i-1], str) and isinstance(docx_list[i], list) and config[title] in docx_list[i-1]:
            for j in range(1, len(docx_list[i])):
                temp = {}
                for k in range(0, len(docx_list[i][j])):
                    if docx_list[i][0][k] in config[content]:
                        temp[docx_list[i][0][k]] = docx_list[i][j][k]
                result.append(temp)
    return result


# 提取规则
def read_docx(file_name):
    docx_file = docx.Document(file_name)
    paragraphs_content = '\n'.join([para.text for para in docx_file.paragraphs])
    # 身份信息
    BUSINESS_INFO = {}
    for i in config['identity_info']:
        BUSINESS_INFO[i] = re.search("%s：(.*?)\n" % i, paragraphs_content).group(1).strip()
    INFO_DICT = {}
    for info in config['info_list']:
        INFO_DICT[info[1]] = get_table_content(docx_file, info[0], info[1])
    BUSINESS_INFO = dict2list(BUSINESS_INFO)
    INFO_LIST = recurse_dict2list(INFO_DICT)
    return BUSINESS_INFO, INFO_LIST