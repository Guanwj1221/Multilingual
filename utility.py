# coding=utf-8
import re
from constant import OSType


# 打开文件，并读取
def open_files_read(all_file_path, output_dir_path):
    files = []
    for path in all_file_path:
        files.append(open(output_dir_path + path, "r+", 1024, "utf-8"))
    return files


# 打开文件，并写入
def open_files_write(all_file_path, output_dir_path):
    files = []
    for path in all_file_path:
        files.append(open(output_dir_path + path, "w+", 1024, "utf-8"))
    return files


# 打开文件，并添加
def open_files_add(all_file_path, output_dir_path):
    files = []
    for path in all_file_path:
        files.append(open(output_dir_path + path, "a+", 1024, "utf-8"))
    return files


def add_new_copy(add_head_annotation, os_type, all_file_path, output_dir_path, keys, values, flag):
    print("新增%d条文案" % len(values))
    print("新增中...")
    files = open_files_add(all_file_path, output_dir_path)
    for i in range(len(values)):
        tem_values = values[i]
        for index in range(len(tem_values)):
            file = files[index]
            if i == 0 and add_head_annotation == "True":
                file.write("\n//MARK: " + flag + "\n")
            value = tem_values[index]
            value = format_value(value)
            if value is None or value == "":
                continue

            # 处理一个文案多个key的情况
            tem_keys = keys[i].splitlines()
            for key in tem_keys:
                line = ""
                if os_type == OSType.android.value:
                    line = "<string name= \"" + key.strip() + "\" >" + format_value(value) + "</string>\n"
                elif os_type == OSType.iOS.value:
                    line = "\"" + key.strip() + "\" = \"" + format_value(value) + "\";\n"
                file.write(line)
    close_files(files)
    print("新增文案成功")


# 更新旧文案
def update_old_copy(os_type, all_file_path, output_dir_path, keys, values):
    print("修改%d条文案" % len(values))
    print("修改中...")
    contents = []
    files = open_files_read(all_file_path, output_dir_path)
    for i in range(len(files)):
        content = ""
        # 所有需要修改的的key组成的一维数组
        # 后面需要修改它的值，所以需要重新赋值
        all_keys = keys.copy()
        # 所有需要修改的key和对应的values组成的二位数组
        keys_values = values.copy()
        # 获取当前语言文件的每一行内容
        for line in files[i]:
            # 遍历key的一维数组
            for index in range(len(all_keys)):
                # 获取一维数组的key
                key = all_keys[index]
                # 获取key对应的values数组
                key_values = keys_values[index]
                # 将key与读取的一行内容进行匹配，如果内容中包含key则表示为需要修改的行
                if re.search("\"%s\"" % key, line):
                    # 取出当前语言的value
                    value = key_values[i]
                    # 修改行内容
                    # <string name="xxx">details</string>
                    key = key.strip()
                    value = format_value(value)
                    if key is None or key == "" or value is None or value == "":
                        break

                    if os_type == OSType.android.value:
                        line = "<string name= \"" + key + "\" >" + value + "</string>\n"
                    else:
                        line = "\"" + key + "\" = \"" + value + "\";\n"
                    # 删除已经匹配过的key,减少比较次数
                    del(keys_values[index])
                    del(all_keys[index])
                    # 不再进行匹配
                    break
            content += line
        contents.append(content)
    close_files(files)

    files = open_files_write(all_file_path, output_dir_path)
    for i in range(len(files)):
        files[i].write(contents[i])
    close_files(files)
    print("修改文案成功")


# 删除久文案
def delete_old_copy(all_file_path, output_dir_path, keys):
    print("删除%d条文案" % len(keys))
    print("删除中...")
    contents = []
    files = open_files_read(all_file_path, output_dir_path)

    for i in range(len(files)):
        content = ""
        # 所有需要删除的的key组成的一维数组
        # 后面需要删除它的值，所以需要重新赋值
        all_keys = keys.copy()
        # 获取当前语言文件的每一行内容
        for line in files[i]:
            # 遍历key的一维数组
            for index in range(len(all_keys)):
                # 获取一维数组的key
                key = all_keys[index]
                # 将key与读取的一行内容进行匹配，如果内容中包含key则表示为需要修改的行
                if re.search(key, line):
                    # 删除已经匹配过的key,减少比较次数
                    all_keys.remove(key)
                    line = "delete"

                    # 不再进行匹配
                    break
            if line != "delete":
                content += line
        contents.append(content)
    close_files(files)

    files = open_files_write(all_file_path, output_dir_path)
    for i in range(len(files)):
        files[i].write(contents[i])
    close_files(files)
    print("删除文案成功")


# 关闭打开的文件
def close_files(files):
    for i in range(len(files)):
        file = files[i]
        file.close()


# 去掉变量后面的说明
def format_variants(original_str):
    match_list = re.findall(r'%\d\$\w{[^{}]*}', original_str, re.S)
    for item in match_list:
        item_we_want = item[0:4]
        original_str = original_str.replace(item, item_we_want)
    return original_str


def format_escape_str(r, original_str):
    match_list = re.findall(r, original_str, re.S)
    for item in match_list:
        item_we_want = ""
        if len(item) == 1:
            item_we_want = "\\" + item[0]
        elif len(item) == 2:
            item_we_want = item[0] + "\\" + item[1]
        original_str = original_str.replace(item, item_we_want)
    return original_str


def format_percent_str(original_str):
    # 查找所有百分号%在字符串中的索引
    percent_str_index_list = []
    for index in range(len(original_str)):
        if original_str[index] == '%':
            percent_str_index_list.append(index)

    # 去除变量的百分号索引
    match_list = re.findall(r'%\d\$\w', original_str, re.S)
    if len(match_list) == 0:
        return original_str

    # 在list中移除变量的百份号的索引
    for item in match_list:
        temp = 0
        for index in range(len(original_str)):
            temp = original_str.find(item, temp, len(original_str))
            if temp != -1:
                if temp in percent_str_index_list:
                    percent_str_index_list.remove(temp)
                temp = temp + 1
            else:
                break

    # 查看找到的百分号是否已经格式化
    replace_index = []
    for index in percent_str_index_list:
        # 字符串长度为1的情况
        if len(original_str) == 1:
            replace_index.append(index)
            continue

        # 字符串长度 >=2 时
        # 第一个字符是百分号
        if index == 0:
            if original_str[0:2] != "%%":
                replace_index.append(index)

        elif index == len(original_str) - 1:
            # 最后一个字符是百分号
            if original_str[len(original_str) - 2:] != "%%":
                replace_index.append(index)

        else:
            # 字符串中间找到百分号，且此时字符串长度 >= 3
            if original_str[index - 1: index + 1] == "%%" or original_str[index: index + 2] == "%%":
                pass
            else:
                replace_index.append(index)

    if len(replace_index) == 0:
        print("优化百分号之后:" + original_str)
        return original_str

    last_index = -1
    result = ""
    for index in replace_index:
        result = result + original_str[last_index + 1:index + 1] + '%'
        last_index = index
    result = result + original_str[last_index + 1:]
    print("优化百分号之后:" + result)
    return result


def format_value(original_str):
    if original_str is None or original_str == "":
        return
    original_str = format_variants(original_str)
    # 转义单引号 '
    original_str = format_escape_str(r'[^\\]\'|^\'', original_str)
    # 转义双引号 "
    original_str = format_escape_str(r'[^\\]\"|^\"', original_str)
    # 转义百分号 %
    original_str = format_percent_str(original_str)

    # excel中的换行符换成 \n转义字符, 文案首尾的换行符认为是翻译人员不小心按出来的，在此去掉
    original_str = re.sub(r'\n+$', "", original_str, re.S)
    original_str = re.sub(r'^\n+', "", original_str, re.S)
    original_str = re.sub(r'\n', "\\\\n", original_str, re.S)
    return original_str



