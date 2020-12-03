# coding=utf-8

# pip install xlrd
import xlrd
import tkinter
from tkinter import *
from tkinter import ttk, messagebox, Button
from enum import Enum
import re


# 可能需要修改的点：
# 1.soundcore_path/anker_work_path: 项目路径
# 2.sheet_name

# 文案Excel路径
global source_path
# 项目路径
soundcore_path = "/Users/Anker/MyProject/Soundcore/iOS/SoundCore/SoundCore/Resource/"
anker_work_path = "/Users/Anker/MyProject/AnkerWork/AnkerWork/AnkerWork/Resource/"

anker_work_excel_path = "/Users/Anker/Downloads/多语言文案-Anker work App.xlsx"
soundcore_excel_path = "/Users/Anker/Downloads/开发使用_SoundCore文案汇总.xlsx"
# 选择需要导出文案的sheet
sheet_name = "德语修改"
# 选择导出文案类型 "iOS","Android"
global os_type
# 是否添加头部注释
global add_head_annotation
# 项目名称 "Anker Work", "Soundcore"
global project_name
# 文案导出的路径（项目文案路径）
global output_path
# 文件路径
global outFilePath

ios_file_path = ["Base.lproj/Localizable.strings",
                 "zh-Hans.lproj/Localizable.strings",
                 "zh-Hant.lproj/Localizable.strings",
                 "zh-HK.lproj/Localizable.strings",
                 "en.lproj/Localizable.strings",
                 "de.lproj/Localizable.strings",
                 "ja.lproj/Localizable.strings",
                 "es.lproj/Localizable.strings",
                 "pt-PT.lproj/Localizable.strings",
                 "it.lproj/Localizable.strings",
                 "fr.lproj/Localizable.strings",
                 "ko.lproj/Localizable.strings",
                 "ru.lproj/Localizable.strings",
                 "ar.lproj/Localizable.strings"]

android_file_path = ["zh", "zh_rTW", "en", "de",
                     "ja", "es", "pt", "it",
                     "fr", "ko", "ru", "ar"]


# 语言列数枚举
class LanguageCol(Enum):
    en = 6
    zh = 7
    zh_tc = 8
    de = 9
    ja = 10
    fr = 11
    it = 12
    es = 13
    ru = 14
    ar = 15
    pt = 16
    ko = 17


# 手机系统枚举
class OSType(Enum):
    iOS = "iOS"
    android = "Android"


# 其他列数枚举
class CommonCol(Enum):
    common = 0
    edit = 1
    os_type = 3
    develop_key = 4


class Application(Frame):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)

        self.master = master
        master.title("Anker Work")

        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        size = '%dx%d+%d+%d' % (320, 240, (screenwidth - 320) / 2, (screenheight - 240) / 2)
        root.geometry(size)

        self.pack()
        self.create_widgets()

    def create_widgets(self):
        name_label = Label(master=self.master, text='--请输入文案的版本号或者产品号--')
        name_label.pack()
        self.name_input = Entry(master=self.master)
        self.name_input.pack()

        os_label = Label(master=self.master, text='--请选择手机系统--')
        os_label.pack()
        self.os_type_value = tkinter.StringVar()
        self.os_type_box_list = ttk.Combobox(master=self.master, textvariable=self.os_type_value)
        self.os_type_name = [OSType.iOS.value, OSType.android.value]
        self.os_type_box_list['value'] = self.os_type_name
        self.os_type_box_list.current(0)
        self.os_type_box_list.pack()

        head_label = Label(master=self.master, text='--是否添加头部注释--')
        head_label.pack()
        self.head_annotation_value = tkinter.StringVar()
        self.head_annotation_box_list = ttk.Combobox(master=self.master, textvariable=self.head_annotation_value)
        self.head_annotation_name = ["False", "True"]
        self.head_annotation_box_list['value'] = self.head_annotation_name
        self.head_annotation_box_list.current(0)
        self.head_annotation_box_list.pack()

        project_label = Label(master=self.master, text='--请选择需要添加的项目--')
        project_label.pack()
        self.project_value = tkinter.StringVar()
        self.project_box_list = ttk.Combobox(master=self.master, textvariable=self.project_value)
        self.project_name = ["Anker Work", "Soundcore"]
        self.project_box_list['value'] = self.project_name
        self.project_box_list.current(0)
        self.project_box_list.pack()

        alert_button = Button(master=self.master, text='确认', command=self.button_click)
        alert_button.pack()

    def button_click(self):
        global add_head_annotation
        global os_type
        global project_name
        global output_path
        global outFilePath
        global source_path
        add_head_annotation = self.head_annotation_box_list.get() or "False"
        os_type = self.os_type_box_list.get() or OSType.iOS.value
        name = self.name_input.get()
        project_name = self.project_box_list.get() or "Anker Work"
        if project_name == "Anker Work":
            output_path = anker_work_path
            source_path = anker_work_excel_path
        else:
            output_path = soundcore_path
            source_path = soundcore_excel_path
        if os_type == OSType.iOS.value:
            outFilePath = ios_file_path
        else:
            outFilePath = android_file_path
        start(name)


# 打开文件，并读取
def open_files_read():
    files = []
    for path in outFilePath:
        files.append(open(output_path + path, "r+", 1024, "utf-8"))
    return files


# 打开文件，并写入
def open_files_write():
    files = []
    for path in outFilePath:
        files.append(open(output_path + path, "w+", 1024, "utf-8"))
    return files


# 打开文件，并添加
def open_files_add():
    files = []
    for path in outFilePath:
        files.append(open(output_path + path, "a+", 1024, "utf-8"))
    return files


def add_new_documents(keys, values, flag):
    print("新增%d条文案" % len(values))
    print("新增中...")
    files = open_files_add()
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
def update_old_documents(keys, values):
    print("修改%d条文案" % len(values))
    print("修改中...")
    contents = []
    files = open_files_read()
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

    files = open_files_write()
    for i in range(len(files)):
        files[i].write(contents[i])
    close_files(files)
    print("修改文案成功")


# 删除久文案
def delete_old_document(keys):
    print("删除%d条文案" % len(keys))
    print("删除中...")
    contents = []
    files = open_files_read()

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

    files = open_files_write()
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


# 开始修改文案
def start(flag):
    add_keys = []
    add_values = []
    update_keys = []
    update_values = []
    delete_keys = []
    if flag is None or flag == "":
        messagebox.showinfo('Message', '输入内容为空')
        return
    workbook = xlrd.open_workbook(source_path)
    table = workbook.sheet_by_name(sheet_name)
    rows = table.nrows
    # 导出文案
    ready = 0
    end = 0
    for i in range(0, rows):
        comment = table.cell(i, CommonCol.common.value).value
        if re.compile(r'(.*%s.*start)' % flag).match(comment):
            ready = 1
        elif re.compile(r'(.*%s.*end)' % flag).match(comment):
            end = 1
        if ready != 1:
            continue
        if end == 1:
            break

        operateType = table.cell(i, CommonCol.edit.value).value
        typeColValue = table.cell(i, CommonCol.os_type.value).value
        keyColValue = table.cell(i, CommonCol.develop_key.value).value

        if keyColValue is None or keyColValue == "":
            continue
        if typeColValue.find(os_type) <= 0:
            continue

        # 取出各个文案的语言
        zhHkColValue = table.cell(i, LanguageCol.zh.value).value
        tcHkColValue = table.cell(i, LanguageCol.zh_tc.value).value
        hkColValue = tcHkColValue
        enColValue = table.cell(i, LanguageCol.en.value).value
        baseColValue = enColValue
        deColValue = table.cell(i, LanguageCol.de.value).value
        jaColValue = table.cell(i, LanguageCol.ja.value).value
        esColValue = table.cell(i, LanguageCol.es.value).value
        ptColValue = table.cell(i, LanguageCol.pt.value).value
        itColValue = table.cell(i, LanguageCol.it.value).value
        frColValue = table.cell(i, LanguageCol.fr.value).value
        koColValue = table.cell(i, LanguageCol.ko.value).value
        ruColValue = table.cell(i, LanguageCol.ru.value).value
        arColValue = table.cell(i, LanguageCol.ar.value).value

        if os_type == OSType.iOS.value:
            col_values = [baseColValue, zhHkColValue, tcHkColValue,
                          hkColValue, enColValue, deColValue,
                          jaColValue, esColValue, ptColValue,
                          itColValue, frColValue, koColValue,
                          ruColValue, arColValue]
        else:
            col_values = [zhHkColValue, tcHkColValue, enColValue, deColValue,
                          jaColValue, esColValue, ptColValue, itColValue,
                          frColValue, koColValue, ruColValue, arColValue]

        if operateType.find("新增") >= 0:
            add_values.append(col_values)
            add_keys.append(keyColValue)

        elif operateType.find("修改") >= 0:
            update_values.append(col_values)
            update_keys.append(keyColValue)

        elif operateType.find("删除") >= 0:
            delete_keys.append(keyColValue)

    if len(update_values) > 0:
        update_old_documents(update_keys, update_values)

    if len(add_values) > 0:
        add_new_documents(add_keys, add_values, flag + "-start")

    if len(delete_keys) > 0:
        delete_old_document(delete_keys)


if __name__ == '__main__':
    root = tkinter.Tk()
    app = Application(master=root)
    app.mainloop()
