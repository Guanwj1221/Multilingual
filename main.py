# coding=utf-8

# pip install xlrd
import xlrd
import tkinter
from tkinter import *
from tkinter import ttk, messagebox
from enum import Enum
import re
import utility
import constant
from constant import OSType, CommonCol, LanguageCol, ProjectName

# 可能需要修改的点：
# 1.soundcore_path/anker_work_path: 项目路径
# 2.sheet_name

# 文案Excel路径
global source_path
# 选择导出文案类型 "iOS","Android"
global os_type
# 是否添加头部注释
global add_head_annotation
# 项目名称
global project_name
# 文案导出的路径（项目文案路径）
global output_dir_path
# 文件路径
global all_file_path


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
        head_annotation_name = ["False", "True"]
        self.head_annotation_box_list['value'] = head_annotation_name
        self.head_annotation_box_list.current(0)
        self.head_annotation_box_list.pack()

        project_label = Label(master=self.master, text='--请选择需要添加的项目--')
        project_label.pack()
        self.project_value = tkinter.StringVar()
        self.project_box_list = ttk.Combobox(master=self.master, textvariable=self.project_value)
        all_project = []
        for name, member in ProjectName.__members__.items():
            all_project.append(member.value)
        self.app_name = all_project
        self.project_box_list['value'] = self.app_name
        self.project_box_list.current(0)
        self.project_box_list.pack()

        self.alert_button = ttk.Button(master=self.master, text='确认', command=self.button_click)
        self.alert_button.pack()

    def button_click(self):
        global add_head_annotation
        global os_type
        global project_name
        global output_dir_path
        global all_file_path
        global source_path
        add_head_annotation = self.head_annotation_box_list.get() or "False"
        os_type = self.os_type_box_list.get() or OSType.iOS.value
        name = self.name_input.get()
        project_name = self.project_box_list.get() or ProjectName.soundcore.value
        if project_name == ProjectName.anker_work.value:
            output_dir_path = constant.anker_work_path
            source_path = constant.anker_work_excel_path
        else:
            output_dir_path = constant.soundcore_path
            source_path = constant.soundcore_excel_path
        if os_type == OSType.iOS.value:
            all_file_path = constant.ios_file_path
        else:
            all_file_path = constant.android_file_path
        start(name)


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
    table = workbook.sheet_by_name(constant.sheet_name)
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

        operate_type = table.cell(i, CommonCol.edit.value).value
        type_col_value = table.cell(i, CommonCol.os_type.value).value
        key_col_value = table.cell(i, CommonCol.develop_key.value).value

        if key_col_value is None or key_col_value == "":
            continue
        if type_col_value.find(os_type) <= 0:
            continue

        # 取出各个文案的语言
        zh_hk_col_value = table.cell(i, LanguageCol.zh.value).value
        tc_hk_col_value = table.cell(i, LanguageCol.zh_tc.value).value
        hk_col_value = tc_hk_col_value
        en_col_value = table.cell(i, LanguageCol.en.value).value
        base_col_value = en_col_value
        de_col_value = table.cell(i, LanguageCol.de.value).value
        ja_col_value = table.cell(i, LanguageCol.ja.value).value
        es_col_value = table.cell(i, LanguageCol.es.value).value
        pt_col_value = table.cell(i, LanguageCol.pt.value).value
        it_col_value = table.cell(i, LanguageCol.it.value).value
        fr_col_value = table.cell(i, LanguageCol.fr.value).value
        ko_col_value = table.cell(i, LanguageCol.ko.value).value
        ru_col_value = table.cell(i, LanguageCol.ru.value).value
        ar_col_value = table.cell(i, LanguageCol.ar.value).value

        if os_type == OSType.iOS.value:
            col_values = [base_col_value, zh_hk_col_value, tc_hk_col_value,
                          hk_col_value, en_col_value, de_col_value,
                          ja_col_value, es_col_value, pt_col_value,
                          it_col_value, fr_col_value, ko_col_value,
                          ru_col_value, ar_col_value]
        else:
            col_values = [zh_hk_col_value, tc_hk_col_value, en_col_value, de_col_value,
                          ja_col_value, es_col_value, pt_col_value, it_col_value,
                          fr_col_value, ko_col_value, ru_col_value, ar_col_value]

        if operate_type.find("新增") >= 0:
            add_values.append(col_values)
            add_keys.append(key_col_value)

        elif operate_type.find("修改") >= 0:
            update_values.append(col_values)
            update_keys.append(key_col_value)

        elif operate_type.find("删除") >= 0:
            delete_keys.append(key_col_value)

    if len(update_values) > 0:
        utility.update_old_copy(os_type, all_file_path,
                                output_dir_path, update_keys,
                                update_values)

    if len(add_values) > 0:
        flag = flag + "-start"
        utility.add_new_copy(add_head_annotation, os_type,
                             all_file_path, output_dir_path,
                             add_keys, add_values, flag)

    if len(delete_keys) > 0:
        utility.delete_old_copy(all_file_path, output_dir_path,
                                delete_keys)


if __name__ == '__main__':
    root = tkinter.Tk()
    app = Application(master=root)
    app.mainloop()
