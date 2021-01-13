# coding=utf-8
from enum import Enum


# 项目路径
soundcore_path = "/Users/Anker/MyProject/Soundcore/iOS/SoundCore/SoundCore/Resource/"
anker_work_path = "/Users/Anker/MyProject/AnkerWork/AnkerWork/AnkerWork/Resource/"

anker_work_excel_path = "/Users/Anker/Downloads/多语言文案-Anker work App.xlsx"
soundcore_excel_path = "/Users/Anker/Downloads/开发使用_SoundCore文案汇总.xlsx"
# 选择需要导出文案的sheet
sheet_name = "开发新增词条"

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


# 其他列数枚举
class CommonCol(Enum):
    common = 0
    edit = 1
    os_type = 3
    develop_key = 4


# 手机系统枚举
class OSType(Enum):
    iOS = "iOS"
    android = "Android"


class ProjectName(Enum):
    soundcore = "Soundcore"
    anker_work = "AnkerWork"
