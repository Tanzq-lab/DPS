import os
import win32api
import win32con
from typing import Dict

from Flyweight.InfoFactory import InfoFactory


def GetProgramPath(extname):
    # 在win 注册表中获取执行该文件的程序地址

    # print(extname)
    try:
        key = win32api.RegOpenKey(win32con.HKEY_CLASSES_ROOT, extname, 0, win32con.KEY_QUERY_VALUE)
    except:
        print("打开{extName}子键失败吖".format(extName=extname))
        return 'Error'

    # print(key)

    try:
        path, typ = win32api.RegQueryValueEx(key, "")
    except:
        print("查询不到对应的值喔!")
        win32api.RegCloseKey(key)
        return 'Error'

    win32api.RegCloseKey(key)
    return path


def GetProgramAbsolutePath(extname):
    # .txt .doc .docx .xlsx
    type_path = GetProgramPath(extname)
    prog_regedit_path = type_path + r'\shell\Printto\command'
    return GetProgramPath(prog_regedit_path).split(' /')[0]


def Openfile(filename, ProgPathList: list):
    # 获取到执行该文件的程序地址
    i = 0
    for key in ProgPathList:
        print(i.__str__() + ":" + key)
        i += 1

    index = int(input("\n请选择您要执行程序的序号："))
    while not 0 <= index < i:
        print("抱歉喔，您输入有误，请重新输入！")
        index = int(input("请输入您想操作的格式的序号："))

    prog_path = ProgPathList[index]
    # print(prog_path)
    if prog_path == "Error":
        return "Error"

    # 调用 os 执行程序
    print('程序正在运行...')
    r = os.system("start /wait  \"\" {exePath} {dir}".format(exePath=prog_path, dir=filename))
    print('程序已经关闭！\n --- \n')
    return "OK"


def init():
    FileTypeDict: Dict[str, list] = {}
    typeList = [".xlsx", ".txt", ".docx"]

    for x in typeList:
        FileTypeDict[x] = [GetProgramAbsolutePath(x)]

    return FileTypeDict


def menu():
    print("------ 欢迎使用数据收集系统 ------")
    print("|     1. 查询当前支持的格式      |")
    print("|     2. 查询格式对应的程序      |")
    print("|     3. 对数据进行读写操作      |")
    print("|     0. 退出系统                |")
    print(" ________________________________ ")
    opt = int(input("请输入您的操作: "))

    if 0 <= opt <= 3:
        return opt

    print("操作有误! \n --- \n")
    return menu()


# 展示当前可支持的格式
def show(FileTypeDict: Dict[str, list]):
    print("当前系统可支持？！")
    for key in FileTypeDict:
        print(key, end=" ")

    print("\n --- \n")


# 查询文件格式对应的程序地址
def query(FileTypeDict: Dict[str, list]):
    show(FileTypeDict)
    filetype = input("请输入您要查询的格式(´▽`ʃ♡ƪ)：")

    ProgList = FileTypeDict.get(filetype)
    if not ProgList:
        print("抱歉哦~该格式不支持！！")
    else:
        print(filetype + " 支持的程序有：")
        for x in ProgList:
            print(x)

        print("使用的是当前系统注册表中默认的程序地址，其他地址请在源码中配置o(*￣︶￣*)o")

    print("\n --- \n")


def dataProcessing(infoFactory: InfoFactory, FileTypeDict: Dict[str, list]):
    # 步骤一 ：获取用户文件想操作的文件后缀名
    print("当前系统支持以下文件格式：")
    FileTypeList = []
    i = 0
    for key in FileTypeDict:
        FileTypeList.append(key)
        print(i.__str__() + ":" + key)
        i += 1

    index = int(input("\n请输入您想操作的格式的序号："))
    while not 0 <= index < i:
        print("抱歉喔，您输入有误，请重新输入！")
        index = int(input("请输入您想操作的格式的序号："))

    # 步骤二 ： 生成可读文件,并得到文件名
    name = input("\n请输入您的名称：")
    filePullPath = infoFactory.Builder(name, FileTypeList[index])

    # 步骤三 ： 用软件打开当前文件
    ProgPathList = []
    for value in FileTypeDict.get(FileTypeList[index]):
        ProgPathList.append(value)

    if Openfile(filePullPath, ProgPathList) == 'Error':
        print("哎呀，程序出现了错误！请联系工作人员！")

    # 步骤四 ： 对数据文件进行序列化，并删除刚打开的文件
    infoFactory.serialization(filePullPath)

