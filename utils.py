import os
import re
import sys

from win32com.shell import shell, shellcon
import win32api, win32con
import win32com

def GetRes(path):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")

    return os.path.join(base, path)

def insertData(data, localizedResourceName: str|None = None, IconResource: str|None = None, IconIndex: int = 0, infoTip: str|None = None) -> bool:
    # TODO: 原本没有的属性不会添加
    for title, datas in data:
        if title == ".ShellClassInfo":
            heads = {data[0]: i for i, data in enumerate(datas)}
            if localizedResourceName:
                if "LocalizedResourceName" not in heads: datas.append(("LocalizedResourceName", localizedResourceName))
                else: datas[heads["LocalizedResourceName"]] = ("LocalizedResourceName", localizedResourceName)
            if IconResource:
                if "IconResource" not in heads: datas.append(("IconResource", f"{IconResource},{IconIndex}"))
                else: datas[heads["IconResource"]] = ("IconResource", f"{IconResource},{IconIndex}")
            if infoTip:
                if "InfoTip" not in heads: datas.append(("InfoTip", infoTip))
                else: datas[heads["InfoTip"]] = ("InfoTip", infoTip)
            return True
    return False

def loadConfig(dir: str, reset_file_attribute: bool = True):
    data = list()
    if os.path.exists(os.path.join(dir, "desktop.ini")):
        with open(os.path.join(dir, "desktop.ini"), "r", encoding="gbk") as f:
            for line in f.readlines():
                if r := re.match(r"^\[(.+)\]$", line):
                    #print(">>> [A]")
                    #print(r.group(1))
                    data.append((r.group(1), list()))
                    continue
                elif r := re.match(r"^(.+)=(.*?)$", line):
                    #print(">>> A=B")
                    #print(r.group(1), r.group(2))
                    data[-1][1].append((r.group(1), r.group(2)))
                    continue
                else:
                    #print(">>> other")
                    #print(line)
                    raise Exception("Invalid desktop.ini file")

        # 设定文件属性，防止后续无法覆盖
        if (reset_file_attribute):
            win32api.SetFileAttributes(os.path.join(dir, "desktop.ini"), 0)
        
    return data

def SaveConfig(dir: str, localizedResourceName: str|None = None, iconResource: str|None = None, iconIndex: int = 0, infoTip: str|None = None):
    if localizedResourceName is None and iconResource is None:
        return
    
    data = loadConfig(dir)

    if not insertData(data, localizedResourceName, iconResource, iconIndex):
        data.insert(0, (".ShellClassInfo", list()))
        if localizedResourceName: data[0][1].append(("LocalizedResourceName", localizedResourceName))
        if iconResource and iconIndex > -1: data[0][1].append(("IconResource", f"{iconResource},{iconIndex}"))
        if infoTip: data[0][1].append(("InfoTip", infoTip))
    
    folder_path = os.path.join(dir, "desktop.ini")

    with open(folder_path, "w", encoding="gbk") as f:
        print("\n")
        for title, datas in data:
            f.write(f"[{title}]\n")
            for key, value in datas:
                f.write(f"{key}={value}\n")

    win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_SYSTEM | win32con.FILE_ATTRIBUTE_HIDDEN| win32con.FILE_ATTRIBUTE_ARCHIVE)
    win32api.SetFileAttributes(dir, win32con.FILE_ATTRIBUTE_SYSTEM)

    shell.SHChangeNotify(shellcon.SHCNE_ATTRIBUTES, shellcon.SHCNF_PATH, dir.encode(), None)
    shell.SHChangeNotify(shellcon.SHCNE_UPDATEITEM, shellcon.SHCNF_PATH, dir.encode(), None)
