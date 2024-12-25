import os
import re

import win32com.shell
import win32com.shell.shell
import win32com.shell.shellcon
import win32api, win32con
import win32com

def insertData(data, localizedResourceName: str|None = None, IconResource: str|None = None, IconIndex: int = 0) -> bool:
    # TODO: 原本没有的属性不会添加
    for title, datas in data:
        if title == ".ShellClassInfo":
            for i in range(len(datas)):
                if datas[i][0] == "LocalizedResourceName" and localizedResourceName: datas[i] = ("LocalizedResourceName", localizedResourceName)
                elif datas[i][0] == "IconResource" and IconResource: datas[i] = ("IconResource", f"{IconResource},{IconIndex}")
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
    # TODO: ico格式的图标表示可能不同，之后要试一下
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

    win32com.shell.shell.SHChangeNotify(win32com.shell.shellcon.SHCNE_ATTRIBUTES, win32com.shell.shellcon.SHCNF_PATH, dir.encode(), None)
    win32com.shell.shell.SHChangeNotify(win32com.shell.shellcon.SHCNE_UPDATEITEM, win32com.shell.shellcon.SHCNF_PATH, dir.encode(), None)
