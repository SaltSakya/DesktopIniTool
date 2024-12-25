import sys
import os
import io

from PyQt6.QtWidgets import QFileDialog
from PyQt6.QtWidgets import (QApplication, QWidget, QMainWindow,
                             QPushButton, QLabel, QLineEdit, QListWidget, QListWidgetItem, QComboBox,
                             QGroupBox, QHBoxLayout, QVBoxLayout)
from PyQt6.QtGui import QIcon, QPixmap, QImage
import win32gui, win32ui, win32con
from PIL import Image

from utils import SaveConfig

class MainWidget(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        mainWidget = QWidget(self)

        # 按钮
        pickDirButton = QPushButton("选择文件夹")
        pickIconButton = QPushButton("选择图标")
        okButton = QPushButton("确定")
        cancelButton = QPushButton("取消")

        # 按钮绑定
        pickDirButton.clicked.connect(self.choose_directory)
        pickIconButton.clicked.connect(self.choose_icon)
        okButton.clicked.connect(self.confirm)
        cancelButton.clicked.connect(self.close)

        # 标签
        name = QLabel('别名')
        infoTip = QLabel('备注')

        # 单行输入框
        self.dirEdit = QLineEdit()
        self.nameEdit = QLineEdit()
        self.infoTipEdit = QLineEdit()
        self.iconEdit = QLineEdit()

        # 输入框绑定
        self.iconEdit.editingFinished.connect(self.fill_icon_list)

        # 预设图标路径
        self.presetIconPaths = QComboBox()
        self.presetIconPaths.addItem('')
        self.presetIconPaths.addItem('Ubuntu')
        self.presetIconPaths.addItem('Mandriva')
        self.presetIconPaths.addItem('Fedora')

        # TODO: 预设图标路径绑定
        #self.presetIconPaths.currentIndexChanged.connect(self.fill_iconEdit_with_preset)

        # 列表视图
        self.iconList = QListWidget()
        self.iconList.setViewMode(QListWidget.ViewMode.IconMode)

        # 目标文件夹
        dirBox = QHBoxLayout()
        dirBox.addWidget(self.dirEdit)
        dirBox.addWidget(pickDirButton)

        targetGroup = QGroupBox("目标文件夹")
        targetGroup.setLayout(dirBox)

        # 基本信息
        nameBox = QHBoxLayout()
        nameBox.addWidget(name)
        nameBox.addWidget(self.nameEdit)

        infoTipBox = QHBoxLayout()
        infoTipBox.addWidget(infoTip)
        infoTipBox.addWidget(self.infoTipEdit)

        baseBox = QVBoxLayout()
        baseBox.addLayout(nameBox)
        baseBox.addLayout(infoTipBox)

        baseGroup = QGroupBox("基本信息")
        baseGroup.setLayout(baseBox)

        # 图标

        iconFileBox = QHBoxLayout()
        iconFileBox.addWidget(self.iconEdit)
        iconFileBox.addWidget(self.presetIconPaths)
        iconFileBox.addWidget(pickIconButton)

        iconBox = QVBoxLayout()
        iconBox.addLayout(iconFileBox)
        iconBox.addWidget(self.iconList)

        iconGroup = QGroupBox("图标")
        iconGroup.setLayout(iconBox)


        # 确认
        confirmBox = QHBoxLayout()
        confirmBox.addStretch(1)
        confirmBox.addWidget(okButton)
        confirmBox.addWidget(cancelButton)

        vbox = QVBoxLayout()
        vbox.addWidget(targetGroup)
        vbox.addWidget(baseGroup)
        vbox.addWidget(iconGroup)
        vbox.addStretch(1)
        vbox.addLayout(confirmBox)

        mainWidget.setLayout(vbox)
        self.statusBar()

        self.setCentralWidget(mainWidget)

        self.setGeometry(300, 300, 350, 400)
        self.setWindowTitle('别名工具')
        self.show()

    def choose_directory(self):
        dir = QFileDialog.getExistingDirectory(self, "选择目标文件夹", os.path.expanduser("~"))
        self.dirEdit.setText(dir)

    def choose_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择图标", os.path.expanduser("~"), "图标 (*.ico *.dll *.exe)")
        self.iconEdit.setText(path)
        self.fill_icon_list()
        
    def fill_icon_list(self):
        self.iconList.clear()
        if self.iconEdit.text() == "":
            return
        
        ext = self.iconEdit.text().split(".")[-1]
        if ext == "ico":
            ico = QListWidgetItem()
            ico.setIcon(QIcon('test.ico'))
            self.iconList.addItem(ico)
        else:
            large, small = win32gui.ExtractIconEx(self.iconEdit.text(), 0, 99999)
            for i in large:
                ico = QListWidgetItem()
                ico.setIcon(QIcon(self.__getPixmapFromHicon(i)))
                self.iconList.addItem(ico)
                win32gui.DestroyIcon(i)
                
            for i in small:
                # ico = QListWidgetItem()
                # ico.setIcon(QIcon(self.__getPixmapFromHicon(i)))
                # self.iconList.addItem(ico)
                win32gui.DestroyIcon(i)

    def __getPixmapFromHicon(self, hicon) -> QPixmap:
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, 32, 32)

        hdc = hdc.CreateCompatibleDC()
        hdc.SelectObject(hbmp)
        win32gui.DrawIconEx(hdc.GetHandleOutput(), 0, 0, hicon, 32, 32, 0, None, win32con.DI_NORMAL)
        bmpinfo = hbmp.GetInfo()
        img = Image.frombuffer('RGBA', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), hbmp.GetBitmapBits(True), 'raw', 'BGRA', 0, 1)
        buffer = io.BytesIO()
        img.save(buffer, format="png")

        ico = QPixmap()
        ico.loadFromData(buffer.getvalue())
        win32gui.DeleteObject(hbmp.GetHandle())
        win32gui.DeleteDC(hdc.GetHandleOutput())
        return ico

    def LogMessage(self, msg:str, msec:int=3000):
        self.statusBar().setStyleSheet("color: black")
        self.statusBar().showMessage(msg, msec)
    
    def LogSuccess(self, msg:str, msec:int=3000):
        self.statusBar().setStyleSheet("color: green")
        self.statusBar().showMessage(msg, msec)

    def LogError(self, msg:str, msec:int=3000):
        self.statusBar().setStyleSheet("color: red")
        self.statusBar().showMessage(msg, msec)

    def confirm(self):
        dir = self.dirEdit.text()
        name = self.nameEdit.text()
        iconPath = self.iconEdit.text()
        iconIdx = self.iconList.currentIndex().row()
        if dir == "":
            self.LogError("请选择文件夹")
            return
        if not os.path.exists(dir):
            self.LogError("目标路径不存在")
            return
        if os.path.isdir(dir) == False:
            self.LogError("目标不是文件夹")
            return

        SaveConfig(dir, name, iconPath, iconIdx)
        
        self.LogSuccess("保存成功")

def main():

    app = QApplication(sys.argv)
    mw = MainWidget()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()