# 2022年11月16日开始. 17日主要功能实现. 功能是枚举所有打开的窗口,按参数调整窗口位置和尺寸
import sys,os,json,time
from PyQt5.QtWidgets import QMainWindow,QApplication,QMessageBox,QCheckBox,QListWidgetItem
import win32gui,win32con
from Ui_WindowPosAdjust import Ui_Form
 
# 定义全局常量
ghwndDict = dict()

#定义全局函数
#获得所有打开的窗口的句柄和名称，存在ghwndDict。在win32gui.EnumWindows(getHwnd, 0)中调用
def getHwnd(hwnd, mouse):
    global ghwndDict
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        ghwndDict.update({hwnd:win32gui.GetWindowText(hwnd)})

#主程序的class   WindowPosAdjust 简称 WPA
class WPA( QMainWindow, Ui_Form): 
    def __init__(self,parent =None):
        super( WPA,self).__init__(parent)
        self.setupUi(self)

        #打开配置文件，初始化界面数据
        if os.path.exists( "./WPA.ini"):
            try:
                iniFileDir = os.getcwd() + "\\"+ "WPA.ini"
                with open( iniFileDir, 'r', encoding="utf-8") as iniFile:
                    iniDict = json.loads( iniFile.read())
                if iniDict:
                    self.sbPosX.setValue( iniDict['PosX'])
                    self.sbPosY.setValue( iniDict['PosY'])
                    self.sbSizeX.setValue( iniDict['SizeX'])
                    self.sbSizeY.setValue( iniDict['SizeY'])
                    self.sbOffsetX.setValue( iniDict['OffsetX'])
                    self.sbOffsetY.setValue( iniDict['OffsetY'])
                    
                    if iniDict['SortZ'] == 'first':
                        self.rbFirstToTop.setChecked( True)
                    elif iniDict['SortZ'] == 'last':
                        self.rbLastToTop.setChecked( True)
                    else:
                        self.rbFirstToTop.setChecked( True)
                        
                    if iniDict['PosAdjust'] == True:
                        self.cbPos.setChecked( True)
                    else:
                        self.cbPos.setChecked( False)

                    if iniDict['SizeAdjust'] == True:
                        self.cbSize.setChecked( True)
                    else:
                        self.cbSize.setChecked( False)

                    if iniDict['OffsetAdjust'] == True:
                        self.cbOffset.setChecked( True)
                    else:
                        self.cbOffset.setChecked( False)

            except:
                QMessageBox.about( self, "提示", "打开初始化文件WPA.ini异常, 软件关闭时会自动重新创建WPA.ini文件")

        #绑定槽函数
        self.btnRefresh.clicked.connect( self.mfRefresh)
        self.btnExecute.clicked.connect( self.mfExecute)

    # 刷新窗口列表框
    def mfRefresh( self):
        global ghwndDict
        self.lwWindows.clear()

        win32gui.EnumWindows(getHwnd, 0)
        for k, t in ghwndDict.items():
            if t != '' and t != 'Microsoft Store' and t != 'Microsoft Text Input Application' \
            and t != 'Windows Shell Experience 主机' and t != 'Program Manager' \
            and t != '设置':               #去除title为空 和无用的句柄
                tempStr = t + '->' + str(k)
                tempCheckBox = QCheckBox( tempStr)
                tempItem = QListWidgetItem()
                self.lwWindows.addItem( tempItem)
                self.lwWindows.setItemWidget( tempItem, tempCheckBox)

        ghwndDict.clear()       # 刷新之后记得清空全局字典, 以便下次刷新

    def mfExecute( self):
        #获得ListWidght选中项组成的列表
        tempList = list()
        for i in range( 0, self.lwWindows.count()):
            if self.lwWindows.itemWidget( self.lwWindows.item(i)).isChecked():
                tempList.append( self.lwWindows.itemWidget( self.lwWindows.item(i)).text())

        #获得所有参数存入字典
        paramDict = { 'PosX':self.sbPosX.value(), 'PosY':self.sbPosY.value(), 'SizeX':self.sbSizeX.value(),
                'SizeY':self.sbSizeY.value(), 'OffsetX':self.sbOffsetX.value(), 'OffsetY':self.sbOffsetY.value(),
                'PosAdjust':self.cbPos.isChecked(), 'SizeAdjust':self.cbSize.isChecked(), 
                'OffsetAdjust':self.cbOffset.isChecked()}
        if self.rbFirstToTop.isChecked():
            paramDict['SortZ'] = 'first'
        else:
            paramDict['SortZ'] = 'last'

        #如果没有选择checkBox,直接返回
        if paramDict['PosAdjust'] == False and paramDict['SizeAdjust'] == False and paramDict['OffsetAdjust'] == False:
            QMessageBox.about( self, '提示', '请选择要进行的操作')
            return

        #调整窗口偏移值
        if not paramDict['OffsetAdjust']:
            paramDict['OffsetX'] = 0
            paramDict['OffsetY'] = 0

        #调整排序参数
        if paramDict['SortZ'] == 'first':
            sortEnum = win32con.HWND_BOTTOM
        else:
            sortEnum = win32con.HWND_TOPMOST

        #根据paramDict['PosAdjust'] 和 paramDict['SizeAdjust'] 设置SetWindowPos函数的参数uFlags
        if paramDict['PosAdjust'] == True and paramDict['SizeAdjust'] == True:
            tempFlags = win32con.SWP_NOACTIVATE     # SWP_NOACTIVATE是不激活窗口
        elif paramDict['PosAdjust'] == True and paramDict['SizeAdjust'] == False:
            tempFlags = win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE       # SWP_NOSIZE是不调整尺寸
        elif paramDict['PosAdjust'] == False and paramDict['SizeAdjust'] == True:
            tempFlags = win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE      # SWP_NOMOVE是不调整位置
        elif paramDict['PosAdjust'] == False and paramDict['SizeAdjust'] == False:
            tempFlags = win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE
            # 既不调整位置,也不调整尺寸,只使用窗口偏移,就以第一个窗口的坐标为基准,使用偏移量来调整
            paramDict['PosX'], paramDict['PosY'], r, b = win32gui.GetWindowRect( tempList[0].split('->')[1])

        #调整窗口
        try:
            for i in tempList:
                win32gui.SetWindowPos( i.split('->')[1], sortEnum, paramDict['PosX'], paramDict['PosY'],
                    paramDict['SizeX'], paramDict['SizeY'], tempFlags)
                paramDict['PosX'] = paramDict['PosX'] + paramDict['OffsetX']
                paramDict['PosY'] = paramDict['PosY'] + paramDict['OffsetY']
                time.sleep(1.5)
        except:
            QMessageBox.about( self, '提示', '执行出错,可能存在被选中但已经关闭的窗口,请点击刷新重试')

#主程序入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = WPA()
    myWin.show()
    myWin.mfRefresh()

    appExit = app.exec_()
    #退出程序之前，保存界面上的设置
    tempDict = { 'PosX':myWin.sbPosX.value(), 'PosY':myWin.sbPosY.value(), 'SizeX':myWin.sbSizeX.value(),
                'SizeY':myWin.sbSizeY.value(), 'OffsetX':myWin.sbOffsetX.value(), 'OffsetY':myWin.sbOffsetY.value(),
                'PosAdjust':myWin.cbPos.isChecked(), 'SizeAdjust':myWin.cbSize.isChecked(), 
                'OffsetAdjust':myWin.cbOffset.isChecked()}
    if myWin.rbFirstToTop.isChecked():
        tempDict['SortZ'] = 'first'
    else:
        tempDict['SortZ'] = 'last'

    saveIniJson = json.dumps( tempDict, indent=4)
    try:
        saveIniFile = open( "./WPA.ini", "w",  encoding="utf-8")
        saveIniFile.write( saveIniJson)
        saveIniFile.close()
    except:
        QMessageBox.about( myWin, "提示", "保存配置文件WPA.ini失败")

    sys.exit( appExit)
# sys.exit(app.exec_())  