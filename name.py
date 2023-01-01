

import win32com.client as win32
import os
import PyQt5
from pathlib import Path
import re
import math
import sys
from PyQt5.QtWidgets import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):
        self.setGeometry(800, 200, 600, 400)

        btn1 = QPushButton("보고서 사진 자동넣기", self)
        btn2= QPushButton("이름표 만들기-한글 한쪽만 만드세요!",self)
        btn2.move(200,0)
        btn1.resize(200, 100)
        btn2.resize(400,100)
        btn1.clicked.connect(self.btn_fun_FileLoad)
        btn2.clicked.connect(self.btn_fun_FileLoad2)

    def btn_fun_FileLoad2(self):
        fname1 = QFileDialog.getOpenFileName(self, "FileLoad", 'D:/ubuntu/disks/swap.disk',
                                            'All File(*);; Text File(*.txt);; PPtx file(*ppt *pptx)',
                                            'PPtx file(*ppt *pptx)')
        fname2 = QFileDialog.getOpenFileName(self, "FileLoad", 'D:/ubuntu/disks/swap.disk',
                                            'All File(*);; Text File(*.txt);; PPtx file(*ppt *pptx)',
                                            'PPtx file(*ppt *pptx)')

        if fname1[0]:

            hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
            hwp.XHwpWindows.Item(0).Visible = True
            path = os.getcwd()
            hwp.Open(fname1[0])

            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(fname2[0])
            ws = wb.Worksheets(1)

            xlsx_values = [list(i) for i in ws.UsedRange()]

            hwp.Run("CopyPage")

            def insert(index, value):
                field_list = list(
                    ws.Range(ws.Cells(1, 1),
                             ws.Cells(1, 4)).Value[0]
                )

                print(field_list,"@@")
                for idx, field in enumerate(field_list):
                    print(idx,field_list)

                    hwp.PutFieldText(f"{field}{{{{{index}}}}}", value[idx])

            row = 2
            while True:
                if not ws.Cells(row, 1).Value:
                    hwp.Run("DeletePage")
                    break
                else:
                    data = list(
                        ws.Range(ws.Cells(row, 1),
                                 ws.Cells(row, 4)).Value[0]
                    )
                    insert(row - 2, data)
                    hwp.Run("PastePage")
                    row += 1

    def btn_fun_FileLoad(self):
        fname = QFileDialog.getOpenFileName(self, "FileLoad", 'D:/ubuntu/disks/swap.disk',
                                            'All File(*);; Text File(*.txt);; PPtx file(*ppt *pptx)',
                                            'PPtx file(*ppt *pptx)')

        if fname[0]:
            print("파일 선택됨 파일 경로는 아래와 같음")
            print(fname[0])
            hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
            hwp.XHwpWindows.Item(0).Visible = True
            path = os.getcwd()
            hwp.Open(fname[0])
            list_picture = []
            path = Path(os.getcwd())

            for file in path.glob('*.jpg'):
                list_picture.append(file)

            for file in path.glob('*.png'):
                list_picture.append(file)

            print(list_picture)

            number = math.ceil(len(list_picture) / 3)

            print(number, "@@")

            hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)  # 액션생성
            hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)  # 적용범위 구분. 없어도 됨
            hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 적용범위. 필수
            hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)  # 해당액션 실행(파라미터셋 적용)
            hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(30.0)  # 파라미터셋 설정
            hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(30.0)

            table = hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
            hwp.HParameterSet.HTableCreation.Rows = 2
            hwp.HParameterSet.HTableCreation.Cols = 3
            hwp.HParameterSet.HTableCreation.WidthType = 2
            hwp.HParameterSet.HTableCreation.HeightType = 1
            hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(148.0)
            hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(150)
            hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 3)
            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(47.0))

            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(44.0))

            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(47.0))
            hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)
            hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(40.0))
            hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(5.0))

            hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
            hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(148)
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
            hwp.Run("CopyPage")

            if(len(list_picture)>3):
                for i in range(3):
                    hwp.InsertPicture(list_picture[i].absolute(), sizeoption=3)
                    hwp.HAction.Run("MoveRight")
                    p = i

            else:
                for i in range(len(list_picture)):
                    hwp.InsertPicture(list_picture[i].absolute(), sizeoption=3)
                    hwp.HAction.Run("MoveRight")
                    p = i


            print(p)

            print(len(list_picture))


            if number >= 2:
                for i in range(number):
                    r = 3 + i
                    print(r)

                    if (r < len(list_picture) or r == len(list_picture)):
                        p = 0
                        hwp.Run("PastePage")
                        hwp.Run("MoveUp")
                        hwp.Run("MoveUp")
                        p = r
                        if (p >= len(list_picture)):
                            break
                        hwp.InsertPicture(list_picture[r].absolute(), sizeoption=3)
                        hwp.HAction.Run("MoveRight")
                        p = r + 1
                        if (p >= len(list_picture)):
                            break
                        hwp.InsertPicture(list_picture[r + 1].absolute(), sizeoption=3)
                        hwp.HAction.Run("MoveRight")
                        if (p >= len(list_picture)):
                            break
                        p = r + 2
                        if (p >= len(list_picture)):
                            break
                        hwp.InsertPicture(list_picture[r + 2].absolute(), sizeoption=3)



















        else:
            print("파일 안 골랐음")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()

