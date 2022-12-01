from tkinter import *
from tkinter import filedialog as fd
import openpyxl
from formMethods import Methods
from settings import EXCEL_START_COLL
from settings import EXCEL_START_ROW


def btnGeomClicked():
    fileGeom = fd.askopenfilename()
    lblGeom.config(text=fileGeom)

def btnExcelClicked():
    fileExcel = fd.askopenfilename()
    lblExcel.config(text=fileExcel)

def btnOkClicked():
    """
    Основной метод работы
    Подсчет количества строк, относительно стартовой ячейки в настройках
    Определение повторяющейся геометрии 
    """
    wb_obj = openpyxl.load_workbook(lblExcel.cget('text'))
    sheet_obj = wb_obj.active

    methods = Methods()

    row_number = methods.numberOfcullums(sheet_obj,EXCEL_START_ROW,EXCEL_START_COLL)

    print(row_number)

    repeat_row = methods.checkRepeatRow(sheet_obj,13,EXCEL_START_COLL)

    print(repeat_row)
    

    for i in range(row_number):
        methods.changeGeom(lblGeom.cget('text'),sheet_obj,EXCEL_START_ROW+i)
        methods.macroGeomCgange(EXCEL_START_ROW + i,lblGeom.cget('text')[:-4].replace('/','\\\\'),sheet_obj)

window = Tk()

window.title("Оптимизация расчетов")
window.geometry('600x300')

lblGeom = Label(window,text='Выбор геометрии:')
lblGeom.grid(row=0,column=0)

btnGeom = Button(window,text="Выбрать геометрию",command = btnGeomClicked)
btnGeom.grid(row=0,column=1)

lblExcel = Label(window,text='Выбор Excel:')
lblExcel.grid(row=1,column=0)

btnExcel = Button(window,text="Выбрать Excel",command = btnExcelClicked)
btnExcel.grid(row=1,column=1)

btnOk = Button(window,text='Ok',command = btnOkClicked)
btnOk.grid(row=10,column=10)

window.mainloop()