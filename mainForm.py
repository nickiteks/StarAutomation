from tkinter import *
from tkinter import filedialog as fd
import os
import openpyxl
from formMethods import Methods
from settings import EXCEL_START_COLL
from settings import EXCEL_START_ROW
from settings import STAR_PATH
from settings import MACROS_PATH


def btnSaveClicked():
    directory_save = fd.askdirectory()
    lblSave.config(text=directory_save)

def btnGeomClicked():
    file_geom = fd.askopenfilename()
    lblGeom.config(text=file_geom)

def btnExcelClicked():
    file_excel = fd.askopenfilename()
    lblExcel.config(text=file_excel)

def btnOkClicked():
    """
    Основной метод работы
    Подсчет количества строк, относительно стартовой ячейки в настройках
    
    Определение повторяющейся геометрии

    Если геометрия не повторялась строим новую и выполняем все расчеты

    Если геометрия повторялась подгружаем ее и меняем значения

    ВАЖНО!!!!! базовый фал стара должен лежать в папке сохранения и называться star.sim

    """
    wb_obj = openpyxl.load_workbook(lblExcel.cget('text'))
    sheet_obj = wb_obj.active

    methods = Methods()

    row_number = methods.numberOfcullums(sheet_obj,EXCEL_START_ROW,EXCEL_START_COLL)

    print(row_number)
    

    for i in range(4):
        repeat_row = methods.checkRepeatRow(sheet_obj,EXCEL_START_ROW+i,EXCEL_START_COLL)

        if repeat_row == 0:
            methods.changeGeom(lblGeom.cget('text'),sheet_obj,EXCEL_START_ROW+i)
            methods.macroGeomCgange(EXCEL_START_ROW + i,
                                    lblGeom.cget('text')[:-4].replace('/','\\\\'),
                                    sheet_obj,
                                    lblSave.cget('text').replace('/','\\\\'))

            # sim = lblSave.cget('text').replace('/','\\')+'\\'
            # sim = f"{sim}star.sim"
            # java = f"{MACROS_PATH}macros{EXCEL_START_ROW + i}.java"
            # os.system(f'start /wait cmd /c " cd {STAR_PATH} & starccm+ -locale en -np 8 {sim} -batch {java}"')
        
        else:
            methods.macroGeomDontChange(EXCEL_START_ROW+i,sheet_obj,lblSave.cget('text').replace('/','\\\\'))
            # sim = lblSave.cget('text').replace('/','\\')+'\\'
            # sim = f"{sim}star{repeat_row}.sim"
            # java = f"{MACROS_PATH}macros{EXCEL_START_ROW + i}.java"
            # os.system(f'start /wait cmd /c " cd {STAR_PATH} & starccm+ -locale en -np 8 {sim} -batch {java}"')


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

lblSave = Label(window,text="Сохранение .sim файлов")
lblSave.grid(row=2,column=0)

btnSave = Button(window,text='Выбрать путь',command=btnSaveClicked)
btnSave.grid(row=2,column=1)

btnOk = Button(window,text='Ok',command = btnOkClicked)
btnOk.grid(row=10,column=10)

window.mainloop()