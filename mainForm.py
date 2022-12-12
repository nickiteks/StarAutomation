from tkinter import *
from tkinter import filedialog as fd
import os
import openpyxl
from formMethods import Methods
from configParser import Settings


def settingsClick():

    def changeTxt(txt,value):
        txt.delete(0,END)
        txt.insert(0,value)

    def btnGrimech30Path():
        file = fd.askopenfilename()
        file = file.replace('/','\\\\')
        changeTxt(txtGrimech30Path,f'"{file}"')
    
    def btnThermo30Path():
        file = fd.askopenfilename()
        file = file.replace('/','\\\\')
        changeTxt(txtThermo30Path,f'"{file}"')

    def btnTransportPath():
        file = fd.askopenfilename()
        file = file.replace('/','\\\\')
        changeTxt(txtTransportPath,f'"{file}"')

    def btnStarPath():
        file = fd.askopenfilename()
        file = file.replace('/','\\\\')
        changeTxt(txtStarPath,f'"{file}"')

    def btnMacrosPath():
        file = fd.askdirectory()
        file = file.replace('/','\\\\')+"\\\\"
        changeTxt(txtMacrosPath,f'"{file}"')

    def btnCSVPath():
        file = fd.askdirectory()
        file = file.replace('/','\\\\')+"\\\\"
        changeTxt(txtCSVPath,f'"{file}"')


    def btnSaveConfig():
        settings.set_to_settings('excel_start_row',txtExcelStartRow.get())
        settings.set_to_settings('excel_start_coll',txtExcelStartColl.get())
        settings.set_to_settings('grimech30_path',txtGrimech30Path.get())
        settings.set_to_settings('thermo30_path',txtThermo30Path.get())
        settings.set_to_settings('transport_path',txtTransportPath.get())
        settings.set_to_settings('star_path',txtStarPath.get())
        settings.set_to_settings('macros_path',txtMacrosPath.get())
        settings.set_to_settings('csv_path',txtCSVPath.get())
        settings.set_to_settings('core_number',txtCoreNumber.get())
   
    settings = Settings()

    settingsWindow = Toplevel(window)
 
    settingsWindow.title("Настройки")
 
    settingsWindow.geometry("400x500")

    lblExcelStartRow = Label(settingsWindow,text='Стартовая строка Excel')
    lblExcelStartRow.grid(row = 0,column = 0)

    txtExcelStartRow = Entry(settingsWindow,width=5)
    txtExcelStartRow.insert(0,settings.get_from_settings("excel_start_row"))
    txtExcelStartRow.grid(row = 0,column = 1)

    lblExcelStartColl = Label(settingsWindow,text='Стартовая колонка Excel')
    lblExcelStartColl.grid(row = 1,column = 0)

    txtExcelStartColl = Entry(settingsWindow,width=5)
    txtExcelStartColl.insert(0,settings.get_from_settings("excel_start_coll"))
    txtExcelStartColl.grid(row = 1,column = 1)

    #Grimech30____________
    lblGrimech30Path = Label(settingsWindow,text='Путь Grimech30')
    lblGrimech30Path.grid(row = 2,column = 0)

    txtGrimech30Path = Entry(settingsWindow,width=20)
    txtGrimech30Path.insert(0,settings.get_from_settings('grimech30_path'))
    txtGrimech30Path.grid(row = 2,column = 1)

    btnGrimech30Path = Button(settingsWindow,text='...',command=btnGrimech30Path)
    btnGrimech30Path.grid(row=2,column=3)

    #Thermo30____________
    lblThermo30Path = Label(settingsWindow,text='Путь Thermo30')
    lblThermo30Path.grid(row = 3,column = 0)

    txtThermo30Path = Entry(settingsWindow,width=20)
    txtThermo30Path.insert(0,settings.get_from_settings('thermo30_path'))
    txtThermo30Path.grid(row = 3,column = 1)

    btnThermo30Path = Button(settingsWindow,text='...',command=btnThermo30Path)
    btnThermo30Path.grid(row=3,column=3)

    #Transport____________
    lblTransportPath = Label(settingsWindow,text='Путь Transport')
    lblTransportPath.grid(row = 4,column = 0)

    txtTransportPath = Entry(settingsWindow,width=20)
    txtTransportPath.insert(0,settings.get_from_settings('transport_path'))
    txtTransportPath.grid(row = 4,column = 1)

    btnTransportPath = Button(settingsWindow,text='...',command=btnTransportPath)
    btnTransportPath.grid(row=4,column=3)

    #star-ccm+____________
    lblStarPath = Label(settingsWindow,text='Путь к star-ccm+')
    lblStarPath.grid(row = 5,column = 0)

    txtStarPath = Entry(settingsWindow,width=20)
    txtStarPath.insert(0,settings.get_from_settings('star_path'))
    txtStarPath.grid(row = 5,column = 1)

    btnStarPath = Button(settingsWindow,text='...',command=btnStarPath)
    btnStarPath.grid(row=5,column=3)

    #папка с макросами____________
    lblMacrosPath = Label(settingsWindow,text='Путь к папке с макросами')
    lblMacrosPath.grid(row = 6,column = 0)

    txtMacrosPath = Entry(settingsWindow,width=20)
    txtMacrosPath.insert(0,settings.get_from_settings('macros_path'))
    txtMacrosPath.grid(row = 6,column = 1)

    btnMacrosPath = Button(settingsWindow,text='...',command=btnMacrosPath)
    btnMacrosPath.grid(row=6,column=3)

    #папка с csv____________
    lblCSVPath = Label(settingsWindow,text='Путь к папке с CSV')
    lblCSVPath.grid(row = 7,column = 0)

    txtCSVPath = Entry(settingsWindow,width=20)
    txtCSVPath.insert(0,settings.get_from_settings('csv_path'))
    txtCSVPath.grid(row = 7,column = 1)

    btntxtCSVPath = Button(settingsWindow,text='...',command=btnCSVPath)
    btntxtCSVPath.grid(row=7,column=3)

    lblCoreNumber = Label(settingsWindow,text='Количество ядер')
    lblCoreNumber.grid(row = 8,column = 0)

    txtCoreNumber = Entry(settingsWindow,width=5)
    txtCoreNumber.insert(0,settings.get_from_settings('core_number'))
    txtCoreNumber.grid(row = 8,column = 1)

    btnSaveConfig = Button(settingsWindow,text='Save Config',command=btnSaveConfig)
    btnSaveConfig.grid(row=10,column=3)




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

    ВАЖНО!!!!! базовый файл стара должен лежать в папке сохранения и называться star.sim
    ВАЖНО!!!!! в геометрии поле Body должно называться Body 1

    """
    settings = Settings()

    wb_obj = openpyxl.load_workbook(lblExcel.cget('text'))
    sheet_obj = wb_obj.active

    methods = Methods()

    row_number = methods.numberOfcullums(sheet_obj,
                                         int(settings.get_from_settings("excel_start_row")),
                                         int(settings.get_from_settings("excel_start_coll")))
    

    for i in range(row_number):
        repeat_row = methods.checkRepeatRow(sheet_obj,
                                            int(settings.get_from_settings("excel_start_row"))+i,
                                            int(settings.get_from_settings("excel_start_coll")))

        if repeat_row == 0:
            methods.changeGeom(lblGeom.cget('text'),sheet_obj,int(settings.get_from_settings("excel_start_row"))+i)
            methods.macroGeomCgange(int(settings.get_from_settings("excel_start_row")) + i,
                                    lblGeom.cget('text')[:-4].replace('/','\\\\'),
                                    sheet_obj,
                                    lblSave.cget('text').replace('/','\\\\'))

            sim = lblSave.cget('text').replace('/','\\')+'\\'
            sim = f"{sim}star.sim"
            java = f"{settings.get_from_settings('macros_path')}macros{int(settings.get_from_settings('excel_start_row')) + i}.java"
            os.system(f'start /wait cmd /c " cd {settings.get_from_settings("star_path")} & starccm+ -locale en -np {settings.get_from_settings("core_number")} {sim} -batch {java}"')

        
        else:
            methods.macroGeomDontChange(int(settings.get_from_settings("excel_start_row"))+i,sheet_obj,lblSave.cget('text').replace('/','\\\\'))
            sim = lblSave.cget('text').replace('/','\\')+'\\'
            sim = f"{sim}star{repeat_row}.sim"
            java = f"{settings.get_from_settings('macros_path')}macros{int(settings.get_from_settings('excel_start_row')) + i}.java"
            os.system(f'start /wait cmd /c " cd {settings.get_from_settings("star_path")} & starccm+ -locale en -np {settings.get_from_settings("core_number")} {sim} -batch {java}"')

        methods.reportGenerate(int(settings.get_from_settings("excel_start_row"))+i,
                                sheet_obj,
                                int(settings.get_from_settings("excel_start_coll")))


window = Tk()

mainmenu = Menu(window)
window.config(menu=mainmenu)
mainmenu.add_command(label='Настройки', command=settingsClick) 

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