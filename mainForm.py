from tkinter import *
from tkinter import filedialog as fd
import os
import openpyxl
from formMethods import Methods
from configParser import Settings
import customtkinter
import tkinter


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
        file = fd.askdirectory()
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
        settings.set_to_settings('stop_criterion',txtStopCriterion.get())
        settings.set_to_settings('base_size',txtBaseSize.get())
   
    settings = Settings()

    settingsWindow = Toplevel(window)

    settingsWindow.focus()

    settingsWindow.config(bg="#242424")

 
    settingsWindow.title("Настройки")
 
    settingsWindow.geometry("400x400")

    lblExcelStartRow = Label(settingsWindow,text='Стартовая строка Excel',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblExcelStartRow.grid(row = 0,column = 0,sticky=E)

    txtExcelStartRow = Entry(settingsWindow,width=20)
    txtExcelStartRow.insert(0,settings.get_from_settings("excel_start_row"))
    txtExcelStartRow.grid(row = 0,column = 1)

    lblExcelStartColl = Label(settingsWindow,text='Стартовая колонка Excel',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblExcelStartColl.grid(row = 1,column = 0,sticky=E)

    txtExcelStartColl = Entry(settingsWindow,width=20)
    txtExcelStartColl.insert(0,settings.get_from_settings("excel_start_coll"))
    txtExcelStartColl.grid(row = 1,column = 1)

    #Grimech30____________
    lblGrimech30Path = Label(settingsWindow,text='Путь Grimech30',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblGrimech30Path.grid(row = 2,column = 0,sticky=E)

    txtGrimech30Path = Entry(settingsWindow,width=20)
    txtGrimech30Path.insert(0,settings.get_from_settings('grimech30_path'))
    txtGrimech30Path.grid(row = 2,column = 1)

    btnGrimech30Path = Button(settingsWindow,text='...',command=btnGrimech30Path)
    btnGrimech30Path.grid(row=2,column=3,sticky=W)

    #Thermo30____________
    lblThermo30Path = Label(settingsWindow,text='Путь Thermo30',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblThermo30Path.grid(row = 3,column = 0,sticky=E)

    txtThermo30Path = Entry(settingsWindow,width=20)
    txtThermo30Path.insert(0,settings.get_from_settings('thermo30_path'))
    txtThermo30Path.grid(row = 3,column = 1)

    btnThermo30Path = Button(settingsWindow,text='...',command=btnThermo30Path)
    btnThermo30Path.grid(row=3,column=3,sticky=W)

    #Transport____________
    lblTransportPath = Label(settingsWindow,text='Путь Transport',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblTransportPath.grid(row = 4,column = 0,sticky=E)

    txtTransportPath = Entry(settingsWindow,width=20)
    txtTransportPath.insert(0,settings.get_from_settings('transport_path'))
    txtTransportPath.grid(row = 4,column = 1)

    btnTransportPath = Button(settingsWindow,text='...',command=btnTransportPath)
    btnTransportPath.grid(row=4,column=3,sticky=W)

    #star-ccm+____________
    lblStarPath = Label(settingsWindow,text='Путь к star-ccm+',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblStarPath.grid(row = 5,column = 0,sticky=E)

    txtStarPath = Entry(settingsWindow,width=20)
    txtStarPath.insert(0,settings.get_from_settings('star_path'))
    txtStarPath.grid(row = 5,column = 1,sticky=W)

    btnStarPath = Button(settingsWindow,text='...',command=btnStarPath)
    btnStarPath.grid(row=5,column=3,sticky=W)

    #папка с макросами____________
    lblMacrosPath = Label(settingsWindow,text='Путь к папке с макросами',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblMacrosPath.grid(row = 6,column = 0,sticky=E)

    txtMacrosPath = Entry(settingsWindow,width=20)
    txtMacrosPath.insert(0,settings.get_from_settings('macros_path'))
    txtMacrosPath.grid(row = 6,column = 1)

    btnMacrosPath = Button(settingsWindow,text='...',command=btnMacrosPath)
    btnMacrosPath.grid(row=6,column=3,sticky=W)

    #папка с csv____________
    lblCSVPath = Label(settingsWindow,text='Путь к папке с CSV',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblCSVPath.grid(row = 7,column = 0,sticky=E)

    txtCSVPath = Entry(settingsWindow,width=20)
    txtCSVPath.insert(0,settings.get_from_settings('csv_path'))
    txtCSVPath.grid(row = 7,column = 1)

    btntxtCSVPath = Button(settingsWindow,text='...',command=btnCSVPath)
    btntxtCSVPath.grid(row=7,column=3,sticky=W)

    lblCoreNumber = Label(settingsWindow,text='Количество ядер',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblCoreNumber.grid(row = 8,column = 0,sticky=E)

    txtCoreNumber = Entry(settingsWindow,width=20)
    txtCoreNumber.insert(0,settings.get_from_settings('core_number'))
    txtCoreNumber.grid(row = 8,column = 1)

    lblStopCriterion = Label(settingsWindow,text='Критерий остановки',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblStopCriterion.grid(row = 9,column = 0,sticky=E)

    txtStopCriterion = Entry(settingsWindow,width=20)
    txtStopCriterion.insert(0,settings.get_from_settings('stop_criterion'))
    txtStopCriterion.grid(row = 9,column = 1)

    lblBaseSize = Label(settingsWindow,text='Базовый размер',background="#242424",font=('Times New Roman italic', 11),foreground="white")
    lblBaseSize.grid(row = 10,column = 0,sticky=E)

    txtBaseSize = Entry(settingsWindow,width=20)
    txtBaseSize.insert(0,settings.get_from_settings('base_size'))
    txtBaseSize.grid(row = 10,column = 1)

    btnSaveConfig = customtkinter.CTkButton(settingsWindow,text='Save Config',command=btnSaveConfig)
    btnSaveConfig.place(relx=0.8, rely=0.9, anchor=tkinter.CENTER)




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


window = customtkinter.CTk()


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")

btnSettings = customtkinter.CTkButton(window,text='Настройки',command = settingsClick,width=5,height=1)
btnSettings.place(relx=0.05, rely=0.025, anchor=tkinter.CENTER)

window.title("Оптимизация расчетов")
window.minsize(650,300)
window.maxsize(650,300)


lblGeom = Label(window,text='Геометрия',width=51,background="#242424",font=('Times New Roman italic', 11),foreground="white")
lblGeom.grid(row=0,column=0,sticky=E)

btnGeom = customtkinter.CTkButton(window,text="Выбрать геометрию",command = btnGeomClicked,)
btnGeom.grid(row=0,column=1,pady=(15, 10))

lblExcel = Label(window,text='Excel',width=51,background="#242424",font=('Times New Roman italic', 11),foreground="white")
lblExcel.grid(row=1,column=0,sticky=E)

btnExcel = customtkinter.CTkButton(window,text="Выбрать Excel",command = btnExcelClicked,)
btnExcel.grid(row=1,column=1,pady=(15, 10))

lblSave = Label(window,text="Сохранение .sim файлов",width=51,background="#242424",font=('Times New Roman italic', 11),foreground="white")
lblSave.grid(row=2,column=0,sticky=E)

btnSave = customtkinter.CTkButton(window,text='Выбрать путь',command=btnSaveClicked,)
btnSave.grid(row=2,column=1,pady=(15, 10))

btnOk = customtkinter.CTkButton(window,text='Ok',command = btnOkClicked)
btnOk.place(relx=0.85, rely=0.9, anchor=tkinter.CENTER)

window.mainloop()
