import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const

def stamp(doc7):
    characteristic = {}

    stamp = doc7.LayoutSheets.Item(0).Stamp
    characteristic["applicable_stamp"] = stamp.Text(25).Str
    characteristic["decimal_stamp"] = stamp.Text(2).Str
    characteristic["name_stamp"] = (stamp.Text(1).Str).replace("\n", " ")
    characteristic["Designer"] = stamp.Text(110).Str
    return characteristic

def parse_design_documents(paths):
    module7, api7, const7 = get_kompas_api7()
    app7 = api7.Application
    app7.Visible = True
    app7.HideMessage = const7.ksHideMessageNo

    table = []
    for path in paths:
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=False,
                                   ReadOnly=True)
        table.append(stamp(doc7))

        doc7.Close(const7.kdDoNotSaveChanges)
    app7.Quit()
    return table

def print_to_excel(result):
    excel = Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Add()
    sheet = wb.ActiveSheet

    sheet.Range("A1:D1").value = ["Перв.прим. ", "Дец.номер", "Наименование", "Разработчик"]
    for i, row in enumerate(result):
        sheet.Cells(i + 2, 1).value = row['applicable_stamp']
        sheet.Cells(i + 2, 2).value = row['decimal_stamp']
        sheet.Cells(i + 2, 3).value = row['name_stamp']
        sheet.Cells(i + 2, 4).value = row['Designer']

if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    filenames = askopenfilenames(title="Выберети чертежи деталей", filetypes=[('Файлы Компас3D ', '*.spw;*.cdw'), ])

    result = parse_design_documents(filenames)
    print_to_excel(result)

    root.destroy()
root.mainloop()