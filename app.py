import win32com.client as win32
import shutil, os


ROOT_PATH = r'C:\Users\mathe\Documents\Aula1_ProjetoDeveloper'
FILENAME_DEST = 'India_Menu_copy.xlsm'
COPY_FILE_PATH = os.path.join(ROOT_PATH, FILENAME_DEST)
SHEET_NAME = "Menu"

MENUS = [
    'Beverages Menu', 'Breakfast Menu', 'Condiments Menu',
    'Desserts Menu', 'Gourmet Menu', 'McCafe Menu',
    'Regular Menu'
]

shutil.copyfile(
    os.path.join(ROOT_PATH, 'India_Menu.xlsm'),
    COPY_FILE_PATH
    )

xl = win32.gencache.EnsureDispatch("Excel.Application")
xl.Visible = True

wb = xl.Workbooks.Open(COPY_FILE_PATH)
sheet_menu = wb.Sheets(SHEET_NAME)

for menu in MENUS:
    deu_certo =  xl.Application.Run('SplitMenu', menu, menu)
    wb.SaveAs(os.path.join(ROOT_PATH, menu + '.xlsm'))
    if not deu_certo:
        break


wb.Close(False)
xl.Quit()
del xl


print("debug")