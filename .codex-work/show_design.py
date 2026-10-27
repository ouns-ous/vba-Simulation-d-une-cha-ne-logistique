from pathlib import Path
from win32com.client import dynamic
root=Path(__file__).resolve().parent.parent
xl=dynamic.Dispatch('Excel.Application')
xl.Visible=True
xl.WindowState=-4137
xl.EnableEvents=False
wb=xl.Workbooks.Open(str(root/'DataCo_Logistique.xlsm'))
xl.EnableEvents=True
xl.Run("'DataCo_Logistique.xlsm'!OpenMain")
print('New design opened in Excel')
