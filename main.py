import pandas as pd
import shutil
import win32com.client as win32
import os

excel_file_path = "./excel.xlsx"
df = pd.read_excel(excel_file_path)

template_hwp_path = "./template.hwp"
file_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))

title = input("사업명을 입력해주세요: ")

for index, row in df.iterrows():
    if row["name"] is None:
        print("not name")
    else:
        new_file_path = f"./dist/{row["name"]}.hwp"

        shutil.copy(template_hwp_path, new_file_path)

        hwp = win32.gencache.EnsureDispatch("Hwpframe.hwpobject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        file_path = file_root + new_file_path

        hwp.Open(file_path)

        fields = ["name", "number", "artist", "date", "size", "framesize", "material"]
        
        hwp.PutFieldText("title", title)

        for field in fields:
            if field in df.columns: 
                if row[field] is None:
                    hwp.PutFieldText(field, "-")
                else:
                    hwp.PutFieldText(field, str(row[field]))
            else:
                hwp.PutFieldText(field, "-")

        hwp.Save()
        hwp.Quit()
