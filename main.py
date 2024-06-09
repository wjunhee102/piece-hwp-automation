import pandas as pd
import shutil
import win32com.client as win32
import os

def sanitize_name(name):
    invalid_chars = '<>:"/\\|?*'
    new_name = name

    for char in invalid_chars:
        new_name = name.replace(char, "")
    return new_name

def create_unique_directory(base_path, dir_name):
    try:
        new_dir_name = sanitize_name(dir_name)
        directory_path = os.path.join(base_path, new_dir_name)

        if os.path.exists(directory_path):
            counter = 1
            new_dir_name = f"{dir_name}_{counter}"
            new_directory_path = os.path.join(base_path, new_dir_name)
            
            while os.path.exists(new_directory_path):
                counter += 1
                new_dir_name = f"{dir_name}_{counter}"
                new_directory_path = os.path.join(base_path, new_dir_name)
            
            os.mkdir(new_directory_path)
            print(f"Directory created at: {new_directory_path}")
        else:
            os.mkdir(directory_path)
            print(f"Directory created at: {directory_path}")
        
        return new_dir_name
    except ValueError:
        return None

def main():
    excel_file_path = "./excel.xlsx"
    df = pd.read_excel(excel_file_path)
    template_hwp_path = "./template.hwp"
    file_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    fields = ["name", "number", "artist", "date", "size", "framesize", "material"]

    if os.path.exists(os.path.join(file_root, template_hwp_path)) is False:
        print("template.hwp파일이 존재하지 않습니다. template.hwp을 해당 프로그램 위치에 배치하여 다시 시도해주세요.")

        return

    title = input("사업명을 입력해주세요: ")
    dir_name = create_unique_directory(file_root, title)

    if dir_name is None:
        print("사업명이 폴더명에 적합하지 않습니다. 다시 시도해주세요.")
        
        return

    failed_name = []

    for index, row in df.iterrows():
        if row["name"] is None:
            print("not name")
        else:
            try:
                file_name = sanitize_name(row["name"])
                new_file_path = f"./{dir_name}/{file_name}.hwp"

                shutil.copy(template_hwp_path, new_file_path)

                hwp = win32.gencache.EnsureDispatch("Hwpframe.hwpobject")
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

                file_path = file_root + new_file_path

                hwp.Open(file_path)
                
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
            except ValueError:
                failed_name.push(row["name"])
    
    print("실패한 그림 목록입니다.")
    print(failed_name.join(", "))

main()