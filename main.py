import pandas as pd
import shutil
import win32com.client as win32
import os
from tqdm import tqdm
import keyboard

def sanitize_name(name):
  invalid_chars = '<>:"/\\|?*\n'
  new_name = f"{name}"

  for char in invalid_chars:
    new_name = new_name.replace(char, "")
  return new_name

def get_excel_filename():
  file_name = input("Excel 파일의 이름을 입력하세요 (확장자 .xlsx 포함 가능): ")
  
  if file_name.endswith('.xlsx'):
    file_name = file_name[:-5]
  
  return file_name

def get_numeric_input(prompt, default=0):
  number = default + 1

  while True:
    user_input = input(prompt)
    
    if user_input == "":
      return default
    try:
      number = int(user_input)
      
      break
    except ValueError:
      print("유효하지 않은 입력입니다. 숫자를 입력해주세요.")

  if number < 1:
    return default
  
  return number - 1

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
        print(f"폴더가 생성되었습니다: {new_directory_path}")
      else:
          os.mkdir(directory_path)
          print(f"폴더가 생성되었습니다:  {directory_path}")
      
      return new_dir_name
    except Exception as e:
        print("사용중 오류가 발생했습니다. 하단의 에러 메세지를 개발자에게 전달하여 문제를 해결할 수 있습니다.")
        print(f"{e}")
        return None

def main():

  template_hwp_path = "./template.hwp"
  file_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))
  fields = ["name", "number", "artist", "date", "size", "framesize", "material"]

  if os.path.exists(os.path.join(file_root, template_hwp_path)) is False:
    print("template.hwp이 존재하지 않습니다. template.hwp을 해당 프로그램 위치에 배치하여 다시 시도해주세요.")

    return

  title = input("사업명을 입력해주세요: ")
  excel_file_name = get_excel_filename()
  excel_file_path = f"./{excel_file_name}.xlsx"
  start_point = get_numeric_input("시작지점을 입력해주세요. (빈값으로 입력시 1): ")
  selected_count = get_numeric_input("생성할 파일의 수를 입력해주세요. (빈값으로 입력시 전체 생성): ")

  dir_name = create_unique_directory(file_root, title)

  if dir_name is None:

    print("사업명이 폴더명에 적합하지 않습니다. 다시 시도해주세요.")
    
    return

  if os.path.exists(os.path.join(file_root, excel_file_path)) is False:
    print(f"{excel_file_name}.xlsx이 존재하지 않습니다. {excel_file_name}.xlsx을 해당 프로그램 위치에 배치하여 다시 시도해주세요.")

    return

  df = pd.read_excel(excel_file_path)

  progress_count = 0;
  generated_count = 0;
  failed_names = []

  print("\n - 정보 -")
  print(f"사업명: {title}")
  print(f"설정된 시작지점: {start_point + 1}")
  print(f"설정된 생성할 파일의 수: {selected_count + 1}")
  print("\n 작업을 시작합니다.")

  for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="진행 중"):
    if keyboard.is_pressed('esc'):
      response = input("작업을 중단하시겠습니까? (y/n): ")
      if response.lower() == 'y':
          print("작업이 중단되었습니다.")
          break
      
    if start_point > index:
      continue

    if selected_count != 0 and index > start_point + selected_count:
      break

    progress_count += 1

    if pd.isna(row['name']) or row['name'] == '':
      failed_names.append(f"[index: {index + 1}, name: 이름 없음]")
    
    else:
      try:
        file_name = sanitize_name(row["name"])
        new_file_path = f"./{dir_name}/{index + 1}-{file_name}.hwp"

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

        generated_count += 1
        hwp.Save()
        hwp.Quit()

      except ValueError:
          failed_names.append(f"[index: {index + 1}, name: {row["name"]}]")

  print(f"\n 총 {progress_count}중 {generated_count} 생성완료. \n")
  
  if not failed_names:
    print("성공적으로 완료됐습니다!")
    return
  else:
    print("실패한 그림 목록입니다.")
    print(", ".join(failed_names))

try:
  main()
except Exception as e:
  print("\n 사용중 오류가 발생했습니다. 하단의 에러 메세지를 개발자에게 전달하여 문제를 해결할 수 있습니다. \n")
  print(f"{e}")