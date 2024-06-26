import pandas as pd
import shutil
import win32com.client as win32
import os
import sys
from tqdm import tqdm
import keyboard
import threading
import time
import datetime

def sanitize_name(name):
  invalid_chars = '<>:"/\\|?*\n'
  new_name = f"{name}"

  for char in invalid_chars:
    new_name = new_name.replace(char, "")
  return new_name

def get_title_input():
  while True:
    title = input("사업명을 입력해주세요 (필수작성): ")

    if title.strip() == "":
      print("사업명은 필수로 작성해주셔야 합니다.")
    else:
      return title
    
def get_current_date_input():
  today = datetime.datetime.now().strftime("%Y.%m.%d")
  
  user_input = input(f"날짜를 YYYY.MM.DD 형식으로 입력해주세요 (비워두면 오늘 날짜인 {today} 사용): ")
  
  if not user_input: 
    return today
  else:
    try:
      return datetime.datetime.strptime(user_input, "%Y.%m.%d").date()
    except ValueError:
      print("잘못된 날짜 형식입니다. 다시 시도해주세요.")
      return get_current_date_input()

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

def create_txt(name, str):
  try:
    file_name = f"{name}.txt"

    with open(file_name, "w") as file:
      file.write(str)
      print(f"{file_name}이 생성되었습니다. \n")

  except Exception as e:
    print(e)
    print(f"{name}.txt 생성에 실패하였습니다. \n")

def get_settings(base_path):
  settings_path = os.path.join(base_path, "settings.txt")
  default_target_name = "name"
  default_sub_target_name = "artist"
  default_fields = ["name", "number", "artist", "date", "size", "framesize", "material", "place"]
  target_name = default_target_name
  sub_target_name = ""
  fields = default_fields

  try:
    with open(settings_path, "r") as file:
      for line in file:
        line = line.strip()

        if line:
          key, value = line.split("=", 1)
          key = key.strip()
          
          if key == "target":
            if value.isalpha():
              target_name = value.strip()

          elif key == "subtarget":
            if value.isalpha():
              sub_target_name = value.strip()

          elif key == "fields":
            try:
              fields = [item.strip() for item in value.split(",")]
            except:
              print("Fields 설정을 파싱할 수 없습니다.")

              return [default_target_name, default_sub_target_name, default_fields]
                    
    return [target_name, sub_target_name, fields]

  except FileNotFoundError:
    print("\nsettings.txt가 존재하지 않아 새로 생성합니다.")
    
    settings = f"target={default_target_name}\nsubtarget={default_sub_target_name}\nfields={','.join(default_fields)}"
    
    create_txt("settings", settings)

    return [default_target_name, default_sub_target_name, default_fields]
  except Exception as e:
    print("사용중 오류가 발생했습니다. 하단의 에러 메세지를 개발자에게 전달하여 문제를 해결할 수 있습니다.")
    print(f"{e}")

    return [default_target_name, default_sub_target_name, default_fields]


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
  print("\n** Piece hwp automation **\n")

  template_hwp_path = "./template.hwp"
  file_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))

  if getattr(sys, 'frozen', False):
    file_root = os.path.dirname(os.path.abspath(sys.executable))

  settings = get_settings(file_root)

  target_name = settings[0]
  sub_target_name = settings[1]
  fields = settings[2]

  if os.path.exists(os.path.join(file_root, template_hwp_path)) is False:
    print("template.hwp이 존재하지 않습니다. template.hwp을 해당 프로그램 위치에 배치하여 다시 시도해주세요.")

    return

  excel_file_name = get_excel_filename()
  excel_file_path = f"./{excel_file_name}.xlsx"

  if os.path.exists(os.path.join(file_root, excel_file_path)) is False:
    print(f"{excel_file_name}.xlsx이 존재하지 않습니다. {excel_file_name}.xlsx을 해당 프로그램 위치에 배치하여 다시 시도해주세요.")

    return
  
  df = pd.read_excel(excel_file_path)
  print(f"\n{excel_file_name}를 성공적으로 불러왔습니다.")
  print(f"총 라인의 수: {df.shape[0]} \n")

  title = get_title_input()
  current_date = get_current_date_input()
  start_point = get_numeric_input("시작지점을 입력해주세요. (비워두면 1로 입력): ")
  selected_count = get_numeric_input("생성할 파일의 수를 입력해주세요. (비워두면 전체 생성): ", -1)

  dir_name = create_unique_directory(file_root, title)

  if dir_name is None:

    print("사업명이 폴더명에 적합하지 않습니다. 다시 시도해주세요.")
    
    return

  progress_count = 0
  generated_count = 0
  failed_names = []

  print("\n- 정보 -")
  print(f"가져올 데이터 파일: {excel_file_name}.xlsx")
  print(f"사업명: {title}")
  print(f"날짜: {current_date}")

  if sub_target_name != "":
    print(f"파일명이 될 필드: {target_name}, {sub_target_name}")
  else:
    print(f"파일명이 될 필드: {target_name}")

  print(f"입력될 필드: {fields}")
  print(f"시작지점: {start_point + 1}")

  if selected_count < 0:
    print(f"생성할 파일의 수: 전체")
  else:
    print(f"생성할 파일의 수: {selected_count + 1}")

  print("\n작업을 시작합니다. (작업을 중단하려면 ESC를 눌러주세요.)\n")

  for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="진행 중"):      
    if start_point > index:
      continue

    if selected_count > -1 and index > start_point + selected_count:
      break

    progress_count += 1

    if pd.isna(row[target_name]) or row[target_name] == "":
      failed_names.append(f"[index: {index + 1}, {target_name}: 없음]")
    
    else:
      try:
        name = sanitize_name(row[target_name])
        new_file_path = f"./{dir_name}/{index + 1}-{name}.hwp"
        subname = ""

        if sub_target_name != "":
          if pd.isna(row[sub_target_name]) is False and row[sub_target_name] != "":
            subname = sanitize_name(row[sub_target_name])

        if subname != "":
          new_file_path = f"./{dir_name}/{index + 1}-{name}-{subname}.hwp"
          

        shutil.copy(template_hwp_path, new_file_path)

        hwp = win32.gencache.EnsureDispatch("Hwpframe.hwpobject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        file_path = file_root + new_file_path

        hwp.Open(file_path)
        
        hwp.PutFieldText("title", title)
        hwp.PutFieldText("currentdate", current_date)

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
          failed_names.append(f"[index: {index + 1}, {target_name}: {row[target_name]}]")

  print(f"\n총 {progress_count} 중 {generated_count} 생성완료. \n")
  
  if not failed_names:
    print("성공적으로 완료됐습니다!")
    return
  else:
    print("실패한 그림 목록입니다.")
    print(", ".join(failed_names))

def check_esc():
  while True:
    if keyboard.is_pressed('esc'):
      print("\n 작업이 중단되었습니다.")
      time.sleep(0.2)
      os._exit(1) 
        
    time.sleep(0.1)
    

esc_thread = threading.Thread(target=check_esc)
esc_thread.daemon = True
esc_thread.start()

try:
  main()
except Exception as e:
  print("\n사용중 오류가 발생했습니다.\n하단의 에러 메세지를 개발자에게 전달하여 문제를 해결할 수 있습니다.(email: wjunhee102@gmail.com)\n")
  print(f"{e}")

input("\n종료하려면 enter 키를 눌러주세요.\n")