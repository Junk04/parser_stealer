import re
import pandas as pd
from zipfile import ZipFile, BadZipFile
from rarfile import RarFile, BadRarFile, Error as RarError
import os
import rarfile

rarfile.UNRAR_TOOL = r"D:\\WinRar\\UnRAR.exe" # Укажите путь к UnRAR.exe для работы с rar архивами

# Поиск файлов с кредами
def find_password_files_in_archive(archive, file_list):
    password_files = []
    password_pattern = re.compile(r'.*passwords.*\.txt$', re.IGNORECASE)
    for file_name in file_list:
        if password_pattern.search(file_name):
            password_files.append(file_name)
    return password_files

# Извлечение данных из файла с кредами
def parse_password_file_in_archive(archive, file_name, password=None):
    extracted_data = []

    url_pattern = re.compile(r"URL\s*:\s*(.*)", re.IGNORECASE)
    login_pattern = re.compile(r"(Login|USER|USERNAME|user)\s*:\s*(.*)", re.IGNORECASE)
    password_pattern = re.compile(r"(Password|PASS|password)\s*:\s*(.*)", re.IGNORECASE)

    with archive.open(file_name, pwd=password) as file:
        content = file.read().decode('utf-8', errors='ignore')
        
        records = content.split("\n\n")
        for record in records:
            record_data = {
                "URL": "",
                "Login": "",
                "Password": ""
            }
            
            for line in record.splitlines():
                if url_match := url_pattern.match(line):
                    record_data["URL"] = url_match.group(1).strip()
                elif login_match := login_pattern.match(line):
                    record_data["Login"] = login_match.group(2).strip()
                elif password_match := password_pattern.match(line):
                    record_data["Password"] = password_match.group(2).strip()
            
            if record_data["URL"] or record_data["Login"]:
                extracted_data.append([
                    record_data["URL"], 
                    record_data["Login"], 
                    record_data["Password"]
                ])
    
    return extracted_data

# Очищает строку от недопустимых символов для Excel
def clean_string(value):
    return ''.join(c for c in value if ord(c) >= 32 and ord(c) <= 126)

# Сохранение данных в Excel
def save_to_excel(data, output_path):
    cleaned_data = []
    for row in data:
        cleaned_row = [clean_string(value) for value in row]
        cleaned_data.append(cleaned_row)
    
    df = pd.DataFrame(cleaned_data, columns=["URL", "Login", "Password"])
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Данные сохранены в файл {output_path}")

# Обработка архива
def process_archive(archive_path, password=None):
    if archive_path.endswith('.zip'):
        archive_class = ZipFile
    elif archive_path.endswith('.rar'):
        archive_class = RarFile
    else:
        print("Неизвестный формат архива. Поддерживаются только .zip и .rar.")
        return
    
    try:
        with archive_class(archive_path) as archive:
            file_list = archive.namelist()
            password_files = find_password_files_in_archive(archive, file_list)
            if not password_files:
                print("Файлы с паролями не найдены в архиве.")
                return
            
            all_data = []
            for file_name in password_files:
                print(f"Обработка файла: {file_name}")
                try:
                    file_data = parse_password_file_in_archive(archive, file_name, password=password)
                    all_data.extend(file_data)
                except RuntimeError as e:
                    print(f"Ошибка при чтении {file_name}: {e}")
            
            if all_data:
                output_excel_path = os.path.join(os.getcwd(), "data.xlsx")
                save_to_excel(all_data, output_excel_path)
            else:
                print("Не удалось извлечь данные из файлов.")
    except (BadZipFile, BadRarFile, RarError) as e:
        print(f"Ошибка при открытии архива: {e}")

def main():
    print("Программа для извлечения паролей из архивов.")
    archive_path = input("Введите путь к архиву (.zip или .rar): ").strip()
    
    password = input("Введите пароль для архива (если не требуется, оставьте пустым): ").strip()
    password = None if password == "" else password.encode('utf-8')
    
    process_archive(archive_path, password)

if __name__ == "__main__":
    main()
