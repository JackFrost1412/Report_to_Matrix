from Report_to_Matrix import report_to_matrix
import configparser

# Đọc đường dẫn thư mục từ file config.ini
config = configparser.ConfigParser()
try:
    config.read('config.ini', encoding='utf-8')
    file = config['files']['file_name']
except (FileNotFoundError, KeyError) as e:
    print(f"Error reading config file: {e}")
    exit(1)

file_path = fr"Matrix_BRD\{file}"

report_to_matrix(file_path)