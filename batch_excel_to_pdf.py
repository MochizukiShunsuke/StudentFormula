import os
import win32com.client

# 相対パスやユーザー入力に変更
folder_path = input("PDF化したいExcelファイルが入ったフォルダのパスを入力してください: ")

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False

for file in os.listdir(folder_path):
    if file.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file)
        file_base_name = os.path.splitext(file)[0]
        combined_name_pdf = f"{file_base_name}.pdf"

        try:
            wb = excel.Workbooks.Open(file_path)
            wb.ExportAsFixedFormat(0, os.path.join(folder_path, combined_name_pdf))
            wb.Close()
        except Exception as e:
            print(f"Failed to convert {file_path}: {e}")

excel.Quit()
