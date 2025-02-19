import win32com.client

def add_option_explicit_to_vba_modules(excel_file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    workbook = excel.Workbooks.Open(excel_file_path)

    vba_project = workbook.VBProject

    # 모든 모듈 탐색 (module.Type = Standard_Module('1'))
    for module in vba_project.VBComponents:
        if module.Type == 1:
            code_module = module.CodeModule
            existing_code = code_module.Lines(1, code_module.CountOfLines)

            # 'Option Explicit' 존재 여부 확인 (첫 줄에 추가, 기존 코드 삭제)
            if "Option Explicit" not in existing_code:
                new_code = "Option Explicit\n" + existing_code
                code_module.DeleteLines(1, code_module.CountOfLines)
                code_module.AddFromString(new_code)
                print(f"✔ 'Option Explicit' 추가: {module.Name}")

            else:
                print(f"이미 'Option Explicit'가 포함됨: {module.Name}")

    # 변경사항 저장 후 엑셀 닫기
    workbook.Save()
    workbook.Close()
    excel.Quit()

excel_file_path = r"C:\Users\DAPH-L\Desktop\eb7_project\1_3.10.5세만기_PV산출_최종_송부.xlsm"
add_option_explicit_to_vba_modules(excel_file_path)
