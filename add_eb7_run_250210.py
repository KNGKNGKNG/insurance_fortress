import win32com.client

def add_vba_module(excel_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        workbook = excel.Workbooks.Open(excel_file)

        vb_project = workbook.VBProject

        # 기존에 동일한 모듈이 있다면 삭제, 없으면 pass
        try:
            vb_component = vb_project.VBComponents("EB7_RUN")
            vb_project.VBComponents.Remove(vb_component)
        except:
            pass

        # 새 모듈 추가(Standard_Module='1')
        new_module = vb_project.VBComponents.Add(1)
        new_module.Name = "EB7_RUN"

        # VBA 코드 추가 (한글이 들어갈 경우, MBCS 인코딩 필요)
        vba_code = "Option Explicit\n''START\n"

        # 모듈에 코드 추가
        new_module.CodeModule.AddFromString(vba_code)

        # 엑셀 파일 저장
        workbook.Save()
        workbook.Close(SaveChanges=True)
        print("VBA 모듈 'EB7_RUN'이 성공적으로 추가되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        excel.Quit()


excel_file_path = r"C:\Users\DAPH-L\Desktop\eb7_project\1_3.10.5세만기_PV산출_최종_송부.xlsm"
add_vba_module(excel_file_path)
