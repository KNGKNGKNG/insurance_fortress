import os
import win32com.client

def add_vba_module_from_bas(file_path, output_dir):
    """
    저장된 .bas 파일을 VBA 프로젝트에 추가하고, 파일명과 동일한 모듈명으로 설정하는 함수
    """

    excel = None
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        # 엑셀 워크북 열기
        workbook = excel.Workbooks.Open(file_path)
        if workbook is None:
            raise Exception("Workbook 객체 생성 실패")

        # VBA 프로젝트 접근 확인
        if not hasattr(workbook, 'VBProject'):
            raise Exception("VBA 프로젝트 접근 불가. 매크로 사용이 허용되었는지 확인하세요.")

        vba_project = workbook.VBProject

        # output_dir 내의 모든 .bas 파일을 가져옴
        for bas_file in os.listdir(output_dir):
            if bas_file.endswith(".bas"):
                bas_file_path = os.path.join(output_dir, bas_file)
                
                # .bas 파일명에서 확장자 제거하여 모듈명 생성
                module_name = os.path.splitext(bas_file)[0]

                # 기존 모듈 삭제 (같은 이름이 있으면 삭제)
                for component in vba_project.VBComponents:
                    if component.Name.lower() == module_name.lower():
                        vba_project.VBComponents.Remove(component)
                        print(f"기존 모듈 '{module_name}' 삭제 완료")

                # 새로운 모듈 추가
                new_component = vba_project.VBComponents.Import(bas_file_path)
                print(f"새로운 모듈 '{module_name}' 추가 완료")

                # 모듈명 변경
                new_component.Name = module_name
                #print(f"모듈명을 '{module_name}'으로 변경 완료")

        # 엑셀 저장 후 닫기
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        print("모든 VBA 모듈 추가 완료.")

    except Exception as e:
        print(f"오류 발생: {e}")

# 실행 코드
file_path = r"C:\Users\DAPH-L\Desktop\eb7_project\1_3.10.5세만기_PV산출_최종_송부.xlsm"
output_dir = r"C:\Users\DAPH-L\Desktop\eb7_project\output"

add_vba_module_from_bas(file_path, output_dir)
