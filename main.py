import os
import re
import win32com.client


# 엑셀 매크로 파일로부터 VBA코드 가져오기
def get_vba_code(file_path, module_type="StandardModule"):
    '''
    module_type:
    - "StandardModule" : 일반 모듈
    - "ClassModule" : 클래스 모듈
    - "Document" : 시트 또는 워크북 관련 코드
    '''

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

        # VBA 프로젝트 접근
        if not hasattr(workbook, 'VBProject'):
            raise Exception("VBA 프로젝트 접근 불가.")

        vba_project = workbook.VBProject

        # VBA 코드 가져오기
        vba_code = {}
        for component in vba_project.VBComponents:
            if component.Type == 1 and module_type == 'StandardModule':
                line_count = component.CodeModule.CountOfLines
                vba_code[component.Name] = component.CodeModule.Lines(1, line_count)

        # 엑셀 Close
        workbook.Close(SaveChanges = False)
        excel.Quit()

        return vba_code
    
    except Exception as e:
        print(f"오류 발생: {e}")
    
## 단일 변수 선언으로 전환하는 메소드 (250123)##
# def tune_vba_code(vba_code):
    # pattern = r"Public\s+(.*)\s+As\s+(Integer|Double|Long|String|Variant|Worksheet|Range)"

    # tuned_code = []

    # for line in vba_code.splitlines():
        # match = re.match(pattern, line.strip())
        # if match:
            # variables = match.group(1).split(",")
            # data_type = match.group(2)

            # tuned_variables = []
            # for var in variables:
                # var = var.strip()
                # if "(" in var and ")" in var:
                    # tuned_variables.append(f"{var} As {data_type}")
                # else:
                    # tuned_variables.append(f"{var} As {data_type}")

            # tuned_line = "Public " + ", ".join(tuned_variables)
            # tuned_code.append(tuned_line)
        # else:
            # tuned_code.append(line)

    # return "\n".join(tuned_code)

def tune_vba_code(vba_code):
    pattern = r"Public\s+(.*)\s+As\s+(Integer|Double|Long|String|Variant|Worksheet|Range)"

    tuned_code = []

    for line in vba_code.splitlines():
        match = re.match(pattern, line.strip())
        if match:
            variables = match.group(1)
            data_type = match.group(2)

            # 배열 - 괄호 안의 쉼표를 임시로 보호
            protected_variables = re.sub(r"\([^)]*\)", lambda x: x.group(0).replace(",", "|"), variables)

            # 코드를 쉼표로 나누고, 다시 괄호 안 쉼표를 복원 후 데이터 타입 추가
            split_variables = protected_variables.split(",")
            tuned_variables = [
                f"{var.strip().replace('|', ',')} As {data_type}" for var in split_variables
            ]

            tuned_line = "Public " + ", ".join(tuned_variables)
            tuned_code.append(tuned_line)
        else:
            # 매칭되지 않은 줄은 그대로 추가
            tuned_code.append(line)

    return "\n".join(tuned_code)


def result_print_out(file_path, output_dir):
    vba_code_dict = get_vba_code(file_path)
    if not vba_code_dict:
        print(f"[ERROR] VBA 코드를 가져오지 못했습니다.")
        return

    for module_name, code in vba_code_dict.items():
        print(f"모듈 '{module_name}' 튜닝 중...")
        tuned_code = tune_vba_code(code)

        output_file = f"{output_dir}/{module_name}_tuned_for_eb7.bas"

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(tuned_code)
        print(f"모듈 '{module_name}'의 튜닝된 코드가 저장되었습니다: {output_file}")



file_path = r"C:\Users\user\Desktop\project\vba_tuner_to_eb7\1_3.10.5세만기_PV산출_최종_송부.xlsm"
output_dir = r"C:\Users\user\Desktop\project\vba_tuner_to_eb7\output"
result_print_out(file_path, output_dir)

