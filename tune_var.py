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
    
## 단일 변수 선언으로 전환하는 메소드 (250207)##
def tune_var(vba_code):
    tuned_code = []
    
    for line in vba_code.splitlines():
        line = line.strip()
        if not line.startswith("Public"):
            tuned_code.append(line)
            continue

        # VBA 주석 제거
        line = re.split(r"'", line)[0].strip()

        # "Public" 제거 후 전체 문자열 처리
        line_content = line.replace("Public ", "", 1)
        
        # "As 데이터타입"을 기준으로 분리
        parts = re.split(r"\s+As\s+", line_content)
        
        # 데이터타입이 지정되지 않은 변수 선언은 그대로 유지
        if len(parts) < 2:
            tuned_code.append(line)
            continue

        parsed_variables = []
        
        for i in range(len(parts) - 1):
            var_segment = parts[i].strip()
            dtype = parts[i + 1].split(",")[0].strip()
            
            # 쉼표를 기준으로 개별 변수 추출 (배열 보호 처리)
            protected_vars = re.sub(r"\([^)]*\)", lambda x: x.group(0).replace(",", "|"), var_segment)
            variable_list = [v.strip().replace("|", ",") for v in protected_vars.split(",")]
            
            # 데이터 타입 문자열이 아닌 경우만 추가
            if dtype not in ["Integer", "Double", "Long", "String", "Variant", "Worksheet", "Range"]:
                continue
            
            # 각 변수에 해당하는 데이터 타입 매칭하여 저장
            for var in variable_list:
                if var and var not in ["Integer", "Double", "Long", "String", "Variant", "Worksheet", "Range"]:
                    parsed_variables.append(f"{var} As {dtype}")
        
        # 변환된 줄을 추가
        tuned_code.append("Public " + ", ".join(parsed_variables))
    
    return "\n".join(tuned_code)


def result_print_out(file_path, output_dir):
    vba_code_dict = get_vba_code(file_path)
    if not vba_code_dict:
        print(f"[ERROR] VBA 코드를 가져오지 못했습니다.")
        return

    for module_name, code in vba_code_dict.items():
        print(f"모듈 '{module_name}' 튜닝 중...")
        tuned_code = tune_var(code)

        output_file = f"{output_dir}/{module_name}_tuned_for_eb7.bas"

        with open(output_file, 'w', encoding='mbcs') as file:
            file.write(tuned_code)
        print(f"모듈 '{module_name}'의 튜닝된 코드가 저장되었습니다: {output_file}")



file_path = r"C:\Users\DAPH-L\Desktop\eb7_project\1_3.10.5세만기_PV산출_최종_송부.xlsm"
output_dir = r"C:\Users\DAPH-L\Desktop\eb7_project\output"
result_print_out(file_path, output_dir)

