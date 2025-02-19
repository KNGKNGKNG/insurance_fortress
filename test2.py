import re

def tune_vba(vba_code):
    # 패턴: Public 키워드 다음에 변수 목록과 데이터 타입이 오는 형태
    pattern = r"Public\s+(.*)"
    
    tuned_code = []
    
    for line in vba_code.splitlines():
        match = re.match(pattern, line.strip())
        if match:
            variables_segment = match.group(1)
            
            # 변수들을 ',' 기준으로 분할하되, 배열의 경우 괄호 안의 ','는 임시 보호
            protected_variables = re.sub(r"\([^)]*\)", lambda x: x.group(0).replace(",", "|"), variables_segment)

            # 변수를 데이터 타입과 함께 분할하기 위한 정규식
            var_type_pairs = re.findall(r"(\S+(?:\([^)]*\))?)\s+As\s+(\w+)", protected_variables)
            print(protected_variables)
            print(var_type_pairs)
            
            # 변수를 개별적으로 나누고, 적절한 데이터 타입을 부여
            formatted_variables = []
            for var_group, dtype in var_type_pairs:
                var_list = [v.strip().replace("|", ",") for v in var_group.split(",")]
                formatted_variables.extend([f"{var} As {dtype}" for var in var_list])
            
            # 최종 라인 조합
            tuned_line = "Public " + ", ".join(formatted_variables)
            tuned_code.append(tuned_line)
        else:
            tuned_code.append(line)  # 매칭되지 않는 줄은 그대로 유지
    
    return "\n".join(tuned_code)

# 테스트 실행
test_input = '''Public start, last, 산출종류, s산출여부, s사용여부, youl_s, youl_e, youl, a, b, c, 무해지 As Integer
Public alpha1(1163, 6, 2, 3, 4), alpha2(1163, 6, 2, 30), beta(1163, 6, 2), beta5, ce, beta1, ce1 As Double
Public 순한도(1), 순한도_표준(1) As Double, 순한도1원(1), 순한도1원_표준(1) As Long, 해약공제계수, 해지공제기간 As Long
Public test1(1), test1_1(1), test_array1(1, 2, 3, 4) As Double, test2(1), test2_1(1), test_array2(10, 20, 30, 40) As Long, test3, test3_1(1), test_array4(100, 200, 300, 400) As Long'''

output = tune_vba(test_input)
# print(output)
