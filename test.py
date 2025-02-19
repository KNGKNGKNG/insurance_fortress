import re

def tune_vba(vba_code):
    pattern = r"Public\s+(.*)\s+As\s+(Integer|Double|Long|String|Variant|Worksheet|Range)"
    # dup_type_pattern = r"(As\s+(Integer|Double|Long|String|Variant|Worksheet|Range)),"
    # remaining_type_pattern = r"(As\s+(Integer|Double|Long|String|Variant|Worksheet|Range))"

    tuned_code = []

    for line in vba_code.splitlines():
        match = re.match(pattern, line.strip())
        # dup_match = re.match(dup_type_pattern, line.strip())

        if match:
            variables = match.group(1)
            data_type = match.group(2)
            print(variables)
            print(data_type)
            

            ## 배열 구조를 가지고 있는 코드를 정규식으로 찾고, 괄호 안의 쉼표를 임시로 보호
            protected_variables = re.sub(r"\([^)]*\)", lambda x: x.group(0).replace(",", "|"), variables)

            ## 코드를 쉼표로 나누고, 다시 괄호 안 쉼표를 복원 후 데이터 타입 추가
            split_variables = protected_variables.split(",")
            tuned_variables = [
                f"{var.strip().replace('|', ',')} As {data_type}" for var in split_variables
            ]

            tuned_line = "Public " + ", ".join(tuned_variables)
            tuned_code.append(tuned_line)
        # elif dup_match:

        else:
            ## 매칭되지 않은 줄은 그대로 추가
            tuned_code.append(line)

    return "\n".join(tuned_code)

    # matches = list(re.finditer(dup_type_pattern, vba_code))
    # print(matches)

    # tuned_variables = []
    # start_index = 0

    # for match in matches:
        # data_type = match.group(2)
        # end_index = match.start()

        # variables_part = vba_code[start_index:end_index].strip()
        # variables = [var.strip() for var in variables_part.split(",")]
        # print(variables_part)
        # print(variables)

        # typed_variables = [f"{var} As {data_type}" if " As " not in var else var for var in variables]
        # print(typed_variables)

        # tuned_variables.extend(typed_variables)
        # start_index = match.end()

    # remaining_part = vba_code[start_index:].strip()
    # remain_match = re.match(remaining_type_pattern, line.strip())
    # print(remain_match)

    # if remaining_part:
        # remaining_variables = [var.strip() for var in remaining_part.split(",")]
        # tuned_variables.extend(remaining_variables)

    # 최종 output
    # return ", ".join(tuned_variables)

origin_vba_code = """Public start, last, 산출종류, s산출여부, s사용여부, youl_s, youl_e, youl, a, b, c, 무해지 As Integer
Public alpha1(1163, 6, 2, 3, 4), alpha2(1163, 6, 2, 30), beta(1163, 6, 2), beta5, ce, beta1, ce1 As Double
Public 순한도(1), 순한도_표준(1) As Double, 순한도1원(1), 순한도1원_표준(1) As Long, 해약공제계수, 해지공제기간 As Long
Public test1(1), test1_1(1), test_array1(1, 2, 3, 4) As Double, test2(1), test2_1(1), test_array2(10, 20, 30, 40) As Long, test3, test3_1(1), test_array4(100, 200, 300, 400) As Long
"""

tuned_vba_code = tune_vba(origin_vba_code)
# print(f"origin_code: \n{origin_vba_code}")
# print(f"tuned_code: \n{tuned_vba_code}")
