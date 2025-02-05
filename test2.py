import re

line = """
Public 해지율(1, 4, 110), w_rate(1, 4, 110) As Double
Public 상품V(110) As Long
"""

tuned_code = []

# 입력 문자열을 줄 단위로 처리
for l in line.splitlines():
    l = l.strip()  # 공백 제거
    if not l:  # 빈 줄은 무시
        continue

    # 정규식을 사용하여 변수와 데이터 타입 매칭
    match = re.match(r"Public\s+(.*)\s+As\s+(Integer|Double|Long|String|Variant|Worksheet|Range)", l)
    if match:
        variables = match.group(1)  # 변수 부분
        data_type = match.group(2)  # 데이터 타입

        # 괄호 안의 쉼표를 임시로 보호하고, 쉼표로 나누기
        protected_variables = re.sub(r"\([^)]*\)", lambda x: x.group(0).replace(",", "|"), variables)
        split_variables = protected_variables.split(",")
        tuned_variables = []

        # 다시 원래 쉼표 복원하고 데이터 타입 추가
        for var in split_variables:
            var = var.strip().replace("|", ",")
            tuned_variables.append(f"{var} As {data_type}")
        
        # 새로운 코드 작성
        tuned_line = "Public " + ", ".join(tuned_variables)
        tuned_code.append(tuned_line)
    else:
        # 매칭되지 않는 줄은 그대로 유지
        tuned_code.append(l)

# 결과 출력
for line in tuned_code:
    print(line)
