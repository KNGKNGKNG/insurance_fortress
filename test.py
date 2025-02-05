import re

def tune_vba_code(vba_code):
    tuned_code = []

    for line in vba_code.splitlines():
        line = line.strip()
        if not line:
            continue
        

        match = re.match(r"Public\s+(.*)\s+As\s+(\w+)", line)
        if match:
            variables = match.group(1)
            data_type = match.group(2)
            
            tuned_variables = []
            for var in variables.split(","):
                var = var.strip()
                # 배열 변수 처리
                array_match = re.match(r"(.*?)(\(.*\))", var)
                if array_match:
                    var_name = array_match.group(1).strip()
                    array_dims = array_match.group(2).strip()
                    tuned_variables.append(f"{var_name}{array_dims} As {data_type}")
                else:
                    # 일반 변수 처리
                    tuned_variables.append(f"{var} As {data_type}")
            

            tuned_line = "Public " + ", ".join(tuned_variables)
            tuned_code.append(tuned_line)
        else:
            tuned_code.append(line)

    return "\n".join(tuned_code)



vba_code = """
Public 해지율(1, 4, 110), w_rate(1, 4, 110) As Double
Public 상품V(110)  As Long
"""

tuned_code = tune_vba_code(vba_code)

print("원본 VBA 코드:")
print(vba_code)
print("\n튜닝된 VBA 코드:")
print(tuned_code)
