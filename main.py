from summarizing import patient_affiliation_check
import os


with open("path.txt", "r", encoding="UTF8") as f:
    path = f.read()

path_file = os.path.join(path, r"ZLHIS\吉林省人民医院病人记帐费用汇总清单.xlsx")

patient_affiliation_check(path)
