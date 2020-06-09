import pandas as pd

# 读取Excel
DistanceFile = pd.ExcelFile("./Ques1.xlsx")
RepoFile = pd.ExcelFile("附件1：仓库数据.xlsx")
DistanceData = pd.read_excel(DistanceFile)

