import pandas as pd

# rows = []  # 同一排编号
# cols = []  # 同一列编号
# for i in range(1, 8, 2):
#     rows.append(list(range(i, i + 194, 8)) + list(range(i + 1, i + 195, 8)))
# for i in range(1, 195, 8):
#     cols.append(list(range(i, i + 8, 2)))
#     cols.append(list(range(i + 1, i + 9, 2)))
# for i in range(len(rows)):
#     print(rows[i])
# for i in range(len(cols)):
#     print(cols[i])
RepoFile = pd.ExcelFile("附件4：计算结果.xlsx")
ans = pd.read_excel(RepoFile, "Ques1")
for i in range(213):
    for j in range(213):
        if list(ans.iloc[i])[j] != list(ans.iloc[j])[i]:
            print(i, j, list(ans.iloc[i])[j], list(ans.iloc[j])[i])
