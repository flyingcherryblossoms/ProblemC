import pandas as pd


# 读取Excel
distanceFile = pd.ExcelFile("./Ques1.xlsx")
sourceFile = pd.ExcelFile("附件1：仓库数据.xlsx")
distanceData = pd.read_excel(distanceFile)  # 距离数据
ReviewStation = pd.read_excel(sourceFile, "复核台")
taskLists = pd.read_excel(sourceFile, "任务单")  # 任务单

taskList = []  # 任务单T0001
for i in range(23):
    taskList.append(list(taskLists.iloc[i]))
    # print("货架：", int(taskList[i][2][-5:-2]))
    # print("货格", int(taskList[i][2][-2:]))
    # print("序号：", (int(taskList[i][2][-5:-2])-1)*15 + int(taskList[i][2][-2:]))
    taskList[i].append((int(taskList[i][2][-5:-2]) - 1) * 15 + int(taskList[i][2][-2:]))
    print(taskList[i])

distance = []  # 任务单每个货格到其它任务单货格的距离和所有复核台的距离
# 保存计算所需数据，避免去所有距离中读取，太浪费时间了
for i in range(len(taskList)):
    distancei = []
    indexi = taskList[i][4]-1
    # print("indexi= ", indexi)
    for j in range(len(taskList)):
        indexj = taskList[j][4]-1
        # print("indexj= ", indexj)
        distancei.append(list(distanceData.iloc[indexi])[indexj])
    for j in range(3000, 3013):
        indexj = j
        distancei.append(list(distanceData.iloc[indexi])[indexj])
    print(distancei)
    distance.append(distancei)


def Manhattan(x1, y1, x2, y2):
    # 计算曼哈顿距离
    return abs(int(x1) - int(x2)) + abs(int(y1) - int(y2))

# 计算复核台之间距离
ans2 = []
for i in range(len(ReviewStation)):
    ansi = []
    rsi = list(ReviewStation.iloc[i])  # 复核台i信息
    for j in range(len(ReviewStation)):
        rsj = list(ReviewStation.iloc[j])  # 复核台j信息
        dist = Manhattan(rsi[1], rsi[2], rsj[1], rsj[2])
        ansi.append(dist)
    ans2.append(ansi)
    # print(len(ansi), ansi)

lenansj = len(distance)
lenansi = len(distance[1])
k = 0
for i in range(lenansi-13, lenansi):  # 列循环
    ansi = []
    for j in range(0, lenansj):  # 行循环
        ansi.append(distance[j][i])
    ansi.extend(ans2[k])
    k += 1
    # print(len(ansi), ansi)
    distance.append(ansi)

pd.DataFrame(distance).to_excel("P2_distance.xlsx", sheet_name='P2_distance', index=False)
