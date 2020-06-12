import numpy as np
from scipy import spatial
import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 读取Excel
sourceFile = pd.ExcelFile("../题目/附件1：仓库数据.xlsx")
CargoContainer = pd.read_excel(sourceFile, "货格")
taskLists = pd.read_excel(sourceFile, "任务单")  # 任务单
ReviewStation = pd.read_excel(sourceFile, "复核台")  # 复核台
distanceP2 = pd.ExcelFile("./P2_distance.xlsx")

# 任务单T0001
taskList = []
for i in range(23):
    taskList.append(list(taskLists.iloc[i]))
    taskList[i].append((int(taskList[i][2][-5:-2]) - 1) * 15 + int(taskList[i][2][-2:]))
for i in range(1, 14):
    if i < 10:
        taskList.append(['T0001', '', 'FH0' + str(i), 0, 3000 + i])
    else:
        taskList.append(['T0001', '', 'FH' + str(i), 0, 3000 + i])

# 货格和复核台的位置
taskPoints = []
for i in range(23):
    for j in range(len(CargoContainer)):
        if taskList[i][2] == list(CargoContainer.iloc[j])[0]:
            taskPoints.append([int(list(CargoContainer.iloc[j])[1]), int(list(CargoContainer.iloc[j])[2])])
for i in range(len(ReviewStation)):
    taskPoints.append([int(list(ReviewStation.iloc[i])[1]), int(list(ReviewStation.iloc[i])[2])])
taskPoints = np.array(taskPoints)

# 任务单每个货格到其它任务单货格的距离和所有复核台的距离
DistanceAll = []
distanceP2Data = pd.read_excel(distanceP2)
for i in range(len(distanceP2Data)):
    DistanceAll.append(list(distanceP2Data.iloc[i]))


num_points = 23
origin = 23
# 总距离
def get_total_distance(x):
    distance = 0
    distance += DistanceAll[origin][x[0]]
    for i in range(len(x)):
        if i == len(x) - 1:
            distance += DistanceAll[origin][x[i]]
        else:
            distance += DistanceAll[x[i]][x[i + 1]]
    return distance

# %% do GA

from sko.GA import GA_TSP

ga_tsp = GA_TSP(func=get_total_distance, n_dim=num_points, size_pop=50, max_iter=500, prob_mut=1)
best_points, best_distance = ga_tsp.run()
print(best_points)
print(best_distance)

# %% plot
fig, ax = plt.subplots(1, 2)
best_points_ = np.concatenate([best_points, [best_points[0]]])
best_points_coordinate = taskPoints[best_points_, :]
ax[0].plot(best_points_coordinate[:, 0], best_points_coordinate[:, 1], 'o-r')
ax[1].plot(ga_tsp.generation_best_Y)
plt.show()