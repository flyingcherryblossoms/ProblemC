import pandas as pd
import glob
import os

def findExcel():
    # 读取已经计算并保存过的数据
    # path_file_number=glob.glob('D:/case/test/testcase/checkdata/*.py')#或者指定文件下个数
    path_file_number = glob.glob(pathname='./Save_Files/*.xlsx')  # 获取当前文件夹下个数
    # print(path_file_number)
    # print(len(path_file_number))
    return len(path_file_number)
def writeExcel(index, ans):
    directory = "Save_Files"
    if not os.path.exists(directory):
        os.mkdir(directory)
    pd.DataFrame(ans).to_excel(directory +"/" + str(int(index / 100)) + '.xlsx', sheet_name='Ques1', index=False)
def Manhattan(x1, y1, x2, y2):
    # 计算曼哈顿距离
    return abs(int(x1) - int(x2)) + abs(int(y1) - int(y2))

RepoFile = pd.ExcelFile("附件1：仓库数据.xlsx")
Shelf = pd.read_excel(RepoFile, "货架")
CargoContainer = pd.read_excel(RepoFile, "货格")
ReviewStation = pd.read_excel(RepoFile, "复核台")

counts = 1
ans = []
d = 750  # 拐角的边距
cclen = 800  # 货格长宽
rsflen = 1000  # 复核台长宽
lencargo = len(CargoContainer)  # 货格长宽
lenRS = len(ReviewStation)  # 复核台长款
rows = []  # 同一排编号
cols = []  # 同一列编号
for i in range(1, 8, 2):
    rows.append(list(range(i, i + 194, 8)) + list(range(i + 1, i + 195, 8)))
for i in range(1, 195, 8):
    cols.append(list(range(i, i + 8, 2)))
    cols.append(list(range(i + 1, i + 9, 2)))



# 查找已经计算的数据条数
alreadyFilesCounts = findExcel()
if alreadyFilesCounts != 0:
    counts = alreadyFilesCounts * 100 + 1


# 开始外循环
for i in range(alreadyFilesCounts * 100, lencargo):

    ansi = []
    cci = CargoContainer.iloc[i]  # i货格信息
    shelfi = int(list(cci)[5][-3:])  # i货格所属的货架

    # 计算货格i的中点以及两个拐点
    x = 0
    y = 0
    for k in range(len(Shelf)):
        if list(cci)[5] == list(Shelf.iloc[k])[0]:
            x = int(list(Shelf.iloc[k])[1])
            y = int(list(Shelf.iloc[k])[2])
            break
    # print("货格所属货架左下角坐标：", x, y)
    downXi = 0  # 上拐点
    downYi = 0
    upXi = 0  # 下拐点
    upYi = 0
    midXi = 0  # 中点
    midYi = 0
    if shelfi % 2 == 0 and shelfi > 1:
        # print("货格在右边")
        downXi = x + d + cclen
        downYi = y - d
        upXi = x + d + cclen
        upYi = y + d + cclen * 15

        midXi = int(cci[1]) + cclen
        midYi = int(cci[2]) + cclen / 2
    else:
        # print("货格在左边")
        downXi = x - d
        downYi = y - d
        upXi = x - d
        upYi = y + d + cclen * 15
        midXi = int(cci[1])
        midYi = int(cci[2]) + cclen / 2

    distanceIDown = Manhattan(midXi, midYi, downXi, downYi)  # i货格到下拐点的距离
    distanceIUp = Manhattan(midXi, midYi, upXi, upYi)  # i货格到上拐点的距离

    # 计算货格之间距离
    for j in range(lencargo):
        # if i > j:
        #     ansi.append(0)
        #     continue
        ccj = CargoContainer.iloc[j]  # j货格信息
        shelfj = int(list(ccj)[5][-3:])  # j货格所属的货架

        # print(cci[0], ccj[0])

        isSpecial = False  # 此标志用于判断是否是属于特殊情况，若是(flag为True则是)特殊情况则计算，不是特殊情况的一起考虑通用解法
        # 同一个货架
        if shelfi == shelfj:
            # print("在同一个货架")
            # 同一个货架直接曼哈顿计算曼哈顿距离+2d，为同一个则为距离为0
            if i == j:
                distance = 0
            else:
                distance = Manhattan(cci[1], cci[2], ccj[1], ccj[2]) + 2 * d
            ansi.append(distance)
            isSpecial = True
            # print("距离：", distance)
        # 不在一个货架
        else:
            # print("不在同一个货架")
            flag = False
            for k in range(4):
                if (shelfi in rows[k]) and (shelfj in rows[k]):
                    flag = True
                    break
            # 在同一排
            if flag:
                # print("在同一排")
                if abs(shelfj - shelfi) == 7:
                    # print("相邻")
                    midXj = 0  # 中点
                    midYj = 0
                    if shelfj % 2 == 0 and shelfj > 1:
                        # print("货格在右边")
                        midXj = int(ccj[1]) + cclen
                        midYj = int(ccj[2]) + cclen / 2
                    else:
                        # print("货格在左边")
                        midXj = int(ccj[1])
                        midYj = int(ccj[2]) + cclen / 2

                    distance = Manhattan(midXi, midYi, midXj, midYj)
                    ansi.append(distance)
                    isSpecial = True
                    # print("距离：", distance)
            # 不在同一排
            else:
                # print("不在同一排")
                flag = False
                for k in range(50):
                    if (shelfi in cols[k]) and (shelfj in cols[k]):
                        flag = True
                        break
                if flag:
                    # print("在同一列")
                    # 相邻货架直接曼哈顿计算曼哈顿距离 + 2d - cclen
                    distance = Manhattan(cci[1], cci[2], ccj[1], ccj[2]) + 2 * d - cclen
                    ansi.append(distance)
                    isSpecial = True
                    # print("距离：", distance)

        # 其它一般情况
        if not isSpecial:
            # print("不特殊")
            # 计算拐点
            for k in range(len(Shelf)):
                if list(ccj)[5] == list(Shelf.iloc[k])[0]:
                    x = int(list(Shelf.iloc[k])[1])
                    y = int(list(Shelf.iloc[k])[2])
                    break
            # print("货格所属货架左下角坐标：", x, y)
            downXj = 0  # 上拐点
            downYj = 0
            upXj = 0  # 下拐点
            upYj = 0
            midXj = 0  # 中点
            midYj = 0
            if shelfj % 2 == 0 and shelfj > 1:
                # print("货格在右边")
                downXj = x + d + cclen
                downYj = y - d
                upXj = x + d + cclen
                upYj = y + d + cclen * 15

                midXj = int(ccj[1]) + cclen
                midYj = int(ccj[2]) + cclen / 2
            else:
                # print("货格在左边")
                downXj = x - d
                downYj = y - d
                upXj = x - d
                upYj = y + d + cclen * 15
                midXj = int(ccj[1])
                midYj = int(ccj[2]) + cclen / 2

            distanceJDown = Manhattan(midXj, midYj, downXj, downYj)  # j货格到下货格拐点距离
            distanceJUp = Manhattan(midXj, midYj, upXj, upYj)   # j货格到上货格拐点

            distance1 = distanceIUp + distanceJUp + Manhattan(upXi, upYi, upXj, upYj)
            distance2 = distanceIUp + distanceJDown + Manhattan(upXi, upYi, downXj, downYj)
            distance3 = distanceIDown + distanceJUp + Manhattan(downXi, downYi, upXj, upYj)
            distance4 = distanceIDown + distanceJDown + Manhattan(downXi, downYi, downXj, downYj)
            distance = min(distance1, distance2, distance3, distance4)
            # print("i上拐点坐标：", upXi, upYi, "i下挂点坐标：", downXi, downYi, "i中点坐标：", midXi, midYi)
            # print("j上拐点坐标：", upXj, upYj, "j下挂点坐标：", downXj, downYj, "j中点坐标", midXj, midYj)
            # print("i到上拐点距离：", distanceIUp, "i到下拐点距离：", distanceIDown)
            # print("j到上拐点距离：", distanceJUp, "j到下拐点距离：", distanceJDown)
            # print(distance1, distance2, distance3, distance4)
            # print("最短距离：", distance)
            ansi.append(distance)

    # 计算复核台和货格距离
    for j in range(lenRS):
        rsj = list(ReviewStation.iloc[j])  # 复核台信息
        # 复核台左下角坐标
        Xj = int(rsj[1])
        Yj = int(rsj[2])
        # 右边中点
        rMidXj = Xj + rsflen
        rMidYj = Yj + rsflen / 2
        # 上面中点
        uMidXj = Xj + rsflen / 2
        uMidYj = Yj + rsflen

        # 复核台右下角拐点
        downX = Xj + d + rsflen
        downY = Yj - d
        # 复核台右上角拐点
        upX = Xj + d + rsflen
        upY = Yj + d + rsflen
        # 左上角拐点
        leftX = Xj - d
        leftY = Yj + d + rsflen
        # 右上角拐点
        rightX = Xj + d + rsflen
        rightY = Yj + d + rsflen

        RS_Turn_Lenth = d * 2 + rsflen / 2  # 复核台到拐点的距离
        # 是否在第一排
        if (shelfi in rows[0] and (rsj[0] == "FH09" or rsj[0] == "FH10" or rsj[0] == "FH11")) \
                or (shelfi in rows[1] and (rsj[0] == "FH12" or rsj[0] == "FH13")):
            # print("在第一或二排", cci[0], rsj[0])
            # 是否相邻
            if (shelfi == 1 and (rsj[0] == "FH09" or rsj[0] == "FH10" or rsj[0] == "FH11")) \
                    or (shelfi == 3 and (rsj[0] == "FH12" or rsj[0] == "FH13")):
                # print("相邻")
                distance = 0
                if int(cci[2]) > downY and int(cci[2]) < upY:
                    distance = Manhattan(int(cci[1]), int(cci[2]) + cclen / 2, rMidXj, rMidYj)
                elif int(cci[2]) < downY:
                    distance = Manhattan(int(cci[1]), int(cci[2]) + cclen / 2, downX, downY) + d * 2 + rsflen / 2
                else:
                    distance = Manhattan(int(cci[1]), int(cci[2]) + cclen / 2, upX, upY) + d * 2 + rsflen / 2
                ansi.append(distance)
                # print("距离：", distance)
            else:
                distance1 = RS_Turn_Lenth + Manhattan(downX, downY, downXi, downYi) + distanceIDown
                distance2 = RS_Turn_Lenth + Manhattan(upX, upY, downXi, downYi) + distanceIDown
                distance3 = RS_Turn_Lenth + Manhattan(downX, downY, upXi, upYi) + distanceIUp
                distance4 = RS_Turn_Lenth + Manhattan(upX, upY, upXi, upYi) + distanceIUp
                # print(distance1, distance2, distance3, distance4)
                distance = min(distance1, distance2, distance3, distance4)
                ansi.append(distance)
                # print("距离：", distance)
        elif (rsj[0] == "FH09" or rsj[0] == "FH10" or rsj[0] == "FH11" or rsj[0] == "FH12" or rsj[0] == "FH13"):
            # print("复核台在一二排，但是货格不在对应的台")
            distance1 = RS_Turn_Lenth + Manhattan(downX, downY, downXi, downYi) + distanceIDown
            distance2 = RS_Turn_Lenth + Manhattan(upX, upY, downXi, downYi) + distanceIDown
            distance3 = RS_Turn_Lenth + Manhattan(downX, downY, upXi, upYi) + distanceIUp
            distance4 = RS_Turn_Lenth + Manhattan(upX, upY, upXi, upYi) + distanceIUp
            # print(distance1, distance2, distance3, distance4)
            distance = min(distance1, distance2, distance3, distance4)
            ansi.append(distance)
            # print("距离：", distance)
        else:
            # print("不在一二排", cci[0], rsj[0])
            distance = 0
            if upXi > leftX and upXi < rightX:
                distance1 = Manhattan(uMidXj, uMidYj, downXi, downYi) + Manhattan(midXi, midYi, downXi, downYi)
                distance2 = Manhattan(uMidXj, uMidYj, upXi, upYi) + Manhattan(midXi, midYi, upXi, upYi)
                distance = min(distance1, distance2)
            else:
                distance1 = RS_Turn_Lenth + Manhattan(leftX, leftY, downXi, downYi) + distanceIDown
                distance2 = RS_Turn_Lenth + Manhattan(rightX, rightY, downXi, downYi) + distanceIDown
                distance3 = RS_Turn_Lenth + Manhattan(leftX, leftY, upXi, upYi) + distanceIUp
                distance4 = RS_Turn_Lenth + Manhattan(rightX, rightY, upXi, upYi) + distanceIUp
                # print(distance1, distance2, distance3, distance4)
                distance = min(distance1, distance2, distance3, distance4)
            ansi.append(distance)
            # print("距离：", distance)

    ans.append(ansi)
    # print(counts, len(ansi), ansi)
    print(counts)
    # 每100条数据写入一个excel表格保存
    if counts % 100 == 0 and counts >= 100:
        writeExcel(counts, ans)
        ans.clear()
    counts += 1


# 计算复核台之间的距离
ans2 = []
for i in range(lenRS):
    ansi = []
    rsi = list(ReviewStation.iloc[i])  # 复核台i信息
    for j in range(lenRS):
        rsj = list(ReviewStation.iloc[j])  # 复核台j信息
        distance = Manhattan(rsi[1], rsi[2], rsj[1], rsj[2])
        ansi.append(distance)
    ans2.append(ansi)
    # print(len(ansi), ansi)


ans.clear()
#读取存储的数据
filesNames = glob.glob(pathname='./Save_Files/*.xlsx')
for i in range(len(filesNames)):
    fileName = "./Save_Files/" + str(i+1) + ".xlsx"
    print("读取: "+fileName)
    excelFile = pd.ExcelFile(fileName)
    file = pd.read_excel(excelFile, "Ques1")
    for i in range(len(file)):
        ans.append(list(file.iloc[i]))


#合并数据
lenansj = len(ans)
lenansi = len(ans[1])
k = 0
print("开始合并数据")
for i in range(lenansi-13, lenansi):  # 列循环
    ansi = []
    for j in range(0, lenansj):  # 行循环
        ansi.append(ans[j][i])
    ansi.extend(ans2[k])
    k += 1
    # print(len(ansi), ansi)
    ans.append(ansi)
print("开始补充三角阵")
leni = len(ans)-13
for i in range(leni):
    lenj = len(ans[1])-13
    for j in range(i, lenj):
        ans[j][i] = ans[i][j]
print("补充完毕")
# 写入最终结果
print("开始写入Excel文件")

# ans3 = []
# # 添加表头
# ans3.append([])
# for i in range(lencargo):
#     cci = CargoContainer.iloc[i]
#     ans3[0].append(list(cci)[0])
# for i in range(lenRS):
#     cci = ReviewStation.iloc[i]
#     ans3[0].append(list(cci)[0])
# for i in range(len(ans)):
#     ans3.append(ans[i])
# ans = ans3
pd.DataFrame(ans).to_excel('附件4：计算结果.xlsx', sheet_name='Ques1', index=False)
print("写入完成")