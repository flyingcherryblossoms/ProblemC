import random
import matplotlib.pyplot as plt


class GATSP:
    def __init__(self, point_count=0, pop_count=100, improve_count=2000, itter_time=500, retain_rate=0.3,
                 random_select_rate=0.5, mutation_rate=0.1, origin=0, index=None, taskList=None, taskPoint=None,
                 Distance=None):
        # 点数
        self.point_count = point_count
        # 种群数
        self.pop_count = pop_count
        # 改良次数
        self.improve_count = improve_count
        # 进化次数
        self.itter_time = itter_time
        # 设置强者的定义概率，即种群前30%为强者
        self.retain_rate = retain_rate
        # 设置弱者的存活概率
        self.random_select_rate = random_select_rate
        # 变异率
        self.mutation_rate = mutation_rate
        # 设置起点
        self.origin = origin
        # index是初始的坐标的个数按0到index排序
        self.index = index
        # 任务单信息
        self.taskList = taskList
        # 任务点位置
        self.taskPoint = taskPoint
        # 距离矩阵
        self.Distance = Distance
        # 最短路径
        self.minPath = []
        # 最短路径长度
        self.minDistance = 9999999999
        # 最短的那次的迭代收敛过程的路径长
        self.minRegister = []
        # 最短路径的地点名称
        self.ccPath = []

    # 总距离
    def get_total_distance(self, x):
        distance = 0
        distance += self.Distance[self.origin][x[0]]
        for i in range(len(x)):
            if i == len(x) - 1:
                distance += self.Distance[self.origin][x[i]]
            else:
                distance += self.Distance[x[i]][x[i + 1]]
        return distance

    # 改良圈子算法
    def improve(self, x):
        indexi = 0
        distance = self.get_total_distance(x)
        while indexi < self.improve_count:
            # randint [a,b]
            u = random.randint(0, len(x) - 1)
            v = random.randint(0, len(x) - 1)
            if u != v:
                new_x = x.copy()
                t = new_x[u]
                new_x[u] = new_x[v]
                new_x[v] = t
                new_distance = self.get_total_distance(new_x)
                if new_distance < distance:
                    distance = new_distance
                    x = new_x.copy()
            else:
                continue
            indexi += 1
        return x

    # 自然选择
    def selection(self, population):
        """
        选择
        先对适应度从大到小排序，选出存活的染色体
        再进行随机选择，选出适应度虽然小，但是幸存下来的个体
        """
        # 对总距离从小到大进行排序
        graded = [[self.get_total_distance(x), x] for x in population]
        graded = [x[1] for x in sorted(graded)]
        # 选出适应性强的染色体
        retain_length = int(len(graded) * self.retain_rate)
        parents = graded[:retain_length]
        # 选出适应性不强，但是幸存的染色体
        for chromosome in graded[retain_length:]:
            if random.random() < self.random_select_rate:
                parents.append(chromosome)
        return parents

    # 交叉繁殖
    def crossover(self, parents):
        # 生成子代的个数,以此保证种群稳定
        target_count = self.pop_count - len(parents)
        # 孩子列表
        children = []
        while len(children) < target_count:
            male_index = random.randint(0, len(parents) - 1)
            female_index = random.randint(0, len(parents) - 1)
            if male_index != female_index:
                male = parents[male_index]
                female = parents[female_index]

                left = random.randint(0, len(male) - 2)
                right = random.randint(left + 1, len(male) - 1)

                # 交叉片段
                gene1 = male[left:right]
                gene2 = female[left:right]

                child1_c = male[right:] + male[:right]
                child2_c = female[right:] + female[:right]
                child1 = child1_c.copy()
                child2 = child2_c.copy()

                for o in gene2:
                    child1_c.remove(o)

                for o in gene1:
                    child2_c.remove(o)

                child1[left:right] = gene2
                child2[left:right] = gene1

                child1[right:] = child1_c[0:len(child1) - right]
                child1[:left] = child1_c[len(child1) - right:]

                child2[right:] = child2_c[0:len(child1) - right]
                child2[:left] = child2_c[len(child1) - right:]

                children.append(child1)
                children.append(child2)

        return children

    # 变异
    def mutation(self, children):
        for i in range(len(children)):
            if random.random() < self.mutation_rate:
                child = children[i]
                u = random.randint(1, len(child) - 4)
                v = random.randint(u + 1, len(child) - 3)
                w = random.randint(v + 1, len(child) - 2)
                child = children[i]
                child = child[0:u] + child[v:w] + child[u:v] + child[w:]
        return children

    # 得到最佳纯输出结果
    def get_result(self, population):
        graded = [[self.get_total_distance(x), x] for x in population]
        graded = sorted(graded)
        return graded[0][0], graded[0][1]

    # 总距离
    def get_distance(self, x):
        distance = 0
        for k in range(len(x) - 1):
            distance += self.Distance[x[k]][x[k + 1]]
        return distance

    # 绘图
    def draw(self, Path, Register=None):
        X = []
        Y = []
        # plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
        plt.rcParams.update({'font.size': 8})
        for i in Path:
            x = self.taskPoint[i, 0]
            y = self.taskPoint[i, 1]
            X.append(x)
            Y.append(y)
            plt.annotate(self.taskList[i][2], xy=(x, y), xytext=(x + 600, y))
        plt.plot(X, Y, '-o')
        plt.show()
        if Register is not None:
            plt.plot(list(range(len(Register))), Register)
            plt.show()

    # 运行
    def run(self, times=1, RS=None):
        self.minDistance = 99999999999
        for timei in range(times):
            # 使用改良圈算法初始化种群
            population = []
            for i in range(self.pop_count):
                # 随机生成个体
                x = self.index.copy()
                random.shuffle(x)
                x = self.improve(x)
                population.append(x)

            register = []  # 记录遗传过程的每一组序列的总长度变化
            i = 0
            distance, result_path = self.get_result(population)
            while i < self.itter_time:
                # 选择繁殖个体群
                parents = self.selection(population)
                # 交叉繁殖
                children = self.crossover(parents)
                # 变异操作
                children = self.mutation(children)
                # 更新种群
                population = parents + children

                distance, result_path = self.get_result(population)
                register.append(distance)
                i = i + 1

            # print("当前路径：", distance, result_path)

            result_path = [self.origin] + result_path + [self.origin]
            self.draw(result_path, register)
            ccPath = []  # 记录货格名称
            if self.minDistance > distance:
                self.minDistance = distance
                self.minPath = result_path
                self.minRegister = register
                for j in range(len(self.minPath)):
                    ccPath.append(self.taskList[self.minPath[j]][2])
                self.ccPath = ccPath

            tempPath = self.minPath.copy()  # 算出的最小回路
            tempPathReverse = self.minPath.copy()  # 首尾倒置的最小回路
            tempPathReverse.reverse()
            for rs in RS:  # 去掉最后一个点（终点，或者起点，因为是同一个点，所以要倒过来算一次），环取最短路径，一共13个复核台，26种情况
                ccPath = []  # 记录货格名称
                tempPath.pop()  # 去尾
                tempPath.append(rs + self.point_count)  # 加复核台
                tempPathReverse.pop()  # 去尾
                tempPathReverse.append(rs + self.point_count)  # 加复核台
                pathlenth1 = self.get_distance(tempPath)
                pathlenth2 = self.get_distance(tempPathReverse)
                # print(pathlenth1, pathlenth2, 'FH', i + 1)
                if pathlenth1 < pathlenth2:
                    if self.minDistance > pathlenth1:
                        self.minDistance = pathlenth1
                        self.minPath = tempPath.copy()
                        for j in range(len(self.minPath)):
                            ccPath.append(self.taskList[self.minPath[j]][2])
                        self.ccPath = ccPath
                        # print("当前minDistance", self.minDistance, "当前minPath：", self.minPath)
                else:
                    if self.minDistance > pathlenth2:
                        self.minDistance = pathlenth2
                        self.minPath = tempPathReverse.copy()
                        for j in range(len(self.minPath)):
                            ccPath.append(self.taskList[self.minPath[j]][2])
                        self.ccPath = ccPath
                        # print("当前minDistance", self.minDistance, "当前minPath：", self.minPath)
        return self.minDistance, self.minPath, self.ccPath, self.minRegister
