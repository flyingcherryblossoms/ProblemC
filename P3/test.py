minDistances = [5, 2, 9, 6, 3, 4, 7, 1]
minPaths = [5, 2, 9, 6, 3, 4, 7, 10]
ccPaths = [5, 2, 9, 6, 3, 4, 7, 4]

for indexi in range(len(minDistances)):
    minIndex = indexi
    for indexj in range(indexi+1, len(minDistances)):
        if minDistances[minIndex] > minDistances[indexj]:
            minIndex = indexj
    if minIndex != indexi:
        tempDistance = minDistances[minIndex]
        tempPath = minPaths[minIndex]
        tempccPath = ccPaths[minIndex]
        minDistances.pop(minIndex)
        minPaths.pop(minIndex)
        ccPaths.pop(minIndex)
        minDistances.insert(indexi, tempDistance)
        minPaths.insert(indexi, tempPath)
        ccPaths.insert(indexi, tempccPath)


print(minDistances)
print(minPaths)
print(ccPaths)
