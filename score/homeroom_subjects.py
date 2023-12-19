import pandas as pd
import random

data = '교과.xlsx'
subjectsDict = pd.read_excel('PreData/' + data, engine='openpyxl', sheet_name=None, header=None)
for i in subjectsDict:
    subjectsDict[i] = subjectsDict[i].dropna()

index = {}
finalDict = {}
subjects = list(subjectsDict.keys())
class_count = 8

for class_num in range(1, class_count + 1):
    for subject in subjects:
        temp = []

        for value in subjectsDict[subject].values:
            if len(value) > 0:
                temp.append(value[0])

        index[subject] = []
        for i in range(len(temp)):
            if '#' in temp[i]:
                index[subject].append(i)
        index[subject].append(len(temp))

        for i in range(len(index[subject]) - 1):
            globals()["temp{}".format(i)] = temp[index[subject][i] + 1:index[subject][i + 1]]

        finalDict[subject] = []
        students_num = 35
        for i in range(students_num):
            sentence = ''
            for j in range(len(index[subject]) - 1):
                sentence += random.choice(globals()["temp{}".format(j)]).strip() + ' '
                sentence = sentence.replace('\t', '')
            sentence = sentence.strip()
            finalDict[subject].append(sentence)

    df = pd.DataFrame(finalDict)
    df.to_excel('PostData/' + str(class_num) + '반_' + data.split('.')[0] + '_done.xlsx')
