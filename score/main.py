import pandas as pd
import random

data = 'data.xlsx'
subjects = ['국어', '사회', '도덕', '수학', '과학', '체육', '미술']
subjectsDict = pd.read_excel(data, engine='openpyxl', sheet_name=None)
for i in subjectsDict:
  subjectsDict[i] = subjectsDict[i].dropna()

index = {}
finalDict = {}

for subject in subjects:
    temp = []

    for value in subjectsDict[subject].values:
        if len(value) > 0:
            temp.append(value[0])

    index[subject] = [0, ]
    for i in range(len(temp)):
        if '[' in temp[i]:
            index[subject].append(i)
    index[subject].append(len(temp))

    for i in range(len(index[subject]) - 1):
        globals()["temp{}".format(i)] = temp[index[subject][i] + 1:index[subject][i + 1]]

    finalDict[subject] = []
    for i in range(30):
        sentence = ''
        for j in range(len(index[subject]) - 1):
            sentence += random.choice(globals()["temp{}".format(j)]) + ' '
            sentence = sentence.replace('\t', '')
        finalDict[subject].append(sentence)

df = pd.DataFrame(finalDict)
df.to_excel('done.xlsx')