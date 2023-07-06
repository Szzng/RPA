import pandas as pd
import random

data = 'data.xlsx'
subjectsDict = pd.read_excel(data, engine='openpyxl', sheet_name=None, header=None)
for i in subjectsDict:
  subjectsDict[i] = subjectsDict[i].dropna()

index = {}
finalDict = {}
subjects = list(subjectsDict.keys())

for subject in subjects:
    for num in range(1, 8):
      temp = []

      for value in subjectsDict[subject].values:
          if len(value) > 0:
              temp.append(value[0])

      index[subject] = []
      for i in range(len(temp)):
          if '[' in temp[i]:
              index[subject].append(i)
      index[subject].append(len(temp))

      for i in range(len(index[subject]) - 1):
          globals()["temp{}".format(i)] = temp[index[subject][i] + 1:index[subject][i + 1]]

      finalDict[num] = []
      students_num= 35
      for i in range(students_num):
          sentence = ''
          for j in range(len(index[subject]) - 1):
              sentence += random.choice(globals()["temp{}".format(j)]).strip() + ' '
              sentence = sentence.replace('\t', '')
          sentence = sentence.strip()
          finalDict[num].append(sentence)

    df = pd.DataFrame(finalDict)
    df.to_excel(f'final{subject}.xlsx')