import pandas as pd
import random

data = '창체.xlsx'
subjectsDict = pd.read_excel(data, engine='openpyxl', sheet_name=None, header=None)

for i in subjectsDict:
    subjectsDict[i] = subjectsDict[i].dropna()

subject_sentences = {}
final_with_students = {}
subjects = list(subjectsDict.keys())
sentence_num = {'진로': 1, '동아리': 1, '자율': 2}

for subject in subjects:
    subject_sentences[subject] = []

    for value in subjectsDict[subject].values:
        if len(value) > 0:
            subject_sentences[subject].append(value[0].replace('\t', '').strip())

    final_with_students[subject] = []
    students_num = 35
    for i in range(students_num):
        sentences = random.sample(subject_sentences[subject], k=sentence_num[subject])
        merged_sentence = ' '.join(sentences)
        final_with_students[subject].append(merged_sentence)

df = pd.DataFrame(final_with_students)
df.to_excel(data.split('.')[0] + 'done.xlsx')
