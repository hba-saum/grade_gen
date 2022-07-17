import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import glob
from pathlib import Path

# getting the assignment file
assignmentPath = input("Path of your Assignment.csv file: ")
AssData = Path(assignmentPath)
df = pd.read_csv(AssData)
df.drop_duplicates(inplace=True)  # checking dublicates


# lab exam
labExamPath = input("Path of your Lab Exam.xlsx file: ")
labData = Path(labExamPath)
labExamDf = pd.read_excel(labData, sheet_name='Sheet1')


df['lab exam'] = 0


for i in range(len(df)):
    id = df['Student ID'][i]
    for j in range(len(labExamDf)):
        onlyID = labExamDf['Email'][j]

        if(id == onlyID[0:10]):
            df.at[i, 'lab exam'] = labExamDf['Total points'][j]


# quizes
quizesPath = input("Directory of quizes Exel files: ")
qulizCols = []
for quiz in glob.glob(quizesPath + '\*.xlsx'):
    quizData = Path(quiz)
    quizDf = pd.read_excel(quizData, sheet_name='Sheet1')

    quizColName = quiz.replace(quizesPath + "\\", '').replace('.xlsx', '')

    df[quizColName] = 0
    qulizCols.append(quizColName)

    for i in range(len(df)):
        id = df['Student ID'][i]
        for j in range(len(quizDf)):
            onlyID = quizDf['Email'][j]

            if(id == onlyID[0:10]):
                df.at[i, quizColName] = quizDf['Total points'][j]

# count best two
df['Total Quiz Mark'] = 0
for i in range(len(df)):
    allQuizNums = []
    allQuizNums.append(df[qulizCols[0]][i])
    allQuizNums.append(df[qulizCols[1]][i])
    allQuizNums.append(df[qulizCols[2]][i])

    allQuizNums.sort(reverse=True)

    df.at[i, 'Total Quiz Mark'] = allQuizNums[0] + allQuizNums[1]


# attendance
attendancesPath = input("Directory of attendances Exel files: ")
attendanceCols = []

for attendance in glob.glob(attendancesPath + '\*.csv'):
    attendanceData = Path(attendance)
    attendanceDf = pd.read_csv(attendanceData, sep='\t', encoding = 'utf-16')

    attendanceColName = attendance.replace(attendancesPath + "\\", '').replace('.csv', '')
    df[attendanceColName] = 0
    attendanceCols.append(attendanceColName)

    for i in range(len(df)):
        name = df['Name'][i]
        newName = ''
        if ', ' in name:
            newName = name.split(", ")
            name = newName[1] + " " + newName[0]


        for j in range(len(attendanceDf)):
            if (name == attendanceDf['Full Name'][j]):
                df.at[i, attendanceColName] = 1
            

# total attendance mark
df['Attendance Mark'] = 0

for i in range(len(df)):
    total = 0
    for j in range(len(attendanceCols)):
        total = total + df[attendanceCols[j]][i]

    attendancePoint = (total * 10.0)/len(attendanceCols)

    df.at[i, 'Attendance Mark'] = attendancePoint


# total marks
df['Total'] = 0
df['Grade'] = ''
gradeCnt = [0, 0, 0, 0, 0, 0, 0, 0, 0]

for i in range(len(df)):
    grade = ''
    total = df['Ass.'][i] + df['lab exam'][i] + df['Total Quiz Mark'][i] + df['Attendance Mark'][i] 
    df.at[i, 'Total'] = total

    if total >= 90:
        grade = 'A+'
        gradeCnt[0] += 1
    elif total >= 85 and total < 90:
        grade = 'A'
        gradeCnt[1] += 1
    elif total >= 80 and total < 85:
        grade = 'B+'
        gradeCnt[2] += 1
    elif total >= 75 and total < 80:
        grade = 'B'
        gradeCnt[3] += 1
    elif total >= 70 and total < 75:
        grade = 'C+'
        gradeCnt[4] += 1
    elif total >= 65 and total < 70:
        grade = 'C'
        gradeCnt[5] += 1
    elif total >= 60 and total < 65:
        grade = 'D+'
        gradeCnt[6] += 1
    elif total >= 50 and total < 60:
        grade = 'D'
        gradeCnt[7] += 1
    elif total > 50:
        grade = 'F'
        gradeCnt[8] += 1
    
    df.at[i, 'Grade'] = grade

print(df.to_string())

#Export
finalPath = input("Directory to save final exel file: ")
finalData = Path(finalPath + '\\final.csv')
df.to_csv(finalData)

#plt
x = np.array(["A+", "A", "B+", "B", "C+", "C", "D+", "D", "F"])
y = np.array(gradeCnt)

plt.bar(x,y)
plt.show()
