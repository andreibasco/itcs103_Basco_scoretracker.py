import openpyxl
from openpyxl import Workbook
import os

FILE_NAME = "student_scores.xlsx"


def get_student_input():
    name = input("Enter student name: ")
    score = float(input("Enter student score: "))
    return name, score

def create_excel_file():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.title = "Scores"
        ws.append(["Name", "Score", "Result"])
        wb.save(FILE_NAME)

def add_student_score(name, score):
    if score >=75:
        print (f"{name } Pass!")
    else:
        print(f"{name } Fail")
    result = "Pass" if score >= 75 else "Fail"
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb["Scores"]
    ws.append([name, score, result])
    wb.save(FILE_NAME)
    print(f"Added: {name} - {score} - {result}")

def main():
    create_excel_file()

    while True:
        name, score = get_student_input()
        add_student_score(name, score)
        addmore = input("Add another student? (y/n): ").lower()
        if addmore != 'y':
            break

if __name__ == "__main__":
    main()