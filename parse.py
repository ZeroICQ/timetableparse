import sys
import xlrd
import json


LESSON_TYPES = ('ПЗ', 'Л', 'ЛР')

def isDivided(col):
    a = 1
    return col[0].value in LESSON_TYPES and len(col[2].value) == 0 \
        or col[2].value in LESSON_TYPES \


def parseLesson(sheet, start_row):
    lesson = {
    }
    lesson['order'] = sheet.cell(start_row + 2, 1).value
    if not isDivided(sheet.col(3, start_row, start_row + 3)):
        lesson['name'] = sheet.cell(start_row, 2).value
        print(lesson)

    return lesson


def parse(infile):
    UPPER_LEFT_CELL = (14, 1)
    LOWER_RIGHT_CELL = (UPPER_LEFT_CELL[0] + 167, UPPER_LEFT_CELL[1] + 2)

    wb = xlrd.open_workbook(infile)
    sheet = wb.sheet_by_index(0)

    j = UPPER_LEFT_CELL[1]
    data = []

    # 28 - step for one day in cells
    for i in range(UPPER_LEFT_CELL[0], LOWER_RIGHT_CELL[0], 28):
        day = []
        for j in range(i, i + 28, 4):
            day.append(parseLesson(sheet, j))

        data.append(day)

    print(data)


def main():
    infile = None
    outfile = None

    if len(sys.argv) > 1:
        infile = sys.argv[1]

    if len(sys.argv) > 2:
        outfile = sys.argv[2]

    if not infile or not outfile:
        print("Input infile and outfile")
        return

    parse(infile)


if __name__ == "__main__":
    main()
