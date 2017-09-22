import sys
import xlrd
import json


LESSON_TYPES = ('ПЗ', 'Л', 'ЛР')


def isDivided(col):
    return not (col[0].value in LESSON_TYPES and len(col[2].value) > 0
                and col[2].value not in LESSON_TYPES)


def parse_splitted(sheet, start_row):
    lessons = []

    order = sheet.cell(start_row + 2, 1).value
    for i in range(0, 4, 2):
        lesson = {}
        lesson['order'] = order
        lesson['name'] = sheet.cell(start_row + i, 2).value
        lesson['teacher'] = sheet.cell(start_row + i + 1, 2).value
        lesson['type'] = sheet.cell(start_row + i, 3).value
        lesson['room'] = sheet.cell(start_row + i + 1, 3).value
        lessons.append(lesson)

    return lessons


def parse_lesson(sheet, start_row):
    lesson = {}

    if not isDivided(sheet.col(3, start_row, start_row + 3)):
        lesson['order'] = sheet.cell(start_row + 2, 1).value
        lesson['name'] = sheet.cell(start_row, 2).value
        lesson['teacher'] = sheet.cell(start_row + 2, 2).value
        lesson['type'] = sheet.cell(start_row, 3).value
        lesson['room'] = sheet.cell(start_row + 2, 3).value
    else:
        return parse_splitted(sheet, start_row)

    return lesson


def parse(infile, outfile):
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
            day.append(parse_lesson(sheet, j))

        data.append(day)
    json_data = json.dumps(data, indent=1)
    with open(outfile, 'w', encoding='UTF-8') as f:
        f.write(json_data)

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

    parse(infile, outfile)


if __name__ == "__main__":
    main()
