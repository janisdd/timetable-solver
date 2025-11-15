# https://openpyxl.readthedocs.io/en/stable/tutorial.html
import openpyxl
from openpyxl.cell import MergedCell

wb = openpyxl.load_workbook('example/SJ 25-26_Gesamtübersicht Einsatz Lehrkräfte EA Halle 2025-05-21_3.xlsx',
                            read_only=False)

teacher_row_start = 10

# teachers start at row 10, col B - E
# subjects start at col J, row 6
# class is in col J, row 3
#  subjects from other classes are separated by empty cells or start with (...)
# required lessons in UE for subjects are in row 9 (below subjects)


def get_excel_rect_data_as_array(ws, min_row, max_row, min_col, max_col):
    array = []
    for row in range(min_row, max_row + 1):
        array.append([])
        all_none = True
        for col in range(min_col, max_col + 1):
            data = ws.cell(row=row, column=col).value
            if data is not None:
                all_none = False
            array[-1].append(data)

        if all_none:
            # remove last row
            del array[-1]
            break

    return array


def get_cell_range_from_merged_cells(ws, cell):
    merged_cells = ws.merged_cells.ranges
    for merged_cell in merged_cells:
        if cell.coordinate in merged_cell:
            return merged_cell.bounds  # row, col, row, col


def extract_classes_with_data_from_sheet(ws):
    start_row = 3
    start_col = 10  # J
    subjects_start_row = 6
    subjects_hours_year = 7
    subjects_hours_term = 9 # for one half year (one term/semester)

    all_classes = []
    curr_class = ws.cell(row=start_row, column=start_col)  # J3

    # some classes have 1 / 2 / 3 rows ...
    # the subjects for all classes start in row 6
    while curr_class.value is not None:
        range_class_name = get_cell_range_from_merged_cells(ws, curr_class)
        class_name_lines = []

        class_name_lines.append(curr_class.value)
        # extract class data
        if range_class_name[1] == range_class_name[3]:
            # whole class in one merged cell (same row)
            # so check the next rows until we get to row 6
            class_data_2 = ws.cell(row=range_class_name[1] + 1, column=range_class_name[0])
            range_class_data_2 = get_cell_range_from_merged_cells(ws, class_data_2)
            class_name_lines.append(class_data_2.value)

            if range_class_data_2[1] == range_class_data_2[3]:
                # one row
                class_data_3 = ws.cell(row=range_class_data_2[1] + 1, column=range_class_data_2[0])
                range_class_data_3 = get_cell_range_from_merged_cells(ws, class_data_3)
                class_name_lines.append(class_data_3.value)

            else:
                # data is a merged cell with more than one row -> there is no data 3
                pass
        else:
            # data 1 is more than one row high
            if range_class_name[3] >= subjects_start_row:
                # we only have one data in a big merged cell
                raise Exception("TODO")
            else:
                # data 1 is 2 rows high, so we can have data 3
                class_data_3 = ws.cell(row=range_class_name[1] + 1, column=range_class_name[0])
                range_class_data_3 = get_cell_range_from_merged_cells(ws, class_data_3)
                class_name_lines.append(class_data_3.value)

        class_obj = {
            "name": str.join("\n", class_name_lines),
            "name_fields": class_name_lines,
            "col_range": [range_class_name[0], range_class_name[2]],
            "subjects" : [], # {"name", "col", "coord"}
            "fake_subjects": []
        }
        all_classes.append(class_obj)
        curr_class = ws.cell(row=start_row, column=range_class_name[2] + 1)

    # now we have all classes
    # extract subjects for each class
    for class_obj in all_classes:
        class_col_range = class_obj["col_range"]
        for col in range(class_col_range[0], class_col_range[1] + 1):
            subject_cell = ws.cell(row=subjects_start_row, column=col)

            # some misc cell, e.g. sum of hours or something
            if subject_cell.value is None:
                continue

            # we have fake subjects e.g. managing a class
            subject_hours_total = ws.cell(row=subjects_hours_year, column=col)
            subject_hours_term = ws.cell(row=subjects_hours_term, column=col)

            subject_obj = {
                "name": subject_cell.value,
                "col": subject_cell.column,
                "coord": subject_cell.coordinate,
                "hours_total": subject_hours_total.value,
                "hours_term": subject_hours_term.value,
                "teachers_with_hours": [] # {"teacher_key", "hours"}
            }

            if subject_hours_term.value is None:
                # not a real subject...
                class_obj["fake_subjects"].append(subject_obj)
            else:
                class_obj["subjects"].append(subject_obj)

    return all_classes

def extract_and_set_teacher_hours_from_sheet(ws, all_classes, all_teachers_list):

    for class_obj in all_classes:
        subject_objs = class_obj["subjects"]
        for subject_obj in subject_objs:
            subject_col = subject_obj["col"]

            for row_index, teacher_obj in enumerate(all_teachers_list):
                hours_for_teacher_cell = ws.cell(row=teacher_row_start + row_index, column=subject_col)

                if hours_for_teacher_cell.value is None:
                    continue

                subject_obj["teachers_with_hours"].append({
                    "teacher_key": teacher_obj["key"],
                    "hours": hours_for_teacher_cell.value
                })


def get_all_teachers_from_rect_data(teacher_datas):

    all_teachers_dict = {}
    all_teachers_list = []

    for teacher_data in teacher_datas:
        i = 0
        contract_form = teacher_data[i]
        i+=1
        last_name = teacher_data[i]
        i += 1
        first_name = teacher_data[i]
        i += 1
        teacher_key = teacher_data[i]
        teacher_obj = {
            "contract_form": contract_form,
            "last_name": last_name,
            "first_name": first_name,
            "key": teacher_key
        }
        all_teachers_dict[teacher_key] = teacher_obj
        all_teachers_list.append(teacher_obj)

    return all_teachers_list, all_teachers_dict

def extract_all_data_from_sheet(ws):
    # key is "teacher key" (or nummer in excel/german)
    # value is hash with data
    teachers = dict()

    teacher_end_row = 100
    teacher_start_col = 2
    teacher_end_col = 5
    teacher_datas = get_excel_rect_data_as_array(ws, teacher_row_start, teacher_end_row, teacher_start_col,
                                                 teacher_end_col)
    all_teachers_list, all_teachers_dict = get_all_teachers_from_rect_data(teacher_datas)
    all_classes = extract_classes_with_data_from_sheet(ws)
    extract_and_set_teacher_hours_from_sheet(ws, all_classes, all_teachers_list)

    for teacher_data in teacher_datas:
        first_name = teacher_data[0]
        last_name = teacher_data[1]
        teacher_key = teacher_data[2]
        teachers[teacher_key] = {
            "first_name": first_name,
            "last_name": last_name,
            "key": teacher_key
        }

    ws.cell(row=4, column=2)

    pass


def main():
    # for sheet in wb:
    #     print(sheet.title)

    ws = wb['KP 25_26']
    data = extract_all_data_from_sheet(ws)
    # print(ws['A4'])
    # print(ws.cell(row=4, column=2))
    # print(data)
    print("finished")


main()
