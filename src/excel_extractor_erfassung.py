# https://openpyxl.readthedocs.io/en/stable/tutorial.html

import openpyxl

class ExcelExtractorStundenerfassung:
    def __init__(self, all_excel_files_stundenerfassung, year_index):
        self.all_excel_files_stundenerfassung = all_excel_files_stundenerfassung
        self.year_index = year_index
        # the soll index of the year
        # 5 = J
        self.year_columns = [5, 10, 15, 20]


    def read_all(self):
        for excel_file_stundenerfassung_path in self.all_excel_files_stundenerfassung:
            self.read_single_stundenerfassung(excel_file_stundenerfassung_path, self.year_index)
        # we are only allowed to write to "PLAN" sheet
        print("finished reading excel files")


    def read_single_stundenerfassung(self, excel_files_stundenerfassung_path, year_index):
        workbook = openpyxl.load_workbook(excel_files_stundenerfassung_path, read_only=True)
        ws = workbook["LehrerStundenübersicht"]

        soll_datas = []

        lfd_nr_col = 1 # A just some number
        name_col = 2 # B
        teacher_subject_pair_col = 3 # C

        start_row = 33
        soll_start_col = self.year_columns[year_index]
        curr_row = start_row

        curr_nr_cell = ws.cell(row=curr_row, column=lfd_nr_col)

        while curr_nr_cell.value is not None:
            print(curr_nr_cell.value)
            lfd_nr_value = curr_nr_cell.value
            name_value = ws.cell(row=curr_row, column=name_col).value
            teacher_subject_pair_value = ws.cell(row=curr_row, column=teacher_subject_pair_col).value
            soll_value = ws.cell(row=curr_row, column=soll_start_col).value
            # soll ws x. jahr
            ist_akkum_value = ws.cell(row=curr_row, column=soll_start_col + 2).value

            curr_row += 1
            curr_nr_cell = ws.cell(row=curr_row, column=lfd_nr_col)

            if teacher_subject_pair_value is None:
                print(f"warning: teacher subject pair is None for lfd nr {lfd_nr_value}, ignoring")
                continue

            if soll_value is None:
                print(f"warning: soll is None for lfd nr {lfd_nr_value}, ignoring")
                continue

            soll_info = {
                "lfd_nr": lfd_nr_value,
                "name": name_value,
                "teacher_subject_pair": teacher_subject_pair_value,
                "soll": soll_value,
                "ist": ist_akkum_value
            }
            soll_datas.append(soll_info)

        workbook.close()
        print(soll_datas)



if __name__ == '__main__':
    excel_extractor = ExcelExtractorStundenerfassung(
        ['example/Std-erfassung_Erz24A_1. und 2. Sj.xlsx'],
        0
    )
    excel_extractor.read_all()
