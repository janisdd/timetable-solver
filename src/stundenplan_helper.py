from pulp import *
from tabulate import tabulate

from src.excel_extractor_gesamtuebersicht import ExcelExtractorGesamtuebersicht
from src.excel_extractor_stunden_erfassung import ExcelExtractorStundenerfassung
from src.logger import Logger


# TODO blocks (2 slots) is respected in existing teacher's timetables & existing class's timetables?

class StundenplanHelper:
    # the lp problem
    problem = None
    # all lp vars
    all_vars = None
    all_var_names = None
    max_days = -1
    max_slots_per_day = -1
    all_class_timetables_tuples = None

    def __init__(self, excel_extractor, excel_stundenerfassung):
        self.log_name = "StundenplanHelper"
        self.excel_extractor = excel_extractor
        self.excel_stundenerfassung = excel_stundenerfassung

    def read_excel_data(self):
        self.excel_extractor.read_all()

    def read_excel_stundenerfassung(self):
        self.excel_stundenerfassung.read_all()

    def prepare_excel_data(self):
        self.excel_stundenerfassung.add_soll_data_to_class_subject_teachers(self.excel_extractor.all_classes, self.excel_extractor.all_teachers_list)

    def _excel_data_sanity_checks(self):
        # e.g. our vars use _ as a separator
        # make sure no data includes this!
        pass

    def init_new_timetable_problem(self):
        # Create the 'prob' variable to contain the problem data
        self.problem = LpProblem("Stundenplan", LpMinimize)

        self.max_days = self.excel_extractor.get_max_days()
        self.max_slots_per_day = self.excel_extractor.get_max_slots_per_day()

        all_var_names = create_all_var_names(
            self.excel_extractor.all_teachers_list,
            self.excel_extractor.all_classes,
            self.max_days,
            self.max_slots_per_day,
        )
        self.all_var_names = all_var_names
        self.all_vars = LpVariable.dicts("var", all_var_names, cat="Binary")

    def setup_fixed_vars(self):
        print('setup_fixed_vars')
        # some vars are fixed/known
        # - some classes slots/days should be skipped -> set all corresponding vars to 0 (see 'PLAN' sheet)
        #   (ignored days just set all slots to ignored)
        # - some classes slots are already filled -> set all corresponding vars to 0 (see 'PLAN' sheet) # TODO 0 -> not set
        # - some teachers have white/black lists for slots
        #   - convert all to black list (so we can set vars to 0) (because 1 would mean we know the solution)

        Logger.log(f"[{self.log_name}][FIXED Setup] setting up fixed variables based on pre-filled values and skip flag")

        c_count = 0
        for class_obj in self.excel_extractor.all_classes:
            # one list for each day
            table_dates = class_obj['table_dates']
            class_key = class_obj['key']

            # [0] are just the heads (dates), [2] is the start coord (start_col, start_row)
            table_body = table_dates[1]

            for day_index, table_info_list in enumerate(table_body):
                for slot_index, info_obj in enumerate(table_info_list):
                    filled_value = info_obj['entry']  # e.g. LF5(Gru) -> teacher already teachers subject here
                    ignore = info_obj['ignore']

                    should_not_process = (filled_value is not None and filled_value != '') or ignore == True

                    if should_not_process:
                        var_names = get_all_vars_with_preset(self.all_var_names, class_key=class_key,
                                                             day_index=day_index, slot_index=slot_index)

                        # for a prefilled value, we just set our variable to 0, so that we don't place something else there
                        if ignore:
                            Logger.log(
                                f"[{self.log_name}][FIXED Setup] skipping day '{day_index}', slot '{slot_index}' for class '{class_key}' because of ignore flag (prefilled value: {filled_value})")
                        else:
                            Logger.log(
                                f"[{self.log_name}][FIXED Setup] skipping day '{day_index}', slot '{slot_index}' for class '{class_key}' because of non-empty prefilled value: {filled_value}")

                        if len(var_names) == 0:
                            print("TODO TODO")

                        var = [self.all_vars[var_name] for var_name in var_names]
                        # self.problem += lpSum(var) == 0, f"skip prefilled slots in plan {var_names}, {day_index}, {slot_index}"
                        self.problem += lpSum(
                            var) == 0, f"skip prefilled/skipped slots in plan, {day_index}, {slot_index} {c_count}"
                        c_count += 1


        Logger.log(f"[{self.log_name}][FIXED Setup] setting up fixed variables based on teachers' availability preferences")

        c_count = 0
        for i, teacher_obj in enumerate(self.excel_extractor.all_teachers_list):
            teacher_key = teacher_obj["key"]
            availability_preference_table = teacher_obj['availability_preference_table']

            for day_index, day_slots_list in enumerate(availability_preference_table):
                for slot_index, slot_obj in enumerate(day_slots_list):
                    class_key = slot_obj['class_key']
                    allowed = slot_obj['allowed']
                    has_fixed_class = class_key is not None

                    if not allowed or has_fixed_class:
                        # teacher has some fixed class in this slot...
                        # not allowed -> blacklisted, teach cannot do any subject in this slot
                        # for us that just means, set all related vars to 0, so we don't place anything else there

                        var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key,
                                                             day_index=day_index, slot_index=slot_index)

                        Logger.log(f"[{self.log_name}][FIXED Setup] skip fixed teacher day '{day_index}', slots '{slot_index}' for teacher '{teacher_key}' because blacklist (class: '{class_key}')")

                        if len(var_names) == 0:
                            print("TODO TODO")

                        var = [self.all_vars[var_name] for var_name in var_names]
                        # self.problem += lpSum(var) == 0, f"skip fixed teacher slots {var_names}, {day_index}, {slot_index}"
                        self.problem += lpSum(
                            var) == 0, f"skip fixed teacher slots, {day_index}, {slot_index} {c_count}"
                        c_count += 1

                        # because of this, the class has no other option too
                        # make sure we don't set the class to some other subject on that slot

                        if has_fixed_class:
                            var_names_class = get_all_vars_with_preset(self.all_var_names, class_key=class_key,
                                                                       day_index=day_index, slot_index=slot_index)

                            Logger.log(
                                f"[{self.log_name}][FIXED Setup] skip fixed class day '{day_index}', slot '{slot_index}' class: '{class_key}' because prefilled class (teacher '{teacher_key}')")

                            if len(var_names_class) == 0:
                                print("TODO TODO")

                            var_class = [self.all_vars[var_name] for var_name in var_names_class]
                            # self.problem += lpSum(var_class) == 0, f"skip teacher slots for classes {var_names_class}, {day_index}, {slot_index} {c_count}"
                            self.problem += lpSum(
                                var_class) == 0, f"skip teacher slots for classes, {day_index}, {slot_index} {c_count}"
                            c_count += 1

    def setup_constraints(self):

        # every day can only have the max number of slots for one class
        # monday-1 has max 4 slots, so 5a can only have 4 slots (and any other class too)
        # [teacher]_[subject]_montag-1_slot_1_5a + ... <= 4
        c_count = 0

        Logger.log(f"[{self.log_name}][Constraint Setup]: max slots per day: {self.max_slots_per_day}")
        for day_index in range(self.max_days):
            for class_obj in self.excel_extractor.all_classes:
                class_key = class_obj['key']
                var_names = get_all_vars_with_preset(self.all_var_names, day_index=day_index, class_key=class_key)

                if len(var_names) == 0:
                    print("TODO TODO")

                var = [self.all_vars[var_name] for var_name in var_names]
                c1 = lpSum(
                    var) <= self.max_slots_per_day, f"every day can only have a max number of slots for one class {c_count}"
                c_count += 1
                self.problem += c1

        # every teacher can only teach the max number of slots per day
        # day 4 slots -> teacher can only teacher 4
        # l1_[subject]_montag-1_1_[class] + ... <= 4
        c_count = 0
        Logger.log(f"[{self.log_name}][Constraint Setup]: max slots per teacher per day")
        for day_index in range(self.max_days):
            for teacher_obj in self.excel_extractor.all_teachers_list:
                teacher_key = teacher_obj["key"]
                var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key, day_index=day_index)

                if len(var_names) == 0:
                    print("TODO TODO")

                var = [self.all_vars[var_name] for var_name in var_names]
                c1 = lpSum(
                    var) <= self.max_slots_per_day, f"every teacher can only teach the max number of slots per day {c_count}"
                c_count += 1
                self.problem += c1

        # one teacher can only teach max one subject per (any) class at one slot
        # e.g. l1 can only teach deutsch to 5a at montag-1 in slot 1 (but not to 5b in the same slot)
        # l1_[subject]_montag-1_1_[class] + ... = 1
        c_count = 0
        Logger.log(f"[{self.log_name}][Constraint Setup]: max one subject per class at one slot")
        for teacher_obj in self.excel_extractor.all_teachers_list:
            teacher_key = teacher_obj["key"]
            for day_index in range(self.max_days):
                for slot_index in range(self.max_slots_per_day):
                    var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key,
                                                         slot_index=slot_index,
                                                         day_index=day_index)

                    if len(var_names) == 0:
                        print("TODO TODO")

                    var = [self.all_vars[var_name] for var_name in var_names]
                    c1 = lpSum(var) <= 1, f"teacher can only teach max one subject per class at one slot {c_count}"
                    c_count += 1
                    self.problem += c1

        # every class can only have one subject with one (any) teacher at a specific slot
        # e.g. 5a at montag-1 in slot 1 can only have l1 teaching deutsch
        # [teacher]_[subject]_montag-1_1_5a + ... = 1
        c_count = 0
        Logger.log(f"[{self.log_name}][Constraint Setup]: every class can only have one subject with one teacher at a specific slot")
        for class_obj in self.excel_extractor.all_classes:
            class_key = class_obj['key']
            for day_index in range(self.max_days):
                for slot_index in range(self.max_slots_per_day):
                    var_names = get_all_vars_with_preset(self.all_var_names, class_key=class_key, slot_index=slot_index,
                                                         day_index=day_index)
                    var = [self.all_vars[var_name] for var_name in var_names]
                    c1 = lpSum(
                        var) <= 1, f"every class can only have one subject with one teacher at a specific slot {c_count}"
                    c_count += 1
                    self.problem += c1

    def solve_timetable_problem_2(self):

        c_count = 0
        all_rest_varialbes = []
        all_rest_varialbes_infos = []
        big_sum_SOLL_hours_term = 0
        for class_obj in self.excel_extractor.all_classes:
            class_key = class_obj['key']
            subjects_info = class_obj["subjects"]
            for subject_info in subjects_info:
                subject_name = subject_info["name"]

                # SOLL for the subject e.g. 80
                SOLL_hours_term = subject_info["hours_term"]
                # this is our 'mde' (missing deu)
                MISSING_hours_term = SOLL_hours_term

                Logger.log(f"[{self.log_name}][OBJ] class '{class_key}' needs '{SOLL_hours_term}' times subject '{subject_name}'")

                sum_IST_hours_teacher = 0
                sum_SOLL_hours_term_teachers = 0
                # check all connected teachers
                for teacher_tracking_hours_obj in self.excel_stundenerfassung.all_soll_data_dict[class_key]:
                    tracking_subject_name = teacher_tracking_hours_obj['subject_name']
                    if subject_name == tracking_subject_name:
                        teacher_key = teacher_tracking_hours_obj['teacher_key']
                        hours_ist = teacher_tracking_hours_obj['ist']
                        hours_soll = teacher_tracking_hours_obj['soll']

                        sum_SOLL_hours_term_teachers += hours_soll
                        sum_IST_hours_teacher += hours_ist

                        Logger.debug(
                            f"[{self.log_name}][OBJ] teacher key '{teacher_key}' has IST hours '{hours_ist}' and SOLL hours '{hours_soll}'")

                if sum_SOLL_hours_term_teachers != SOLL_hours_term:
                    Logger.warn(
                        f"[{self.log_name}][OBJ] class '{class_key}' subject '{subject_name}' has SOLL hours term '{SOLL_hours_term}' but sum of SOLL hours from teachers is '{sum_SOLL_hours_term_teachers}', using teachers' SOLL hours '{sum_SOLL_hours_term_teachers}'!")
                    SOLL_hours_term = sum_SOLL_hours_term_teachers

                MISSING_hours_term = SOLL_hours_term - sum_IST_hours_teacher
                needs_hours = MISSING_hours_term > 0

                big_sum_SOLL_hours_term += SOLL_hours_term

                if needs_hours:
                    Logger.log(
                        f"[{self.log_name}][OBJ] class '{class_key}' misses '{MISSING_hours_term}' times subject '{subject_name}' ({sum_IST_hours_teacher}/{SOLL_hours_term})")
                else:
                    Logger.log(
                        f"[{self.log_name}][OBJ] class '{class_key}' has '{-MISSING_hours_term}' times subject '{subject_name}' TOO MUCH ({sum_IST_hours_teacher}/{SOLL_hours_term}), ignoring subject for that class")
                    continue

                var_names = get_all_vars_with_preset(self.all_var_names, class_key=class_key, subject_key=subject_name)
                var = [-self.all_vars[var_name] for var_name in var_names]

                # slack variable for how many ours are still to be filled
                # mde - sd1 - sd2 - sd3 - rde >= 0;
                # mde - sd1 - sd2 - sd3 - rde <= 0;
                # sd1, sd2, ... , are our vars, rde is the new var
                rde = pulp.LpVariable(f"r_{class_key}_{subject_name}", lowBound=0)

                # this is the same as ... == 0
                constraint1 = MISSING_hours_term + lpSum(var) - rde >= 0
                constraint2 = MISSING_hours_term + lpSum(var) - rde <= 0
                self.problem += constraint1, f"res variable for '{class_key}', subject '{subject_name}' >= 0, {c_count}"
                c_count += 1
                self.problem += constraint2, f"res variable for '{class_key}', subject '{subject_name}' <= 0, {c_count}"
                c_count += 1

                all_rest_varialbes.append(rde)
                all_rest_varialbes_infos.append([class_key, subject_name, rde])

        # every r_ represents the new missing hours (the hours we are still missing with the new solution)
        # we try to minimize the sum of them, so every difference var should be enough on its own to overrule them
        safe_diff_var_multiplier = big_sum_SOLL_hours_term + 1

        # add all abs difference varialbes
        # e.g. deu, eng -> r_deu, r_eng
        # bd_deu_eng + r_deu - r_eng >= 0
        # bd_deu_eng - r_deu + r_eng >= 0

        # bd_deu_eng  >= r_eng - r_deu
        # bd_deu_eng  >= r_deu - r_eng

        # e.g. r_deu = 4, r_eng = 2 --> bd_deu_eng = 2 (this always gives us the difference)
        # now do that for all combinations...
        all_diff_vars = []
        all_diff_constraint_names = []
        c_count = 0
        for i in range(len(all_rest_varialbes_infos)):
            for j in range(i+1, len(all_rest_varialbes_infos)):
                deu_info = all_rest_varialbes_infos[i]
                eng_info = all_rest_varialbes_infos[j]

                deu_class = deu_info[0]
                deu_subj = deu_info[1]
                r_deu = deu_info[2]

                eng_class = eng_info[0]
                eng_subj = eng_info[1]
                r_eng = eng_info[2]

                bd_deu_eng = pulp.LpVariable(f"bd_{deu_class}_{deu_subj}_{eng_class}_{eng_subj}", lowBound=0)
                all_diff_vars.append(bd_deu_eng * safe_diff_var_multiplier)

                diff_constraint_1 = bd_deu_eng + r_deu - r_eng >= 0
                diff_constraint_2 = bd_deu_eng - r_deu + r_eng >= 0
                # names can get too long and are then truncated...
                # name_1 = f"difference for class '{deu_class}', subject '{deu_subj}' with class {eng_class}, subject {eng_subj}, {c_count}"
                name_1 = f"difference for class constraint, {c_count}"
                self.problem += diff_constraint_1, name_1
                c_count += 1
                # name_2 = f"difference for class '{deu_class}', subject '{deu_subj}' with class {eng_class}, subject {eng_subj}, {c_count}"
                name_2 = f"difference for class constraint, {c_count}"
                self.problem += diff_constraint_2, name_2
                c_count += 1

                all_diff_constraint_names.append(name_1.replace(" ", "_"))
                all_diff_constraint_names.append(name_2.replace(" ", "_"))


        self.problem += lpSum(all_rest_varialbes) + lpSum(all_diff_vars)
        self.problem.solve()

        # The status of the solution is printed to the screen
        print("Status:", LpStatus[self.problem.status])

        # because we try to minimize the differences
        #   it can happen that we don't proceed further because the missing values are the same
        #   then reducing it further would make the obj value worse again
        # so, a second run to place entries after that should work to fill the missing gaps 

        # for constraint_name in all_diff_constraint_names:
        #     del self.problem.constraints[constraint_name]
        #
        # self.problem.solve()

        # The status of the solution is printed to the screen
        # print("Status:", LpStatus[self.problem.status])

        if self.problem.status != 1:
            print("No solution found")
            exit()

        # Each of the variables is printed with it's resolved optimum value
        for v in self.problem.variables():
            if v.name.startswith("r_"):
                print(v.name, "=", v.varValue)
            if v.name.startswith("var") and v.varValue == 1:
                print(v.name, "=", v.varValue)

        print("---------------")

        all_class_timetables_tuples = get_stundenplan_tuples_from_vars(self)
        self.write_timetable_solution_to_excel(all_class_timetables_tuples, 'example_real/OUT.xlsm')


    def solve_timetable_problem(self):

        all_slack_vars_for_class_lesson_requirements = []

        # TODO test all teachers should have 10 lessons / slots
        # required_slots = 10
        # for index, teacher_obj in enumerate(self.excel_extractor.all_teachers_list):
        #     teacher_key = teacher_obj["key"]
        #
        #     s1 = pulp.LpVariable(f"{teacher_key}_missing_required_lessons", lowBound=0) # Slack variable
        #     all_slack_vars_for_class_lesson_requirements.append(s1)
        #
        #     var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key)
        #     var = [self.all_vars[var_name] for var_name in var_names]
        #
        #     # TODO ist akkum + slack = soll --> welches jahr???
        #     # dazu brauchen wir alle stundenerfassungen...
        #     # außerdem haben wir nur name vom lehrer + abkürzung D/K(Hmue) und Erz24A
        #     # finden wir dadurch das fach?
        #     c1 = lpSum(var) + s1 == required_slots, f"teacher should get {required_slots}, {index}"
        #     self.problem += c1

        # optimization function
        _opt_index = 0
        opt_objs_dict = {}
        for class_obj in self.excel_extractor.all_classes:
            class_key = class_obj['key']
            subjects_info = class_obj["subjects"]
            for subject_info in subjects_info:
                subject_name = subject_info["name"]
                teachers_with_hours = subject_info["teachers_with_hours"]
                for teachers_with_hour_tuple in teachers_with_hours:
                    teacher_key = teachers_with_hour_tuple['teacher_key']

                    if 'ist' not in teachers_with_hour_tuple or 'soll' not in teachers_with_hour_tuple:
                        print(f"WARNING: missing ist/soll for teacher '{teacher_key} 'in class '{class_key}' for subject '{subject_name}' -> skipping")
                        continue

                    hours_ist = teachers_with_hour_tuple['ist']
                    hours_soll = teachers_with_hour_tuple['soll']

                    opt_obj = {
                        "class_key": class_key,
                        "subject_name": subject_name,
                        "teacher_key": teacher_key,
                        "hours_ist": hours_ist,
                        "hours_soll": hours_soll,
                    }
                    hours_missing = hours_soll - hours_ist

                    if hours_missing < 0:
                        print(f"[INFO] {str(hours_missing)} hours missing < 0 for teacher {teacher_key} in class {class_key} in subject '{subject_name}' -> skipping")
                        continue

                    s1 = pulp.LpVariable(f"{class_key}_{teacher_key}_{subject_name}", lowBound=0) # Slack variable
                    all_slack_vars_for_class_lesson_requirements.append(s1)

                    opt_objs_dict[s1.name] = opt_obj

                    var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key, subject_key=subject_name, class_key=class_key)
                    var = [self.all_vars[var_name] for var_name in var_names]

                    c1 = lpSum(var) + s1 == hours_missing, f"{teacher_key}_{subject_name}_lessons, {_opt_index}"
                    _opt_index += 1
                    self.problem += c1

        print('solve_timetable_problem')
        # dummy
        self.problem += lpSum(all_slack_vars_for_class_lesson_requirements)
        # self.problem.writeLP("StundenplanReal.lp", max_length=1000)
        self.problem.solve()

        # The status of the solution is printed to the screen
        print("Status:", LpStatus[self.problem.status])

        if self.problem.status != 1:
            print("No solution found")
            exit()

        # Each of the variables is printed with it's resolved optimum value
        for v in self.problem.variables():
            if v.varValue == 1:
                print(v.name, "=", v.varValue)

        for slack_var in all_slack_vars_for_class_lesson_requirements:
            if slack_var.varValue > 0:
                opt_obj = opt_objs_dict[slack_var.name]
                class_key = opt_obj["class_key"]
                teacher_key = opt_obj["teacher_key"]
                subject_name = opt_obj["subject_name"]
                hours_ist = opt_obj["hours_ist"]
                hours_soll = opt_obj["hours_soll"]

                subject_key = self.excel_extractor.subject_name_to_key_dict[subject_name]
                hours_diff = hours_soll - hours_ist

                # teacher_Key = get_slack_var_parts__teacher_offered_lessons_total(slack_var.name)
                # offered_lessons = self.excel_extractor.all_teachers_dict[teacher_Key]["offered_lessons"]
                # print("WARNING: Slack variable > 0, not all constraints could be satisfied!")
                print(
                    f"In Klasse '{class_key}' in Fach '{subject_key}' mit Lehrer {teacher_key} hatte vorher {hours_ist}/{hours_soll} Stunden, jetzt {hours_ist + hours_diff - slack_var.varValue}/{hours_soll}, verteilt: {hours_diff - slack_var.varValue}/{hours_diff}")

        print("finished")
        self.all_class_timetables_tuples = get_stundenplan_tuples_from_vars(self)

    # clone the file with PLAN
    def write_timetable_solution_to_excel(self, all_class_timetables_tuples, output_file_path):
        print(f"writing timetable solution to excel file: {output_file_path}")
        self.excel_extractor.write_timetable_solution_to_excel_impl(output_file_path, all_class_timetables_tuples)

def get_slack_var_parts__teacher_offered_lessons_total(var_name):
    var_parts = var_name.split("_")
    return var_parts[0]

# create binary variables
# e.g. jan_math_0_slot_0_ERZ24.
# [teacher]_[subject]_[day]_slot_[lesson hour slot index]_[class]
def _get_var_name(teacher_key, subject_key, day_index, lesson_hour_slot_index, class_key):
    return f"{teacher_key}_{subject_key}_{day_index}_slot_{lesson_hour_slot_index}_{class_key}"


# [teacher]_[subject]_[day]_slot_[lesson hour slot index]_[class]
def _get_var_parts(var_name):
    var_parts = var_name.split("_")
    return var_parts[0], var_parts[1], int(var_parts[2]), int(var_parts[4]), var_parts[5]


# slack var for class subject requirements: [class]_[subject]
def _get_slack_var_parts__class_subject_requirements(var_name):
    var_parts = var_name.split("_")
    return var_parts[0], var_parts[1]


# f"{teacher_name}_offered_lessons"
def _get_slack_var_parts__teacher_offered_lessons_total(var_name):
    var_parts = var_name.split("_")
    return var_parts[0]


def _get_var_parts_obj(var_name):
    parts = _get_var_parts(var_name)
    return {
        "teacher_key": parts[0],
        "subject_key": parts[1],
        "day_index": parts[2],
        "lesson_hour_slot_index": parts[3],
        "class_key": parts[4],
    }


def create_all_var_names(all_teachers_list, all_classes, max_days, max_slots_per_day):
    all_var_names = []

    # classes determine which it needs
    # even when we don't need all vars, create them
    # we set them to fixed values in other methods
    #  (e.g. skip slot -> set all corresponding vars to 0)

    for class_obj in all_classes:
        class_key = class_obj['key']
        subjects_info = class_obj["subjects"]

        for subject_info in subjects_info:
            subject_key = subject_info["name"]
            teachers_with_hours = subject_info["teachers_with_hours"]  # list

            for teachers_with_hour_tuple in teachers_with_hours:
                teacher_key = teachers_with_hour_tuple['teacher_key']
                total_required_hours = teachers_with_hour_tuple['hours']

                # all slots for this teacher and subject
                for day_index in range(max_days):
                    for slot_index in range(max_slots_per_day):
                        var_name = _get_var_name(teacher_key, subject_key, day_index, slot_index, class_key)
                        all_var_names.append(var_name)

    return all_var_names


def get_all_vars_with_preset(all_var_names, teacher_key=None, subject_key=None, day_index=None, slot_index=None,
                             class_key=None):
    var_names_with_preset = all_var_names.copy()

    if teacher_key is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 teacher_key == _get_var_parts_obj(var_name)["teacher_key"]]

    if subject_key is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 subject_key == _get_var_parts_obj(var_name)["subject_key"]]

    if day_index is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 day_index == _get_var_parts_obj(var_name)["day_index"]]

    if slot_index is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 slot_index == _get_var_parts_obj(var_name)["lesson_hour_slot_index"]]

    if class_key is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 class_key == _get_var_parts_obj(var_name)["class_key"]]

    return var_names_with_preset

def create_empty_timetable(max_days, max_slots_per_day):
    # every day with the max number of lessons
    # key is the day
    timetable = {}
    for day_index in range(max_days):
        timetable[day_index] = [None] * max_slots_per_day
        for lesson_hour in range(max_slots_per_day):
            timetable[day_index][lesson_hour] = None

    return timetable
#
#
def print_timetable(class_obj, stundenplan, stundenplanHelper, at_least_one_lesson):
    # max_days = 0
    # for day_name, max_lesson_hours in days.items():
    #     if max_lesson_hours > max_days:
    #         max_days = max_lesson_hours



    data = []
    table_header = []
    table_body = [[]] * stundenplanHelper.max_slots_per_day

    for i, _ in enumerate(table_body):
        table_body[i] = [""] * (stundenplanHelper.max_days + 1)

    table_header.append("Stunden")
    for lesson_hour in range(stundenplanHelper.max_slots_per_day):
        table_body[lesson_hour][0] = lesson_hour

    for day_index, lessons_list in stundenplan.items():
        table_header.append(day_index)
        for lesson_hour_0, lesson_obj in enumerate(lessons_list):
            if lesson_obj is not None:
                table_body[lesson_hour_0][day_index + 1] = f"'{lesson_obj['subject_key']}' [{lesson_obj['teacher_key']}]"

    data.append(table_header)
    data.extend(table_body)
    print(f"\nStundenplan für: {class_obj['key']}")
    if at_least_one_lesson:
        print(tabulate(data, headers="firstrow"))
    else:
        print("Keine Lehrveranstaltungen gefunden.")
    print()

def is_real_var(var):
    return var.name.startswith("var_")

def is_slack_var(var):
    return var.name.startswith("r_")

def get_stundenplan_tuples_from_vars(stundenplanHelper):
    # we need a timetable for each teacher
    # we need a timetable for each class

    all_class_timetables = []
    solution_variables_obj = []

    # for some reason var_name expluced the
    for var_name, var in stundenplanHelper.all_vars.items():
        if not is_real_var(var):
            continue

        if var.varValue == 1:
            parts = _get_var_parts_obj(var_name)
            solution_variables_obj.append(parts)

    # for teacher_key in all_teachers:

    # timetable for each class

    for class_obj in stundenplanHelper.excel_extractor.all_classes:
        at_least_one_lesson = False
        class_key = class_obj['key']
        table_start_coord_tuple = class_obj['table_dates'][2]

        # one row per day, then all slots in array
        class_timetable = create_empty_timetable(stundenplanHelper.max_days, stundenplanHelper.max_slots_per_day)
        for day_index in range(stundenplanHelper.max_days):
            for slot_index in range(stundenplanHelper.max_slots_per_day):
                # find correct entry
                for var_obj in solution_variables_obj:
                    if (var_obj["class_key"] == class_key and
                            var_obj["day_index"] == day_index and
                            var_obj["lesson_hour_slot_index"] == slot_index
                    ):
                        class_timetable[day_index][slot_index] = var_obj
                        at_least_one_lesson = True

        if at_least_one_lesson:
            all_class_timetables.append([class_obj, class_timetable, table_start_coord_tuple])

        print_timetable(class_obj, class_timetable, stundenplanHelper, at_least_one_lesson)

    # # timetable for each teacher
    # for teacher_name in stundenplanHelper.excel_extractor.all_teachers_list:
    #     teacher_timetable = create_empty_timetable()
    #     for day_name in all_days:
    #         max_lesson_hours = days[day_name]
    #         for lesson_hour in range(1, max_lesson_hours + 1):
    #             # find correct entry
    #             for var_obj in solution_variables_obj:
    #                 if (var_obj["teacher"] == teacher_name and
    #                         var_obj["day"] == day_name and
    #                         var_obj["lesson_hour"] == lesson_hour
    #                 ):
    #                     teacher_timetable[day_name][lesson_hour - 1] = var_obj
    #
    #     print_timetable(teacher_name, teacher_timetable)
    return all_class_timetables


# TODO output to new file
excel_extractor = ExcelExtractorGesamtuebersicht(
    'example_real/SJ 25-26_Gesamtübersicht Einsatz Lehrkräfte EA Halle 2025-05-21_3.xlsx',
    "example_real/Mappings.xlsx",
    'example_real/03_KW 10_02.03.-06.03.2026_gesetzteverpflichtende Stunden drin ohne Reste aus Vorwoche_IN.xlsm'
)
excel_extractor.read_all()

excel_stundenerfassung = ExcelExtractorStundenerfassung(
    "example_real/",
)
excel_stundenerfassung.apply_filter_found_files_with_real_classes(excel_extractor.all_classes)
excel_stundenerfassung.read_all()
print("finished")

#
stundenplaner = StundenplanHelper(excel_extractor, excel_stundenerfassung)
# stundenplaner.read_excel_data()
# stundenplaner.read_excel_stundenerfassung()
# stundenplaner.prepare_excel_data()
stundenplaner.init_new_timetable_problem()
stundenplaner.setup_fixed_vars()
stundenplaner.setup_constraints()
# stundenplaner.solve_timetable_problem()
stundenplaner.solve_timetable_problem_2()
# stundenplaner.write_timetable_solution_to_excel('example/07_KW 45_03.11.-07.11.2025_OUT.xlsm')
