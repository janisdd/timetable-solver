from pulp import *
from tabulate import tabulate

from src.excel_extractor import ExcelExtractor


# TODO internship not needed -> set all slots to skip
# TODO teacher_sick_data not needed -> set all slots to skip
# fixed_teacher_subjects_for_class not needed because we have a white list for a class which teachers are allowed for which subject
# teacher_fixed_data is not needed because stored in 'PLAN' sheet in excel table

class StundenplanHelper:
    # the lp problem
    problem = None
    # all lp vars
    all_vars = None
    all_var_names = None
    max_days = -1
    max_slots_per_day = -1

    def __init__(self, excel_extractor):
        self.excel_extractor = excel_extractor

    def read_excel_data(self):
        self.excel_extractor.read_all()

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
        # - some classes slots should be skipped -> set all corresponding vars to 0 (see 'PLAN' sheet)
        # - some classes slots are already filled -> set all corresponding vars to 0 (see 'PLAN' sheet) # TODO 0 -> not set
        # - some teachers have white/black lists for slots
        #   - convert all to black list (so we can set vars to 0) (because 1 would mean we know the solution)

        c_count = 0
        for class_obj in self.excel_extractor.all_classes:
            # one list for each day
            table_dates = class_obj['table_dates'][1]
            class_key = class_obj['key']

            for day_index, table_info_list in enumerate(table_dates):
                for slot_index, info_obj in enumerate(table_info_list):
                    filled_value = info_obj['entry']  # e.g. LF5(Gru) -> teacher already teachers subject here
                    ignore = info_obj['ignore']

                    should_not_process = (filled_value != '' and filled_value is not None) or ignore == True

                    if should_not_process:
                        var_names = get_all_vars_with_preset(self.all_var_names, class_key=class_key,
                                                             day_index=day_index, slot_index=slot_index)

                        print(f"skip prefilled slot {day_index}, {slot_index} for class {class_key} (prefilled value: {filled_value})")

                        if len(var_names) == 0:
                            print("")

                        var = [self.all_vars[var_name] for var_name in var_names]
                        # self.problem += lpSum(var) == 0, f"skip prefilled slots in plan {var_names}, {day_index}, {slot_index}"
                        self.problem += lpSum(
                            var) == 0, f"skip prefilled slots in plan, {day_index}, {slot_index} {c_count}"
                        c_count += 1

        c_count = 0
        # TODO set prefilled to 1?
        for i, teacher_obj in enumerate(self.excel_extractor.all_teachers_list):
            teacher_key = teacher_obj["key"]
            availability_preference_table = teacher_obj['availability_preference_table']

            for day_index, day_slots_list in enumerate(availability_preference_table):
                for slot_index, slot_obj in enumerate(day_slots_list):
                    class_key = slot_obj['class_key']
                    allowed = slot_obj['allowed']

                    if not allowed:
                        # teacher has some fixed class in this slot...
                        var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key,
                                                             day_index=day_index, slot_index=slot_index)

                        # print(f"skip fixed teacher slots {day_index}, {slot_index} for teacher {teacher_key} (class: {class_key})")

                        if len(var_names) == 0:
                            print("")

                        var = [self.all_vars[var_name] for var_name in var_names]
                        # self.problem += lpSum(var) == 0, f"skip fixed teacher slots {var_names}, {day_index}, {slot_index}"
                        self.problem += lpSum(
                            var) == 0, f"skip fixed teacher slots, {day_index}, {slot_index} {c_count}"
                        c_count += 1

                        # because of this, the class has no other option too

                        if class_key is not None:
                            var_names_class = get_all_vars_with_preset(self.all_var_names, class_key=class_key,
                                                                       day_index=day_index, slot_index=slot_index)

                            print(
                                f"skip fixed class slots day {day_index}, slot {slot_index} class: {class_key} (teacher {teacher_key})")

                            if len(var_names_class) == 0:
                                print("")

                            var_class = [self.all_vars[var_name] for var_name in var_names_class]
                            # self.problem += lpSum(var_class) == 0, f"skip teacher slots for classes {var_names_class}, {day_index}, {slot_index} {c_count}"
                            self.problem += lpSum(
                                var_class) == 0, f"skip teacher slots for classes, {day_index}, {slot_index} {c_count}"
                            c_count += 1

    def setup_constraints(self):
        print('setup_constraints')

        # every day can only have the max number of slots for one class
        # monday-1 has max 4 slots, so 5a can only have 4 slts (and any other class too)
        # [teacher]_[subject]_montag-1_slot_1_5a + ... <= 4
        c_count = 0

        for day_index in range(self.max_days):
            for class_obj in self.excel_extractor.all_classes:
                class_key = class_obj['key']
                var_names = get_all_vars_with_preset(self.all_var_names, day_index=day_index, class_key=class_key)
                var = [self.all_vars[var_name] for var_name in var_names]
                c1 = lpSum(
                    var) <= self.max_slots_per_day, f"every day can only have a max number of slos for one class {c_count}"
                c_count += 1
                self.problem += c1

        # every teacher can only teach the max number of slots per day
        # day 4 slots -> teacher can only teacher 4
        # l1_[subject]_montag-1_1_[class] + ... <= 4
        c_count = 0
        for day_index in range(self.max_days):
            for teacher_obj in self.excel_extractor.all_teachers_list:
                teacher_key = teacher_obj["key"]
                var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key, day_index=day_index)
                var = [self.all_vars[var_name] for var_name in var_names]
                c1 = lpSum(
                    var) <= self.max_slots_per_day, f"every teacher can only teach the max number of slots per day {c_count}"
                c_count += 1
                self.problem += c1

        # a teacher can only teach max one subject per class at one slot
        # e.g. l1 can only teach deutsch to 5a at montag-1 in slot 1 (but not to 5b in the same slot)
        # l1_[subject]_montag-1_1_[class] + ... = 1
        c_count = 0
        for teacher_obj in self.excel_extractor.all_teachers_list:
            teacher_key = teacher_obj["key"]
            for day_index in range(self.max_days):
                for slot_index in range(self.max_slots_per_day):
                    var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key,
                                                         slot_index=slot_index,
                                                         day_index=day_index)
                    var = [self.all_vars[var_name] for var_name in var_names]
                    c1 = lpSum(var) <= 1, f"teacher can only teach max one subject per class at one slot {c_count}"
                    c_count += 1
                    self.problem += c1

        # every class can only have one subject with one teacher at a specific slot
        # e.g. 5a at montag-1 in slot 1 can only have l1 teaching deutsch
        # [teacher]_[subject]_montag-1_1_5a + ... = 1
        c_count = 0
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

    def solve_timetable_problem(self):

        all_slack_vars_for_class_lesson_requirements = []

        # TODO test all teachers should have 10 lessons / slots
        required_slots = 10
        for index, teacher_obj in enumerate(self.excel_extractor.all_teachers_list):
            teacher_key = teacher_obj["key"]

            s1 = pulp.LpVariable(f"{teacher_key}_missing_required_lessons", lowBound=0) # Slack variable
            all_slack_vars_for_class_lesson_requirements.append(s1)

            var_names = get_all_vars_with_preset(self.all_var_names, teacher_key=teacher_key)
            var = [self.all_vars[var_name] for var_name in var_names]

            c1 = lpSum(var) + s1 == required_slots, f"teacher should get {required_slots}, {index}"
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

        print("finished")

        for slack_var in all_slack_vars_for_class_lesson_requirements:
            if slack_var.varValue > 0:
                teacher_Key = get_slack_var_parts__teacher_offered_lessons_total(slack_var.name)
                # offered_lessons = self.excel_extractor.all_teachers_dict[teacher_Key]["offered_lessons"]
                # print("WARNING: Slack variable > 0, not all constraints could be satisfied!")
                print(
                    f"Lehrer {teacher_Key} hat {int(slack_var.varValue)}x zu wenig Unterricht (hat {int(required_slots) - int(slack_var.varValue)}/{int(required_slots)})")


        get_stundenplan_from_vars(self)

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
                table_body[lesson_hour_0][day_index + 1] = f"{lesson_obj['subject_key']}/{lesson_obj['teacher_key']}"

    data.append(table_header)
    data.extend(table_body)
    print(f"\nStundenplan für: {class_obj['key']}")
    if at_least_one_lesson:
        print(tabulate(data, headers="firstrow"))
    else:
        print("Keine Lehrveranstaltungen gefunden.")
    print()

def get_stundenplan_from_vars(stundenplanHelper):
    # we need a timetable for each teacher
    # we need a timetable for each class

    solution_variables_obj = []

    for var_name, var in stundenplanHelper.all_vars.items():
        if var.varValue == 1:
            parts = _get_var_parts_obj(var_name)
            solution_variables_obj.append(parts)

    # for teacher_key in all_teachers:

    # timetable for each class

    for class_obj in stundenplanHelper.excel_extractor.all_classes:
        at_least_one_lesson = False
        class_key = class_obj['key']
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



# TODO output to new file
excel_extractor = ExcelExtractor(
    'example/SJ 25-26_Gesamtübersicht Einsatz Lehrkräfte EA Halle 2025-05-21_3.xlsx',
    'example/07_KW 45_03.11.-07.11.2025_IN.xlsm'
)

stundenplaner = StundenplanHelper(excel_extractor)
stundenplaner.read_excel_data()
stundenplaner.init_new_timetable_problem()
stundenplaner.setup_fixed_vars()
stundenplaner.setup_constraints()
stundenplaner.solve_timetable_problem()

# # TODO
# teacher_avg_lessons = 1000
#
# # all days can have a different number of lessons
# # but for display purposes we need to know the max number of lessons among all days
# maximal_lessons_per_day = -1
# for day_name, max_lesson_hours in days.items():
#     if max_lesson_hours > maximal_lessons_per_day:
#         maximal_lessons_per_day = max_lesson_hours
#

######## constraints

# # the sum of lessons for one teacher cannot exceed the total number of lessons among all subjects
# not necessarily...
# for teacher_name, teacher_info in teachers.items():
#     offered_lessons = teacher_info["offered_lessons"]
#     teacher_subjects = teacher_info["subjects"]
#
#     var_names = get_all_vars_with_preset(teacher=teacher_name)
#     var = [all_vars[var_name] for var_name in var_names]
#     c1 = lpSum(var) <= offered_lessons, f"the sum of lessons for one teacher cannot exceed the total number of lessons among all subjects {teacher_name}"
#     prob += c1

# internships
# for class_name in all_internships:
#     days_off_names = internships[class_name]
#     for day_name in days_off_names:
#         var_names = get_all_vars_with_preset(class_key=class_name, day_index=day_name)
#         var = [all_vars[var_name] for var_name in var_names]
#         prob += lpSum(var) == 0, f"internships {var_names}, {day_name}"

# TODO is combined with favor monday over friday...
# coefficient = -1000
# fixe_teacher_subjects_for_class = LpAffineExpression(None)
# # fixed teachers with subjects for classes
# for class_name in fixed_teacher_subjects_for_class.keys():
#     subject_teacher_tuple_map = fixed_teacher_subjects_for_class[class_name]
#     for subject_name, teacher_name in subject_teacher_tuple_map.items():
#         var_names = get_all_vars_with_preset(teacher=teacher_name, subject=subject_name, class_=class_name)
#
#         for var_name in var_names:
#             var = all_vars[var_name]
#             fixe_teacher_subjects_for_class += var * coefficient

# TODO??
# one class should always have the same subject with the same teacher
# because we don't know which day the teacher will teach, we set all other options to 0
# e.g. 5a should always be teaching deutsch
# [teacher]_[subject]_montag-1_1_5a + ... = 1
# c_count = 0
# for class_name in fixed_teacher_subjects_for_class.keys():
#     subject_teacher_tuple_map = fixed_teacher_subjects_for_class[class_name]
#     for subject_name, teacher_name in subject_teacher_tuple_map.items():
#         var_names_with_all_teachers = get_all_vars_with_preset(class_key=class_name, subject_key=subject_name)
#         # now get all other teachers
#         var_names_with_other_teachers = [var_name for var_name in var_names_with_all_teachers if
#                                          teacher_name != get_var_parts_obj(var_name)["teacher"]]
#         var = [all_vars[var_name] for var_name in var_names_with_other_teachers]
#         c1 = lpSum(var) == 0, f"one class should always have the same subject with the same teacher {c_count}"
#         c_count += 1
#         prob += c1

# teacher sick data (for day X for [list] lesson hours)
# set the variables to 0
# e.g. l1_[subject]_montag-1_1_[class] = 0´
# c_count = 0
# for teacher_name, sick_map in teacher_sick_data.items():
#     for day_name, sick_lesson_hours_list in sick_map.items():
#         for sick_lesson_hour in sick_lesson_hours_list:
#             var_names = get_all_vars_with_preset(teacher_key=teacher_name, day_index=day_name, slot_index=sick_lesson_hour)
#             var = [all_vars[var_name] for var_name in var_names]
#             c1 = lpSum(var) == 0, f"teacher sick data (for day X for [list] lesson hours) {c_count}"
#             c_count += 1
#             prob += c1


# teacher fixed data (teacher X want subject Y for class Z at day A in slots [list])
# set the variables to 1
# e.g. l1_deu_montag-1_2_5a = 1
# c_count = 0
# for teacher_name, fixed_map in teacher_fixed_data.items():
#     for day_name, data in fixed_map.items():
#         class_name = data["_class"]
#         subject_name = data["subject"]
#         slots_list = data["slots"]
#         for slot in slots_list:
#             var_names = get_all_vars_with_preset(teacher_key=teacher_name, subject_key=subject_name, day_index=day_name,
#                                                  slot_index=slot, class_key=class_name)
#             var = [all_vars[var_name] for var_name in var_names]
#             c1 = lpSum(var) == 1, f"teacher fixed data (teacher X want subject Y for class Z at day A in slots [list]) {c_count}"
#             c_count += 1
#             prob += c1

# every class should get all subjects in one year
# because this is a slack var, we minimize it -> we try to meet this goal
# e.g. 5a must get 4x deutsch
# [teacher]_deu_[day]_[class hour]_5a + ... = 4
# c_count = 0
# all_slack_vars_for_class_lesson_requirements = []
# for class_name in all_classes:
#     class_with_subjects = classes[class_name]
#
#     for subject_name, required_lessons in class_with_subjects.items():  # e.g. deutsch
#         var_names = get_all_vars_with_preset(class_key=class_name, subject_key=subject_name)
#         var = [all_vars[var_name] for var_name in var_names]
#
#         s1 = pulp.LpVariable(f"{class_name}_{subject_name}_missing_required_lessons", lowBound=0)  # Slack variable
#         all_slack_vars_for_class_lesson_requirements.append(s1)
#
#         c1 = lpSum(var) + s1 == required_lessons, f"every class must get all subjects in one year {c_count}"
#         c_count += 1
#         prob += c1

#############################################################################################
# every teacher should have his offered lessons
# simple: every teacher should do his required lessons, everything too much is captured in slack var -> minimize slack
# c_count = 0
# all_slack_vars_for_teacher_lesson_requirements = []
# sub_teacher_lesson_requirements = []
#
# for teacher_name, teacher_info in teachers.items():
#     offered_lessons = teacher_info["offered_lessons"]
#     teacher_subjects = teacher_info["subjects"]
#
#     var_names = get_all_vars_with_preset(teacher_key=teacher_name)
#     var = [all_vars[var_name] for var_name in var_names]
#
#     s1 = pulp.LpVariable(f"{teacher_name}_missing_offered_lessons", lowBound=0)  # Slack variable
#     all_slack_vars_for_teacher_lesson_requirements.append(s1)
#
#     # sub_teacher_lesson_requirements.append(s1 * offered_lessons)
#     sub_teacher_lesson_requirements.append(s1)
#
#     c1 = lpSum(var) + s1 == offered_lessons, f"every teacher should have approx his offered lessons in total {c_count}"
#     c_count += 1
#     prob += c1


# dummy target function
# prob += 0 # no slack
# prob += lpSum(all_slack_vars)


# # favor first days first
# def sort_by_lesson(var_name):
#     parts_obj = get_var_parts_obj(var_name)
#     return parts_obj["lesson_hour"]
#
#
# # TODO this also handles "no holes" like free lessons hours in the middle?
# all_days_list = list(all_days)
# favor_first_days_obj_func_vars = LpAffineExpression(None)
# # var = [all_vars[var_name] for var_name in var_names]
# day_count = 0
# lessons_per_day_counter = 0
# for day_name, max_lessons in days.items():
#
#     # all_lessons_on_day = get_all_vars_with_preset(day=day_name)
#     # all_lessons_on_day.sort(key=sort_by_lesson)
#     day_count += 1
#
#     for lesson_hour in range(1, max_lessons + 1):
#         vars_lesson_slot_day = get_all_vars_with_preset(day_index=day_name, slot_index=lesson_hour)
#
#         lessons_per_day_counter += 1
#         for var_name in vars_lesson_slot_day:
#             var = all_vars[var_name]
#             favor_first_days_obj_func_vars += var * lessons_per_day_counter
#
#         print()

# for slack_var in all_slack_vars:
# make sure slack variables are used last by multiplying with a large number
# prob += favor_first_days_obj_func_vars + fixe_teacher_subjects_for_class + lpSum(all_slack_vars) * len(all_var_names)
# prob += favor_first_days_obj_func_vars + lpSum(all_slack_vars_for_class_lesson_requirements) * len(all_var_names)
# prob += (favor_first_days_obj_func_vars +
#          lpSum(all_slack_vars_for_class_lesson_requirements) * len(all_var_names) +
#          lpSum(sub_teacher_lesson_requirements)
#          )
# prob += favor_first_days_obj_func_vars

# prob.writeLP("Stundenplan2.lp")
# prob.solve()
#
# # The status of the solution is printed to the screen
# print("Status:", LpStatus[prob.status])
#
# if prob.status != 1:
#     print("No solution found")
#     exit()
#
# # Each of the variables is printed with it's resolved optimum value
# for v in prob.variables():
#     if v.varValue == 1:
#         print(v.name, "=", v.varValue)
#
# print()
# get_stundenplan_from_vars(prob.variables())
#
# # check slack variables
# for slack_var in all_slack_vars_for_class_lesson_requirements:
#     if slack_var.varValue > 0:
#         class_name, subject_name = get_slack_var_parts__class_subject_requirements(slack_var.name)
#         # print("WARNING: Slack variable > 0, not all constraints could be satisfied!")
#         required_lessons = classes[class_name][subject_name]
#         print(
#             f"Klasse {class_name} hat {int(slack_var.varValue)}x zu wenig {subject_name} (hat {int(required_lessons) - int(slack_var.varValue)}/{int(required_lessons)})")
#
#
# for slack_var in all_slack_vars_for_teacher_lesson_requirements:
#     if slack_var.varValue > 0:
#         teacher_name = get_slack_var_parts__teacher_offered_lessons_total(slack_var.name)
#         offered_lessons = teachers[teacher_name]["offered_lessons"]
#         # print("WARNING: Slack variable > 0, not all constraints could be satisfied!")
#         print(
#             f"Lehrer {teacher_name} hat {int(slack_var.varValue)}x zu wenig Unterricht (hat {int(offered_lessons) - int(slack_var.varValue)}/{int(offered_lessons)})")
#
#
# print()

