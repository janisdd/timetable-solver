from pulp import *
from tabulate import tabulate

# how many lessons per subject a teacher can/must teach
teachers = {
    "l1": {
        "deutsch": 3,
        "mathe": 3,
    },
    # "l2": {
    #     "mathe": 5,
    #     "sport": 3,
    # },
    # "l3": {
    #     "deutsch": 3,
    # },
    # "l4": {
    #     "deutsch": 3,
    #     "mathe": 4,
    # },
}

classes = {
    # "5a": {
    #     "deutsch": 4,
    #     "mathe": 4,
    # },
    # "5b": {
    #     "deutsch": 4,
    #     "mathe": 4,
    # },
    "5a": {
        "deutsch": 3,
        "mathe": 2,
    },
    # "5b": {
    #     "deutsch": 1,
    #     "mathe": 1,
    # },
}

# day_week
days = {
    "montag-1": 4,
    "dienstag-1": 4,
    "mittwoch-1": 4,
    # "donnerstag-1": 4,
    # "freitag-1": 4,
}

# whole day for classes off
internships = {
    "5a": [
        "dienstag-1",
    ]
}

# TODO sanity checks
# - same subjects
# - there must be at least one teacher per subject for a class
# - warning, when a teacher has a subject that is not needed by any class

all_teachers = teachers.keys()
all_classes = classes.keys()
all_days = days.keys()
all_internships = internships.keys()

maximal_lessons_per_day = -1
for day_name, max_lesson_hours in days.items():
    if max_lesson_hours > maximal_lessons_per_day:
        maximal_lessons_per_day = max_lesson_hours


# Create the 'prob' variable to contain the problem data
prob = LpProblem("Stundenplan", LpMinimize)


# create binary variables
# [teacher]_[subject]_[day]_slot_[class hour]_[class]
def get_var_name(teacher, subject, day, lesson_hour, class_):
    return f"{teacher}_{subject}_{day}_slot_{lesson_hour}_{class_}"

# [teacher]_[subject]_[day]_slot_[class hour]_[class]
def get_var_parts(var_name):
    var_parts = var_name.split("_")
    return var_parts[0], var_parts[1], var_parts[2], int(var_parts[4]), var_parts[5]

# slack var: [class]_[subject]
def get_slack_var_parts(var_name):
    var_parts = var_name.split("_")
    return var_parts[0], var_parts[1]

def get_var_parts_obj(var_name):
    parts = get_var_parts(var_name)
    return {
        "teacher": parts[0],
        "subject": parts[1],
        "day": parts[2],
        "lesson_hour": parts[3],
        "class_": parts[4],
    }


def create_all_var_names():
    all_var_names = []
    for teacher_key in all_teachers:
        for subject in teachers[teacher_key].keys():
            for day in all_days:
                for class_ in all_classes:
                    max_lesson_hours = days[day]
                    for lesson_hour in range(1, max_lesson_hours + 1):
                        var_name = get_var_name(teacher_key, subject, day, lesson_hour, class_)
                        # print(var_name)
                        all_var_names.append(var_name)

    return all_var_names


all_var_names = create_all_var_names()
all_vars = LpVariable.dicts("var", all_var_names, cat="Binary")


def get_all_vars_with_preset(teacher=None, subject=None, day=None, lesson_hour=None, class_=None):
    var_names_with_preset = all_var_names.copy()

    if teacher is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 teacher == get_var_parts_obj(var_name)["teacher"]]

    if subject is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 subject == get_var_parts_obj(var_name)["subject"]]

    if day is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 day == get_var_parts_obj(var_name)["day"]]

    if lesson_hour is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 lesson_hour == get_var_parts_obj(var_name)["lesson_hour"]]

    if class_ is not None:
        var_names_with_preset = [var_name for var_name in var_names_with_preset if
                                 class_ == get_var_parts_obj(var_name)["class_"]]

    return var_names_with_preset


def create_empty_timetable():
    # every day with the max number of lessons
    # key is the day
    timetable = {}
    for day_name, max_lesson_hours in days.items():
        timetable[day_name] = [None] * max_lesson_hours
        for lesson_hour in range(1, max_lesson_hours + 1):
            timetable[day_name][lesson_hour - 1] = None

    return timetable


def print_timetable(class_, stundenplan):
    # max_days = 0
    # for day_name, max_lesson_hours in days.items():
    #     if max_lesson_hours > max_days:
    #         max_days = max_lesson_hours

    data = []
    table_header = []
    table_body = [[]] * maximal_lessons_per_day

    for i, _ in enumerate(table_body):
        table_body[i] = [""] * (len(all_days) + 1)

    table_header.append("Stunden")
    for lesson_hour in range(1, maximal_lessons_per_day + 1):
        table_body[lesson_hour - 1][0] = lesson_hour

    all_days_list = list(all_days)
    for day_name, lessons_list in stundenplan.items():
        day_index = all_days_list.index(day_name)
        table_header.append(day_name)
        for lesson_hour_0, lesson_obj in enumerate(lessons_list):
            if lesson_obj is not None:
                table_body[lesson_hour_0][day_index + 1] = f"{lesson_obj['teacher']}/{lesson_obj['subject']}"

    data.append(table_header)
    data.extend(table_body)
    print(f"\nStundenplan für: {class_}")
    print(tabulate(data, headers="firstrow"))
    print()


def get_stundenplan_from_vars(all_vars_solved):
    # we need a timetable for each teacher
    # we need a timetable for each class

    solution_variables_obj = []

    solution_variables_obj_by_class = {}
    solution_variables_obj_by_teacher = {}

    for var_name, var in all_vars.items():
        if var.varValue == 1:
            parts = get_var_parts_obj(var_name)
            solution_variables_obj.append(parts)
            solution_variables_obj_by_class[parts["class_"]] = parts
            solution_variables_obj_by_teacher[parts["teacher"]] = parts

    # for teacher_key in all_teachers:

    # timetable for each class
    for class_name in all_classes:
        class_timetable = create_empty_timetable()
        for day_name in all_days:
            max_lesson_hours = days[day_name]
            for lesson_hour in range(1, max_lesson_hours + 1):
                # find correct entry
                for var_obj in solution_variables_obj:
                    if (var_obj["class_"] == class_name and
                            var_obj["day"] == day_name and
                            var_obj["lesson_hour"] == lesson_hour
                    ):
                        class_timetable[day_name][lesson_hour - 1] = var_obj

        print_timetable(class_name, class_timetable)

    # timetable for each teacher
    for teacher_name in all_teachers:
        teacher_timetable = create_empty_timetable()
        for day_name in all_days:
            max_lesson_hours = days[day_name]
            for lesson_hour in range(1, max_lesson_hours + 1):
                # find correct entry
                for var_obj in solution_variables_obj:
                    if (var_obj["teacher"] == teacher_name and
                            var_obj["day"] == day_name and
                            var_obj["lesson_hour"] == lesson_hour
                    ):
                        teacher_timetable[day_name][lesson_hour - 1] = var_obj

        print_timetable(teacher_name, teacher_timetable)


######## constraints


# every day can only have a max number of lessons for one class
# monday 1 has max 4 lessons, so 5a can only have 4 lessons (and any other class too)
# [teacher]_[subject]_montag-1_1_5a + ... <= 4
c_count = 0
for day_name, max_lessons in days.items():
    for class_name in all_classes:
        var_names = get_all_vars_with_preset(day=day_name, class_=class_name)
        var = [all_vars[var_name] for var_name in var_names]
        c1 = lpSum(var) <= max_lessons, f"every day can only have a max number of lessons for one class {c_count}"
        c_count += 1
        prob += c1

# every teacher can only teach the max number of lessons per day
# day 4 slots -> teacher can only teacher 4
# l1_[subject]_montag-1_1_[class] + ... <= 4
c_count = 0
for day_name, max_lessons in days.items():
    for teacher_name in all_teachers:
        var_names = get_all_vars_with_preset(teacher=teacher_name, day=day_name)
        var = [all_vars[var_name] for var_name in var_names]
        c1 = lpSum(var) <= max_lessons, f"every teacher can only teach the max number of lessons per day {c_count}"
        c_count += 1
        prob += c1

# a teacher can only teach max one subject per class at one slot
# e.g. l1 can only teach deutsch to 5a at montag-1 in slot 1 (but not to 5b in the same slot)
# l1_[subject]_montag-1_1_[class] + ... = 1
c_count = 0
for teacher_name in all_teachers:
    for day_name, max_lessons in days.items():
        for lesson_hour in range(1, max_lessons + 1):
            var_names = get_all_vars_with_preset(teacher=teacher_name, lesson_hour=lesson_hour, day=day_name)
            var = [all_vars[var_name] for var_name in var_names]
            c1 = lpSum(var) <= 1, f"teacher can only teach max one subject per class at one slot {c_count}"
            c_count += 1
            prob += c1

# every class must get all subjects in one year
# e.g. 5a must get 4x deutsch
# [teacher]_deu_[day]_[class hour]_5a + ... = 4
c_count = 0
all_slack_vars = []
for class_name in all_classes:
    class_with_subjects = classes[class_name]

    for subject_name, required_lessons in class_with_subjects.items():  # e.g. deutsch
        var_names = get_all_vars_with_preset(class_=class_name, subject=subject_name)
        var = [all_vars[var_name] for var_name in var_names]

        s1 = pulp.LpVariable(f"{class_name}_{subject_name}", lowBound=0)  # Slack variable
        all_slack_vars.append(s1)

        c1 = lpSum(var) + s1 == required_lessons, f"every class must get all subjects in one year {c_count}"
        c_count += 1
        prob += c1


# internships

for class_name in all_internships:
    days_off_names = internships[class_name]
    for day_name in days_off_names:
        var_names = get_all_vars_with_preset(class_=class_name, day=day_name)
        var = [all_vars[var_name] for var_name in var_names]
        prob += lpSum(var) == 0, f"internships {var_names}, {day_name}"


# dummy target function
# prob += 0 # no slack
# prob += lpSum(all_slack_vars)


# favor first days first
def sort_by_lesson(var_name):
    parts_obj = get_var_parts_obj(var_name)
    return parts_obj["lesson_hour"]


all_days_list = list(all_days)
favor_first_days_obj_func_vars = LpAffineExpression(None)
# var = [all_vars[var_name] for var_name in var_names]
day_count = 0
lessons_per_day_counter = 0
for day_name, max_lessons in days.items():

    # all_lessons_on_day = get_all_vars_with_preset(day=day_name)
    # all_lessons_on_day.sort(key=sort_by_lesson)
    day_count += 1

    for lesson_hour in range(1, max_lessons + 1):
        vars_lesson_slot_day = get_all_vars_with_preset(day=day_name, lesson_hour=lesson_hour)

        lessons_per_day_counter += 1
        for var_name in vars_lesson_slot_day:
            var = all_vars[var_name]
            favor_first_days_obj_func_vars += var * lessons_per_day_counter

        print()


# for slack_var in all_slack_vars:
# make sure slack variables are used last by multiplying with a large number
prob += favor_first_days_obj_func_vars + lpSum(all_slack_vars) * len(all_var_names)
# prob += favor_first_days_obj_func_vars

prob.writeLP("Stundenplan2.lp")
prob.solve()

# The status of the solution is printed to the screen
print("Status:", LpStatus[prob.status])

if prob.status != 1:
    print("No solution found")
    exit()


# Each of the variables is printed with it's resolved optimum value
for v in prob.variables():
    print(v.name, "=", v.varValue)

print()
get_stundenplan_from_vars(prob.variables())

# check slack variables
for slack_var in all_slack_vars:
    if slack_var.varValue > 0:
        class_name, subject_name = get_slack_var_parts(slack_var.name)
        # print("WARNING: Slack variable > 0, not all constraints could be satisfied!")
        required_lessons = classes[class_name][subject_name]
        print(f"Klasse {class_name} hat {slack_var.varValue}x zu wenig {subject_name} ({required_lessons-slack_var.varValue}/{required_lessons}.0)")

print()
print("finished")
