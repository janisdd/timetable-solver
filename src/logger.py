from openpyxl.cell import cell
from yachalk import chalk



class Logger:
    @staticmethod
    def debug(message: str):
        print("[DEBUG] " + message)

    @staticmethod
    def log(message: str):
        print("[LOG] " + message)

    @staticmethod
    def warn(message: str):
        print(chalk.yellow("[WARN] " + message))

    @staticmethod
    def error(message: str):
        print(chalk.red("[ERROR] " + message))

    # row, col is 1 based!
    @staticmethod
    def get_cell_full_coord(cell):
        return f"{cell.coordinate} ({cell.row}, {cell.column})"

    @staticmethod
    def get_teacher_full(teacher_obj):
        return f"{teacher_obj['teacher_full_name']} ({teacher_obj['key']})"