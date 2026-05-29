from openpyxl.cell import cell
from yachalk import chalk

import os
from datetime import datetime

class Logger:
    output_file = None
    _output_handle = None
    output_timestamp = True

    @staticmethod
    def _close_output_handle():
        if Logger._output_handle is not None:
            Logger._output_handle.close()
            Logger._output_handle = None

    @staticmethod
    def set_output_file(file_path: str):
        Logger._close_output_handle()
        Logger.output_file = file_path
        print(f"[Logger] output will be written to '{file_path}', existing content will be overwritten")

        parent_dir = os.path.dirname(file_path)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)

        Logger._output_handle = open(file_path, "w", encoding="utf-8")

    @staticmethod
    def close_output_file():
        Logger._close_output_handle()
        Logger.output_file = None

    @staticmethod
    def write_to_output_file(message: str):
        if Logger._output_handle is None:
            return
        if Logger.output_timestamp:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        else:
            timestamp = ""
        Logger._output_handle.write(f"{timestamp} {message}\n")

    @staticmethod
    def debug(message: str):
        _message = f"[DEBUG] {message}"
        print(_message)
        Logger.write_to_output_file(_message)

    @staticmethod
    def log(message: str):
        _message = f"[LOG] {message}"
        print(_message)
        Logger.write_to_output_file(_message)

    @staticmethod
    def warn(message: str):
        _message = f"[WARN] {message}"
        print(chalk.yellow(_message))
        Logger.write_to_output_file(_message)

    @staticmethod
    def error(message: str):
        _message = f"[ERROR] {message}"
        print(chalk.red(_message))
        Logger.write_to_output_file(_message)

    # row, col is 1 based!
    @staticmethod
    def get_cell_full_coord(cell):
        return f"{cell.coordinate} ({cell.row}, {cell.column})"

    @staticmethod
    def get_teacher_full(teacher_obj):
        return f"{teacher_obj['teacher_full_name']} ({teacher_obj['key']})"