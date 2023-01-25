import openpyxl

def retrieve_students(config):
    workbook_path = config["workbook"]
    sheet_path = config["sheet"]
    info = config["info"]
    output = config["output"]

    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    sheet = workbook[sheet_path]

    students = []

    cur_index = 2
    while True:
        student_info = {}
        not_empty = False
        for item in info:
            student_info[item["key"]] = sheet[item["col"] + str(cur_index)].value
            if student_info[item["key"]] is not None:
                not_empty = True
        if not not_empty:
            break
        student_info["path"] = output.format(info=student_info)
        students.append(student_info)
        cur_index += 1

    return students

