import openpyxl
from copy import copy

# Helper functions for merged cells
# Taken from https://gist.github.com/tchen/01d1d61a985190ff6b71fc14c45f95c9
def is_merged(cell):
    return not isinstance(cell, openpyxl.cell.cell.Cell)

def parent_of_merged_cell(cell):
    """ Find the parent of the merged cell by iterating through the range of merged cells """
    sheet = cell.parent
    child_coord = cell.coordinate

    # Note: if there are many merged cells in a large spreadsheet, this may become inefficient
    for merged in sheet.merged_cells.ranges:
        if child_coord in merged:
            return merged, merged.start_cell.coordinate
    return None

def retrieve_questions(question_conf):
    workbook_path = question_conf["workbook"]
    sheet_path = question_conf["sheet"]
    question_cols = question_conf["question_cols"]
    question_height = question_conf["question_height"]

    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    sheet = workbook[sheet_path]

    question_data = []
    index = 2
    while True:
        if sheet[question_cols[0] + str(index)].value is None:
            break
        question_data.append(get_question_data(sheet, index, **question_conf))
        index += question_height
    return question_data

def get_question_data(wb, index, question_cols, question_height, difficulty_col,
                      tags_cols, indexing_col, **kwargs):
    data = {
        "question_cells": {},
        "difficulty": None,
        "tags": [],
        "index": None,
    }
    # Question contents
    for cell_col in question_cols:
        for i in range(question_height):
            key = cell_col + str(i + index)
            cell = wb[key]
            data["question_cells"][key] = {}
            if cell.has_style:
                data["question_cells"][key]["style"] = {
                    "font": copy(cell.font),
                    "border": copy(cell.border),
                    "fill": copy(cell.fill),
                    "number_format": copy(cell.number_format),
                    "protection": copy(cell.protection),
                    "alignment": copy(cell.alignment),
                }
            data["question_cells"][key]["merged"] = parent_of_merged_cell(cell)[1] if is_merged(cell) else None
            if data["question_cells"][key]["merged"]:
                continue
            data["question_cells"][key]["value"] = cell.value
    data["difficulty"] = wb[difficulty_col + str(index)].value
    data["tags"] = {
        tag_col: wb[tag_col + str(index)].value for tag_col in tags_cols
    }
    data["index"] = wb[indexing_col + str(index)].value

    return data
