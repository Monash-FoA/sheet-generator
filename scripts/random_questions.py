import copy
import itertools
import random
import yaml

def create_random_questions(n_students, question_data, n_sheets, n_questions_per_sheet, conf):
    sets = select_question_pools(question_data, n_sheets, n_questions_per_sheet, conf)
    choices = load_choices(conf["choices"])
    return assign_sets_to_students(n_students, choices, sets)

def select_question_pools(question_data, n_sheets, n_questions_per_sheet, conf):
    # Generate all combinations possible sheets.
    possible = list(sorted(filter(lambda x: x[1] > 0, map(lambda x: (x, question_set_score(x, conf)), itertools.combinations(question_data, n_questions_per_sheet))), key=lambda x:x[1], reverse=True))
    if len(possible) < n_sheets:
        raise ValueError("Not enough questions to generate this many sheets.")
    lowest_score = possible[n_sheets-1][1]
    possible = list(filter(lambda x: x[1] >= lowest_score, possible))
    random.shuffle(possible)
    return possible

def question_set_score(question_set, conf):
    """
    Scores a question set from 0-10.
    sets with score 0 should not be included.
    """
    if len(set(q["index"] for q in question_set)) != len(question_set):
        return 0 # Two questions have same index.
    total_diff = sum(q["difficulty"] for q in question_set)
    if not conf["selection"]["min_diff"] <= total_diff <= conf["selection"]["max_diff"]:
        return 0 # Outside of accepted difficulty range
    tags_sums = [
        sum(q["tags"][col] for q in question_set)
        for col in conf["selection"]["required_tags"]
    ]
    for v in tags_sums:
        if v < 1:
            return 0
    # Otherwise, check how close we are to average of min/max diff.
    distance = abs((conf["selection"]["min_diff"] + conf["selection"]["max_diff"]) / 2 - total_diff)
    return 10 - distance * 18/((conf["selection"]["max_diff"] + conf["selection"]["min_diff"]))

def assign_sets_to_students(n_students, choices, sets):
    final_sets = []
    for _ in range(n_students):
        question_set, _score = random.choice(sets)
        question_set = copy.deepcopy(list(question_set))

        random.shuffle(question_set)

        choice_dict = {
            key: random.choice(value)
            for key, value in choices.items()
        }
        for question in question_set:
            for info in question["question_cells"].values():
                if "value" in info:
                    info["value"] = info["value"].format(**choice_dict)

        final_sets.append(question_set)
    return final_sets

def load_choices(choice_path):
    with open(choice_path, "r") as f:
        choices = yaml.safe_load(f)
    for key in choices:
        # Every element of the choices dict should be a list
        if type(choices[key]) != list or len(choices[key]) < 1:
            raise ValueError("Choices config format incorrect")
        for i in range(len(choices[key])):
            if type(choices[key][i]) == dict:
                t = choices[key][i]["type"]
                if t == "file":
                    with open(choices[key][i]["path"], "r") as f:
                        choices[key][i] = f.read()
                else:
                    raise ValueError(f"Choices config type invalid: {t}")
    return choices
