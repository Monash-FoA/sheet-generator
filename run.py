import argparse
import logging
import yaml

from scripts.student_info import retrieve_students
from scripts.question_ingest import retrieve_questions
from scripts.random_questions import create_random_questions
from scripts.create_sheet import create_sheets
from scripts.sheets_to_pdf import workbooks_to_pdf

logging.basicConfig()
LOGGER = logging.getLogger(__file__)
LOGGER.setLevel(logging.INFO)

def main():
    p = argparse.ArgumentParser()
    p.add_argument("config")

    args = p.parse_args()
    with open(args.config, "r") as f:
        conf = yaml.safe_load(f)

    students = retrieve_students(conf["students"])
    LOGGER.info(f"Retrieved {len(students)} students.")

    question_data = retrieve_questions(conf["questions"])
    LOGGER.info(f"Retrieved {len(question_data)} questions.")

    n_sheets = len(students)
    student_sheets = create_random_questions(len(students), question_data, n_sheets, len(conf["template"]["question_locations"]), conf["randomisation"])
    LOGGER.info(f"Generated {n_sheets} random question sets and assigned to {len(students)} students.")

    create_sheets(students, student_sheets, conf["template"], conf.get("images", {}))
    LOGGER.info(f"Generated {len(students)} workbooks.")

    workbooks_to_pdf(students)
    LOGGER.info(f"Generated {len(students)} PDFs.")

if __name__ == "__main__":
    main()
