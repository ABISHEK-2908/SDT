from flask import Flask, render_template, request
import csv
import os
from openpyxl import Workbook, load_workbook

app = Flask(__name__)

CSV_FILE = "daily_status.csv"
EXCEL_FILE = "daily_status.xlsx"


def save_to_csv(name, date, work_done, blockers, plan):

    file_exists = os.path.isfile(CSV_FILE)

    with open(CSV_FILE, mode='a', newline='') as file:
        writer = csv.writer(file)

        if not file_exists:
            writer.writerow(["Name", "Date", "Work Done", "Blockers", "Plan"])

        writer.writerow([name, date, work_done, blockers, plan])


def save_to_excel(name, date, work_done, blockers, plan):

    if os.path.isfile(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Date", "Work Done", "Blockers", "Plan"])

    ws.append([name, date, work_done, blockers, plan])
    wb.save(EXCEL_FILE)


def save_report(name, date, work_done, blockers, plan):
    save_to_csv(name, date, work_done, blockers, plan)
    save_to_excel(name, date, work_done, blockers, plan)


@app.route("/", methods=["GET", "POST"])
def index():

    report = None

    if request.method == "POST":

        name = request.form["name"]
        date = request.form["date"]
        status = request.form["status"]

        lines = status.split("\n")

        work_done = lines[0] if len(lines) > 0 else ""
        blockers = lines[1] if len(lines) > 1 else ""
        plan = lines[2] if len(lines) > 2 else ""

        save_report(name, date, work_done, blockers, plan)

        report = {
            "name": name,
            "date": date,
            "work_done": work_done,
            "blockers": blockers,
            "plan": plan
        }

    return render_template("index.html", report=report)


if __name__ == "__main__":
    app.run(debug=True)
