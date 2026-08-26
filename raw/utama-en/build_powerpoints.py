# -*- coding: utf-8 -*-
"""
Build PowerPoint decks for the English 7-meeting module:
Introduction to Computers, Excel, and VBA.

Run:
    uv run --with python-pptx --with pillow python build_powerpoints.py
"""
import os
import shutil
import sys

sys.path.insert(0, r"D:\OneDrive\Bahan-Kuliah\_Slide-Template-UGM")

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

from mn_template_common import (
    pick_source,
    delete_all_slides,
    title_slide,
    agenda_slide,
    divider_slide,
    content_slide,
    closing_slide,
    add_note_box,
    add_table,
    build_and_save,
    NAVY,
    DGRAY,
    LGRAY,
    PALE_GOLD,
)


HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = r"D:\OneDrive\Bahan-Kuliah\_Slide-Template-UGM\MN1-Pengantar.pptx"
OUT_DIR = os.path.join(HERE, "00-Slide-Kuliah")
os.makedirs(OUT_DIR, exist_ok=True)

COURSE = "Algorithms and Computer Programming"
SUBTITLE = "Excel and VBA"
LECTURER = "Department of Civil and Environmental Engineering, UGM"
DATE_LABEL = "August 2026"


def bullet(text, size=None, level=0, kind="dot", color=None, bold=False):
    item = {"text": text, "bullet": kind, "level": level}
    if size:
        item["size"] = size
    if color:
        item["color"] = color
    if bold:
        item["bold"] = True
    return item


def plain(text, size=None, color=None, align=None, bold=False):
    item = {"text": text, "bullet": None}
    if size:
        item["size"] = size
    if color:
        item["color"] = color
    if align:
        item["align"] = align
    if bold:
        item["bold"] = True
    return item


def add_schedule(slide, rows):
    add_table(
        slide,
        left=Inches(0.70),
        top=Inches(1.47),
        width=Inches(8.55),
        height=Inches(2.95),
        headers=["Minutes", "Activity", "Small product"],
        rows=rows,
        col_widths=[Inches(1.15), Inches(4.15), Inches(3.25)],
        header_size=12,
        body_size=10.5,
        left_cols={1, 2},
        highlight_rows=[len(rows) - 1],
        highlight_fill=PALE_GOLD,
    )


def schedule_slide(prs, title, rows):
    slide = content_slide(prs, "100-Minute Flow", [
        plain(title, size=15, color=NAVY, bold=True),
    ], base_size=15)
    add_schedule(slide, rows)
    return slide


def quiz_slide(prs, questions, title="Short Quiz"):
    content_slide(prs, title, [
        plain("Work individually for 6 minutes, then discuss quickly for 7 minutes.", size=15, color=NAVY, bold=True),
        bullet(questions[0], kind="num", size=14.5),
        bullet(questions[1], kind="num", size=14.5),
        bullet(questions[2], kind="num", size=14.5),
    ], base_size=14.5)


def code_slide(prs, title, code_lines, note=None):
    slide = content_slide(prs, title, [
        plain("Read from top to bottom: data in, process, then result out.", size=13.5, color=NAVY, bold=True),
    ], base_size=13.5)
    add_note_box(
        slide,
        left=Inches(0.65),
        top=Inches(1.52),
        width=Inches(8.70),
        height=Inches(3.00),
        fill=LGRAY,
        line=NAVY,
        base_size=10.5,
        items=[plain(line, size=10.5, color=DGRAY) for line in code_lines],
    )
    if note:
        add_note_box(
            slide,
            left=Inches(0.65),
            top=Inches(4.58),
            width=Inches(8.70),
            height=Inches(0.47),
            fill=PALE_GOLD,
            line=NAVY,
            base_size=10,
            items=[plain(note, size=10, color=DGRAY, bold=True)],
        )


DECKS = [
    {
        "file": "01-computers-and-operating-systems.pptx",
        "meeting": "Meeting 1",
        "topic": "Computers and Operating Systems",
        "agenda": [
            "Course aims, scope, and learning method",
            "Computer history and how an OS works",
            "Binary numbers, bits, and bytes",
            "Computer programs in civil engineering",
            "Script demo: Excel VBA and AutoCAD VBA",
        ],
        "outcomes": [
            "Describe the course aim and the predict-run-explain learning cycle.",
            "Map input, CPU, memory, storage, output, and the OS.",
            "Convert a simple binary number to decimal.",
            "Distinguish existing applications, programs, and automation scripts.",
        ],
        "schedule": [
            ["0-10", "Learning contract and civil engineering cases", "Opening question"],
            ["10-22", "A concise history of computers", "Four broad stages"],
            ["22-38", "Computer components and the OS", "Processing diagram"],
            ["38-53", "Binary, bits, bytes, rounding", "Convert 1101"],
            ["53-68", "Software vs scripts in civil engineering", "Use-adapt-build choice"],
            ["68-82", "Excel VBA and AutoCAD VBA demo", "Observe input-process-output"],
            ["82-100", "Quiz, discussion, exit ticket", "Three core answers"],
        ],
        "sections": [
            ("How We Learn", [
                bullet("Main aim: computational thinking for civil engineering problems."),
                bullet("Class cycle: **predict** the result, **run** the example, then **explain** why it happens."),
                bullet("The focus is not syntax memorisation; it is problem breakdown and result checking."),
                bullet("Each meeting ends with a small example students can run themselves."),
            ]),
            ("Computer and OS", [
                bullet("The CPU executes instructions; memory holds active data."),
                bullet("Storage keeps files; input-output devices connect users and the system."),
                bullet("The OS manages files, processes, memory, users, devices, and the interface."),
                bullet("Excel and AutoCAD are applications; VBA is automation inside those applications."),
            ]),
            ("Numbers in Computers", [
                bullet("One bit has two states: 0 or 1; eight bits form one byte."),
                bullet("Binary digits use weights 1, 2, 4, 8, 16, and so on."),
                bullet("Example: 1101 = 8 + 4 + 0 + 1 = 13."),
                bullet("Fractional calculations may have tiny rounding differences; engineering checks need tolerance."),
            ]),
            ("Software vs Script", [
                bullet("Use validated software for high-risk analysis and standard engineering tasks."),
                bullet("Create a small script for repeated, local, easily verifiable work."),
                bullet("A practical route is to use existing software and automate only repetitive parts."),
                bullet("Common risks: wrong units, wrong formulas, and overwritten data."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Sub ExcelDemo()",
            "    Range(\"A1\").Value = \"Length (m)\"",
            "    Range(\"B1\").Value = 10",
            "    Range(\"A2\").Value = \"Width (m)\"",
            "    Range(\"B2\").Value = 2",
            "    Range(\"A3\").Value = \"Area (m2)\"",
            "    Range(\"B3\").Value = Range(\"B1\").Value * Range(\"B2\").Value",
            "End Sub",
        ],
        "activity": [
            bullet("Change length and width, then predict which cell will change."),
            bullet("Mark the input, process, and output on the spreadsheet."),
            bullet("Discuss: when is a formula enough, and when does it deserve a macro?"),
        ],
        "quiz": [
            "Convert binary 10110 to decimal and show the weight of every digit.",
            "For 500 channel segments every week, choose software, a large application, or a small Excel script. Give two reasons and one risk.",
            "In the Excel demo, identify input, process, output, application, scripting language, and OS.",
        ],
    },
    {
        "file": "02-excel-fundamentals.pptx",
        "meeting": "Meeting 2",
        "topic": "Excel Fundamentals",
        "agenda": ["Cells, worksheets, workbooks", "Relative and absolute references", "Data and number types", "Errors, zero, and iteration", "Basic Excel functions"],
        "outcomes": [
            "Distinguish values, labels, formulas, and display formats.",
            "Use relative and absolute references correctly.",
            "Read Excel errors as debugging clues.",
            "Build a small calculation table that can be checked manually.",
        ],
        "schedule": [
            ["0-10", "Warm-up: spreadsheet as a calculation model", "One example cell"],
            ["10-28", "Cells, ranges, worksheets, workbooks", "Workbook map"],
            ["28-45", "Relative and absolute references", "Copied formula"],
            ["45-60", "Data and number types", "Number vs text check"],
            ["60-75", "Errors, zero, and iteration", "Trace #DIV/0!"],
            ["75-90", "Basic Excel functions", "Mini table"],
            ["90-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("Excel as a Model", [
                bullet("A cell stores a value, label, or formula; formatting only controls display."),
                bullet("A formula always starts with an equal sign."),
                bullet("A range helps us read data as a table, not isolated cells."),
                bullet("A good model separates inputs, processes, and outputs."),
            ]),
            ("Cell References", [
                bullet("A1 changes when a copied formula moves because it is relative."),
                bullet("$A$1 locks both column and row."),
                bullet("A$1 locks the row; $A1 locks the column."),
                bullet("In Excel, F4 toggles reference locking."),
            ]),
            ("Errors and Zero", [
                bullet("#DIV/0! means the divisor is blank or zero."),
                bullet("#VALUE! often means a number is being read as text."),
                bullet("Zero is a valid value; blank means no data yet."),
                bullet("Debugging starts from input cells, not the final output."),
            ]),
            ("Iteration in Excel", [
                bullet("Iteration means repeated calculation until a condition is met or the change is small."),
                bullet("Civil example: adjust a dimension until capacity satisfies design discharge."),
                bullet("Limit iteration so the process cannot run forever."),
                bullet("Always compare the result with a simple manual check."),
            ]),
        ],
        "activity": [
            bullet("Create a length-width-area table for three rows of data."),
            bullet("Copy the area formula and observe which references change."),
            bullet("Create one #DIV/0! error, then explain its cause."),
        ],
        "quiz": [
            "What is the difference among the cell contents 10, '10', and =5+5?",
            "When should $B$1 be used in a formula copied across many rows?",
            "Why should zero not always be treated the same as a blank cell?",
        ],
    },
    {
        "file": "03-excel-if.pptx",
        "meeting": "Meeting 3",
        "topic": "Excel Functions and IF",
        "agenda": ["Built-in Excel functions", "TRUE/FALSE logic", "Single IF", "Nested IF", "Student grade case"],
        "outcomes": [
            "Write Excel functions with correct arguments.",
            "Read logical expressions as TRUE or FALSE.",
            "Use single IF and nested IF formulas.",
            "Create A, B, C, D, E grade rules consistently.",
        ],
        "schedule": [
            ["0-12", "Formula and reference review", "Prediction"],
            ["12-28", "SUM, AVERAGE, MIN, MAX", "Data summary"],
            ["28-43", "Logical operators", "TRUE/FALSE"],
            ["43-60", "Single IF", "Pass/fail"],
            ["60-78", "Nested IF", "A-B-C-D-E"],
            ["78-92", "Student grade exercise", "Grade table"],
            ["92-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("Excel Functions", [
                bullet("A function turns inputs into an output: =FUNCTION_NAME(arguments)."),
                bullet("SUM adds; AVERAGE finds the mean; MIN/MAX find limits."),
                bullet("Arguments may be numbers, cells, or ranges."),
                bullet("Function names make the intent of a calculation visible."),
            ]),
            ("TRUE/FALSE Logic", [
                bullet("An expression such as B2>=60 returns TRUE or FALSE."),
                bullet("AND requires all conditions to be true; OR only needs one."),
                bullet("Grade boundaries must be consistent and non-overlapping."),
                bullet("Clear logic matters more than a short-looking formula."),
            ]),
            ("Single IF", [
                bullet("Pattern: =IF(condition, value_if_true, value_if_false)."),
                bullet("Example: =IF(B2>=60, \"Pass\", \"Fail\")."),
                bullet("Always test boundaries: 59, 60, and 61."),
                bullet("Boundary cases are where many bugs appear."),
            ]),
            ("Nested IF", [
                bullet("Nested IF is used when there are more than two categories."),
                bullet("Example: A >= 80, B >= 70, C >= 60, D >= 50, otherwise E."),
                bullet("Start from the highest boundary to keep the formula readable."),
                bullet("Use a lookup table when the rules become long."),
            ]),
        ],
        "activity": [
            bullet("Create final score and letter grade columns for 8 students."),
            bullet("Test scores 49, 50, 59, 60, 69, 70, 79, and 80."),
            bullet("Explain one IF formula line in ordinary language."),
        ],
        "quiz": [
            "Predict the result of =IF(75>=70,\"B\",\"C\").",
            "Why must score 80 be tested when creating grade A?",
            "Write the nested IF grading rule A, B, C, D, E in sentences.",
        ],
    },
    {
        "file": "04-algorithms.pptx",
        "meeting": "Meeting 4",
        "topic": "Algorithms and Elements",
        "agenda": ["Problems into steps", "Algorithm definition", "Input-process-output", "Flowcharts", "Pseudocode and programs"],
        "outcomes": [
            "Break a calculation problem into ordered steps.",
            "Recognise input, process, decision, loop, and output.",
            "Write simple pseudocode before Excel formulas or VBA.",
            "Connect an algorithm to a computer program.",
        ],
        "schedule": [
            ["0-10", "Daily tasks and civil calculations", "Manual steps"],
            ["10-25", "Algorithm definition", "Good algorithm traits"],
            ["25-43", "Input-process-output", "IPO sketch"],
            ["43-62", "Flowcharts", "Main symbols"],
            ["62-80", "Pseudocode", "Human-readable version"],
            ["80-93", "Translate to Excel/VBA", "Small program"],
            ["93-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("Problem to Steps", [
                bullet("A computer does not understand a general intention; it executes specific instructions."),
                bullet("Our job is to turn a problem into an unambiguous sequence of steps."),
                bullet("An algorithm may be written as sentences, a flowchart, or pseudocode."),
                bullet("Before coding, make sure the manual example is correct."),
            ]),
            ("Algorithm Elements", [
                bullet("Input: required data, including units and reasonable limits."),
                bullet("Process: formulas, transformations, decisions, or repetition."),
                bullet("Output: the result that the user will use."),
                bullet("Validation: a reference check to decide whether the result makes sense."),
            ]),
            ("Flowcharts", [
                bullet("Oval: start/end; parallelogram: input/output."),
                bullet("Rectangle: process; diamond: decision."),
                bullet("Arrows show execution order."),
                bullet("A flowchart helps us see branches and loops before coding."),
            ]),
            ("Pseudocode", [
                bullet("Pseudocode is an algorithm written in semi-formal language."),
                bullet("Example: read length, read width, area = length * width, display area."),
                bullet("The syntax need not be perfect, but the steps must be executable."),
                bullet("Pseudocode bridges an idea to Excel or VBA."),
            ]),
        ],
        "activity": [
            bullet("Use a rectangular cross-section area case."),
            bullet("Write input, process, and output in three lines."),
            bullet("Convert it into pseudocode and a simple flowchart."),
        ],
        "quiz": [
            "Identify input, process, and output for concrete beam volume.",
            "Why should an algorithm be tested with simple numbers before coding?",
            "When is a flowchart more helpful than pseudocode?",
        ],
    },
    {
        "file": "05-vba-and-linear-macros.pptx",
        "meeting": "Meeting 5",
        "topic": "VBA and Linear Macros",
        "agenda": ["Excel components", "Cell formulas", "VBA editor", "Record macro", "Linear programs"],
        "outcomes": [
            "Explain workbook, worksheet, range, macro, module, and subroutine.",
            "Record a simple macro and read the generated code.",
            "Modify recorded code into a cleaner linear program.",
            "Run an Excel calculation sequence from VBA.",
        ],
        "schedule": [
            ["0-12", "Excel as a VBA host", "Object map"],
            ["12-28", "Cell formula vs macro", "Workflow comparison"],
            ["28-45", "VBA editor and module", "First Sub"],
            ["45-62", "Record macro", "Recorded code"],
            ["62-80", "Clean a linear macro", "Readable macro"],
            ["80-93", "Simple calculation exercise", "Run macro/button"],
            ["93-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("Excel Components", [
                bullet("A workbook contains worksheets; a worksheet contains cells and ranges."),
                bullet("Formulas are good for open, visible calculation models."),
                bullet("Macros are good for repeated steps: clean data, format, calculate, report."),
                bullet("VBA runs inside a host application such as Excel or AutoCAD."),
            ]),
            ("Recording Macros", [
                bullet("Record Macro captures manual operations as VBA code."),
                bullet("Recorded code is often too long, but useful for learning objects and commands."),
                bullet("Read the recording as a sequence of actions, not as final code."),
                bullet("After understanding it, rename, remove noise, and structure the code."),
            ]),
            ("Linear Programs", [
                bullet("A linear program runs from the first line to the last line without branching."),
                bullet("Starter pattern: write labels, read inputs, calculate, write output."),
                bullet("Use `Option Explicit` so mistyped variables are detected early."),
                bullet("Test with small numbers that can be checked manually."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Sub CalculateSlabArea()",
            "    Range(\"A1\").Value = \"Length\"",
            "    Range(\"A2\").Value = \"Width\"",
            "    Range(\"A3\").Value = \"Area\"",
            "    Range(\"B3\").Value = Range(\"B1\").Value * Range(\"B2\").Value",
            "End Sub",
        ],
        "activity": [
            bullet("Record a macro that creates a table header and simple formatting."),
            bullet("Open the recorded code and remove unnecessary steps."),
            bullet("Add one area calculation line using inputs B1 and B2."),
        ],
        "quiz": [
            "What is the difference between a cell formula and a VBA macro?",
            "Why does recorded macro code usually need cleaning?",
            "What does `Option Explicit` do while learning VBA?",
        ],
    },
    {
        "file": "06-input-output-and-modularity.pptx",
        "meeting": "Meeting 6",
        "topic": "Input-Output and Modularity",
        "agenda": ["Cells and Range", "Local and public variables", "Arrays", "Functions", "Subroutines"],
        "outcomes": [
            "Read and write data using Cells and Range.",
            "Use variables to store temporary data.",
            "Distinguish a Sub from a Function.",
            "Create a more modular small program that is easier to test.",
        ],
        "schedule": [
            ["0-10", "Linear macro review", "Input-output pattern"],
            ["10-28", "Cells, Range, read, write", "Output cell"],
            ["28-45", "Local and public variables", "Scope"],
            ["45-60", "Arrays", "Repeated data"],
            ["60-78", "Function", "Modular formula"],
            ["78-92", "Subroutine", "Main procedure"],
            ["92-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("VBA Input-Output", [
                bullet("`Range(\"B2\").Value` reads or writes one specific cell."),
                bullet("`Cells(row, column)` is convenient for loops because indices can change."),
                bullet("Separate input and output areas so old data is not overwritten accidentally."),
                bullet("Every small program needs reference input and output examples."),
            ]),
            ("Variables and Scope", [
                bullet("A variable stores temporary values: length, width, area, discharge."),
                bullet("A local variable exists only inside the procedure where it is declared."),
                bullet("A public variable can be shared by procedures, but is harder to trace."),
                bullet("For beginners, local variables usually make debugging clearer."),
            ]),
            ("Arrays", [
                bullet("An array stores many values under one name and index."),
                bullet("It fits lists of lengths, discharges, elevations, or student scores."),
                bullet("Indices let a program read repeated data inside a loop."),
                bullet("Record the first and last index to avoid out-of-range errors."),
            ]),
            ("Function and Sub", [
                bullet("A Function returns a value, for example `Area(L, W)`."),
                bullet("A Subroutine performs actions, for example reading a table and writing a report."),
                bullet("Modularity lets us test code part by part."),
                bullet("Clear function names are the first layer of documentation."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Function RectangleArea(l As Double, w As Double) As Double",
            "    RectangleArea = l * w",
            "End Function",
            "",
            "Sub WriteArea()",
            "    Dim l As Double, w As Double",
            "    l = Range(\"B1\").Value",
            "    w = Range(\"B2\").Value",
            "    Range(\"B3\").Value = RectangleArea(l, w)",
            "End Sub",
        ],
        "activity": [
            bullet("Create a Function for block volume: length * width * height."),
            bullet("Create a Sub that reads B1:B3 and writes the result to B4."),
            bullet("Test with 2, 3, and 4 so the manual result is 24."),
        ],
        "quiz": [
            "When is Cells(i, 2) better than Range(\"B2\")?",
            "What is the main difference between a Function and a Subroutine?",
            "Why should public variables be used carefully?",
        ],
    },
    {
        "file": "07-userforms-and-controls.pptx",
        "meeting": "Meeting 7",
        "topic": "UserForms and Controls",
        "agenda": ["UserForm", "Controls", "Properties", "Events", "Pre-midterm mini app"],
        "outcomes": [
            "Explain a UserForm as a small user interface.",
            "Recognise common controls: Label, TextBox, ComboBox, CommandButton.",
            "Set important properties and write a simple event.",
            "Build a small calculation app that can be verified manually.",
        ],
        "schedule": [
            ["0-10", "Function/Sub review", "Program map"],
            ["10-25", "UserForm as an interface", "Form sketch"],
            ["25-42", "Controls and properties", "Control names"],
            ["42-60", "Events", "Button click"],
            ["60-78", "Area/volume form example", "Demo"],
            ["78-93", "Pre-midterm integrative practice", "Mini app"],
            ["93-100", "Quiz and exit ticket", "Three core answers"],
        ],
        "sections": [
            ("UserForm", [
                bullet("A UserForm is a small input-output window for users."),
                bullet("A form is useful when users do not need to see the whole worksheet."),
                bullet("A form does not replace validation; inputs still need checking."),
                bullet("Simple design is better: few inputs, clear button, readable output."),
            ]),
            ("Controls and Properties", [
                bullet("Label explains; TextBox receives input; Button runs a command."),
                bullet("Properties set the name, caption, initial value, colour, and size."),
                bullet("Control names should be meaningful: txtLength, txtWidth, cmdCalculate."),
                bullet("Caption is for users; Name is for code."),
            ]),
            ("Events", [
                bullet("An event is something that triggers code, such as clicking a button."),
                bullet("Event code usually reads input, validates, calls a function, then displays output."),
                bullet("If input is blank or not numeric, show a clear message."),
                bullet("Short event procedures are easier to test and maintain."),
            ]),
            ("Pre-Midterm Practice", [
                bullet("Choose one small case: area, volume, unit conversion, or student grading."),
                bullet("Prepare the algorithm, worksheet/form, code, and manual reference result."),
                bullet("Verify at least one easy case and one boundary case."),
                bullet("Submit the small program plus notes on how the result was checked."),
            ]),
        ],
        "code": [
            "Private Sub cmdCalculate_Click()",
            "    Dim l As Double, w As Double",
            "    l = CDbl(txtLength.Value)",
            "    w = CDbl(txtWidth.Value)",
            "    lblResult.Caption = \"Area = \" & Format(l * w, \"0.00\")",
            "End Sub",
        ],
        "activity": [
            bullet("Design a form with two inputs, one button, and one output label."),
            bullet("Keep the calculation function separate from the button event."),
            bullet("Test normal input, blank input, and zero input."),
        ],
        "quiz": [
            "What is the difference between Name and Caption on a Button?",
            "Why should event code stay short?",
            "Name two test cases for a small area calculation app.",
        ],
    },
]


def build_deck(data):
    src = pick_source(TEMPLATE)
    out_path = os.path.join(OUT_DIR, data["file"])
    tmp_path = out_path + ".tmp"
    shutil.copy(src, tmp_path)
    prs = Presentation(tmp_path)
    delete_all_slides(prs)

    title_slide(
        prs,
        COURSE,
        SUBTITLE,
        data["topic"],
        LECTURER,
        data["meeting"],
        DATE_LABEL,
        dept="Department of Civil and Environmental Engineering, Faculty of Engineering, UGM",
    )
    agenda_slide(prs, "Learning Topics", data["agenda"])
    content_slide(prs, "Learning Outcomes", [bullet(x, kind="num", size=15.5) for x in data["outcomes"]], base_size=15.5)
    schedule_slide(prs, data["topic"], data["schedule"])

    divider_slide(prs, 1, "Core Concepts")
    for title, items in data["sections"]:
        content_slide(prs, title, items, base_size=15.5)

    if data.get("code"):
        divider_slide(prs, 2, "Short Demo")
        code_slide(prs, "VBA Script Example", data["code"], note="Observation focus: input, process, output, and data location.")

    content_slide(prs, "Class Exercise", data["activity"], base_size=16)
    quiz_slide(prs, data["quiz"])
    content_slide(prs, "Exit Ticket", [
        bullet("Write one concept that is now clear."),
        bullet("Write one part that is still confusing."),
        bullet("Write one civil engineering use case you want to try."),
    ], base_size=17)
    closing_slide(prs, "_Start from a small problem, make the steps clear, then validate the result._", "- Algorithms and Computer Programming")

    build_and_save(prs, tmp_path, out_path)
    return out_path


def main():
    print("Building English PowerPoint decks...")
    for deck in DECKS:
        build_deck(deck)


if __name__ == "__main__":
    main()
