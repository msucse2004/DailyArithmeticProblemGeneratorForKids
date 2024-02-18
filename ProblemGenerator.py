import copy
import os.path
import subprocess
from datetime import datetime
from fractions import Fraction
import numpy as np
import openpyxl
import pandas as pd
import random
import StringTable
from PyQt5.QtGui import QIcon, QFont
from openpyxl import Workbook
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QStatusBar, QCheckBox, QHBoxLayout, QVBoxLayout, \
    QToolTip
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, qApp
from PyQt5.QtGui import QIcon

from StringTable import *
__INPUT_DEBUG_MODE__ = False
__NEGATIVE_CONSTRAINT__ = True
__FRACTION_CONSTRAINT__ = True

__OPERATOR_PRINT_MAP__:dict = {"+": "+", "-": "-", "*": "x", "/": "÷"}
__OPERATOR_MAP__ = {"덧셈": "+", "뺄셈": "-", "곱셈": "*", "나눗셈": "/"}
__OPERATOR_INPUT_MAP__ = {1:"덧셈", 2:"뺄셈", 3:"곱셈", 4:"나눗셈"}
__OUTPUT_CSV_PATH__ = "./output"
# 중위표기법 후위표기법
from openpyxl.worksheet.worksheet import Worksheet


def Infix2Postfix(infix_list: list):  # expr: 입력 리스트(중위 표기식)
    """
    중위 표기법 리스트를 후위 표기법 리스트로 변환하는 함수

    Args:
        infix_list: 숫자와 기호가 들어있는 중위 표기법 리스트

    Returns:
        postfix_list: 후위 표기법 리스트
    """
    prec = {'+': 1, '-': 1, '*': 2, '/': 2, '^': 3}
    stack = []
    postfix_list = []
    for token in infix_list:
        if isinstance(token, int) or isinstance(token, float):
            postfix_list.append(token)
        elif token == '(':
            stack.append(token)
        elif token == ')':
            while stack[-1] != '(':
                postfix_list.append(stack.pop())
            stack.pop()
        else:
            while len(stack) > 0 and prec[stack[-1]] >= prec[token]:
                postfix_list.append(stack.pop())
            stack.append(token)
    while len(stack) > 0:
        postfix_list.append(stack.pop())
    return postfix_list

def CalculationDecimal(x):
    stack = []
    for i in x:
        if i == '+':
            stack.append(stack.pop()+stack.pop())
        elif i == '-':
            stack.append(-(stack.pop()-stack.pop()))
        elif i == '*':
            stack.append(stack.pop()*stack.pop())
        elif i == '/':
            divide = stack.pop()
            try:
                stack.append(stack.pop()/divide)
            except ZeroDivisionError:
                return None
        else:
            #stack.append(int(i))
            stack.append(i)
    return stack.pop()

def CalculationFraction(x):
    stack = []
    for i in x:
        if i == '+':
            operand2 = stack.pop()
            operand1 = stack.pop()
            if isinstance(operand1, Fraction):
                fraction1 = operand1
            else:
                fraction1 = Fraction(operand1)
            if isinstance(operand2, Fraction):
                fraction2 = operand2
            else:
                fraction2 = Fraction(operand2)
            stack.append(fraction1 + fraction2)
        elif i == '-':
            operand2 = stack.pop()
            operand1 = stack.pop()
            if isinstance(operand1, Fraction):
                fraction1 = operand1
            else:
                fraction1 = Fraction(operand1)
            if isinstance(operand2, Fraction):
                fraction2 = operand2
            else:
                fraction2 = Fraction(operand2)
            stack.append(fraction1 - fraction2)
        elif i == '*':
            operand2 = stack.pop()
            operand1 = stack.pop()
            if isinstance(operand1, Fraction):
                fraction1 = operand1
            else:
                fraction1 = Fraction(operand1)
            if isinstance(operand2, Fraction):
                fraction2 = operand2
            else:
                fraction2 = Fraction(operand2)
            stack.append(fraction1 * fraction2)
        elif i == '/':
            divisor = stack.pop()
            if divisor == 0:
                #raise ZeroDivisionError("Division by zero")
                return None
            operand = stack.pop()
            if isinstance(operand, Fraction):
                fraction = operand
            else:
                fraction = Fraction(operand)
            quotient = fraction / divisor
            # Simplify the quotient to its lowest terms
            stack.append(quotient.limit_denominator())
        else:
            try:
                # Attempt to convert to a fraction first
                stack.append(Fraction(i))
            except ValueError:
                # If conversion fails, treat as an integer
                stack.append(int(i))

    return stack.pop()

def convert_string_equation(problem_list: list) -> str:
    equation = ""
    equation_printable = ""

    for token in problem_list:
        if isinstance(token, (int, float, complex)):
            #equation += str(token)
            equation_printable += str(token)
        else:
            #print(f"tocken {token} -> {__OPERATOR_PRINT_MAP__.get(token)}")
            #equation += token
            equation_printable += f" {__OPERATOR_PRINT_MAP__.get(token)} "

    equation_printable += " ="
    return equation_printable


def check_constraint(ans: float) -> bool:
    if __NEGATIVE_CONSTRAINT__:
        if ans < 0:
            return False
    if __FRACTION_CONSTRAINT__:
        if ans % 1 != 0:
            return False

    return True

def adjust_column_style(filepath: str) -> None:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    # 칼럼 너비 조절
    for i, col in enumerate(ws.columns):
        if i % 2 == 0:
            max_length = 0
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)

                except TypeError:
                    pass
                cell.font = openpyxl.styles.Font(name='Arial', size=18, bold=True)
                cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='dotted'))
                #cell.font = openpyxl.styles.Font(name='Arial', size=20)
            adjusted_width = (max_length + 2) * 1.5
            ws.column_dimensions[col[i].column_letter].width = adjusted_width
            ws.column_dimensions[col[i].column_letter].best_fit = True
        else:
            max_length = 0
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)

                except TypeError:
                    pass
                # cell.font = openpyxl.styles.Font(name='Arial', size=14, bold=True)

                cell.font = openpyxl.styles.Font(name='Arial', size=14)
                cell.alignment = openpyxl.styles.Alignment(horizontal='left')
            adjusted_width = (max_length + 5) * 1.7
            if (i < 5) :
                ws.column_dimensions[col[i].column_letter].width = adjusted_width

    page_setup = ws.page_setup
    page_setup.fitToWidth = True

    page_margins = ws.page_margins
    page_margins.left = 0.0
    page_margins.right = 0.5
    page_margins.top = 1.0
    page_margins.bottom = 0.5
    page_margins.header = 0.5
    page_margins.footer = 0.5

    curr_date = datetime.now()
    formatted_date = curr_date.strftime("%Y-%m-%d")
    ws.oddHeader.center.text = f"Date: {formatted_date}                  Name:                           Score:"
    ws.oddHeader.center.size = 14

    # 엑셀 파일 저장
    wb.save(filepath)


def main():
    num_operand = 0
    problem_count = 0
    pd_problem_set = pd.DataFrame()
    if __INPUT_DEBUG_MODE__:
        num_operand = 3
    else:
        num_operand = int(input("몇개의 숫자 연산을 할까요? "))

    if num_operand <= 1:
        print(f"땡.. 1개는 안되요!!! 다시 해주세요..")
    else:
        print(f"\n당신은 {num_operand} 개의 연산을 입력 하셨습니다")

    operands_list = []
    if __INPUT_DEBUG_MODE__:
        #operands_list = [1, 1, 1]
        operands_list = [3, 2]
    else:
        for index in range(num_operand):
            tmp_num = int(input(f"{index+1} 번째 숫자는 몇자리까지 할까요? "))
            operands_list.append(tmp_num)

    print(operands_list)

    operator_list = []

    if __INPUT_DEBUG_MODE__:
        #operator_list = ["덧셈", "뺄셈", "곱셈", "나눗셈"]
        operator_list = ["나눗셈"]
    else:
        tmp_op = input("어떤 연산을 할까요? 덧셈-1, 뺄셈-2, 곱셈-3, 나눗셈-4 예) 덧셈과 뺄셈: 1, 2 ")
        tmp_operator_list = tmp_op.split(",")
        for op in tmp_operator_list:
            operator_list.append(__OPERATOR_INPUT_MAP__.get(int(op)))

    print(f"{operator_list} 의 연산 조합 문제를 만들겠습니다.")

    if __INPUT_DEBUG_MODE__:
        problem_count = 10
    else:
        problem_count = int(input("몇문제를 만들까요? "))
    print(f"네~ {problem_count} 개의 문제를 만들겠습니다.")

    problem = []
    problem_set = {}
    problem_cnt = 0
    #for id in range(problem_count):
    while(1):
        problem.clear()
        problem_generation_violation = False

        for cnt in operands_list:
            rand_num = random.randrange(0, 1 * pow(10, cnt))
            problem.append(rand_num)
            problem.append(__OPERATOR_MAP__.get(random.choice(operator_list)))

        problem = problem[0:-1]
        problem_validation = copy.deepcopy(problem)
        equation_print= convert_string_equation(problem_validation)
        infix_notation = problem_validation

        postfix_notation = Infix2Postfix(infix_notation)
        #answer = CalculationDecimal(postfix_notation)
        answer = CalculationFraction(postfix_notation)


        if answer is not None and check_constraint(answer):
            #print(f"str problem: {equation_print} ans: {int(answer)}")
            problem_cnt += 1
            if __FRACTION_CONSTRAINT__:
                problem_set = {"number": int(problem_cnt), "problem": equation_print, "answer": int(answer)}
            else:
                if answer.denominator == 1:
                    problem_set = {"number": int(problem_cnt), "problem": equation_print, "answer": int(answer)}
                else:
                    problem_set = {"number": int(problem_cnt), "problem": equation_print, "answer": f"{int(answer.numerator/answer.denominator)}과 {answer.numerator%answer.denominator} / {answer.denominator}"}
            pd_problem_set = pd_problem_set.append(problem_set, ignore_index=True)

        if problem_cnt == problem_count:
             break;


        #problem_set.append(copy.deepcopy(problem))
    pd_problem_set['number'] = pd_problem_set['number'].astype(int)
    #if __FRACTION_CONSTRAINT__:
    #    pd_problem_set['answer'] = pd_problem_set['answer'].astype(int)
    #else:
        #pd_problem_set['answer'] = pd_problem_set['answer'].dt.format(f"{0} / {1}")
        #pd_problem_set['answer'] = f"{pd_problem_set['answer'].numerator} / {pd_problem_set['answer'].denominator}"


    pd_problem_set = pd_problem_set[["number", "problem", "answer"]]
    print(pd_problem_set)


    if not os.path.exists(__OUTPUT_CSV_PATH__):
        os.mkdir(__OUTPUT_CSV_PATH__)
    if os.path.exists(os.path.join(__OUTPUT_CSV_PATH__, "problemset.csv")):
        os.remove(os.path.join(__OUTPUT_CSV_PATH__, "problemset.csv"))
    if os.path.exists(os.path.join(__OUTPUT_CSV_PATH__, "problemset.xlsx")):
        os.remove(os.path.join(__OUTPUT_CSV_PATH__, "problemset.xlsx"))
    if os.path.exists(os.path.join(__OUTPUT_CSV_PATH__, "problemset_teacher.xlsx")):
        os.remove(os.path.join(__OUTPUT_CSV_PATH__, "problemset_teacher.xlsx"))
    if os.path.exists(os.path.join(__OUTPUT_CSV_PATH__, "problemset_student.xlsx")):
        os.remove(os.path.join(__OUTPUT_CSV_PATH__, "problemset_student.xlsx"))


    pd_problem_set_1 = pd_problem_set.loc[pd_problem_set.index %3 == 0]
    pd_problem_set_1 = pd_problem_set_1.reset_index()
    pd_problem_set_2 = pd_problem_set.loc[pd_problem_set.index % 3 == 1]
    pd_problem_set_2 = pd_problem_set_2.reset_index()
    pd_problem_set_3 = pd_problem_set.loc[pd_problem_set.index % 3 == 2]
    pd_problem_set_3 = pd_problem_set_3.reset_index()



    #pd_problem_set_first = pd_problem_set.loc[0:pd_problem_set.shape[0]/3-1]
    #pd_problem_set_first = pd_problem_set_first.reset_index()
    #pd_problem_set_second = pd_problem_set.loc[pd_problem_set.shape[0] / 2:]
    #pd_problem_set_second = pd_problem_set_second.reset_index()

    pd_problem_set_final = pd.concat([pd_problem_set_1, pd_problem_set_2, pd_problem_set_3], axis=1, ignore_index=True)
    pd_problem_set_final_teacher = pd_problem_set_final[[2, 3, 6, 7, 10, 11]]
    pd_problem_set_final_student = pd_problem_set_final[[2, 3, 6, 7, 10, 11]]
    pd_problem_set_final_student.loc[:, 3] = np.NaN
    pd_problem_set_final_student.loc[:, 7] = np.NaN
    pd_problem_set_final_student.loc[:, 11] = np.NaN
    pd_problem_set_final_teacher.columns = ["문제", "정답", "문제", "정답", "문제", "정답"]
    pd_problem_set_final_student.columns = ["문제", "정답", "문제", "정답", "문제", "정답"]




    print(f"origin:\n{pd_problem_set}")
    #print(f"first:\n{pd_problem_set_first}")
    #print(f"second:\n{pd_problem_set_second}")
    print(f"final:\n{pd_problem_set_final}")
    #pd_problem_set.to_csv(os.path.join(__OUTPUT_CSV_PATH__, "problemset.csv"), encoding='utf-8', index=False)
    #pd_problem_set.to_excel(os.path.join(__OUTPUT_CSV_PATH__, "problemset.xlsx"), encoding='utf-8', index=False)
    #pd_problem_set_final.to_excel(os.path.join(__OUTPUT_CSV_PATH__, "problemset.xlsx"), encoding='utf-8', index=False, header=False)

    # https://traumees.tistory.com/39

    pd_problem_set_final_teacher.to_excel(os.path.join(__OUTPUT_CSV_PATH__, "problemset_teacher.xlsx"),
                                          encoding='utf-8', index=False, engine="openpyxl")
    pd_problem_set_final_student.to_excel(os.path.join(__OUTPUT_CSV_PATH__, "problemset_student.xlsx"),
                                          encoding='utf-8', index=False, engine="openpyxl")

    adjust_column_style(os.path.join(__OUTPUT_CSV_PATH__, "problemset_teacher.xlsx"))
    adjust_column_style(os.path.join(__OUTPUT_CSV_PATH__, "problemset_student.xlsx"))

    subprocess.call(["explorer", f"{os.getcwd()}\\output"])

class DailyArithmeticGenerator(QWidget):
    def __init__(self):
        # https://wikidocs.net/21933
        super().__init__()
        self.statusBar = QStatusBar()
        self.systemLanguage = "KR"
        self.stringTbl = StringTable()

        self.initUI()

    def initUI(self):
        QToolTip.setFont(QFont('SansSerif', 10))
        self.setWindowTitle(self.stringTbl.findString("windowtitle", self.systemLanguage))
        self.setWindowIcon(QIcon('math_icon.ico'))

        curr_date = datetime.now()
        formatted_date = curr_date.strftime("%Y-%m-%d")
        self.updateStatusBar(formatted_date)

        layout_list = []

        # Options for Problem Generation Constraints
        options_problem_generation_list = []
        checkbox_opt_negative = QCheckBox(self.stringTbl.findString("option_negative", self.systemLanguage), self)
        checkbox_opt_negative.setToolTip(self.stringTbl.findString("option_negative_tip", self.systemLanguage))
        options_problem_generation_list.append(checkbox_opt_negative)

        checkbox_opt_decimal = QCheckBox(self.stringTbl.findString("option_decimal", self.systemLanguage), self)
        checkbox_opt_decimal.setToolTip(self.stringTbl.findString("option_decimal_tip", self.systemLanguage))
        options_problem_generation_list.append(checkbox_opt_decimal)

        checkbox_opt_fraction = QCheckBox(self.stringTbl.findString("option_fraction", self.systemLanguage), self)
        checkbox_opt_fraction.setToolTip(self.stringTbl.findString("option_fraction_tip", self.systemLanguage))
        options_problem_generation_list.append(checkbox_opt_fraction)

        layout_options_prob_gen = QHBoxLayout()
        layout_list.append(layout_options_prob_gen)
        for widget in options_problem_generation_list:
            layout_options_prob_gen.addWidget(widget)

        # Parameters for Problem Generation


        # Options for Exel
        checkbox_opt_bold = QCheckBox(self.stringTbl.findString("option_bold", self.systemLanguage), self)
        layout_options_excel = QHBoxLayout()
        layout_options_excel.addWidget(checkbox_opt_bold)
        layout_list.append(layout_options_excel)

        # Deploy layout
        layout_widget = QVBoxLayout()
        for layout in layout_list:
            layout_widget.addLayout(layout)

        self.setLayout(layout_widget)



        self.move(300, 300)
        self.resize(400, 200)
        self.show()

    def updateStatusBar(self, msg):
        if msg is not None:
            self.statusBar.showMessage(msg)

if __name__ == "__main__":
    #main()
    app = QApplication(sys.argv)
    ex = DailyArithmeticGenerator()
    sys.exit(app.exec_())
