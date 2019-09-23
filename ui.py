# -*- coding: utf-8 -*-

from __future__ import print_function, unicode_literals
import regex

from pprint import pprint
from PyInquirer import style_from_dict, Token, prompt
from PyInquirer import Validator, ValidationError
import xlwt
import xlrd
from xlutils.copy import copy



#get the attendance spreadsheet

#open file
path = ('attendance.xls')

#reading
rb = xlrd.open_workbook(path)
r_sheet = rb.sheet_by_index(0)

wb = copy(rb)
w_sheet = wb.get_sheet(0)


#current meeting
current_meeting = 2

style = style_from_dict({
    Token.QuestionMark: '#E91E63 bold',
    Token.Selected: '#673AB7 bold',
    Token.Instruction: '',  # default
    Token.Answer: '#2196f3 bold',
    Token.Question: '',
})

class NumberValidator(Validator):
    def validate(self, document):
        try:
            int(document.text)
        except ValueError:
            raise ValidationError(
                message='Please enter a number',
                cursor_position=len(document.text))  # Move cursor to end


class EmailValidator(Validator):
    def validate(self, document):
        ok = False

        t = document.text

        t += "@students.d211.org"

        for i in range(r_sheet.nrows):
            email = r_sheet.cell_value(i, 0)
            if t == email: #same email
                ok = True

        if not ok:
            raise ValidationError(
                message='please enter an actual email',
                cursor_position=len(document.text))  # Move cursor to end


def getRowFromEmail(e):

    e += "@students.d211.org"

    for i in range(r_sheet.nrows):
        email = r_sheet.cell_value(i, 0)
        if email == e:
            return i
    return -1

def updateData():
    global rb, r_sheet, wb, w_sheet

    rb = xlrd.open_workbook(path)
    r_sheet = rb.sheet_by_index(0)

    wb = copy(rb)
    w_sheet = wb.get_sheet(0)

    print(r_sheet.nrows)


def login():
    print('LOGIN TO THE DATAFRAME FELLOW HACKER')
    print('')

    login_question = [
        {
            'type': 'list',
            'name': 'logged_in',
            'message': 'have you logged in before?',
            'choices': ["yes", "no"],
            'filter': lambda val: val.lower()
        }
    ]
    answers = prompt(login_question)

    if(answers['logged_in'] == "yes"):
        name = [
            {
                'type': 'input',
                'name': 'email',
                'message': 'what is your email (without @students.d211.org)',
                'validate': EmailValidator
            }
        ]

        answers = prompt(name)

        email = answers['email']

        #code here for adding to excel

        row = getRowFromEmail(email)

        column = current_meeting + 2

        w_sheet.write(row, column, "1")

        #code for getting name

        name = r_sheet.cell_value(row, 1)

        wb.save(path)

        updateData()

        print("")
        print("Welcome " + name + "!")
        print("")
        print("")
        print("")
        login()

    else:
        print("hello, newcomer")
        new_info = [
            {
                'type': 'input',
                'name': 'name',
                'message': 'what is your name',
            },
            {
                'type': 'input',
                'name': 'email',
                'message': 'what is your school email (with @students.d211.org)',
            },
            {
                'type': 'list',
                'name': 'grade',
                'message': 'what is your grade',
                'choices': ['Freshmen', 'Sophomore', 'Junior', 'Senior'],
                'filter': lambda val: val.lower()
            }
        ]

        new_person = prompt(new_info)

        #add code for creating new entry

        new_row = r_sheet.nrows

        print(new_row)

        w_sheet.write(new_row, 0, new_person["email"])
        w_sheet.write(new_row, 1, new_person["name"])
        w_sheet.write(new_row, 2, new_person["grade"])

        for z in range(current_meeting):

            #this updates the meeting header, not needed since we will do it maunually
            # m = "meeting " + str(z + 1)
            #
            # w_sheet.write(0, 3 + z, m) #add the header on the first line

            val = "0"

            if z == current_meeting - 1: #current meeting
                val = "1"

            w_sheet.write(new_row, 3 + z, val) #change the value underneath to either 0, since they missed all the prvious meetings, or 1

        wb.save(path)

        updateData()

        print("")
        print("Hope you have fun during this meeting!")
        print("")
        print("")
        print("")
        login()

login()
