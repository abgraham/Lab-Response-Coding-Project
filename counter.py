#!/usr/bin/python -tt
import sys
import re
import string
import xlwt
import docx
from docx import Document

class color:
    BOLD = '\033[1m'
    END = '\033[0m'

# Reads the responses from the given file and checks to see if they contain
# any of the target words. If a target word is found, it codes a "1" into the appropriate
# worksheet cell.
def read_responses(f, targetList, sheet):
    lineCount = 0
    for line in f:
        lineCount += 1
        sheet.write(lineCount, 0, lineCount)
        print "\n" + line,
        print color.BOLD + "Words coded for this response:" + color.END
        targetCount = 1
        for target in targetList:
            if string.find(line, target) != -1:
                print target
                sheet.write(lineCount, targetCount, "1")
            targetCount += 1

# Reads all the targets from the given file and puts them into a list.
def read_targets(f):
     targetList = []
     for line in f:
         for word in line.split():
             targetList.append(word)
     return targetList

# Sets up the workbook with all of the targets listed on the x axis.
def initialize_workbook(targetList, sheet):
    count = 1
    for target in targetList:
        sheet.write(0, count, target)
        count += 1

# Main should be given two arguments. The first is a file with the participants'
# answers, and the second is a file with a list of target words.
def main():
    args = sys.argv
f1 = open(args[1], 'r')
    f2 = open(args[2], 'r')
    targetList = read_targets(f2)
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet 1", cell_overwrite_ok = True)
    initialize_workbook(targetList, sheet)
    read_responses(f1, targetList, sheet)
    workbook.save("Response Coding");
    document = Document('Test-doc.docx')
    count = 1

if __name__ == '__main__':
    main()
