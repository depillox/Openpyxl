#Name: Zavier DePillo
#Email: depillzj@mail.uc.edu
#Assignment Title: Assignment 09
#Course: IS 4010
#Semester/Year: Spring 2023
#Brief Description: In-Class Assignment, openpyxl
#Citations:
#Anything else that's relevant:

from openpyxl import load_workbook
def assignment09():
    wb = load_workbook(filename = 'empl.xlsx')
    sheet = wb['Sheet1']
    #print(sheet['A1'].value), print(sheet['A2'].value)
    #print(sheet['B1'].value), print(sheet['B2'].value)
    #print(sheet['C1'].value), print(sheet['C2'].value)
    
    #I want a list of all the cells in column C
    #That have last names C2:C1001
    names = [] #Create an empty list
    names = [cell.value for cell in sheet['C'][1:]]
    print(names[0]) #print the first cell
    print(names[-1]) #print the last cell
    
    wb.close()