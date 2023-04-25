import os
from openpyxl import load_workbook, Workbook,worksheet
import pandas as pd

def listPasswords():
    myFileName=r'passwords.xlsx'
    print(f"{pd.read_excel(myFileName)}\n\n")
    input("Press Enter to return to menu...\n")

def getInput():
    while True:
        try:
            website = input("Enter website/name for the password i.e., 'google account': ")
            username= input("Enter Username: ")
            password= input("Enter Password: ")
            return [website,username,password]
        except:
            print("wrong input try again")
def addPassword():
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']
    newRowLocation = ws.max_row +1
    myInput = getInput()
    ws.cell(column=1,row=newRowLocation, value=myInput[0])
    ws.cell(column=2,row=newRowLocation, value=myInput[1])
    ws.cell(column=3,row=newRowLocation, value=myInput[2])
    wb.save(filename=myFileName)
    wb.close()

def createXlsxFile():
    filepath = "passwords.xlsx"
    colnames = ['Website', 'Username', 'Password']
    wb = Workbook()
    wb.save(filepath)

    wb = load_workbook(filename=filepath)
    ws = wb['Sheet']
    ws.cell(column=1,row=1, value=colnames[0])
    ws.cell(column=2,row=1, value=colnames[1])
    ws.cell(column=3,row=1, value=colnames[2])
    wb.save(filename=filepath)
    wb.close()

def deletePassword():
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']
    try:
        row=int(input("Enter the row number of the password you wish to delete: "))+2
    except:
        print("Must be an integer")
    ws.delete_rows(row,1)
    wb.save(filename=myFileName)
    wb.close()

def menu():
    try:
        choice = int(input("1. List all passwords\n2. Add a password\n3. Delete a password\n4. Find a password\n5. exit\n\ninput:"))
        print("\n\n")
        if 1<=choice<=5:
            if choice == 1:
                listPasswords()
            elif choice==2:
                addPassword()
            elif choice==3:
                deletePassword()

            elif choice==5:
                return 1
        else:
            print("Invalid number")
    except:
        print("Invalid choice")
def start():

    while True:
        if os.path.isfile('passwords.xlsx'):
            start=menu()
            if start == 1:
                print("bye")
                break
        else:
            createXlsxFile()

if __name__ == "__main__":
    start()