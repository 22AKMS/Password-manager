import os,random,array
from openpyxl import load_workbook, Workbook
import pandas as pd
from cryptography.fernet import Fernet
def generateKey():
    if os.path.isfile('masterKey.key'):
        print('\nYou already have a key')
        input('Press a key to continue...\n')
    else:
        key = Fernet.generate_key()
        with open('masterKey.key', 'wb') as filekey:
            filekey.write(key)
def encryptFile():
    try:
        with open('masterKey.key', 'rb') as filekey:
            key = filekey.read()
        fernet = Fernet(key)
        with open('passwords.xlsx', 'rb') as file:
            original = file.read()
        encrypted = fernet.encrypt(original)
        with open('passwords.xlsx', 'wb') as encrypted_file:
            encrypted_file.write(encrypted)
    except:
        return
def decryptFile():
    try:
        with open('masterKey.key', 'rb') as filekey:
            key = filekey.read()
        fernet = Fernet(key)
        with open('passwords.xlsx', 'rb') as enc_file:
            encrypted = enc_file.read()
        decrypted = fernet.decrypt(encrypted)
        with open('passwords.xlsx', 'wb') as dec_file:
            dec_file.write(decrypted)
    except:
        return
def listPasswords():
    decryptFile()
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']
    if ws.max_row == 1:
        print("No password in the list\n")
        return
    print(f"{pd.read_excel(myFileName)}\n\n")
    encryptFile()
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
    decryptFile()
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
    encryptFile()
def createXlsxFile():
    decryptFile()
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
    encryptFile()
def deletePassword():
    decryptFile()
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']

    if ws.max_row == 1:
        print("No password in the list to delete\n")
        return
    try:
        row=int(input("Enter the row number of the password you wish to delete: "))+2
        if row <2 :
            print("Row number can't be below 0\n")
            return
        else:
            ws.delete_rows(row,1)
            wb.save(filename=myFileName)
            wb.close()
    except:
        print("Must be a positive integer\n")
    encryptFile()
def findPassword():
    decryptFile()
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']
    if ws.max_row == 1:
        print("No password in the list to delete\n")
        return

    passYouWant= input("\n\nEnter the website, username, or password:")
    print('\nWebsite : Username : Password')
    for row in ws.iter_cols(1):
        for cell in row:
            if cell.value == passYouWant:
                print(f"{ws.cell(row=cell.row, column=1).value} : {ws.cell(row=cell.row, column=2).value} : {ws.cell(row=cell.row, column=3).value}")
    encryptFile()
    input("\nPress a key to continue...")
def generatePassword():
    decryptFile()
    try:
        website = input("Enter website: ")
        username= input("Enter Username: ")
        MAX_LEN = int(input("Length of the password: "))
        print("\n\n")
    except:
        print("Invalid input\n")

    DIGITS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    LOCASE_CHARACTERS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h',
                        'i', 'j', 'k', 'm', 'n', 'o', 'p', 'q',
                        'r', 's', 't', 'u', 'v', 'w', 'x', 'y',
                        'z']
    UPCASE_CHARACTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H',
                        'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q',
                        'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
                        'Z']
    SYMBOLS = ['@', '#', '$', '%', '=', ':', '?', '.', '/', '|', '~', '>',
            '*', '(', ')', '<']
    COMBINED_LIST = DIGITS + UPCASE_CHARACTERS + LOCASE_CHARACTERS + SYMBOLS
    rand_digit = random.choice(DIGITS)
    rand_upper = random.choice(UPCASE_CHARACTERS)
    rand_lower = random.choice(LOCASE_CHARACTERS)
    rand_symbol = random.choice(SYMBOLS)

    temp_pass = rand_digit + rand_upper + rand_lower + rand_symbol
    for x in range(MAX_LEN - 4):
        temp_pass = temp_pass + random.choice(COMBINED_LIST)
        temp_pass_list = array.array('u', temp_pass)
        random.shuffle(temp_pass_list)
    password = ""
    for x in temp_pass_list:
            password = password + x
    myFileName=r'passwords.xlsx'
    wb = load_workbook(filename=myFileName)
    ws = wb['Sheet']
    newRowLocation = ws.max_row +1
    ws.cell(column=1,row=newRowLocation, value=website)
    ws.cell(column=2,row=newRowLocation, value=username)
    ws.cell(column=3,row=newRowLocation, value=password)
    wb.save(filename=myFileName)
    wb.close()
    encryptFile()
def menu():
    try:
        choice = int(input("1. List all passwords\n2. Add a password\n3. Delete a password\n4. Find a password\n5. Generate and add password\n6. Exit\n\ninput:"))
        print("\n\n")
        if 1<=choice<=6:
            if choice == 1:
                listPasswords()
            elif choice==2:
                addPassword()
            elif choice==3:
                deletePassword()
            elif choice==4:
                findPassword()
            elif choice == 5:
                generatePassword()
            elif choice==6:
                return 1
        else:
            print("Invalid number")
    except:
        print("Invalid choice")
def start():

    while True:
        if os.path.isfile('passwords.xlsx') and os.path.isfile('masterkey.key'):
            start=menu()
            if start == 1:
                print("bye")
                break
        else:
            if os.path.isfile('passwords.xlsx')==False:
                createXlsxFile()
            if os.path.isfile('masterkey.key') == False:
                generateKey()
                encryptFile()
if __name__ == "__main__":
    start()