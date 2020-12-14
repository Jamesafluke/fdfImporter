import sys # To check if a text file exists.
import openpyxl as xl
import time
import subprocess #To open by dragging a text file onto the script.
import excelWriter as writer
import os.path # To open Excel at the end of the script.


droppedFile = sys.argv[1]
excelPath = "start C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
excelFileName = "OnboardingPython.xlsx"

#For testing in Visual Studio.
#droppedFile = "testssp.txt"
#droppedFile = "testnonssp.txt"


interfaces = ["Interface - Jonathan","Interface - Brittnee","Interface - Chris","Interface - James"]

def checkForSettingsFile():

    if os.path.exists('settings.txt'):
        print("Settings file found.")
        textFile = open('settings.txt')
        interface = textFile.read()
        print(f'Using {interface}.')


    else:
        print("No settings file detected. We'll create one now.")
        

        while True:
            user = input("Enter 1 for Jonathan, 2 for Brittnee, 3 for Chris and 4 for James: ")
            if user == "1":
                interface = interfaces[0]
                break
            elif user == '2':
                interface = interfaces[1]
                break
            elif user == '3':
                interface = interfaces[2]
                break
            elif user == '4':
                interface = interfaces[3]
                break
            else:
                print("Invalid input!")
                pass

        textFile = open('settings.txt','a')
        textFile.write(interface)

    return interface




def readTextFile(nameOfFile):

    print(f"Reading {nameOfFile}")
    textFile = open(nameOfFile, "r")
    importedText = textFile.read()
    textFile.close()

    return importedText




def StripToData(importedText): 

    #First split. (Strip off the text preceding the data.)
    splitTextFile = importedText.split('V(')
    
    #Second split. (Strip off the text after the data.)
    formattedData = [] # An array to contain manager, acceptanceDate, etc.
    strippedText = [] # An array to contain the useful data at index 0.
    for text in splitTextFile:
        strippedText = text.split(')')
        formattedData.append(strippedText[0])
    
    
    #Remove first element because it's garbage.
    formattedData.pop(0)
    
    
    #Return final list.
    return formattedData







def writeToExcelNonSSP(sheet, formattedList):
    print("Non-SSP")

    wb = xl.load_workbook('OnboardingPython.xlsx')
    sheet = wb['Interface - James']
        
    sheet['h3'].value = formattedList[4] #First name.
    sheet['h4'].value = formattedList[8] #Last name.
    sheet['h5'].value = formattedList[9] #Mobile.
    sheet['h6'].value = formattedList[12] #Title
    sheet['h7'].value = formattedList[0] #Manager.
        
    sheet['h12'].value = formattedList[7] #State.
    sheet['h15'].value = formattedList[17] #Mitel queues.
    sheet['h16'].value = formattedList[16] #Template employee.
    
    sheet['h20'].value = ""
    sheet['h21'].value = ""
    sheet['h22'].value = ""
    sheet['h23'].value = ""
    sheet['h25'].value = ""

    sheet['h45'].value = ""
        
    wb.save('OnboardingPython.xlsx')


def writeToExcelSSP(sheet, formattedList):
    print("SSP")

    wb = xl.load_workbook('OnboardingPython.xlsx')
    sheet = wb['Interface - James']

    sheet['h3'].value = formattedList[4] #First name.
    sheet['h4'].value = formattedList[8] #Last name.
    sheet['h5'].value = formattedList[9] #Mobile.
    sheet['h6'].value = formattedList[12] #Title
    sheet['h7'].value = formattedList[0] #Manager.

    sheet['h12'].value = ""
    sheet['h13'].value = ""
    sheet['h14'].value = ""
    sheet['h15'].value = ""
    sheet['h16'].value = ""
    sheet['h17'].value = ""

    sheet['h20'].value = formattedList[6]
    sheet['h21'].value = formattedList[10]
    sheet['h22'].value = formattedList[7]
    sheet['h23'].value = formattedList[11]
    sheet['h25'].value = formattedList[5]

    sheet['h45'].value = ""
    
    wb.save('OnboardingPython.xlsx')










interface = checkForSettingsFile()
importedText = readTextFile(droppedFile)
formattedList = StripToData(importedText)

print(formattedList)

# SSP or non-SSP?
if formattedList[12] == "SSP":
    writeToExcelSSP(interface, formattedList)
else:
    writeToExcelNonSSP(interface, formattedList)

os.system('start "Excel"  "OnboardingPython.xlsx"')

print("Finished!")
time.sleep(5)



