import openpyxl as xl


def writeToExcelNonSSP(sheet, formattedList):
    print("Non-SSP")

    wb = xl.load_workbook('OnboardingPython.xlsx')
    sheet = wb['Interface - James']
    #cell = sheet['a1']

    
    sheet['h3'].value = formattedList[4] #First name.
    sheet['h4'].value = formattedList[5] #Last name.
    sheet['h5'].value = formattedList[9] #Mobile.
    sheet['h6'].value = formattedList[12] #Title
    sheet['h7'].value = formattedList[0] #Manager.
    
    sheet['h12'].value = formattedList[7] #State.
    sheet['h15'].value = formattedList[17] #Mitel queues.
    sheet['h16'].value = formattedList[16] #Template employee.
    


    
    wb.save('OnboardingPython.xlsx')


def writeToExcelSSP(sheet, formattedList):
    print("SSP")

    wb = xl.load_workbook('OnboardingPython.xlsx')
    sheet = wb['Interface - James']

    sheet['h3'].value = formattedList[4] #First name.
    sheet['h4'].value = formattedList[5] #Last name.
    sheet['h5'].value = formattedList[9] #Mobile.
    sheet['h6'].value = formattedList[12] #Title
    sheet['h7'].value = formattedList[0] #Manager.

    sheet['h20'].value = formattedList[6]
    sheet['h21'].value = formattedList[10]
    sheet['h22'].value = formattedList[7]
    sheet['h23'].value = formattedList[11]
    sheet['h25'].value = formattedList[5]
    


    wb.save('OnboardingPython.xlsx')
