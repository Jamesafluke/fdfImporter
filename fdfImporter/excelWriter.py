import openpyxl as xl


wb = xl.load_workbook('OnboardingPython.xlsx')
dataSheet = wb["Data"]
getFromCell = dataSheet['b1']
writeToCell = dataSheet['a1']

print(getFromCell.value)
newText = 5
writeToCell.value = newText
print("I think it worked.")


wb.save('OnboardingPythonUpdated.xlsx')