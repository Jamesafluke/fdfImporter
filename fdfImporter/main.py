import sys
import time
import txtReader as reader
import excelWriter as writer


droppedFile = sys.argv[1]
for p in droppedFile:
    print(p)


nameOfFile = "testssp.txt"

importedText = reader.readTextFile(nameOfFile)
formattedList = reader.StripToData(importedText)

print(formattedList)


if(formattedList[12]=="SSP"):
    writer.writeToExcelSSP("Interface - James", formattedList)
else:
    writer.writeToExcelNonSSP("Interface - James", formattedList)


print("Finished")



