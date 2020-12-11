

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
    counter = 0
    formattedData = [] # An array to contain manager, acceptanceDate, etc.
    strippedText = [] # An array to contain the useful data at index 0.
    for text in splitTextFile:
        strippedText = text.split(')')
        formattedData.append(strippedText[0])
        counter += 1
    
    
    #Remove first element because it's garbage.
    formattedData.pop(0)
    
    
    #Return final list.
    return formattedData







