import os
import fitz
import json
import re
import string
print(fitz.__doc__)

import sys


DEBUG = False


fileName = r''
fileName += sys.argv[1]
print(fileName)

fileName=fileName.replace("\\","\\\\")

configFile = "config.txt"

with open("config.txt") as json_file:
    data = json.load(json_file)

listToSearch = data['sri_search_list']


doc = fitz.open(fileName) 


def getColumnFromX(xPos, text):
	if(xPos < 50):
		print("Error : no column corresponds to " + xpos)
		return -1
	elif (xPos <110):
		return 1
	elif (xPos <180):
		return 2
	elif (xPos <240):
		return 3		
	elif (xPos <300):
		return 4	
	elif (xPos <380):
		return 5
	elif (xPos <430):
		return 6
	elif (xPos <520):
		return 7
	else :
		print("===========================================")
		print("Error : no column corresponds to " + str(xPos))
		print (text)
		print("===========================================")		

		return -1	
		
def lookFor(mytext, name):
	if(name in mytext):
		print("Found " + name)
		foundBlock = myblock
		lineList2 = foundBlock.get('lines')
		tempText = ""
		for myElement in lineList2:
			myspan2 = myElement.get('spans')[0]
			mytext2 = myspan2['text']

			tempText += mytext + "\t"
			if(name not in mytext2):
				print ("new value")
				if(name not in newDict):
					newDict[name] = [mytext2]	
				else:
					newDict[name].append(mytext2)

def isPartOfList(line):
	for item in listToSearch:
		if(item in line):
		    return 1
	return 0

	
	
#TODO  : ask for specific date
#my_date = raw_input('Date :')	
#print("looking for " + my_date)
	
#===============================================================================
#                       Parse SRI specific format exported as XPS
#===============================================================================

count = 0
textToPrint = ""
newLine = 1
resultList = []

fileDebug = open("debug.json","w")

nbPage = 0
listOfDates = []
listOfTimes = []
for page in doc:
	nbPage +=1
	text = page.getText("json")
	fileDebug.write(text)
	parsed_json = json.loads(text)

	myPosition = 0
	currentColumn = 1

	tempLine = ""
	for myblock in parsed_json.get('blocks'):
		
		lineList = myblock.get('lines')

		if  lineList is not None:
			for myline in lineList:
				spanList = myline.get('spans')
				
				for myspan in spanList:
				    
					mytext = myspan.get('text')

					#print(mytext.encode('mbcs'))
					
					previousPosition = myPosition
					myPosition = myspan.get('bbox')
	
					
					
					x = re.search("\d\d\/\d\d\/\d\d\d\d\s$", mytext)
					if x is not None:
						#print("#" + x.string + "#")
						if(nbPage == 1):
							listOfDates.append(x.string) 
					
					x = re.search("\d\d\:\d\d$", mytext)
					if x is not None:
						#print("#" + x.string + "#")
						if(nbPage == 1):
							listOfTimes.append(x.string) 		
						
				
					if(previousPosition == 0):
						previousPosition = myPosition
					
					previousX = int(previousPosition[0])
					currentX = int(myPosition[0])
					previousY = int(previousPosition[1])
					currentY = int(myPosition[1])

					#print("\tPrevious :" + str(previousX))
					#print("\tCurrent :" + str(currentX))
					#print("\tPreviousX :" + str(previousY))
					#print("\tCurrentX :" + str(currentY))
					colNb = getColumnFromX(currentX, tempLine)
					if(currentX < previousX):
						#print("/////////////////////////")
						#print(tempLine.encode('mbcs'))
						currentColumn = 1
						if(isPartOfList(tempLine)):
							textToPrint += tempLine
							textToPrint += "\n"
							resultList.append(tempLine + "\n")
						newLine = 1
						#tempLine = str(count) + "\t|" + mytext + "(" + str(colNb) + ")"
						tempLine = mytext 
						if (DEBUG == True):
							tempLine + "(" + str(colNb) + ")"
						count += 1 
					else:
						if(((currentY - previousY) < 3) ):
							deltaColumn = colNb-currentColumn
							for i in range(0,deltaColumn) :
								tempLine += "\t"
							#print ("\t\tnew column" + "(" + str(colNb) + ")")
							currentColumn += deltaColumn
						tempLine += mytext 
						if (DEBUG == True):
							tempLine + "(" + str(colNb) + ")"						
						newLine = 0

print("\n==========================\n==========================\n===============")
	
#print (textToPrint)


#===============================================================================
#                             Write result to file
#===============================================================================

fileWrite = "output.xls"
file1 = open(fileWrite,"w") 
emptyLine = "\t/\t/\t/\t/\n"

#Build first line with proper dates
first_line = "\t"
for this_date in listOfDates:
	first_line += "\t" + this_date
first_line +="\n"

second_line = "\t"
for this_time in listOfTimes:
	second_line += "\t" + this_time
second_line +="\n"

file1.write(first_line)
file1.write(second_line)
for item in listToSearch:
	itemFound = False
	for line in resultList:
		if item in line:
			#print ("\n!!!!!!!!!!!!!!!!!!\nitem : " + item)
			#print (line.encode('mbcs'))
			file1.write(line.encode('mbcs'))
			itemFound = True
	if itemFound == False:
		file1.write(item.encode('mbcs') + emptyLine )
file1.close()
fileDebug.close

#===============================================================================
#                                  Launch Excel                        
#===============================================================================
os.system('start "excel" "output.xls"')	


