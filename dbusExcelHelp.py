#! python3
#dbusExcelHelp - Reads körjournalen and makes it easier to see who to bill and
#what to bill them for.

import pprint, openpyxl, datetime, os, copy, shelve

#Prints a starting message
print('Opening journal...')
#Save some variables for the different filenames
contact = 'Contact information [PU] (Svar) 2020_2021.xlsx'
journal = 'Körjournal DBus [PU] (Svar) 2020.xlsx'

#Opening the shelf for dbus file paths
dbusShelf = shelve.open('dbus')

#Helper methods used in the script
#A short method for filtering out info that we don't need for the list.
def shortDict(d):
    newDict = copy.deepcopy(d)
    for x in newDict:
        newDict[x].pop('trips')
        if newDict[x]['contact info']['mobil'] == 0:
            newDict[x].pop('contact info')
    return newDict

def validMonth(m):
    months = {'january' : 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5,
                 'june': 6, 'july': 7, 'august': 8, 'september': 9, 'october': 10,
                 'november': 11, 'december': 12}
    if m.lower() in months:
        return months[m.lower()]
    else:
        return 0

def validCheck(d,h,r):
    if (d < 0):
        raise InvalidDistanceError("Invalid distance @ row: " + str(r))
        return
    elif(d/max(h,1)>80):
        raise RatioError("DRAGRACING WITH DBUS!! @ row: " + str(r))
        return
    return


#Searches the users dir for the files that we want to read from. Currently
#quite slow. Should implement something to make it faster
#If the script has been run once before the paths will be saved in a shelf that
#we access instead of running the slow loop again.
if 'journalP' not in dbusShelf.keys() or 'contactP' not in dbusShelf.keys():
    for root, dirs, files in os.walk("C:\\Users\\"):
        for name in files:
            if name == journal:
                dbusShelf['journalP'] =os.path.join(root, name)
            if name == contact:
                dbusShelf['contactP'] = os.path.join(root, name)
    print(str(list(dbusShelf.keys())))
#Opening up the diffrent excel files so that we may work with them
wbF = openpyxl.load_workbook(dbusShelf['journalP'])
sheetF = wbF['Formulärsvar 1']
wbC = openpyxl.load_workbook(dbusShelf['contactP'])
sheetC = wbC['Formularsvar1']
dbusData = {}

dbusShelf.close()

#Input for choosing which month to process
#TODO - Make it possible to input names of months. Also errorhandling
period = input('Choose period to analyze: (Start - End)').split('-')
startMonth = validMonth(period[0])
endMonth = validMonth(period[1])
print(startMonth)
print(period)
print('Analysing data...')

#A scuffed while loop to read the different values from the file.
#Probably possible to make this better and more nice looking.
row = 3
while(sheetF['C' + str(row)].value != None):
    #Each row in the spreadsheet contains data that we want, (except first 2)
    date = (sheetF['A' + str(row)].value)
    #This could be something else probably. Currently filters the data so that
    #we don't save data we don't want to process.
    if(startMonth <= date.month and date.month <= endMonth ):
        #Save different values from each row of the file
        try:
            date = (sheetF['A' + str(row)].value).strftime('%Y/%m/%d - %H:%M:%S')
            booker = sheetF['F' + str(row)].value
            booker = str(booker).lower().strip()
            hours = sheetF['D' + str(row)].value
            distance = sheetF['C' + str(row)].value - sheetF['C' + str(row-1)].value
            #print(str(distance) + " @row" + str(row))
            drove_as = sheetF['G' + str(row)].value
            validCheck(distance, hours, row)

            #Make sure the key for the booker exists
            dbusData.setdefault(booker, {'total hours': 0, 'total distance' :0, 'trips' : {}
                                        , 'contact info': {'personnr': 0,
                                                            'mobil': 0, 'email': ''}})
            dbusData[booker]['trips'].setdefault(date,{'drove as':' ','time rented': 0,
                                                'distance':0})

            #Make sure the key for the date exists
            #each row represents a trip, set everything to the right values
            dbusData[booker]['trips'][date]['drove as']    = drove_as
            dbusData[booker]['trips'][date]['distance']    = distance
            dbusData[booker]['trips'][date]['time rented'] = hours
            dbusData[booker]['total distance']             += distance
            dbusData[booker]['total hours']                += hours
        except TypeError:
            print('Invalid data @ row: ' + str(row))
        except:
            print('Invalid distance @ row: ' + str(row))
    row += 1

#Same scuffed while loop but not quite. This time it is for the contact info.
#It would be possible to save this info in a shelve and access it whenever, but
#that would make it one more file that you'd have to track for GDPR purposes.
#Might be nice, but this is so quickly done anyway.
row = 2
while(sheetC['D'+ str(row)].value != None):
    name = (sheetC['D'+ str(row)].value + ' ' + sheetC['E'+ str(row)].value)
    name = name.lower().strip()

    if(name in dbusData.keys()):
        email = sheetC['C' + str(row)].value
        personnr = int(sheetC['G' + str(row)].value)
        mobil = sheetC['F' + str(row)].value
        dbusData[name]['contact info']['mobil']       = mobil
        dbusData[name]['contact info']['email']       = email
        dbusData[name]['contact info']['personnr']    = personnr
    row += 1

#If dbusData is empty, that means a month that has no data was choosen.
if(dbusData == {}):
    print('No data available for the chosen month')
else:
    #print(pprint.pprint(dbusData))
    print('Data colleced!')
    print('Enter new command:')
    print('Write \'help\' for more information')

#Command line, prints different things depending on what you want to know/do
while True:
    command = input()
    if (command.lower() == 'list'):
        print(pprint.pformat(shortDict(dbusData)))
    elif(command.lower() == 'help'):
        print('''List of commands:
         <list> - prints a list of all billing info for chosen month
         <person> - prints trip data for chosen person''')
    elif command.lower() in dbusData :
        print('Printing data for '+ command)
        print(pprint.pformat(dbusData[command]))
    else:
        print('invalid command')
    print('Enter new command')
