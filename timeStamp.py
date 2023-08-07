# Full imports 
import sys
sys.path.append('.')
import yaml
import urllib.request, urllib.parse
import xlwings as xw
import os
import atexit
import matplotlib.pyplot as plt

import pytz
# Class Specific Imports
from rdrive import RDrive
from io import BytesIO
from PIL import Image
from datetime import datetime, timedelta
from tabulate import tabulate
from requests import get

introduction = """
Welcome to Pakenham IR Findings Punchlist Matching Script
Some things to note before using:
1) Ensure that you have logged in and have a valid token.
2) Connect to the internet through a network that can make http requests.
3) Have the all libraries listed above.
4) Have a config.yaml file.
5) Ran the program through RDrive.py.
6) Ensure that the opened file is correct.
7) Ensure that the trainset is correct before starting to look for anything.
9) Ensure that you close the excel file either through the terminal or manually (Not recommmended).
"""
print(introduction)
# Sheet
class Excel(RDrive):
    def __init__(self, book) -> None:
        super().__init__()
        self.path = book
        self.project = 'Pakenham'
        self.book = xw.Book(book)
        self.sheet = None
        self.trainset = self.identifyTrainset()
        self.photoReferenceList = []
        self.referenceList = []
        self.urlList = {}
        self.trainsetLocation = 'CUETBSA2'
        pass
        
    def read_config(self):
        filename = 'config.yaml'
        with open(filename, "r") as f:
            data = yaml.safe_load(f)
        self.username = data['username']
        self.password = data['password']
        self.token = data['token']
        self.project = data['project']
        return data    

    def processSearch(self,starttime,endtime,drawingId):
        starttime, endtime = urllib.parse.quote(starttime), urllib.parse.quote(endtime)
        processEndpoint = 'https://api-downer-rts.rdrive.io/api/Pakenham/processes/search?drawingIds='+drawingId+'&createdDateFrom='+ starttime +'&createdDateTo='+ endtime +'&pageSize=25&page=0&sort=createdDate&order=desc'
        response = get(processEndpoint, headers = self.token)
        # print(drawingId, response)
        try:
            return response.json()
        except:
            print(response.status_code)

    def mainMenu(self):
        options = {1: 'Start Matching', 2: 'Select Sheet', 3: 'Exit'}
        # print and prompt options
        for key in options.keys():
            print(str(key),') ', options[key])
        option = input('Select from the following options: ')
        option = self.input2int(option)
        # decision tree
        match option:
            case 1:
                self.startMatching()
            case 2:
                self.selectSheet()
            case 3:
                pid = os.getpid()
                os.kill(pid)
            case _:
                print("Invalid input\n")
                self.mainMenu()

    def close(self):
        save = input('do you wish to save the workbook (Y/N)?: ')
        if save == 'Y':
            self.book.save(self.path)
            self.book.app.quit()
        elif save == 'N':
            self.book.app.quit()
        else:
            print('Invalid input')
            self.close()

    def startMatching(self):
        print(self.urlList)
        for idx in self.urlList.keys():
            if idx % 5 == 0:
                print("Matching... ")
            address = find_cell(self.sheet, self.referenceList[idx])
            
            pictureAddress = offset_cell(address,7,0)
            if len(self.urlList[idx]) > 0:
                message = ''
                image = url2plot(self.urlList[idx][0])
                self.sheet.pictures.add(image, 
                                        name = 'Rectification Photo' + pictureAddress, 
                                        update = True, 
                                        left=self.sheet.range(pictureAddress).left, 
                                        top=self.sheet.range(pictureAddress).top)
            else:
                message = 'No matches'
                image = None
                self.sheet.range(pictureAddress).value = message         
        print('Finished Matching')
        self.mainMenu()

    def input2int(self,str):
        """
        Params: str - a string from a input 
        Return: int - an integer converted from the str variable
        """
        try:
            return int(str)
        except:
            print("cannot convert into Integer")
            return 0

    def initializeLists(self):
        self.referenceList = []
        self.photoReferenceList = []
        self.urlList = {}
        # Try looking for photo reference and Reference in current sheet, else inform user that sheet is not compatible
        try: 
            photoReferencePointer = find_cell(self.sheet, 'Photo Reference')
            photoReferencePointer = offset_cell(photoReferencePointer,0,1)
            referencePointer = find_cell(self.sheet,'Reference')
            referencePointer = offset_cell(referencePointer,0,1)
        except:
            print('Selected Sheet is not compatible with this script')
            self.selectSheet()
        # Iterate down the column
        while self.sheet[photoReferencePointer].value != None:
            photoReference = self.sheet[photoReferencePointer].value
            reference = self.sheet[referencePointer].value
            self.referenceList.append(reference)
            self.photoReferenceList.append(photoReference)
            # Increment offset
            referencePointer = offset_cell(referencePointer,0,1)
            photoReferencePointer = offset_cell(photoReferencePointer,0,1)
        # Reset pointer (Maybe not necessary)
        photoReferencePointer = find_cell(self.sheet, 'Photo Reference')
        referencePointer = find_cell(self.sheet,'Reference')

    def identifyTrainset(self):
        try:
            print("\nDetecting Trainset...\n")
            sheets = [sheet for sheet in self.book.sheets if sheet.name == 'Interior']
            Interior = sheets[0]
            address = find_cell(Interior,'Trainset')
            number_address = offset_cell(address,1,0)
            trainsetNumber = Interior[number_address].value
            print("Trainset found is TS-" + trainsetNumber,'\n')
        except:   
            trainsetNumber = input('What is the trainset Number: ') 
        self.trainset = trainsetNumber
        return trainsetNumber
    
    def FormalIRFindingsConnector(self):
        FormalIRDrawingID = 'I8VIKH4T'
        FormalIRDrawingList = self.getDrawingDrillDowns(FormalIRDrawingID)
        for trainset in FormalIRDrawingList:
            trainsetExcel = 'TS' + self.trainset + ' Formal'
            trainsetDrawing = trainset['linkedDrawingTitle']
            if trainsetExcel == trainsetDrawing:
                locationid = trainset['linkedDrawingId']
                break 
        processes = self.getDrawingProcess(locationid)
        return locationid
    def show_table(self):
        processSearchResponses = []
        print('\n')
        for idx in range(len(self.photoReferenceList)):
            if idx % 5 == 0:
                print('searching...')
            # try:
            # from the reference list and photo reference list, generate a time interval
            date, time = self.referenceList[idx], self.photoReferenceList[idx]
            utc, localtime = reference2datetime(date,time)
            imageCount = 0
            searchTimeList = [] 

            processSearchResponse = []
            for time in [utc,localtime]:
                time1, time2 = offset_datetime(time,0,1.5) # change this to offset the time tolerance
                response = [x['forms'][0]['id'] for x in self.processSearch(time1,time2,self.trainsetLocation)]
                processSearchResponse += response
                searchTimeList.append((time1,time2))
            
            self.urlList[idx] = list()
            for process in processSearchResponse:
                formFields = self.getFormFields(process)
                for field in formFields:
                    if field['title'] == 'Rectification Photo' and (field['value']!=''):
                        imageCount += 1
                        self.urlList[idx].append(field['value'])
            processSearchResponses.append({'Range': searchTimeList, 'Matches': imageCount})
            # except:
            #     print("couldn't make http request for", self.referenceList[idx])
        
        timeRangeList = [response['Range'] for response in processSearchResponses]
        matchList = [response['Matches'] for response in processSearchResponses]
        attributeList = [[i for i in range(len(self.referenceList))], self.referenceList, self.photoReferenceList, timeRangeList, matchList]
        attributeList = tranpose_list(attributeList)
        print('Found the following entries in the spreadsheet\n')
        print(tabulate(attributeList, headers = ['Index','Reference','Photo Reference', 'Search Range', 'Potential Matches'],tablefmt='orgtbl'))
        return attributeList

    def selectSheet(self):
        # list all sheet names
        sheet_names = [sheet.name for sheet in self.book.sheets if sheet.name != 'Import']
        sheets = {}
        for index in range(len(sheet_names)):
            sheets[index+1] = sheet_names[index]
        # print sheetnames
        print('The following sheets have been found:')
        for sheet in sheets.keys():
            print(str(sheet) + ')', sheets[sheet])

        sheetNumber = input('\nSelect a sheet to operate on (Number):\n')     
        sheetNumber = self.input2int(str =sheetNumber)
        print('You have selected', sheets[sheetNumber])
        self.sheet = self.book.sheets[sheets[sheetNumber]]
        self.initializeLists()
        self.show_table()
        self.mainMenu()
    
"""
# Global Functions
    Functions that are useful for the program to function that operate on external variables
    and not class attributes
"""
def convertbase24to10(string):
    num = 0
    for c in string:
        num = num * 26 + (ord(c.upper()) - 64)
    return num
    
def convertbase10to24(num):
    result = ''
    while num > 0:
        remainder = (num - 1) % 26
        result = chr(65 + remainder) + result
        num = (num - 1) // 26
    return result

def find_cell(sheet,search_string):
    cell = sheet.api.UsedRange.Find(str(search_string))
    if cell is not None:
        return cell.Address
    else: 
        return 'A1'
        
def offset_cell(cell, x = 0, y = 0):
    col, row = cell[1:].split('$')
    col = convertbase24to10(col)
    ncol = col + x
    nrow = int(row) + y
    ncol = convertbase10to24(ncol)
    naddress = '$' + ncol + '$' + str(nrow)
    return naddress

def reference2datetime(date,time):
    """
    """
    date_obj = datetime.strptime(date.split(':')[0], "%d%m%Y")
    print(date_obj, time)
    date_obj = date_obj.strftime("%Y-%m-%dT") + ":".join([str(time)[:2], str(time)[2:4]])

    local_tz = pytz.timezone('Australia/Melbourne')
    dt_local = datetime.strptime(date_obj, '%Y-%m-%dT%H:%M')
    dt_local = local_tz.localize(dt_local, is_dst=None)

    # convert to UTC
    utc_tz = pytz.utc
    dt_utc = dt_local.astimezone(utc_tz)
    return dt_utc.strftime('%Y-%m-%dT%H:%M'), dt_local.strftime('%Y-%m-%dT%H:%M')

def defloatMinutes(floattime):
    minute = int(floattime)
    seconds = int((floattime - minute) * 60)
    return (minute,seconds)

def offset_datetime(date_time, m2, m1):
    m2, s2 = defloatMinutes(m2)
    m1, s1 = defloatMinutes(m1)

    dt = datetime.strptime(date_time, '%Y-%m-%dT%H:%M')
    delta2 = timedelta(minutes=m2, seconds = s2)
    delta1 = timedelta(minutes=m1, seconds = s1)
    dt1 = dt + delta1
    dt2 = dt + delta2
    return dt2.strftime('%Y-%m-%dT%H:%M:%S'),  dt1.strftime('%Y-%m-%dT%H:%M:%S')

def tranpose_list(lst):
    return [list(i) for i in zip(*lst)]

def findWorkBook():
    current_dir = os.getcwd()
    parent_dir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
    

    for file_name in os.listdir(current_dir):
        if file_name.endswith('.xlsx') or file_name.endswith('.xlsm'):
            print('Excel file found:', file_name)
            return os.path.join(current_dir, file_name)
    for filename in os.listdir(parent_dir):
        if file_name.endswith('.xlsx'):
            print('Excel file found:', file_name)
            return os.path.join(current_dir, file_name)
    filename = input('couldn''t find an excel file in the current directory, input filename: ')
    return filename        

def url2plot(url,size =(5,5)):
    """
    params: image - a web url containing the image
    params: size - an tuple of integers that determines how big the output image is
    return: fig - a matplot subplot containing the image
    """
    with urllib.request.urlopen(url) as url_response:
        image_data = url_response.read()
    img = Image.open(BytesIO(image_data))       
    fig,ax = plt.subplots(frameon= False)
    fig.set_size_inches(size)
    ax.set_axis_off()
    fig.add_axes(ax)
    ax.imshow(img)
    ax.axis('off')
    ax.axes.get_xaxis().set_visible(False)
    ax.axes.get_yaxis().set_visible(False)
    return fig

if __name__ == "__main__": 
    filename = findWorkBook()
    Punchlist = Excel(filename)
    Punchlist.read_config()
    atexit.register(Punchlist.close)
    Process = Punchlist.FormalIRFindingsConnector()
    print(Process)
    Punchlist.trainsetLocation = Process
    Punchlist.selectSheet()
    
    