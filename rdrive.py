import requests
import urllib.request, urllib.parse
import time
from timeStamp import Excel
from datetime import datetime
import yaml
import os 
import pytz
from formTransfer import Form_Migration
Introduction = """
Welcome to RDrive toolbox. 
This Project is maintained by Samuel Zhang (gejia.zhang@downergroup.com)
Start by Login to RDrive account to get API Token.
Then select a Project/Site. Each Project will have certain programs.
Such as Form Migration and Excel Sheet Tools
"""
print(Introduction)
current_dir = os.getcwd()

class User():
    def __init__(self) -> None:
        self.username = None
        self.password = None
        self.token = None
        self.tokenExpiryTime = None
        pass 

    def readYaml(self):
        try: 
            with open('config.yaml','r') as f:
                data = yaml.safe_load(f)
            self.username = data['username']
            self.password = data['password']
            return True
        except:
            return False
    
    def getTokenCred(self):
        tokenEndpoint = "https://api-downer-rts.rdrive.io/api/token"
        data = {
            "username" : self.username,
            "password" : self.password,
            "grant_type" : "password"
        }
        tokenResponse = requests.post(tokenEndpoint, data = data)
        return tokenResponse 
    
    def convertTime(self, time):
        """
        params: self
        params: time - a utc time object
        return: time - a string converted to have a similar format to the time format used by RDrive
        """
        return time.strftime('%Y-%m-%dT%H:%M:%S.%f') + 'Z'    
    
    def login(self, skip = False):
        skip = self.readYaml()
        if skip:
            pass
        else: 
            self.username = input("RDrive Username: ")
            self.password = input("RDrive Password: ")

        tokenResponse = self.getTokenCred()

        if tokenResponse.status_code == 200:
            print("\nLogin Successful\n")
            self.token = {"Authorization": "bearer " + tokenResponse.json()["access_token"]}
            now = self.convertTime(datetime.now())
            self.tokenExpiryTime = tokenResponse.json()['expiried_at']
        else: 
            print("\nLogin Failed\n")
            self.login()

class RDrive(User):
    def __init__(self) -> None:
        print("Initializing RDrive Connectors")
        super().__init__()
        self.companyAPI = 'https://api-downer-rts.rdrive.io/api/' 
        self.project = None

    def getProject(self):
        """
        param: project - takes a user input project
        return: projectCode - returns the project code of a existing project based on user selection
        """
        project_endpoint = 'https://api-downer-rts.rdrive.io/api/projects'
        print("Select from the following list of project")
        response = requests.get(project_endpoint, headers = self.token).json()
        index = 0
        for project in response:
            print(str(index), ") ", project['displayName'])
            index += 1 
        project = input("Which Project to visit?: ")
        try:
            project = int(project)
            print(project)
            print("Initializing Programs for: " , response[project]['displayName'])
        except:
            print("Invalid Selection: ")
            time.sleep(1)
            self.getProject()
        return response[project]['code'] 
    
    def findTimeDifference(self, time1, time2):
        datetime_format = '%Y-%m-%dT%H:%M:%S'
        dt1 = datetime.strptime(time1[:19], datetime_format)
        dt2 = datetime.strptime(time2[:19], datetime_format)
        diff = dt1 - dt2
        return diff
    
    def responseError(self, errorCode):
        match errorCode:
            case 401:
                return "Error: Token is either not valid or expired. Try again"
            case 404:
                return "Error: Endpoint is not found: Please contact Samuel at gejia.zhang@downergroup.com"
            case 405:
                return "Error: Method is not allowed: Please contact Samuel at gejia.zhang@downergroup.com"
            case 429: 
                return "Error: Too many requests, try limiting the number of calls"
        return "OK"
    
    def getProcessInfo(self, processCode):
        processEndpoint = self.companyAPI + '/' + self.project + '/process/' + processCode
        processResponse = requests.get(processEndpoint, headers = self.token)
        try:
            return processResponse.json()
        except:
            print("request Failed")
            print(self.responseError(processResponse.status_code))
        return 
    
    def getFormFields(self,formId = None,formCode = None):
        if formId != None:
            
            formFieldsEndpoint = self.companyAPI + self.project + '/forms/' + formId + '/fields'
            formFieldsResponse = requests.get(formFieldsEndpoint, headers = self.token)
        try:
            return formFieldsResponse.json()
        except:
            print("Request Failed")
            print(self.responseError(formFieldsResponse.status_code))

    def getFormLinkedDoc(self,formId):
        formDocEndpoint = self.companyAPI + '/' + self.project + '/forms/' + formId + '/linked-documents'
        
        formDocResponse = requests.get(formDocEndpoint, headers = self.token)
        try:
            return formDocResponse.json()
        except:
            print("Request Failed")
            print(self.responseError(formDocResponse.status_code))
            return
        
    def getDrawingDrillDowns(self, drawingId):
        drawingDrillDownEndpoint = self.companyAPI + self.project + '/drawings/' + drawingId + '/drill-downs'
        drawingDrillDownResponse = requests.get(drawingDrillDownEndpoint, headers = self.token)
        try:
            return drawingDrillDownResponse.json()
        except:
            print("Request Failed")
            print(self.responseError(drawingDrillDownResponse.status_code))
            return
        
    def getDrawingProcess(self, drawingId):
        drawingProcessEndpoint = self.companyAPI + self.project + '/drawings/' + drawingId + '/processes'
        drawingProcessResponse = requests.get(drawingProcessEndpoint, headers = self.token)
        try:
            return drawingProcessResponse.json()
        except:
            print("Request Failed")
            print(self.responseError(drawingProcessResponse.status_code))
            return
        
    def saveSettings(self):
        filepath = 'config.yaml'
        if not os.path.exists(filepath):
            with open(filepath, 'w') as f:
                f.write('')
        
        data = {}
        data['username'] = self.username
        data['password'] = self.password
        data['token'] = self.token
        data['project'] = self.project
        
        with open(filepath, 'w') as f:
            yaml.safe_dump(data, f)
        return 
        
class Programs(RDrive):
    def __init__(self,site):
        self.site = site
        print(self.site)
        self.programFilename = self.selectProgram()           

    def input2int(self,str):
        """
        :param str: A string from a input 
        :return int: A integer converted from the str variable
        """
        try:
            return int(str)
        except:
            print("cannot convert into Integer")
            return 0
    
    def selectProgram(self):
        programList = {}
        match self.site:
            case 'Pakenham':
                programList[1] = 'formTransfer/formTransfer.py'
                programList[2] = 'timeStamp/timeStamp.py'
            case 'newportbuild':
                programList[1] = 'formTransfer/formTransfer.py'
            case _:
                print("nothing has been built for this site yet")
                return False

        for keys in programList.keys():
            print(keys," ",programList[keys])

        programNumber = input('Select from the following (Number):')
        programNumber = self.input2int(programNumber)
        return programList[programNumber]
        
        


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

def mainMenuFormTransfer():
    options = {
        1: "Migrate a form",
        2: "Migrate multiple forms",
        3: "Exit"
    }
    print("Main Menu")
    for option in options.keys():
        print(str(option)+")",options[option])

    number = input("Enter a number: ")
    match trystr2int(number):
        case 1: transfer_form()
        case 2: transfer_form_skip()
        case 3: exit()
        case _:
            print("Instruction unclear")
            mainMenuFormTransfer
    return

def trystr2int(string):
    try: 
        return int(string)
    except:
        return 0 

def transfer_form():
    oldFormId = input("Input the ID for the old form: ")
    newFormId = input("Input the ID for the new form: ")
    form = Form_Migration(oldFormId,newFormId)
    mainMenuFormTransfer(form)

def transfer_form_skip():
    currentdir = os.path.dirname(os.path.abspath(__file__))
    filename = os.path.join(currentdir,'formList.txt')
    with open(filename, 'r') as f:
        for line in f:
            # Remove any trailing newline character
            line = line.strip()
            # Split the line into two parts using space as the delimiter
            parts = line.split(' ')[:2]
            # Check if the line contains two parts
            print(parts[0],parts[1])
            form = Form_Migration(parts[0],parts[1],True)
    mainMenuFormTransfer()

def parse_yaml():
    return 0 

def mainMenuTimeStamp():
    filename = findWorkBook()
    Punchlist = Excel(filename)
    Punchlist.read_config()
    Process = Punchlist.FormalIRFindingsConnector()
    print(Process)
    Punchlist.trainsetLocation = Process
    Punchlist.selectSheet()



def menuLoop(rdrive):
    rdrive.project = rdrive.getProject()
    rdrive.saveSettings()
    program = Programs(rdrive.project)
    if not program:
        menuLoop(rdrive)
    program.run()

def main():
    DownerRTS = RDrive()
    DownerRTS.login(DownerRTS)
    menuLoop(DownerRTS)

    

if __name__ == '__main__':
    main()