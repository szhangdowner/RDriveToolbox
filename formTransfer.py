from datetime import datetime
import time
import requests
from tabulate import tabulate
import yaml
import os 
from pandas import DataFrame as df
from pandas import merge

class User():
    def __init__(self, username, password):
        self.companyAPI = 'https://api-downer-rts.rdrive.io/api'
        self.username = username
        self.password = password
        self.expire_time = 0
        self.token = self.get_token()

    def get_token(self):
        auth_data = {"username" : self.username,
            "password" : self.password,
            "client_id" : "",
            "refresh_token" : "",
            "grant_type" : "password"
        }
        tokenEndpoint = 'https://api-downer-rts.rdrive.io/api/token'
        print('Logging in as ', self.username)
        tokenResponse = requests.post(tokenEndpoint, data=auth_data) 
        try:
            bearer = tokenResponse.json()['access_token']
            self.expire_time = tokenResponse.json()["expiried_at"]
            return {"Authorization": "bearer " + bearer}
        except:
            print('Failed to get token Error:', tokenResponse.status_code) 
    
    def login(self):
        self.username = input("Username: ")
        self.password = input("Password: ")
        try: 
            self.get_token()
        except:
            print("Incorrect Username or Password")
            self.login()

class Form_Migration():
    def __init__(self, oldFormCode, newFormCode, skip = False):
        self.user = User('gejia.zhang@downergroup.com','chill94')
        self.token = self.user.token
        self.tokenExpireTime = self.user.expire_time
        self.companyAPI = 'https://api-downer-rts.rdrive.io/api'
        self.project = 'newportbuild'
        self.skip = skip
        self.oldFormCode = oldFormCode
        self.newFormCode = newFormCode
        self.oldFormData, self.newFormData = [],[]
        self.oldFormFields, self.oldFormLinkedDocs = [], []
        self.oldProcessData, self.newProcessData = None, None 
        self.mergeMetrics = {}
        self.percentageComplete = 0.0
        self.dataLoss = []
        self.dataUp = []
    
        # Run Migration Process
        try:
            self.getFormData()
            self.getProcessData()
            self.getFormFields()
            self.getFormLinkedDocs()
        except:
            print("Could not fetch form data, please check Form ID")
            if skip:
                return
        else:
            print("Success")
            self.uploadTable = self.createUploadTable()
            self.migrationSafetyCheck(self.skip)
            self.upload()
            print("Finished Uploading " + self.oldFormCode + ' to ' + self.newFormCode)
        # Run Migration Process
        pass

    def rdriveFormatTime(self, time):
        """
        params: self
        params: time - a utc time object
        return: time - a string converted to have a similar format to the time format used by RDrive
        """
        return time.strftime('%Y-%m-%dT%H:%M:%S.%f') + 'Z'
    
    def checkTokenExpiry(self):
        """
        params: self - checks the token attribute
        return: True if token is still valid, False if token is not
        """
        now = self.rdriveFormatTime(datetime.utcnow())
        if now[:20] == self.tokenExpireTime[:20]:
            print('Token Has Expired')
            return False
        else:
            return True
    def rest(self,seconds):
        print("Taking a Break for", seconds, 'seconds')
        for i in range(seconds):
            print("Resuming in", seconds-i, 'seconds')
            time.sleep(1)
        
    def compareDateTime(self,datetime1,datetime2):
        """
        params: datetime1 - contains a string in the format "YYYY-MM-DDTHH:mm:ss.sssssssZ", representing an older date time
        params: datetime2 - contains a string in the format "YYYY-MM-DDTHH:mm:ss.sssssssZ", representing a newer data time
        return: True if datetime1 is before datetime2 and False if not
        """
        dt1 = datetime.fromisoformat(datetime1.rstrip('Z')[:19])
        dt2 = datetime.fromisoformat(datetime2.rstrip('Z')[:19])
        if dt1 < dt2:
            return True
        else:
            return False

    def getFormData(self):
        oldFormDataEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.oldFormCode
        newFormDataEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.newFormCode   
        print('Fetching Form Data') 
        self.oldFormData = requests.get(oldFormDataEndpoint, headers = self.token).json()
        self.newFormData = requests.get(newFormDataEndpoint, headers = self.token).json()

    def getProcessData(self):
        oldProcessEndpoint = self.companyAPI + '/' + self.project + '/processes/' + self.oldFormData['processId']
        newProcessEndpoint = self.companyAPI + '/' + self.project + '/processes/' + self.newFormData['processId']
        print('Fetching Process Data')
        self.oldProcessData = requests.get(oldProcessEndpoint, headers= self.token).json()
        self.newProcessData = requests.get(newProcessEndpoint, headers= self.token).json()

    def migrationSafetyCheck(self, skip = False):
        print('Performing Migration Safety Checks: ')
        now = datetime.utcnow()
        print('Current Time: ', self.rdriveFormatTime(now))
        print('Token Expiry Time:', self.tokenExpireTime)

        tabulate_list = []
        # Check for if the processes are located in the same drawing
        self.mergeMetrics['sameDrawingCode'] = (self.oldProcessData['drawingTitle'] == self.newProcessData['drawingTitle'])
        tabulate_list.append(['Same Drawing?:', str(self.mergeMetrics['sameDrawingCode'])])
        if not self.mergeMetrics['sameDrawingCode']:
            print('Old form belongs to ' + self.oldProcessData['drawingTitle'] + ' and new form belongs to ' + self.newProcessData['drawingTitle'])
        #----------------------------------------------------------------------------------------------------------------------------------------------------
        # Check for if the process are located in the same location
        self.mergeMetrics['sameLocation'] = (self.oldProcessData['locationTitle'] == self.newProcessData['locationTitle'])
        tabulate_list.append(['Same Location?:', str(self.mergeMetrics['sameLocation'])])
        if not self.mergeMetrics['sameLocation']:
            print('Old form belongs to ' + self.oldProcessData['locationTitle'] + ' and new form belongs to ' + self.newProcessData['locationTitle'])
        #----------------------------------------------------------------------------------------------------------------------------------------------------
        # Check for if the two forms share the same process template code
        self.mergeMetrics['sameProcessCode'] = (self.oldProcessData['processTemplateCode'] == self.newProcessData['processTemplateCode'])
        tabulate_list.append(['Same Process?: ', str(self.mergeMetrics['sameProcessCode'])])
        if not self.mergeMetrics['sameProcessCode']:
            print('Old form belongs to ' + self.oldProcessData['processTemplateCode'] + ' and new form belongs to ' + self.newProcessData['processTemplateCode'])
        #----------------------------------------------------------------------------------------------------------------------------------------------------
        # Check for if the two forms share the same form template code
        self.mergeMetrics['sameFormCode'] = (self.oldFormData['formTemplateCode'] == self.newFormData['formTemplateCode'])
        tabulate_list.append(['Same Form?: ', str(self.mergeMetrics['sameFormCode'])])
        if not self.mergeMetrics['sameFormCode']:
            print('Old form belongs to ' + self.oldFormData['formTemplateCode'] + ' and new form belongs to ' + self.newFormData['formTemplateCode'])

        # Check for if the old form is indeed older than the new form
        self.mergeMetrics['validDate'] = self.compareDateTime(self.oldFormData['createdDate'],self.newFormData['createdDate'])
        tabulate_list.append(['Valid Dates?: ', str(self.mergeMetrics['validDate'])])
        if not self.mergeMetrics['sameFormCode']:
            print('Old form is newer than the new form')

        # Checks how much data will be ignored
        dataLossPercentage = str(self.mergeMetrics['dataLossPercentage']) + ' %'
        tabulate_list.append(['Approximate Data Loss after upload: ', dataLossPercentage])

        print("The following data will be transferred over to the new form")
        print(tabulate(self.dataUp,headers = ['New Value','Replacing','In'],tablefmt='orgtbl'))

        print("The following data will be ignored")
        print(tabulate(self.dataLoss,headers=['Title','Values','Document'],tablefmt='orgtbl'))

        print("The following anomalies where found during examination")
        print(tabulate(tabulate_list,headers=['Test', 'Result'],tablefmt='orgtbl'))
        print("\n")

        riskyScore = self.mergeMetrics['validDate'] + self.mergeMetrics['sameFormCode'] + self.mergeMetrics['sameProcessCode'] + self.mergeMetrics['sameDrawingCode']
        
        def userInput(skip):
            if skip == True:
                return

            else:
                if (riskyScore > 1) or self.mergeMetrics['dataLossPercentage'] > 10.0:
                    proceed = input("Type in the form code of the New Form to continue or N to stop: ")
                    if proceed == self.newFormCode:
                        return
                    elif proceed == 'N':
                        exit()            
                    else:
                        print("Please Provide the Form Code again")
                    userInput(skip)
                else: 
                    proceed = input("Do You Wish to Proceed? : Y/N \n")
                    if proceed == 'Y':
                        return
                    elif proceed == 'N':
                        exit()            
                    else:
                        print("Please Provide Y/N reponse")
                    userInput(skip)

        userInput(skip)

    def getFormFields(self):
        oldFormEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.oldFormCode + '/fields'
        newFormEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.newFormCode + '/fields'
        print('Fetching Form Fields')
        self.oldFormFields = requests.get(oldFormEndpoint, headers = self.token).json()
        self.newFormFields = requests.get(newFormEndpoint, headers = self.token).json()

    def getFormLinkedDocs(self):
        oldFormDocEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.oldFormCode + '/linked-documents'
        newFormDocEndpoint = self.companyAPI + '/' + self.project + '/forms/' + self.newFormCode + '/linked-documents'
        print('Fetching Form Documents')
        self.oldFormLinkedDocs = requests.get(oldFormDocEndpoint, headers = self.token).json()
        self.newFormLinkedDocs = requests.get(newFormDocEndpoint, headers = self.token).json()
        
    def createUploadTable(self):
        print('Merging Data')
        time.sleep(1)
        dataframe1 = df(self.oldFormFields)
        dataframe2 = df(self.newFormFields)
        filename = os.path.join(os.path.dirname(os.path.realpath(__file__)),'backup.yaml')
        self.dumpYamlData(filename, {self.oldFormCode: dataframe1.to_dict('index')})
        self.dumpYamlData(filename, {self.newFormCode: dataframe2.to_dict('index')})
        upload_table = merge(dataframe1,dataframe2,on = 'title')
        upload_table['linkedDocId'] = upload_table['meta_x'].apply(lambda x: x.get('documentIds') if 'documentIds' in x else []) 
        upload_table["displayName"] = [[] for x in range(len(upload_table.index))]
        upload_table["storageId"] = [[] for x in range(len(upload_table.index))]
        upload_table["mime"] = [[] for x in range(len(upload_table.index))]
        upload_table["fileSize"] = [[] for x in range(len(upload_table.index))]
        upload_table["filename"] = [[] for x in range(len(upload_table.index))]
        upload_table["extension"] = [[] for x in range(len(upload_table.index))]

        for document in self.oldFormLinkedDocs:
            mask = upload_table['linkedDocId'].apply(lambda x: document['id'] in x if len(x) > 0 else False)
            # After obtaining mask, locate all and append
            upload_table.loc[mask,"displayName"].apply(lambda x: x.append(document['displayName']))
            upload_table.loc[mask,"storageId"].apply(lambda x: x.append(document['storageId']))
            upload_table.loc[mask,"mime"].apply(lambda x: x.append(document['mime']))
            upload_table.loc[mask,"fileSize"].apply(lambda x: x.append(document['fileSize']))
            upload_table.loc[mask,"filename"].apply(lambda x: x.append(document['displayName'].replace("." + document['displayName'].split(".").pop(),"")))
            upload_table.loc[mask,"extension"].apply(lambda x: x.append(document['displayName'].split(".").pop()))

        missingData = dataframe1[~(dataframe1['title'].isin(upload_table['title']))]
        try:
            for i in range(missingData.shape[0]):
                title = missingData.loc[i,"title"]
                value = missingData.loc[i,"value"]
                file =  missingData.loc[i,"meta"]
                if file == {}:
                    file = ""
                list = [title,value,file]
                self.dataLoss.append(list)

        except:
            try:
                for i in range(missingData.shape[0]):
                    title = missingData.loc[i,"title"]
                    value = missingData.loc[i,"value"]
                    files = missingData.loc[i,"meta"]
                    if file == {}:
                        file = ""
                    list = [title,value,file]
                    self.dataLoss.append(list)
            except:
                print('Not enough data to check data loss properly')
        self.dataLoss = missingData
        

        self.mergeMetrics['dataLossPercentage'] = round(missingData.shape[0]/dataframe1.shape[0]*100,2)

        for i in range(upload_table.shape[0]):
            if len(upload_table.loc[i,"displayName"])>0:
                new_value = upload_table.loc[i,"displayName"]
            else:
                new_value = upload_table.loc[i,"value_x"]
            list = [new_value,
                    upload_table.loc[i,"value_y"],
                    upload_table.loc[i,"title"]]
            self.dataUp.append(list)

        return upload_table

    def dumpYamlData(self,filename,data):
        with open(filename,"w") as f:
            if self.searchYamlFile(filename,next(iter(data.keys()))):
                return
            yaml.dump(data,f,default_flow_style=False)
    
    def searchYamlFile(self,filename,id):
        with open(filename,'r') as f:
            yaml_data = yaml.safe_load(f)
        try:
            for key in yaml_data:
                if id in key:
                    return True
        except:
            pass
        return False

    def upload(self):
        documentPostingURL =  'https://api-downer-rts.rdrive.io/api/newportbuild/documents/'
        for i in range(self.uploadTable.shape[0]):
            current_completion = round(i/self.uploadTable.shape[0] * 100,1)
            if current_completion > self.percentageComplete + 5.0:
                self.percentageComplete += 5.0
                print(str(round(self.percentageComplete,1)) + '% Complete')

            if not self.checkTokenExpiry():
                self.rest(20)
                oldToken = self.token()
                self.token = self.user.get_token()
                print(self.token)
            api_value = self.uploadTable.loc[i,"value_x"]
            api_meta_data = self.uploadTable.loc[i,"meta_x"]
            api_formId = self.uploadTable.loc[i,"formId_y"]
            post_endpoint = "https://api-downer-rts.rdrive.io/api/newportbuild/forms/" + api_formId + "/field-histories"
            api_formFieldTemplateId = self.uploadTable.loc[i,"formFieldTemplateId_y"]

            if len(self.uploadTable.loc[i,"displayName"]) > 0:
                for j in range(len(self.uploadTable.loc[i,"displayName"])):
                    displayName, fileSize, mime, filename, extension, documentId = self.uploadTable.loc[i,"displayName"], self.uploadTable.loc[i,"fileSize"], self.uploadTable.loc[i,"mime"], self.uploadTable.loc[i,"filename"], self.uploadTable.loc[i,"extension"], self.uploadTable.loc[i,"linkedDocId"]
                    presign_endpoint= "https://api-downer-rts.rdrive.io/api/newportbuild/documents/pre-signed-url?file-path=%2Flink-field-documents%2F" + api_formId + "%2F" + api_formFieldTemplateId +"%2F" + displayName[j]
                    response_url = requests.get(presign_endpoint, headers = self.token)
                    response = requests.put(url = response_url.text, headers = {"x-ms-blob-type" : "BlockBlob"})
                    documentBody = {"id": "",
                                    "storageId": response.headers["x-ms-version-id"],
                                    "documentId": "",
                                    "fileDownloadLink": "",
                                    "comment": "",
                                    "createdDate": "",
                                    "syncedDate": "",
                                    "createdByTeamId": "",
                                    "createdByTeam": "",
                                    "createdByUserId": "",
                                    "createdByUser": "",
                                    "resourceId": "",
                                    "fullPath": "/link-field-documents/" + api_formId + "/" + api_formFieldTemplateId + "/" + displayName[j],
                                    "mime": mime[j],
                                    "fileSize": fileSize[j],
                                    "displayName": "",
                                    "fileType": "",
                                    "parentFolderId": "",
                                    "isFolder": False,
                                    "colors": [],
                                    "dirname": "/link-field-documents/" + api_formId + "/" + api_formFieldTemplateId,
                                    "basename": displayName[j],
                                    "filename": filename[j],
                                    "extension": extension[j]}
                    
                    patchBody = [{  "id": "",
                                    "formId": api_formId,
                                    "formFieldTemplateId": api_formFieldTemplateId,
                                    "createdDate": "",
                                    "syncedDate": "",
                                    "createdByTeamId": "",
                                    "createdByTeam": "",
                                    "createdByUserId": "",
                                    "createdByUser": "",
                                    "value": "1",
                                    "meta": {
                                        "issueIds": [],
                                        "processIds": [],
                                        "documentIds": [documentId[j]],
                                        "changes": [{
                                            "action": "add",
                                            "type": "document",
                                            "value": documentId[j]}]
                                    }
                                }]
                    documentPostResponse = requests.post(documentPostingURL, headers = self.token, json = documentBody)
                    patchReponse = requests.patch(post_endpoint, headers = self.token, json = patchBody)

            body = {"id": "", 
                    "formId": api_formId, 
                    "formFieldTemplateId": api_formFieldTemplateId,
                    "createdDate": "",
                    "syncedDate": "",
                    "createdByTeamId": "",
                    "createdByTeam": "",
                    "createdByUserId": "",
                    "createdByUser": "",
                    "value": api_value,
                    "meta": api_meta_data,
                }
            response = requests.post(url = post_endpoint, headers = self.token, json = body)

