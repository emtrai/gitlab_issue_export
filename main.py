# Copyright (c) 2019 Ngo Huy Anh
# License type: Apache Licenses
# Author: Ngo Huy Anh
# Email: ngohuyanh@gmail.com, emtrai@gmail.com
# Created date: Apr. 30 2019
# Brief: Get issue from gitlab and export to file 


import sys
import os
import json
import xlwt
import datetime
import requests


#gitlab, we have
#- group: has multi project
#- project: is a git, has multi issue
#- issue: for each project


# True to enable debug log
DEBUG = False

# True to use dummy data
DUMMY_DATA = False #True

# list of default file name
DEFAULT_CONFIG = "config.ini"
GROUPS_TEST_FILE = "groups.json"
ISSUES_GRP_TEST_FILE = "issues_grp.json"
PROJECTS_TEST_FILE = "projects.json"
DEFAULT_URL = "https://google.com"
DEFAULT_API = "3"
DEFAULT_EXPORT_NAME = "export"
# definition
CONFIG_FIELD_SEPARATE = ":"
CONFIG_FIELD_VALUE_SPLIT = ","
CONFIG_FIELD_API = "api"
CONFIG_FIELD_URL = "url"
CONFIG_FIELD_TOKEN = "token"
CONFIG_FIELD_GROUPS = "groups"
CONFIG_FIELD_PROJECTS = "projects"
CONFIG_FIELD_AUTHORS = "authors"
CONFIG_FIELD_LABELS = "labels"
CONFIG_FIELD_EXPORTS = "exports"
CONFIG_FIELD_EXPORTNAME = "exportname"
CONFIG_FIELD_COMMENT = "#"

class Config(object):
    cfg = {}
    def __init__(self):
        self.cfg[CONFIG_FIELD_API] = DEFAULT_API
        self.cfg[CONFIG_FIELD_TOKEN] = ""
        self.cfg[CONFIG_FIELD_GROUPS] = []
        self.cfg[CONFIG_FIELD_PROJECTS] = []
        self.cfg[CONFIG_FIELD_AUTHORS] = []
        self.cfg[CONFIG_FIELD_LABELS] = []
        self.cfg[CONFIG_FIELD_URL] = DEFAULT_URL
        self.cfg[CONFIG_FIELD_EXPORTS] = []
        self.cfg[CONFIG_FIELD_EXPORTNAME] = DEFAULT_EXPORT_NAME
        return super(Config, self).__init__()

    def getToken(self):
        if (CONFIG_FIELD_TOKEN in self.cfg):
            return self.cfg[CONFIG_FIELD_TOKEN]
        return None
    def setToken(self, token):
        if (token is not None) and (len(token) > 0):
            self.cfg[CONFIG_FIELD_TOKEN] = token;
    def isExistIn(self, field, val):
        if (field in self.cfg):
            if ((val in self.cfg[field]) or len(self.cfg[field]) == 0):
                return True
        return False

    def getUrl(self):
        return self.cfg[CONFIG_FIELD_URL]

    def getApi(self):
        return self.cfg[CONFIG_FIELD_API]
        
    def getExports(self):
        if (CONFIG_FIELD_EXPORTS in self.cfg):
            return self.cfg[CONFIG_FIELD_EXPORTS]
        return None
    
    def getExportName(self):
        if (CONFIG_FIELD_EXPORTNAME in self.cfg):
            if (len(self.cfg[CONFIG_FIELD_EXPORTNAME]) > 0):
                return self.cfg[CONFIG_FIELD_EXPORTNAME]
        return DEFAULT_EXPORT_NAME

    def parseFile(self, path):
        """
        Parse configuration file
        """
        print ("parse file " + path)
        try:
            with open (path, 'rt') as f:
                for line in f:
                    #logD("config: " + line)
                    line = str.strip(line)
                    pos = line.find(CONFIG_FIELD_SEPARATE)
                    hdr = str.strip(line[:pos]).lower()
                    logD("hdr: " + hdr)
                    val = str.strip(line[pos+1:])
                    if (hdr.startswith(CONFIG_FIELD_COMMENT)):
                        continue
                    if (len(val) > 0):
                        #logD("val " + val)
                        if (hdr in self.cfg):
                            if (isinstance(self.cfg[hdr], list)):
                                logD("%s is in list type" % hdr)
                                tmpsVals = val.split(CONFIG_FIELD_VALUE_SPLIT)
                                vals = []
                                for item in tmpsVals:
                                    if (len(str.strip(item)) > 0) :
                                        vals.append(str.strip(item))
                                if (len(vals) > 0):
                                    self.cfg[hdr] = vals
                                #logD("vals %s" % vals)
                            else:
                                logD("%s is no list, it's normal value" % hdr)
                                self.cfg[hdr] = val
                        
            f.close()

        except:
            print (sys.exc_info()[0])
     

        return

    def dump(self):
        print ("config")
        print (self.cfg)
    def __repr__(self):
        return "Config class"

class gitlabUser(object):
    name = ""
    username = ""

    def parseData(self, jobj):
        if ("username" in jobj) and (jobj["username"] is not None):
            self.username = jobj["username"]
        if ("name" in jobj) and (jobj["name"] is not None):
            self.name = jobj["name"]  

    def __repr__(self):
        return  self.username

class gitlabObj(object):
    """
    gitlab commob object
    """
    name = None
    id = None
    desc = ""
    path = ""
    iid = None
    def parseData(self, jobj):
        """
        get data from json object
        """
        if ("id" in jobj):
            self.id = jobj["id"];
        if ("iid" in jobj):
            self.iid = jobj["iid"];
        if ("name" in jobj):
            self.name = jobj["name"]
        if ("desc" in jobj):
            self.desc = jobj["desc"]
        return
    
    def isValid(self):
        if (self.id is not None):
            return True
        else:
            return False

    def toString(self):
        str = ""
        if (self.name is not None):
            str += "name: %s, " % self.name
        if (self.id is not None):
            str += "id: %s, " % self.id
        if (self.iid is not None):
            str += "iid: %s, " % self.iid
        
        return str
    def __repr__(self):
        return self.toString()

class gitlabGroup(gitlabObj):
    """
    Git lab group object
    """
    projects = [] # list of projects

    def parseData(self, jobj):
        """
        Parse json data to get group info
        """
        super(gitlabGroup, self).parseData(jobj)
        if ("projects" in jobj):
            __jprj = jobj["projects"]
            for __item in jprj:
                __pj = gitlabProject()
                __pj.parseData(item)
                self.projects.append(__pj)
        if ("shared_projects" in jobj):
            __jprj = jobj["shared_projects"]
            for __item in jprj:
                __pj = gitlabProject()
                __pj.parseData(item)
                self.projects.append(__pj)

    
    def toString(self):
        str = super(gitlabGroup, self).toString()
        if (len(self.projects) > 0):
            for item in self.projects:
                str += "%s\n" % item
        return str

class gitlabProject(gitlabObj):
    """
    Gitlab project object
    """
    grp = None

    def __init__(self, grp):
        self.grp = grp
        return super(gitlabProject, self).__init__()

    def parseData(self, jobj):
        super(gitlabProject, self).parseData(jobj)
    
    def toString(self):
        str = super(gitlabProject, self).toString()
        return str

class gitlabIssue(gitlabObj):
    """
    Issue object of gitlab
    """
    project_id = ""
    milestone_due_date = ""
    author = None
    description = ""
    state = ""
    assignee = None
    labels = []
    title = ""
    updated_at = ""
    create_at = ""
    due_date = ""
    prj = None
    def __init__(self, prj):
        self.prj = prj
        self.labels = []
        return super(gitlabIssue, self).__init__()
    def parseData(self, jobj):
        """
        parse data for issue object
        """
        super(gitlabIssue, self).parseData(jobj)

        if ("project_id" in jobj):
            self.project_id = jobj["project_id"]
        if ("description" in jobj):
            self.description = jobj["description"]
        if ("state" in jobj):
            self.state = jobj["state"]
        if ("labels" in jobj):
            self.labels = jobj["labels"]
        if ("title" in jobj):
            self.title = jobj["title"]
        if ("updated_at" in jobj):
            self.updated_at = jobj["updated_at"]
        if ("created_at" in jobj):
            self.created_at = jobj["created_at"]
        if ("milestone" in jobj):
            if (jobj["milestone"] is not None) and ("due_date" in jobj["milestone"]):
                self.milestone_due_date = jobj["milestone"]["due_date"]
        if ("author" in jobj) and (jobj["author"] is not None):
            self.author = gitlabUser()
            self.author.parseData(jobj["author"])
        if ("assignee" in jobj) and (jobj["assignee"] is not None):
            self.assignee = gitlabUser()
            self.assignee.parseData(jobj["assignee"])
                     
        return

    def toString(self):
        str = super(gitlabIssue, self).toString() + "title %s " % self.title
        return str

    def __repr__(self):
        return self.toString()


class gitlabGroupList(object):
    """
    List of gitlab group
    """
    grpList = []
    def __init__(self):
        self.grpList = []
        return super(gitlabGroupList, self).__init__()
    def parseData(self, data):
        __jobj = json.loads(data)
        if (__jobj):
            for __item in __jobj:
                grp = gitlabGroup()
                grp.parseData(__item)
                if (grp.isValid()):
                    self.grpList.append(grp)

    def __repr__(self):
        str = ""
        if (len(self.grpList) > 0):
            for item in self.grpList:
                str += "*****\n"
                str = str + item.toString()
        else:
            str = "group empty"
        return str

class gitlabIssueList(object):
    """
    List of gitlab issues
    """
    issueList = []
    prj = None

    def __init__(self, prj):
        self.prj = prj 
        self.issueList = []
        return super(gitlabIssueList, self).__init__()

    def parseData(self, data):
        logD("parse data of issue list")
        __jobj = json.loads(data)
        if (__jobj):
            for __item in __jobj:
                issue = gitlabIssue(self.prj)
                issue.parseData(__item)
                
                if (issue.isValid()):
                    self.issueList.append(issue)
                    logD ("parsed issue %s, count %d" % (issue, len(self.issueList)))

    def __repr__(self):
        str = ""
        if (len(self.issueList) > 0):
            for item in self.issueList:
                str += "*****\n"
                str = str + item.toString()
        else:
            str = "issues empty"
        return str

class gitlabProjectList(object):
    """
    List of gitlab project
    """
    prjList = []
    grp = None
    def __init__(self, grp):
        self.grp = grp
        self.prjList = []
        return super(gitlabProjectList, self).__init__()
    def parseData(self, data):
        __jobj = json.loads(data)
        if (__jobj):
            for __item in __jobj:
                prj = gitlabProject(self.grp)
                prj.parseData(__item)
                if (prj.isValid()):
                    self.prjList.append(prj)

    def __repr__(self):
        str = ""
        if (len(self.prjList) > 0):
            for item in self.prjList:
                str += "\n*****\n"
                str = str + item.toString()
        else:
            str = "prj empty"
        return str
##################### FUNCIONS DECLARE #####################


def logD(str):
    if DEBUG:
        print (str)

def getFullFilePath(fileName):
    curDir = os.path.dirname(os.path.abspath(__file__))
    if (os.name is "nt"): #window
        testFile = curDir + "\\" + fileName
    else:
        testFile = curDir + "/" + fileName
    return testFile
def getApiUrl(config, path):
    url = "%s/api/v%s/%s" % (config.getUrl(), config.getApi(), path)
    return url


def getListGroups(config):
    """
    Get list of groups
    """
    data = None
    grpList = None
    if (DUMMY_DATA):
        curDir = os.path.dirname(os.path.abspath(__file__))
        testFile = getFullFilePath(GROUPS_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
                        
        f.close()
    else:
        # TODO: do paging
        # retrieve data from server
        url = getApiUrl(config, "groups")
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        params = {'per_page': 200}
        logD("header %s" % hdrs)
        resp = requests.get(url, headers=hdrs, params=params)
        logD("resp status_code %s" % resp.status_code)
        
        if (resp.status_code == 200):
            data = resp.content

    if (data is not None) and (len(data) > 0):
        logD("data %s" % data)
        __grpList = gitlabGroupList()
        __grpList.parseData(data)
        return __grpList
    return None

def getListProjectsInGroup(config, grp):
    """
    Get list of issue in group
    """
    logD("get list issue of group %s " % grp.id)
    data = None
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()
    else:
        # TODO: do paging
        # retrieve data from server
        url = getApiUrl(config, "groups/%s/projects" % grp.id)
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        params = {'per_page': 200}
        logD("header %s" % hdrs)
        resp = requests.get(url, headers=hdrs, params=params)
        logD("resp status_code %s" % resp.status_code)
        
        if (resp.status_code == 200):
            data = resp.content

    if (data is not None) and len(data) > 0:
        logD("data %s" % data)
        __prjLst = gitlabProjectList(grp)
        __prjLst.parseData(data)
        return __prjLst
    else:
        return None


def getListIssuesInGroup(config, groupId):
    """
    Get list of issue in group
    """
    logD("get list issue of group %s " % groupId)
    data = None
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()
    else:
        # TODO: do paging
        # retrieve data from server
        url = getApiUrl(config, "groups/%s/issues" % groupId)
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        params = {'per_page': 200}
        logD("header %s" % hdrs)
        resp = requests.get(url, headers=hdrs, params=params)
        logD("resp status_code %s" % resp.status_code)
        
        if (resp.status_code == 200):
            data = resp.content

    if (data is not None) and len(data) > 0:
        logD("data %s" % data)
        __issueLst = gitlabIssueList()
        __issueLst.parseData(data)
        return __issueLst
    else:
        return None

def getListIssuesInProject(config, prj):
    """
    Get list of issue in project
    """
    logD("get list issue of project %s " % prj.id)
    data = None
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()
    else:
        # retrieve data from server
        url = getApiUrl(config, "projects/%s/issues" % prj.id)
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        logD("header %s" % hdrs)
        resp = requests.get(url, headers=hdrs)
        logD("resp status_code %s" % resp.status_code)
        
        if (resp.status_code == 200):
            data = resp.content

    if (data is not None) and len(data) > 0:
        logD("data %s" % data)
        __issueLst = gitlabIssueList(prj)
        __issueLst.parseData(data)
        if (__issueLst.issueList is not None):
            logD ("found %d issues" % len(__issueLst.issueList))
            return __issueLst
        
    return None

def retrieveDataFromServer(url):
    return

def exportToExcel(config, issueList, path, sheetName, workbook):
    saveToFile = False
    if (workbook is None):
        workbook = xlwt.Workbook()
        saveToFile = True
      
    sheet = workbook.add_sheet(sheetName) 
    
    count = 0
    col = 0
    row = 0
    col = 0

    count = count + 1
    sheet.write(row, col, "No")

    col += 1
    sheet.write(row, col, "Id")

    col += 1
    sheet.write(row, col, "IId")

    col += 1
    sheet.write(row, col, "title") 

    col += 1
    sheet.write(row, col, "status") 

    col += 1
    sheet.write(row, col, "assignee") 

    col += 1
    sheet.write(row, col, "author")


    col += 1
    sheet.write(row, col, "milestone")

    col += 1
    sheet.write(row, col, "project")
    
    col += 1
    sheet.write(row, col, "group")

    col += 1
    sheet.write(row, col, "created date")

    col += 1
    sheet.write(row, col, "updated date")

    col += 1
    sheet.write(row, col, "label")
    
    col += 1
    sheet.write(row, col, "link")
    
    for __issue in issueList:
        row += 1
        col = 0
        
        sheet.write(row, col, count)

        col += 1
        sheet.write(row, col, __issue.id)
        

        col += 1
        if (__issue.iid is not None):
          sheet.write(row, col, __issue.iid)

        col += 1
        sheet.write(row, col, __issue.title) 

        col += 1
        sheet.write(row, col, "%s" % __issue.state) 

        col += 1
        sheet.write(row, col, "%s" % __issue.assignee) 

        col += 1
        sheet.write(row, col, "%s" % __issue.author) 


        col += 1
        sheet.write(row, col, __issue.milestone_due_date) 

        col += 1
        sheet.write(row, col, __issue.prj.name) 

        col += 1
        sheet.write(row, col, __issue.prj.grp.name) 

        col += 1
        sheet.write(row, col, __issue.created_at) 

        col += 1
        sheet.write(row, col, __issue.updated_at) 


        col += 1
        sheet.write(row, col, "%s" % __issue.labels) 
        
        col += 1
        __link = "%s/%s/%s/issues/%s" % (config.getUrl(), \
                                        __issue.prj.grp.name, \
                                        __issue.prj.name, \
                                        __issue.iid)
        sheet.write(row, col, xlwt.Formula('HYPERLINK("{}", "{}")'.format(__link, __link)))
        #sheet.write(row, col, "%s" % __link) 
        
        count += 1
    
    if (saveToFile):
        try:
            workbook.save(path)
        except:
            print ("FAILED TO WRITE TO FILE " + path)
            print ("ERROR %s" % sys.exc_info()[0])
    
    return workbook
#################################################################
############################# START EXECUTION ###################

def main():
    """
    Entry function
    """
    print (sys.argv)
    print ("os name %s" % os.name)
    #os.chdir(os.path.dirname(__file__))
    #print os.getcwd()
    configFileName = DEFAULT_CONFIG;
    if (len(sys.argv) > 1):
        if (sys.argv[1] is not None):
            configFileName = sys.argv[1]
    
    logD("config name %s" % configFileName)

    configFile = getFullFilePath(configFileName)
    config = Config()
    config.parseFile(configFile)
    

    token = config.getToken()
    if (token is None) or (len(token) == 0):
        inputToken = raw_input('Enter private token: ')
        if (inputToken is not None) and (len(inputToken) > 0):
            config.setToken(inputToken)
            token = inputToken

    config.dump()
    # 1. get list of groups
    # 2. get list of project of a group
    # 3. get list of issues of project
    # or (api v4)
    # 1. get list of groups
    # 2. get list of issues of a groups


    grpList = getListGroups(config)

    if (grpList is not None):
        logD ("list of group %s" % grpList)

        prjList = []
        noPrj = 0
        for grp in grpList.grpList:
            if (config.isExistIn(CONFIG_FIELD_GROUPS, grp.name)):
                print ("found group %s" % grp.name)
                __prjLst = None
                __prjLst = getListProjectsInGroup(config, grp)
                if (__prjLst is not None):
                    print ("group %s has %d project" % (grp.name, len(__prjLst.prjList)))
                    noPrj += len(__prjLst.prjList)
                    prjList.extend(__prjLst.prjList)
            else:
                print ("ignore group %s" % grp.name)

        
        logD ("list of prj %s" % prjList)
        print ("number of prj %d, found %d" % (len(prjList), noPrj))

        issueList = []
        noIssue = 0
        for __prj in prjList:
            __lst = getListIssuesInProject(config, __prj)
            if (__lst is not None):
                print ("project %s has %d issue" % (__prj.name, len(__lst.issueList)))
                noIssue += len(__lst.issueList)
                issueList.extend(__lst.issueList)
            
        
        print ("number of issue %d, found %d" % (len(issueList), noIssue))

        exports = config.getExports()
        currentDT = datetime.datetime.now()
        exportFileName = "%s_%s_%s" % (config.getExportName(), \
                                      os.path.splitext(configFileName)[0], \
                                      currentDT.strftime("%Y%m%d_%H%M%S"))
        if ("xlsx" in exports) or ("xls" in exports):
           
            path = getFullFilePath("%s.xls" % exportFileName)
            print("export to excel, path %s" % path)
            exportToExcel(config, issueList, path, "issueList", None)
            print("export done")
    return

####################################################################################
######################################## START RUNNING #############################
main()