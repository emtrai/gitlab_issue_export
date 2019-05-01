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
#import requests


#gitlab, we have
#- group: has multi project
#- project: is a git, has multi issue
#- issue: for each project


# True to enable debug log
DEBUG = True

# True to use dummy data
DUMMY_DATA = True

# list of default file name
DEFAULT_CONFIG = "config.ini"
GROUPS_TEST_FILE = "groups.json"
ISSUES_GRP_TEST_FILE = "issues_grp.json"
PROJECTS_TEST_FILE = "projects.json"

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


class Config(object):
    cfg = {}
    def __init__(self):
        self.cfg[CONFIG_FIELD_API] = 3
        self.cfg[CONFIG_FIELD_TOKEN] = ""
        self.cfg[CONFIG_FIELD_GROUPS] = []
        self.cfg[CONFIG_FIELD_PROJECTS] = []
        self.cfg[CONFIG_FIELD_AUTHORS] = []
        self.cfg[CONFIG_FIELD_LABELS] = []
        self.cfg[CONFIG_FIELD_URL] = []
        self.cfg[CONFIG_FIELD_EXPORTS] = []
        return super(Config, self).__init__()

    def getToken(self):
        if (self.cfg.has_key(CONFIG_FIELD_TOKEN)):
            return self.cfg[CONFIG_FIELD_TOKEN]
        return None
    def getExports(self):
        if (self.cfg.has_key(CONFIG_FIELD_EXPORTS)):
            return self.cfg[CONFIG_FIELD_EXPORTS]
        return None
    def parseFile(self, path):
        """
        Parse configuration file
        """
        print "parse file " + path
        try:
            with open (path, 'rt') as f:
                for line in f:
                    #logD("config: " + line)
                    line = str.strip(line)
                    pos = line.find(CONFIG_FIELD_SEPARATE)
                    hdr = str.strip(line[:pos]).lower()
                    #logD("hdr: " + hdr)
                    val = str.strip(line[pos+1:])
                    if (len(val) > 0):
                        #logD("val " + val)
                        if (self.cfg.has_key(hdr)):
                            tmpsVals = val.split(CONFIG_FIELD_VALUE_SPLIT)
                            vals = []
                            for item in tmpsVals:
                                if (len(str.strip(item)) > 0) :
                                    vals.append(item)
                            if (vals.count > 0):
                                self.cfg[hdr] = vals
                            #logD("vals %s" % vals)
                        
            f.close()

        except:
            print sys.exc_info()[0]
     

        return

    def dump(self):
        print "config"
        print self.cfg
    def __repr__(self):
        return "Config class"

class gitlabUser(object):
    name = ""
    username = ""

    def parseData(self, jobj):
        if ("username" in jobj):
            self.username = jobj["username"]
        if ("name" in jobj):
            self.name = jobj["name"]   

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
        if (self.projects.count > 0):
            for item in self.projects:
                str += "%s\n" % item
        return str

class gitlabProject(gitlabObj):
    """
    Gitlab project object
    """
    def parseData(self, jobj):
        super(gitlabProject, self).parseData(jobj)
    
    def toString(self):
        str = super(gitlabGroup, self).toString()
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
    label = []
    title = ""
    updated_at = ""
    create_at = ""
    due_date = ""

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
        if ("iid" in jobj):
            self.iid = jobj["iid"]
        if ("labels" in jobj):
            self.labels = jobj["labels"]
        if ("title" in jobj):
            self.title = jobj["title"]
        if ("updated_at" in jobj):
            self.updated_at = jobj["updated_at"]
        if ("created_at" in jobj):
            self.created_at = jobj["created_at"]
        if ("milestone" in jobj):
            if ("due_date" in jobj["milestone"]):
                self.milestone_due_date = jobj["milestone"]["due_date"]
        if ("author" in jobj):
            self.author = gitlabUser()
            self.author.parseData(jobj["author"])
        if ("assignee" in jobj):
            self.assignee = gitlabUser()
            self.assignee.parseData(jobj["assignee"])
                     
        return

    def toString(self):
        return super(gitlabIssue, self).toString()

    def __repr__(self):
        return self.toString()


class gitlabGroupList(object):
    """
    List of gitlab group
    """
    grpList = []
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
        if (self.grpList.count > 0):
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
    issueList = {}
    def parseData(self, data):
        __jobj = json.loads(data)
        if (__jobj):
            for __item in __jobj:
                issue = gitlabIssue()
                issue.parseData(__item)
                if (issue.isValid()):
                    self.issueList[issue.iid] = issue

    def __repr__(self):
        str = ""
        if (self.issueList.count > 0):
            for item in self.issueList:
                str += "*****\n"
                str = str + item.toString()
        else:
            str = "issues empty"
        return str


##################### FUNCIONS DECLARE #####################
config = Config()



def logD(str):
    if DEBUG:
        print str

def getFullFilePath(fileName):
    curDir = os.path.dirname(os.path.abspath(__file__))
    testFile = curDir + "/" + fileName
    return testFile

def getListGroups():
    """
    Get list of groups
    """
    data = None
    if (DUMMY_DATA):
        curDir = os.path.dirname(os.path.abspath(__file__))
        testFile = getFullFilePath(GROUPS_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
                        
        f.close()
    #else:
        # retrieve data from server

    if (data is not None):
        grpList = gitlabGroupList()
        grpList.parseData(data)

    return grpList

def getListProjects(groupId):
    return


def getListIssuesInGroup(groupId):
    """
    Get list of issue in group
    """
    data = None
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()

    if (data is not None):
        issueLst = gitlabIssueList()
        issueLst.parseData(data)

    return issueLst

def retrieveDataFromServer(url):
    return

def exportToExcel(issueList, path, sheetName, workbook):
    saveToFile = False
    if (workbook is None):
        workbook = xlwt.Workbook()
        saveToFile = True
      
    sheet = workbook.add_sheet(sheetName) 
    
    count = 0
    col = 0
    row = 0
    
    for key, val in issueList.items():
        col = 0
        count = count + 1
        sheet.write(row, col, count)
        col += 1
        sheet.write(row, col, key)
        col += 1
        sheet.write(row, col, val.title) 
        row += 1
    
    if (saveToFile):
        workbook.save(path)

    return workbook
#################################################################
############################# START EXECUTION ###################

def main():
    """
    Entry function
    """
    print 'hello'
    print sys.argv
    print "os name %s" % os.name
    #os.chdir(os.path.dirname(__file__))
    #print os.getcwd()
    
    configFile = getFullFilePath(DEFAULT_CONFIG)
    config.parseFile(configFile)
    config.dump()

    token = config.getToken()
    if (token is not None):
        print token


    # 1. get list of groups
    # 2. get list of project of a group
    # 3. get list of issues of project
    # or
    # 1. get list of groups
    # 2. get list of issues of a groups


    grpList = getListGroups()


    print grpList

    issueList = {}
    for grp in grpList.grpList:
        __lst = getListIssuesInGroup(grp.id)
        issueList.update(__lst.issueList)
    print issueList

    exports = config.getExports()
    if ("xlsx" in exports):
        exportToExcel(issueList, getFullFilePath("export.xls"), "issueList", None)
    return

main()