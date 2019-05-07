# Copyright (c) 2019 Ngo Huy Anh
# Author: Ngo Huy Anh
# Email: ngohuyanh@gmail.com, emtrai@gmail.com
# Created date: Apr. 30 2019
# Brief: Retrieve info from gitlab and export to file 

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.



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

# gitlab api supports
APIS_SUPPORT = ["3", "4"]
EXPORTS_SUPPORT = ["xls", "xlsx"]

# True to enable debug log, via logD function
DEBUG = False

# True to use dummy data, read from file
DUMMY_DATA = False #True

# paramameters definition, i.e main.py c=d
PARAM_SPLIT = "="
PARAM_INFO_SPLIT = ","
PARAM_CFG = "c"
PARAM_INFO = "l"
#new param? update this list as well
PARAM_SUPPORT = [PARAM_CFG, PARAM_INFO]
USAGE = "USAGE: \tpython main.py %s=<val> %s=<val>" % (PARAM_CFG, PARAM_INFO)
USAGE_API = "Supported GITLAB APIS: %s" % APIS_SUPPORT
USAGE_PARAM={}
USAGE_PARAM[PARAM_CFG] = "Specify config file to be used, i.e. config_ABC.ini.\n\tNot specify, default is config.ini"
USAGE_PARAM[PARAM_INFO] = "Specify info to be exported, separate by \",\", i.e. prj (project list), iss(issue), grp(group).\n\tNot specify, default is iss,grp,prj"


PARAM_INFO_ISS = "iss" # issue list
PARAM_INFO_PRJ = "prj" # project list
PARAM_INFO_GRP = "grp" # group list

# list of default file name
DEFAULT_CONFIG = "config.ini"
GROUPS_TEST_FILE = "groups.json"
ISSUES_GRP_TEST_FILE = "issues_grp.json"
PROJECTS_TEST_FILE = "projects.json"
DEFAULT_URL = "https://google.com"
DEFAULT_API = "3"
DEFAULT_EXPORT_NAME = "export"
DEFAULT_EXPORTS = "xls"
# definition
CONFIG_FIELD_SEPARATE = ":"
CONFIG_FIELD_VALUE_SPLIT = ","
CONFIG_FIELD_API = "api"
CONFIG_FIELD_URL = "url"
CONFIG_FIELD_TOKEN = "token"
CONFIG_FIELD_GROUPS = "groups"
CONFIG_FIELD_GROUP_MAX = "maxgroup"
CONFIG_FIELD_PROJECT_MAX = "maxproject"
CONFIG_FIELD_ISSUE_MAX = "maxissue"
CONFIG_FIELD_PROJECTS = "projects"
CONFIG_FIELD_AUTHORS = "authors"
CONFIG_FIELD_LABELS = "labels"
CONFIG_FIELD_EXPORTS = "exports"
CONFIG_FIELD_EXPORTNAME = "exportname"
CONFIG_FIELD_COMMENT = "#"

class Config(object):
    """
    Configuration class
    """
    cfg = {}
    def __init__(self):
        self.cfg[CONFIG_FIELD_API] = DEFAULT_API
        self.cfg[CONFIG_FIELD_TOKEN] = ""
        self.cfg[CONFIG_FIELD_GROUPS] = []
        self.cfg[CONFIG_FIELD_PROJECTS] = []
        self.cfg[CONFIG_FIELD_AUTHORS] = []
        self.cfg[CONFIG_FIELD_LABELS] = []
        self.cfg[CONFIG_FIELD_URL] = DEFAULT_URL
        self.cfg[CONFIG_FIELD_EXPORTS] = [DEFAULT_EXPORTS]
        self.cfg[CONFIG_FIELD_EXPORTNAME] = DEFAULT_EXPORT_NAME
        self.cfg[CONFIG_FIELD_GROUP_MAX] = ""
        self.cfg[CONFIG_FIELD_PROJECT_MAX] = ""
        self.cfg[CONFIG_FIELD_ISSUE_MAX] = ""

        return super(Config, self).__init__()

    def getMaxValue(self, hdr):
        """
        get the number of item to be get
        """
        logD("get max number of %s" % hdr)
        if (hdr in self.cfg) and len(self.cfg[hdr]) > 0:
            __val = int(self.cfg[hdr])
            logD("max val %d" % __val)
            return __val
        else:
            logD("max val is.. infinity")
            return None
    
    def getMaxIssue(self):
        """
        get the number of issues to be get
        """
        return self.getMaxValue(CONFIG_FIELD_ISSUE_MAX)

    def getMaxProject(self):
        """
        get the number of issues to be get
        """
        return self.getMaxValue(CONFIG_FIELD_PROJECT_MAX)

    def getMaxGroup(self):
        """
        get the number of issues to be get
        """
        return self.getMaxValue(CONFIG_FIELD_GROUP_MAX)

    def getToken(self):
        """
        get private token
        """
        if (CONFIG_FIELD_TOKEN in self.cfg):
            return self.cfg[CONFIG_FIELD_TOKEN]
        return None
    
    def setToken(self, token):
        """
        Set private token
        """
        if (token is not None) and (len(token) > 0):
            self.cfg[CONFIG_FIELD_TOKEN] = token;

    
    def isExistIn(self, field, val):
        """
        Check if value exist in a configuration.
        If configuration is null, mean it'll exist
        """
        if (field in self.cfg):
            # TODO: consider again about the case that configuration's value is empty
            if ((val in self.cfg[field]) or len(self.cfg[field]) == 0):
                return True
        return False

    def getUrl(self):
        """
        Get host URL of gitlab
        """
        return self.cfg[CONFIG_FIELD_URL]

    def getApi(self):
        """
        Get api version
        """
        return self.cfg[CONFIG_FIELD_API]
        
    def getExports(self):
        """
        Get list of export methods
        """
        if (CONFIG_FIELD_EXPORTS in self.cfg):
            return self.cfg[CONFIG_FIELD_EXPORTS]
        return None
    
    def getExportName(self):
        """
        Get exports name
        """
        if (CONFIG_FIELD_EXPORTNAME in self.cfg):
            if (len(self.cfg[CONFIG_FIELD_EXPORTNAME]) > 0):
                return self.cfg[CONFIG_FIELD_EXPORTNAME]
        return DEFAULT_EXPORT_NAME

    def parseFile(self, path):
        """
        Parse configuration file. return True if success
        """
        print ("parse file " + path)
        try:
            with open (path, 'rt') as f:
                for line in f:
                    logD("config: " + line)

                    # read line by line and split it basing on separator
                    line = str.strip(line)
                    pos = line.find(CONFIG_FIELD_SEPARATE)
                    hdr = str.strip(line[:pos]).lower()
                    logD("hdr: " + hdr)

                    val = str.strip(line[pos+1:])
                    # ignore if it's comment
                    if (hdr.startswith(CONFIG_FIELD_COMMENT)):
                        continue
                    if (len(val) > 0):
                        #logD("val " + val)
                        # check if config is support
                        if (hdr in self.cfg):
                            # if config is list of iss, separate its value
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
                            else: # or use value directly
                                logD("%s is no list, it's normal value" % hdr)
                                self.cfg[hdr] = val
                        
            f.close()

        except:
            print ("PARSING CONFIG ERROR %s" % sys.exc_info()[0])
            return False # parsing error
     
        
        return True

    def dump(self):
        print ("config")
        print (self.cfg)
    def __repr__(self):
        return "Config class"

class gitlabUser(object):
    """
    git lab user object
    """
    name = ""
    username = ""

    def parseData(self, jobj):
        """
        parse data
        """
        if ("username" in jobj) and (jobj["username"] is not None):
            self.username = jobj["username"]
        if ("name" in jobj) and (jobj["name"] is not None):
            self.name = jobj["name"]  

    def __repr__(self):
        return  self.username

class gitlabObj(object):
    """
    gitlab common object
    """
    name = None
    id = None # uniqe id
    desc = ""
    path = ""
    iid = None # id in a group/project, ...

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

        # get list of projects if any
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
    grp = None # group that project belongs to

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
    prj = None # project that issue belong to

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

    def getLen(self):
        logD("grp len %d" % len(self.grpList))
        return len(self.grpList)

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

    def getLen(self):
        logD("issue len %d" % len(self.issueList))
        return len(self.issueList)
        
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

    def getLen(self):
        logD("prj len %d" % len(self.prjList))
        return len(self.prjList)
        
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
    """
    get full file path, basing on path of main.py
    """
    curDir = os.path.dirname(os.path.abspath(__file__))
    if (os.name is "nt"): #window
        testFile = curDir + "\\" + fileName
    else: # posix, like MAC, Linux
        testFile = curDir + "/" + fileName
    return testFile

def getApiUrl(config, path):
    """
    get final url to be used
    """
    url = "%s/api/v%s/%s" % (config.getUrl(), config.getApi(), path)
    return url


def getListGroups(config):
    """
    Get list of groups
    """
    print("Retrieve list of group")
    data = None
    grpList = None

    __grpList = gitlabGroupList()
    if (DUMMY_DATA):
        curDir = os.path.dirname(os.path.abspath(__file__))
        testFile = getFullFilePath(GROUPS_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
                        
        f.close()
    else:
        # retrieve data from server
        url = getApiUrl(config, "groups")
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        
        
        __totalPage = 0
        __page = 1
        while True:
            logD("Page %d" % (__page))
            params = {'page': __page}
            logD("header %s" % hdrs)
            resp = requests.get(url, headers=hdrs, params=params)
            logD("resp status_code %s" % resp.status_code)

            if (resp.status_code == 200):
                data = resp.content
                logD (resp.headers)
                if (len(resp.headers.get('X-Next-Page')) > 0):
                    __page = int(resp.headers.get('X-Next-Page'))
                else:
                    __page = 0
                logD("next page %d" % (__page))
            else:
                __page = 0
                break

            if (data is not None) and (len(data) > 0):
                logD("data %s" % data)
                __grpList.parseData(data)
    
            
            __totalPage += 1
            if (config.getMaxGroup() is not None) and (__grpList.getLen() >= config.getMaxGroup()):
                print("Reach max %s/%s" % (__grpList.getLen(), config.getMaxGroup()))
                break

            if (__page == 0): #ok, reach end, out
                break
            if (__totalPage > 500): # 500 pages? no way, something wrong, out
                print("SOMETHING WRONG, total is to big, out")
                break
    print("Total pages %d" % (__totalPage))
    return __grpList

def getListProjectsInGroup(config, grp):
    """
    Get list of issue in group
    """
    print("Retrieve project of group: %s " % grp.name)
    data = None
    __prjLst = gitlabProjectList(grp)
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()
    else:
        # retrieve data from server
        url = getApiUrl(config, "groups/%s/projects" % grp.id)
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        
        __totalPage = 0
        __page = 1
        while True:
            logD("Page %d" % (__page))
            params = {'page': __page}
            logD("header %s" % hdrs)
            resp = requests.get(url, headers=hdrs, params=params)
            logD("resp status_code %s" % resp.status_code)

            if (resp.status_code == 200):
                data = resp.content
                logD (resp.headers)
                if (len(resp.headers.get('X-Next-Page')) > 0):
                    __page = int(resp.headers.get('X-Next-Page'))
                else:
                    __page = 0
                logD("next page %d" % (__page))
            else:
                __page = 0
                break

            if (data is not None) and len(data) > 0:
                logD("data %s" % data)
                __prjLst.parseData(data)
    
            
            __totalPage += 1

            if (config.getMaxProject() is not None) and (__prjLst.getLen() >= config.getMaxProject()):
                print("Reach max %s/%s" % (__prjLst.getLen(), config.getMaxProject()))
                break

            if (__page == 0): #ok, reach end, out
                break
            if (__totalPage > 500): # 500 pages? no way, something wrong, out
                print("SOMETHING WRONG, total is to big, out")
                break
    print("Total pages %d" % (__totalPage))
    return __prjLst
    

# for API 4 and beyon only
def getListIssuesInGroup(config, groupId):
    """
    Get list of issue in group
    """
    logD("get list issue of group %s " % groupId)
    data = None
    __issueLst = gitlabIssueList()
    if (DUMMY_DATA):
        testFile = getFullFilePath(ISSUES_GRP_TEST_FILE)
        with open (testFile, 'rt') as f:
            data = f.read()
        
        f.close()
    else:
        # retrieve data from server
        url = getApiUrl(config, "groups/%s/issues" % groupId)
        logD("URL " + url)
        token = config.getToken()
        
        hdrs = {"PRIVATE-TOKEN":config.getToken()}
        
        __totalPage = 0
        __page = 1
        while True:
            logD("Page %d" % (__page))
            params = {'page': __page}
            logD("header %s" % hdrs)
            resp = requests.get(url, headers=hdrs, params=params)
            logD("resp status_code %s" % resp.status_code)

            if (resp.status_code == 200):
                data = resp.content
                logD (resp.headers)
                if (len(resp.headers.get('X-Next-Page')) > 0):
                    __page = int(resp.headers.get('X-Next-Page'))
                else:
                    __page = 0
                logD("next page %d" % (__page))
            else:
                __page = 0
                break

            if (data is not None) and len(data) > 0:
                logD("data %s" % data)
                __issueLst.parseData(data)
    
            __totalPage += 1
            logD("Total pages %d" % (__totalPage))

            if (config.getMaxIssue() is not None) and (__issueLst.getLen() >= config.getMaxIssue()):
                print("Reach max %s/%s" % (__issueLst.getLen(), config.getMaxIssue()))
                break

            if (__page == 0): #ok, reach end, out
                break
            if (__totalPage > 500): # 500 pages? no way, something wrong, out
                print("SOMETHING WRONG, total is to big, out")
                break

    return __issueLst

def getListIssuesInProject(config, prj):
    """
    Get list of issue in project
    """
    print("Retrieve issue of project: %s " % prj.name)
    data = None
    __issueLst = gitlabIssueList(prj)
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
        
        
        __totalPage = 0
        __page = 1
        while True:
            logD("Page %d" % (__page))
            params = {'page': __page}
            logD("header %s" % hdrs)
            resp = requests.get(url, headers=hdrs, params=params)
            logD("resp status_code %s" % resp.status_code)

            if (resp.status_code == 200):
                data = resp.content
                logD (resp.headers)
                if (len(resp.headers.get('X-Next-Page')) > 0):
                    __page = int(resp.headers.get('X-Next-Page'))
                else:
                    __page = 0
                logD("next page %d" % (__page))
            else:
                __page = 0
                break


            if (data is not None) and len(data) > 0:
                logD("data %s" % data)
                __issueLst.parseData(data)
                if (__issueLst.issueList is not None):
                    logD ("found %d issues" % len(__issueLst.issueList))

            

            __totalPage += 1

            if (config.getMaxIssue() is not None) and (__issueLst.getLen() >= config.getMaxIssue()):
                print("Reach max %s/%s" % (__issueLst.getLen(), config.getMaxIssue()))
                break
            if (__page == 0): #ok, reach end, out
                break
            if (__totalPage > 500): # 500 pages? no way, something wrong, out
                print("SOMETHING WRONG, total is to big, out")
                break
    print("Total pages %d" % (__totalPage))
    return __issueLst

def exportIssueToExcel(config, issueList, path, workbook):
    """
    Export issue list to excel
    """
    print("Export issue list")
    if (workbook is None):
        workbook = xlwt.Workbook()
      
    sheet = workbook.add_sheet("issues") 
    
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
    
    try:
        workbook.save(path)
    except:
        print ("FAILED TO WRITE TO FILE " + path)
        print ("ERROR %s" % sys.exc_info()[0])
    finally:
        return workbook

def exportProjectToExcel(config, prjList, path, workbook):
    """
    Export project list to excel
    """
    print("Export project list")
    if (workbook is None):
        workbook = xlwt.Workbook()
      
    sheet = workbook.add_sheet("project") 
    
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
    sheet.write(row, col, "name") 
    
    col += 1
    sheet.write(row, col, "group")
    
    col += 1
    sheet.write(row, col, "link")
    
    for __prj in prjList:
        row += 1
        col = 0
        
        sheet.write(row, col, count)

        col += 1
        sheet.write(row, col, __prj.id)
        

        col += 1
        if (__prj.iid is not None):
          sheet.write(row, col, __prj.iid)

        col += 1
        sheet.write(row, col, __prj.name) 

        col += 1
        sheet.write(row, col, __prj.grp.name) 
        
        col += 1
        __link = "%s/%s//%s" % (config.getUrl(), \
                                __prj.grp.name, \
                                __prj.name)
        sheet.write(row, col, xlwt.Formula('HYPERLINK("{}", "{}")'.format(__link, __link)))

        count += 1
    
    try:
        workbook.save(path)
    except:
        print ("FAILED TO WRITE TO FILE " + path)
        print ("ERROR %s" % sys.exc_info()[0])
    finally:
        return workbook

def exportGroupToExcel(config, grpList, path, workbook):
    """
    Export group to excel
    """
    if (workbook is None):
        workbook = xlwt.Workbook()
    
    print("Export group list")
    sheet = workbook.add_sheet("group") 
    
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
    sheet.write(row, col, "name") 
    
    for __grp in grpList:
        row += 1
        col = 0
        
        sheet.write(row, col, count)

        col += 1
        sheet.write(row, col, __grp.id)
        

        col += 1
        if (__grp.iid is not None):
          sheet.write(row, col, __grp.iid)

        col += 1
        sheet.write(row, col, __grp.name) 
        
        col += 1
        __link = "%s/%s" % (config.getUrl(), __grp.name)
        sheet.write(row, col, xlwt.Formula('HYPERLINK("{}", "{}")'.format(__link, __link)))

        count += 1
    
    try:
        workbook.save(path)
    except:
        print ("FAILED TO WRITE TO FILE " + path)
        print ("ERROR %s" % sys.exc_info()[0])
    finally:
        return workbook


def usage():
    """
    Print usages
    """
    print(USAGE)
    print(USAGE_API)

    #print paramegers
    for __key, __val in USAGE_PARAM.items():
        print ("%s\t%s" % (__key, __val))

def parseParameter():
    """
    Parse input parameters
    """
    print ("parse parameters %s" % sys.argv)
    __args = {}
    if (len(sys.argv) > 1):
        for arg in sys.argv[1:]:
            __tmp = str.split(arg, PARAM_SPLIT)
            if (__tmp is not None) and (len(__tmp) > 1):
                if (__tmp[0] in PARAM_SUPPORT):
                    if (__tmp[1] is not None) and (len(str.strip(__tmp[1])) > 0):
                        __args[__tmp[0]] = str.strip(__tmp[1])
                else:
                    usage()
                    sys.exit("PARAM %s not support" % __tmp[0]) 
            else:
                return None
    return __args


#################################################################
############################# START EXECUTION ###################

def main():
    """
    Entry function
    """
    

    print ("os name %s" % os.name)

    # parse input parameters
    __args = parseParameter()
    if (__args is None):
        usage()
        return
    print("param %s" % __args)
    
    configFileName = DEFAULT_CONFIG;
    
    reqs = []
    if (PARAM_CFG in __args):
        configFileName = __args[PARAM_CFG]
    if (PARAM_INFO in __args):
        reqs = str.split(__args[PARAM_INFO], PARAM_INFO_SPLIT)
    else:
        reqs = [PARAM_INFO_ISS, PARAM_INFO_PRJ, PARAM_INFO_GRP]
    
    print("Request inf: %s" % reqs)
    logD("config name %s" % configFileName)

    # parse configuration file
    configFile = getFullFilePath(configFileName)
    config = Config()
    config.parseFile(configFile)
    
    print ("API: %s" % config.getApi())

    if (config.getApi() not in APIS_SUPPORT):
        print ("THIS API VERSION (%s) not support. SUPPORTED APIS IS %s" %(config.getApi(), APIS_SUPPORT))
        return
    
    # get token, ask to input if it's empty
    token = config.getToken()
    if (token is None) or (len(token) == 0):
        if (sys.version_info < (3,0)):
            inputToken = raw_input('Enter private token: ')
        else:
            inputToken = input('Enter private token: ')
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
    useGrp = []
    # get group lis
    if (grpList is not None):
        logD ("list of group %s" % grpList)

        prjList = []
        noPrj = 0
        
        for grp in grpList.grpList:
            if (config.isExistIn(CONFIG_FIELD_GROUPS, grp.name)):
                useGrp.append(grp)
                print ("found group %s" % grp.name)
                if (PARAM_INFO_ISS in reqs) or (PARAM_INFO_PRJ in reqs):
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

        # get exports methods
        exports = config.getExports()
        currentDT = datetime.datetime.now()
        exportFileName = "%s_%s_%s" % (config.getExportName(), \
                                      os.path.splitext(configFileName)[0], \
                                      currentDT.strftime("%Y%m%d_%H%M%S"))
        __issueList = []
        __noIssue = 0
        
        # get issue data
        if (PARAM_INFO_ISS in reqs):
            for __prj in prjList:
                __lst = getListIssuesInProject(config, __prj)
                if (__lst is not None):
                    print ("project %s has %d issue" % (__prj.name, len(__lst.issueList)))
                    __noIssue += len(__lst.issueList)
                    __issueList.extend(__lst.issueList)
        
        print ("number of issue %d, found %d" % (len(__issueList), __noIssue))

        #export to file
        __exports = config.getExports()
        if ("xlsx" in __exports) or ("xls" in __exports):
            path = getFullFilePath("%s.xls" % exportFileName)
            print("export to excel, path %s" % path)
            __workbook = None
            if (PARAM_INFO_GRP in reqs):
                __workbook = exportGroupToExcel(config, useGrp, path, __workbook)
            if (PARAM_INFO_PRJ in reqs):
                 __workbook = exportProjectToExcel(config, prjList, path, __workbook)
            if (PARAM_INFO_ISS in reqs):
                exportIssueToExcel(config, __issueList, path, __workbook)
            print("export done")
        else:
            print("ERROR: EXPORT METHOD %s NOT SUPPORTED YET. SUPPORT ONE IS %s" \
                    % (exports, EXPORTS_SUPPORT))
    return

####################################################################################
######################################## START RUNNING #############################
main()