# gitlab_issue_export
Get issue list via gitlab API

usage: python main.py [configFile]

configFile: config file to be used. use "config.ini" if not specified
* api: ==> gitlab API to be used
* url: ==> URL of gitlab web
* token: ==> private token
* groups: ==> list of groups, separate by comma
* projects: ==> list of project, seperate by comma
* authors: ==> list of author, separate by comma
* labels: ==> list of label, seperate by comma
* exports:xlsx ==> list of export file type, seperate by comma (i.e xlsx)
