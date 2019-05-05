# gitlab_issue_export
Get issue list via gitlab API

usage: python main.py c=[configFile] l=<iss,prj,grp>
c: config file to be used.
l: info to be retrieved, seperated by comma (i.e. l=iss, prj)

configFile: config file to be used. use "config.ini" if not specified
# api: api version of gitlab to be used, i.e. 3
# url: URL of gitlab, this is host domain
# token: private token, so that we can retrieve info, if need. This should be secret
#        if not specified, program will ask for the input from commandline
#        WARNING: consider carefully when input token in config file!!!!
#        TOKEN IS SECRET ONE, CONSIDER CAREFULLY, DO WITH CARE. I DONT CARE YOU
# groups: list of groups to be get (group name), separate by comma. If not specified, all groups will be get
# projects: list of projects to be get (name), separate by comma. If not specified, all projects will be get
# authors: list of authors to be get ( name), separate by comma. If not specified, all authors will be get
# labels: list of labels to be get (name), separate by comma. If not specified, all labels will be get
# exports: exports format, i.e xls, xlsx. separate by comma. If not specified, xls will be used
# exportname: export name (base name). final one will be combination of exporname, config name, date time
# maxgroup: maximum group, not specify, infinity
# maxproject: maximum project in group, not specify, infinity
# maxissue: maximum issue in project, not specify, infinity



