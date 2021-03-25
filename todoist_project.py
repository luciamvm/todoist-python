# Import libraries
import todoist, re, csv, datetime, sys
from operator import itemgetter


taskslist = []
debuglog = []


# Try the variables passed as arguments or use the variables defined by default
try:
    client = sys.argv[1]
    startdate = sys.argv[2] + 'T00:00'
    enddate = sys.argv[3] + 'T23:59'
except:
    client = 'todoist-python'
    startdate = "2021-02-01T14:00"
    enddate = "2021-03-30T23:59"


# Define Team
teamlogin = {'Lúcia Moita': 'c85809e301b6d32847d6ea4a345d225a5b343673',
             'Nuno Tenazinha': '2557865df9c85bba8feb9bf086a1f25371bdac44',
             'Karolina Szmit': 'ffba60cd1d7377810f2f07502dd1c98b56c6271e',
             'Marta Gouveia': '8bac2c52ac124a9dbcf1dbe5baa3c118437b0ad8',
             'Gonçalo Cevadinha': 'fab0edf44f51a39ffa318fd21f3f76dd250dbee5'}


# Validation dates
def validation_data(itemdata, itemname, closedate, warning):
    #se a data tiver tamanho 8 e o ano corresponder ao ano presente na data de fecho da tarefa
    if len(itemdata) == 8 and itemdata[:4] == closedate[:4]:
        res = itemdata[:4]

        #se o mes corresponder ao mes de fecho da tarefa
        if itemdata[4:6] == closedate[5:7]:
            res = res + '-' + itemdata[4:6]

            #se o dia for menor ou igual a 31
            if int(itemdata[-2:]) <= 31:
                itemdata = res + '-' + str(itemdata[-2:])
            else:
                itemdata = itemdata
                warning.append(itemname + ' - ' + 'Error in day')

        else:
            itemdata = itemdata
            warning.append(itemname + ' - ' + 'Error in month')

    else:  #continuar mais tarde validações mais complexas
        itemdata = itemdata
        warning.append(itemname + ' - ' + 'Error in len data or year')

    return itemdata


# Process the tasks names, Defining what is the Name, the Date, the Time and the Category
def process_content(username, data, tasks_name, project, warnings):
    # Get Task Label
    taskarray = tasks_name.split("@")
    if len(taskarray) > 1:
        taskcategory = taskarray[1]
    else:
        taskcategory = 'no-category'

    # Get Task Time Tracking
    taskarraynolabel = taskarray[0]
    tasktrackingarray = re.match(r"[^[]*\[([^]]*)\]", taskarraynolabel)

    if tasktrackingarray:
        taskdescriptiontime = tasktrackingarray.groups()[0]

        taskdescription = taskarraynolabel.replace('[' + taskdescriptiontime + ']','').strip()
        taskdescription = taskdescription

        tasktrackingarray = taskdescriptiontime.split('|')
        #print(tasktrackingarray)
    else:
        taskdescription = taskarraynolabel.strip()
        tasktrackingarray = None

    #Create Array with Results
    if tasktrackingarray:
        for task in tasktrackingarray:
            if ':' in task:
                task = task.replace(':', ';')

            tasktimedata = task.split(';')
            if len(tasktimedata) == 2:
                taskdate = tasktimedata[0]
                tasktime = tasktimedata[1]

                taskdate = validation_data(taskdate, taskdescription, data, debuglog)

            elif len(tasktimedata[0]) > 8:
                taskdate = tasktimedata[0][:8]
                tasktime = tasktimedata[0][8:]

                taskdate = validation_data(taskdate, taskdescription, data, debuglog)

            else:
                taskdate = 'Error'
                tasktime = 'Error'
                warnings.append(taskdescription + ' - ' + 'Error in tasktimedata.')

            taskslist.append([username, data, taskdate, project, taskdescription, tasktime, taskcategory])

    return taskslist


def get_tasks(link, finishdate, beginningdate, i, nome, thename):

    usertasks = link.completed.get_all(project_id=i, limit=200, offset=0, until=finishdate, since=beginningdate)
    tasklist =[]

    for task in usertasks['items']:
        name_tasks = task['content']
        date = task['completed_date']
        tasklist.append(process_content(nome, date, name_tasks, thename, debuglog))

    return tasklist


def connect(teamconnect, projectname):

    # Connect for to Extract the Team Members and All Projects
    api = todoist.TodoistAPI('2557865df9c85bba8feb9bf086a1f25371bdac44')
    response = api.sync()

    team = {}
    collaboratorlist = api['collaborators']
    collaborators = api['collaborator_states']

    for collaborator in collaboratorlist:
        team[collaborator['id']] = collaborator['full_name']

    projectlist = api['projects']

    # Save Id's
    projectid = ''
    subprojectid = {}

    for project in projectlist:
        if project['name'] == projectname:
            projectid = project['id']

        if project['parent_id'] == projectid:
            subprojectid[project['id']] = project['name']


    # Save Names of Users that belong Project
    usersproject = []

    if len(subprojectid) > 0:
        for user in collaborators:
            if user['state'] == 'active' and user['project_id'] in subprojectid:
                usersproject.append(team[user['user_id']])

    else:
        for user in collaborators:
            if user['state'] == 'active' and user['project_id'] == projectid:
                usersproject.append(team[user['user_id']])


    # Connect Team and Get Tasks
    for name, token in teamconnect.items():
        if name in usersproject:

            connect = todoist.TodoistAPI(token)
            answer = connect.sync()

            #Throw error if token fails authentication
            if('error_code' in answer) :
                print("{0}: {1} login unsuccessul".format(answer['error'], name))
                continue


            # Get Id's for distint User (project_id is different for each user)
            projects = connect['projects']
            projectid = ''
            subprojectids = {}
            for results in projects:
                if results['name'] == projectname:
                    projectid = results['id']

                for id, identification in subprojectid.items():
                    if results['name'] == identification:
                        subprojectids[results['id']] = results['name']


            # Get Tasks of Subprojects or Project if the subprojectid is empty
            tasks = []
            if len(subprojectid) >= 1:
                for ids, identification in subprojectids.items():
                    name_project = identification
                    tasks.append(get_tasks(connect, enddate, startdate, ids, name, name_project))

            else:
                name_project = projectname
                tasks.append(get_tasks(connect, enddate, startdate,  projectid, name, name_project))

    return tasks


# Export csv file with data of tasks and log file with errors that have occurred in data processing
def export_csv(tasksinformation, company, errors):

    # Sort Tasks in List by Task Date
    tasksinformation = sorted(tasksinformation, key=itemgetter(2))

    # Add Titles to Generate Report List to Export
    tasksreportlist = [['user','closing date','task date','project name','task description','task time','task category']]
    tasksreportlist.extend(tasksinformation)

    # Generate CSV File
    csvfile = '/Users/Lúcia/Desktop/report_' + company.replace(' ', '_') + '.csv'

    with open(csvfile, "w") as output:
        writer = csv.writer(output, lineterminator='\n')
        writer.writerows(tasksreportlist)

    print('Report for ' + company + ' finished.')


    # Generate debug.log file
    logfile = '/Users/Lúcia/Desktop/debug.log'

    if len(errors) > 0:
        with open(logfile, 'w') as debugs:
            writer_debugs = csv.writer(debugs, delimiter='|', lineterminator='\n')
            for error in errors:
                writer_debugs.writerow([datetime.datetime.now(), error])

        print('File with errors created.')
    else:
        print('Not errors found.')

    return csvfile, logfile


if __name__ == '__main__':
    connect(teamlogin, client)
    export_csv(taskslist, client, debuglog)
