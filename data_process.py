##
# \file data_process.py
#Através da API do Todoist (https://developer.todoist.com/sync/v8), usada no ficheiro ‘data_process.py’ é possível fazer
# a conexão às contas de cada colaborador na aplicação e extrair as tarefas de determinado projeto, caso estas já estejam
# fechadas.
# O programa liga-se à conta do CEO da agência extrai uma lista de projetos, verifica se existem subprojectos e caso
# existam guarda o seu nome duma lista e extrai ainda um dicionário com o ID e Nome de cada colaborador.
# É criada também uma lista com o nome dos colaboradores que fazem parte do projeto que se pretende manipular.
# Esta lista é de seguida usada para fazer a conexão de cada colaborador através do seu token, caso este esteja
# presente no projeto. Durante a conexão é feita a extração de todas as tarefas fechadas pelo user num determinado
# período de tempo que é o utilizador do software que insere.
# As tarefas irão ser processadas fazendo a distinção da descrição da tarefa, a sua categoria, data e o tempo demorado
# a concluí-la. Esta informação é guardada numa lista que logo de seguida é organizada num dicionário agrupando tarefas
# por categoria e depois ordenando-as por data. Feito isto as tarefas são gravadas num documento .csv
##

# Import libraries
from todoist.api import TodoistAPI
import re, datetime, sys, os, locale, logging


# Try the passed variables as arguments or use the variables defined by default
try:
    client = sys.argv[1]
    startdate = sys.argv[2] + 'T00:00'
    enddate = sys.argv[3] + 'T23:59'
except:
    client = 'mar shopping algarve'
    startdate = "2021-04-26T00:00"
    enddate = "2021-05-25T23:59"


# Define Team
teamlogin = {'Nuno Tenazinha': '2557865df9c85bba8feb9bf086a1f25371bdac44',
            'Karolina Szmit': 'ffba60cd1d7377810f2f07502dd1c98b56c6271e',
            'Marta Gouveia': '8bac2c52ac124a9dbcf1dbe5baa3c118437b0ad8',
            'Gonçalo Cevadinha': 'fab0edf44f51a39ffa318fd21f3f76dd250dbee5',
            'Sandra Lopes': 'de2d4b29170fabea48f895c76b3efccfc1c598e9',
            'André Oliveira': 'bee37c06e968e18d01d923d944d9480362013f3f',
            'Beatriz Isabel': '83a658e3478a331474d7d5ab28329c0493b3eb5a',
            'Brígida Guerreiro': '010ac4b5482c687f26f22ee7c1baee57c8f001e0',
            'Cátia Dionísio': '4e41ffba321067b80d76f5dc6a049a1f0c0e35da',
            'Daniel Gomes': '7254bdedf93d56aee7f6dc561455f084fb783529',
            'Isabel Evaristo': 'b2fd4bfb15d526281d3d2370847a0fc6231467a5',
            'Liliana Guerreiro': 'd457564036a9993275c55b07642e05b81aac133c',
            'Miguel Spinola': '1df944000cd6de3bf620b00283aebd656f658e5b',
            'Mónica Loureiro': 'fc4d62c3e45e5b157764aa5a2a23e4465b8a992e',
            'Pedro Santos': 'a95efe65844d280d30f3c5d6a8fec3952f91b0bc',
            'Ramiro Mendes': '8edbddc12e37bb6847bbee38fb069a2cc0856380',
            'Sónia Duarte': '4ca59151ab8ca4805fdfcd4cd8e8c66e93553a3d',
            'Vanda Pereira': 'ea5a9ce4797a498f2107c7f9f1f970453802aad2'}


taskslist = []


# Create a new folder for documents each month
locale.setlocale(locale.LC_ALL, 'pt_PT')
date = datetime.datetime.now()
date_folder = date.strftime("%Y%m")
try:
    new_folder = os.mkdir(f'/Users/Lúcia/Desktop/{client}/{date_folder}')
except FileExistsError:
    pass


# debug.log file configuration
logging.basicConfig(filename=f'/Users/Lúcia/Desktop/{client}/{date_folder}/debug.log',
                    filemode='w',
                    format='%(asctime)s %(message)s',
                    datefmt='%Y/%m/%d %I:%M:%S',
                    level=logging.ERROR)



## 20210318 -> 2021-03-18
#
# Function for simple validations dates that user insert in your task description on todoist
#
# \param itemdata : string with dates inserted in description tasks
# \param itemname : description task used if exist errors
# \param closedate : string with complet date when user closed the task
# \param erro : array for save the errors if exist anything problem with date
# \param times : date of today
# \return string with correct date or incorrect date if it didn't correnpond of rules
def validation_data(itemdata, closedate, opendate):

    #se a data tiver tamanho 8 e o ano corresponder ao ano presente na data de fecho da tarefa
    if len(itemdata) == 8 and itemdata[:4] == closedate[:4]:
        data = itemdata[:4]
        if int(itemdata[4:6]) <= 12 and int(itemdata[-2:]) <= 31:
            itemdata = data + '-' + itemdata[4:6] + '-' + itemdata[-2:]

    elif len(itemdata) == 8 and itemdata[:4] == opendate[:4]:
        data = itemdata[:4]
        if int(itemdata[4:6]) <= 12 and int(itemdata[-2:]) <= 31:
            itemdata = data + '-' + itemdata[4:6] + '-' + itemdata[-2:]

    else:
        itemdata = itemdata

    return itemdata



##
# Process the tasks description, Defining what is the Name, the Date, the Time and the Category
#
# \param username : name of the user to which the task belongs
# \param data : date that user closed the task
# \param tasks_date : complet description of task extracted to todoist before divide category, dates and times
# \param project : name of the project to which task belongs
# \param erro : array for save errors
# \return array with an array with information about each task
def process_content(username, data, tasks_name, project):

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
        tasktrackingarray = taskdescriptiontime.split('|')
    else:
        taskdescription = taskarraynolabel.strip()
        tasktrackingarray = None

    tasklist = []
    #Create Array with Results
    if tasktrackingarray:
        for task in tasktrackingarray:
            if ':' in task:
                task = task.replace(':', ';')

            tasktimedata = task.split(';')
            if len(tasktimedata) == 2:
                taskdate = tasktimedata[0]
                taskdate = validation_data(taskdate, enddate, startdate)

                tasktime = tasktimedata[1].replace(' ', '')
                if '+' in tasktime:
                    times = tasktime.split('+')
                    for time in times:
                        tasktime = sum(int(time))
                else:
                    tasktime = tasktime

                if len(tasktime) < 1:
                    tasktime = 'Error'

            elif len(tasktimedata[0]) > 8:
                taskdate = tasktimedata[0][:8]
                taskdate = validation_data(taskdate, enddate, startdate)

                tasktime = tasktimedata[0][8:]
                if '+' in tasktime:
                    times = tasktime.split('+')
                    for time in times:
                        tasktime = sum(int(time))

            else:
                taskdate = tasktime = 'Error'


            tasklist.append(['          ', username, data, taskdate, project, taskdescription[0].capitalize() +
                             taskdescription[1:].replace(',', (' +')), tasktime, (taskcategory.lower()).replace(' ', '')])

    return tasklist


##
# Get completed tasks in specific project indentify with its id, between defined dates
#
# \param link : connection TodoistAPI with user tokens
# \param finishdate : date until which tasks will be extracted
# \param beginningdate : date since which tasks will be extracted
# \param i : project_id for indentify the project in which the tasks will be estracted
# \param nome : name which user have in todoist
# \param thename : project name where tasks will be extracted
# \return array with an array with information about each task
def get_tasks(link, finishdate, beginningdate, i, nome, thename, array_tasks):

    # Soma dois dias ao dia da finishdate
    finishday = finishdate[8:10]
    correct_finishday = int(finishday) + 2
    correct_finishday = str(correct_finishday)
    correct_finishdate = finishdate[0:8] + finishdate[8:10].replace(finishday, correct_finishday) + finishdate[10:]

    # Subtrai 3 dias à beginningdate
    startday = beginningdate[8:10]
    correct_startday = int(startday) - 3
    correct_startday = str(correct_startday)
    correct_beginningdate = beginningdate[0:8] + beginningdate[8:10].replace(startday, correct_startday) + beginningdate[10:]


    usertasks = link.completed.get_all(project_id=i, limit=200, offset=0, until=correct_finishdate, since=correct_beginningdate)
    for task in usertasks['items']:
        name_tasks = task['content']
        task_date = task['completed_date']
        tasks = process_content(nome, task_date, name_tasks, thename)

        # Verification if task is between period of ultil and since og the usertasks
        for infotask in tasks:
            taskdate = infotask[3]

            if taskdate[5:7] == finishdate[5:7] and int(taskdate[8:10]) <= 20:
                array_tasks.append(infotask)
            elif taskdate[5:7] == beginningdate[5:7] and int(taskdate[8:10]) >= 21:
                array_tasks.append(infotask)
            elif taskdate == 'Error':
                array_tasks.append(infotask)
            elif taskdate.find('-') != 4:
                array_tasks.append(infotask)

    return array_tasks


##
# Connection to todoist with Nuno token for extract names and id of users and an array of projects and child projects if exists
#
# \param teamconnect : dictionary with name and token each user
# \param projectname : name of project in todoist to which to do connection
# \return array with an array with information about each task
def connect(teamconnect, projectname, logs):

    # Connect for to Extract the Team Members and All Projects
    api = TodoistAPI('2557865df9c85bba8feb9bf086a1f25371bdac44')
    print('Connect to Nuno todoist for extract users and projects...')

    team = {}
    collaboratorlist = api['collaborators']
    collaborators = api['collaborator_states']

    for collaborator in collaboratorlist:
        team[collaborator['id']] = collaborator['full_name']


    projectlist = api['projects']

    # Save Id's for Next Validations
    projectid = ''
    subproject_name = []

    for project in projectlist:
        if project['name'] == projectname:
            projectid = project['id']

        if project['parent_id'] == projectid:
            subproject_name.append(project['name'])

    # Save Names of Users that belong Project
    usersproject = []
    for user in collaborators:
        if user['state'] == 'active' and user['project_id'] == projectid:
            usersproject.append(team[user['user_id']])



    # Connect Team and Get Tasks
    print('Getting the tasks for project for each user . . .')
    tasks = []
    for name, token in teamconnect.items():
        if name in usersproject:
            connect = TodoistAPI(token)
            answer = connect.sync()

            # Throw error if token fails authentication
            if 'error_code' in answer:
                logs.error("{0}: {1} login unsuccessul".format(answer['error'], name))
                continue

            # Get Id's for distint User (project_id is different for each user)
            projects = connect['projects']
            projectid = ''
            subprojectids = {}
            for results in projects:
                if results['name'] == projectname:
                    projectid = results['id']

                for identification in subproject_name:
                    if results['name'] == identification:
                        subprojectids[results['id']] = results['name']


            # Get Tasks of Subprojects or Project if the subprojectids is empty
            if len(subprojectids) >= 1:
                for ids, identification in subprojectids.items():
                    tasks.append(get_tasks(connect, enddate, startdate, ids, name, identification, taskslist))
            else:
                tasks.append(get_tasks(connect, enddate, startdate, projectid, name, projectname, taskslist))


    print('Tarefas extraidas.')

    return tasks
