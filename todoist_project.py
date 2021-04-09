# Import libraries
import todoist, re, csv
import datetime, sys
import locale, copy
from operator import itemgetter
from collections import defaultdict

from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT


taskslist = defaultdict(list)
debuglog = []
elements = []


# Try the variables passed as arguments or use the variables defined by default
try:
    client = sys.argv[1]
    startdate = sys.argv[2] + 'T00:00'
    enddate = sys.argv[3] + 'T23:59'
except:
    client = 'tivoli'
    startdate = "2021-02-26T00:00"
    enddate = "2021-03-25T23:59"


# Define Team
teamlogin = {name : token}


locale.setlocale(locale.LC_ALL, 'pt_PT')
date = datetime.datetime.now()
atualYear = date.strftime('%Y')
reportDate = date.strftime('%B %Y').upper()
startdateTable = f'{startdate[8:10]}/{startdate[5:7]}/{startdate[:4]}'
enddateTable = f'{enddate[8:10]}/{enddate[5:7]}/{enddate[:4]}'


# Validation dates
def validation_data(itemdata, itemname, closedate, warning, times):

    year = times.strftime('%Y')
    mes = times.strftime('%m')

    #se a data tiver tamanho 8 e o ano corresponder ao ano presente na data de fecho da tarefa
    if len(itemdata) == 8 and itemdata[:4] == closedate[:4]:
        data = itemdata[:4]

        #se o mes corresponder ao mes de fecho da tarefa
        if itemdata[4:6] == closedate[5:7]:
            data = data + '-' + itemdata[4:6]

            #se o dia for menor ou igual a 31
            if int(itemdata[-2:]) <= 31:
                itemdata = data + '-' + str(itemdata[-2:])
            else:
                itemdata = itemdata
                warning.append(itemname + ' - ' + 'Error in day')

        else:
            itemdata = data + '-' + itemdata[4:6] + '-' + itemdata[-2:]#res + '-' + closedate[5:7] + '-' + itemdata[-2:]
            warning.append(itemname + ' - ' + 'Error in month')

    # Validação de mudança de ano
    elif len(itemdata) == 8 and itemdata[:4] != closedate[:4]:
        if itemdata[4:6] == '01' and int(itemdata[-2:]) <= 31:
            itemdata = year + '-' +'01' + '-' + str(itemdata[-2:])

        elif itemdata[4:6] == '12' and int(itemdata[-2:]) in range(26,31):
            res = int(year) - 1
            itemdata = str(res) + '-' +'12' + '-' + str(itemdata[-2:])

    else:  #continuar mais tarde validações mais complexas
        itemdata = itemdata
        warning.append(itemname + ' - ' + 'Error in len data')

    return itemdata


# Process the tasks names, Defining what is the Name, the Date, the Time and the Category
def process_content(username, data, tasks_name, project, warnings):
    # Get Task Label
    taskarray = tasks_name.split("@")
    if len(taskarray) > 1:
        taskcategory = taskarray[1]
    else:
        taskcategory = 'no-category'
        warnings.append('Uncategorized tasks found')

    # Get Task Time Tracking
    taskarraynolabel = taskarray[0]
    tasktrackingarray = re.match(r"[^[]*\[([^]]*)\]", taskarraynolabel)

    if tasktrackingarray:
        taskdescriptiontime = tasktrackingarray.groups()[0]

        taskdescription = taskarraynolabel.replace('[' + taskdescriptiontime + ']','').strip()
        taskdescription = taskdescription

        tasktrackingarray = taskdescriptiontime.split('|')
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
                taskdate = validation_data(taskdate, taskdescription, data, debuglog, date)

                tasktime = tasktimedata[1]
                if len(tasktime) < 1:
                    tasktime = 'Error'
                    warnings.append(taskdescription + ' - No time')

            elif len(tasktimedata[0]) > 8:
                taskdate = tasktimedata[0][:8]
                taskdate = validation_data(taskdate, taskdescription, data, debuglog, date)

                tasktime = tasktimedata[0][8:]

            else:
                taskdate = tasktime = 'Error'
                warnings.append(taskdescription + ' - Error in tasktimedata.')

            taskslist[taskcategory].append([username, data, taskdate, project, taskdescription, tasktime, taskcategory])

    return taskslist


def get_tasks(link, finishdate, beginningdate, i, nome, thename):
    usertasks = link.completed.get_all(project_id=i, limit=200, offset=0, until=finishdate, since=beginningdate)
    tasklist = []

    for task in usertasks['items']:
        name_tasks = task['content']
        date = task['completed_date']
        tasklist.append(process_content(nome, date, name_tasks, thename, debuglog))

    return tasklist


def connect(teamconnect, projectname):

    # Connect for to Extract the Team Members and All Projects
    api = todoist.TodoistAPI(token)

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

                for identification in subproject_name:
                    if results['name'] == identification:
                        subprojectids[results['id']] = results['name']

            # Get Tasks of Subprojects or Project if the subprojectids is empty
            tasks = []
            if len(subprojectids) >= 1:
                for ids, identification in subprojectids.items():
                    tasks.append(get_tasks(connect, enddate, startdate, ids, name, identification))
            else:
                tasks.append(get_tasks(connect, enddate, startdate,  projectid, name, projectname))

    return tasks


# Sort Tasks in List by Task Date
def organize(tasksinformation):
    orderedtasks = []
    for category, task in tasksinformation.items():
        tasks_ordenadas = sorted(task, key=itemgetter(2))
        orderedtasks.append(tasks_ordenadas)

    return orderedtasks


# Export csv file with data of tasks and log file with errors that have occurred in data processing
def export_csv(tasksinformation, company, errors, report_date):
    ordered = organize(tasksinformation)

    # Add Titles to Generate Report List to Export
    tasksreportlist = [['user', 'closing date', 'task date', 'project name', 'task description', 'task time', 'category']]
    for task in ordered:
        tasksreportlist.extend(task)

    # Generate CSV File
    csvfile = '/Users/Lúcia/Desktop/report_' + company.replace(' ', '_') + '.csv'
    with open(csvfile, "w") as output:
        writer = csv.writer(output, lineterminator='\n')
        writer.writerows(tasksreportlist)

    print('Report for ' + company + ' finished.')


    # Generate debug.log file
    logfile = '/Users/Lúcia/Desktop/' + company + '/debug.log'
    if len(errors) > 0:
        with open(logfile, 'w') as debugs:
            writer_debugs = csv.writer(debugs, delimiter='-', lineterminator='\n')
            for error in errors:
                writer_debugs.writerow([report_date.strftime('%Y/%m/%d %H:%M'), error])
        print('File with errors created.')
    else:
        print('Not errors found.')

    return csvfile, logfile


def read_config(company):
    config = open('/Users/Lúcia/Desktop/' + company + '/config.txt', encoding='utf-8', mode='r')
    lines = config.readlines()
    config.close()

    # Define Contrat Type
    contrat = lines[0].split('-')
    contratType = contrat[1]

    # Select Information
    infoMensal = []
    infoAnual = []
    if 'mensal' in contratType:
        infoMensal.append(contratType)
        infoMensal.append(lines[1][:-1])#title
        infoMensal.append(lines[2][:-1])#title2
        infoMensal.append(lines[4:])#subtitles
        return infoMensal

    elif 'anual' in contratType:
        infoAnual.append(contratType)
        infoAnual.append(lines[1][:-1])#title
        order = []
        for configs in lines:
            spar = configs.split('-')
            order.append(spar)
        infoAnual.append(order[2][1][1:-1])#data de pacote adquirido
        infoAnual.append(order[2][2][1:-1])#data de expiração de pacote
        infoAnual.append(lines[3:])#categorias

        return infoAnual


def prepare_pdf(tasksinformation, company, report_date):
    ordered = organize(tasksinformation)

    # Dict with minutes for project
    min_subproject = defaultdict(list)
    for task in ordered:
        for i in task:
            if i[5] != 'Error': #mudar
                min_subproject[i[3]].append(int(i[5]))

    # Calculate the sum of times in minutes per project and convert to hours, saving to a new dictionary
    total = 0
    hour_subproject = {}
    for project, minutes in min_subproject.items():
        min_subproject[project] = sum(minutes)
        hour_subproject[project] = round(sum(minutes) / 60, 2)
        for min in minutes:
            total += min

    # Is the total time in hours spent per month
    projecthoursmonth = round(total/60, 2)
    projecthoursmonth = str(projecthoursmonth).replace('.', ',')

    # Update Summary Hours
    firstdate = report_date.strftime('%B %Y').capitalize()
    summary = '/Users/Lúcia/Desktop/' + company + '/summary.txt'
    with open(summary, 'a') as file:
        file.write(firstdate + ' - ' + projecthoursmonth + '-' + '\n')

    return projecthoursmonth, hour_subproject


def comments_cover(content):
    titlecomment = ParagraphStyle('paragraph', fontName='Helvetica-Bold', fontSize=12, spaceBefore=20, spaceAfter=10)
    content.append(Paragraph('Este relatório contém:', titlecomment))

    commentstyle = ParagraphStyle('paragraph', fontName='Helvetica',fontSize=10)
    content.append(Paragraph('1.    Resumo de Campanhas Facebook e Instagram Ads', commentstyle))
    content.append(Paragraph('2.    Resumo de Consumo de Horas Mensal Digital e Brand Design', commentstyle))
    content.append(Paragraph('3.    Relação de Projetos/Horas Digital e Brand Design', commentstyle))
    content.append(PageBreak())

    return content


# PDF Variables, Calcules and Export
def coverMensalVersion(content, company, firstdate, startdateT, enddateT):
    variables = prepare_pdf(taskslist, client, date)
    hour = variables[0]
    sub_hour = variables[1]

    firstlinestyle = ParagraphStyle('heading2', fontName='Helvetica', fontSize=12, textColor=colors.blue)
    content.append(Paragraph(f'RELATÓRIO {firstdate}', firstlinestyle))

    config = read_config(company)

    titlestyle = ParagraphStyle('heading1', fontName='Helvetica-Bold', fontSize=20, textColor=colors.black, leading=20)
    content.append(Paragraph('PROJECTOS', titlestyle))
    content.append(Paragraph(config[1], titlestyle))

    titlestyle2 = ParagraphStyle('heading1',fontName='Helvetica-Bold', fontSize=20, textColor=colors.black, leading=70)
    content.append(Paragraph(config[2], titlestyle2))

    # First table (capa)  https://www.youtube.com/watch?v=B3OCXBL4Hxs&t=309s
    data = [['Período a que diz respeito', f'{startdateT} a {enddateT}              '],
           [' ', ' '],
           [f'Horas despendidas no período de {startdateT} a {enddateT}', hour],
           ['Por rubrica:', ' ']]

    # Nomes e tempos dos projetos por rubrica acrescentados à tabela
    for pro in config[3]:
        times_capa = '0,00'
        enumerate = pro.split('-')
        for z, m in sub_hour.items():
            if z in enumerate[0]:
                enumerate.append(str(m))
                times_capa = enumerate[2].replace('.', ',')
        else:
            data.append([enumerate[1][:-1] + (' (horas)'), times_capa])


    table = Table(data)
    style = TableStyle([('BOX', (0,0), (-1,-1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black), #linhas dentro
                        ('ALIGN',(0,3),(0,-1),'RIGHT'), #alinhamento a esquerda na primeira coluna na 3 celula até à ultima
                        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),#tipo de letra
                        ('FONTSIZE', (0,0), (-1,-1), 10),#tamanho de letra
                        ('BACKGROUND', (0,1), (2,1), colors.cornflowerblue),#cor de fundo
                        ('FONTNAME', (1,2), (1,2), 'Helvetica-Bold')])

    # Alternate color BACKGROUND table
    rowNumb = len(data)
    for i in range(1, rowNumb):
        if i % 2 == 0:
            bc = colors.white
        else:
            bc = colors.cornflowerblue

        ts = TableStyle([('BACKGROUND', (0, i), (-1, i), bc)])
        table.setStyle(ts)

    table.setStyle(style)
    content.append(table)

    # Comments
    comments_cover(content)

    return content


def secondPageMensalVersion(content, company, year):
    titlesecond = ParagraphStyle('heading3', fontName='Helvetica-Bold', fontSize=14, textColor=colors.black,
                                 leading=70, alignment=TA_CENTER)
    content.append(Paragraph('Resumo de Consumo de Horas Mensal', titlesecond))

    # Read Summary File for Extrat Times
    config2 = open('/Users/Lúcia/Desktop/' + company + '/summary.txt', encoding='utf-8', mode='r')
    lines2 = config2.readlines()
    config2.close()

    # Tempo total somando as horas dos meses e a média
    n = 0
    timepormes = {}
    for line in lines2:
        x = line.split('-')
        count = x[1].replace(',', '.')
        timepormes[x[0]] = float(count)
        n += 1

    total = 0
    for mounth, time in timepormes.items():
        total += time

    media = round(total/n, 2)

    # Dados da segunda tabela
    data2 = [['  Horas Mensais', '                   Reais', '          Contratadas', '            Diferença'],
             [f'Janeiro {year} '],
             [f'Fevereiro {year} '],
             [f'Março {year} '],
             [f'Abril {year} '],
             [f'Maio {year} '],
             [f'Junho {year} '],
             [f'Julho {year} '],
             [f'Agosto {year} '],
             [f'Setembro {year} '],
             [f'Outubro {year} '],
             [f'Novembro {year} '],
             [f'Dezembro {year} '],
             ['TOTAL', str(total).replace('.', ','), str(total).replace('.', ','), '-'+str(total-total).replace('.',',')],
             ['Média Mensal', str(media).replace('.', ',')]]

    # Add data of table after calcules
    for row in data2:
        for i in row:
            for mounth, time in timepormes.items():
                if mounth == i:
                    row.append(str(time).replace('.', ','))
                    row.append(str(time).replace('.', ','))
                    row.append('0,00')

    table2 = Table(data2)
    style2 = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                         ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black), #linhas dentro
                         ('ALIGN',(0, 0), (-1, -1), 'RIGHT'), #alinhamento a esquerda
                         ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),#tipo de letra
                         ('FONTSIZE', (0, 0), (-1, -1), 10),#tamanho de letra
                         ('BACKGROUND', (0, 0), (0, -1), colors.gray),#cor de fundo
                         ('FONTNAME', (0, 0), (3, 0), 'Helvetica-Bold'),
                         ('FONTNAME', (0, 13), (0, -1), 'Helvetica-Bold'),
                         ('LINEBEFORE', (1, 0), (1, -1), 1, colors.black)])

    table2.setStyle(style2)
    content.append(table2)
    content.append(PageBreak())

    return content


def coverAnualVersion(tasksinformation, firstdate, content, company, startdateT, enddateT, year):
    ordered = organize(tasksinformation)

    min_category = defaultdict(list)
    for task in ordered:
        for i in task:
            if i[5] != 'Error':#mudar
                min_category[i[6]].append(i[5])

    firstlinestyle = ParagraphStyle('heading2', fontName='Helvetica', fontSize=12, textColor=colors.blue, leading=15)
    content.append(Paragraph(f'RELATÓRIO {firstdate}', firstlinestyle))

    # Read config.txt of client
    configs = read_config(company)
    categorys = configs[4]
    catTimes = []
    for category in categorys:
        cat = category.split('-')
        catTimes.append(cat)


    titlestyle = ParagraphStyle('heading1', fontName='Helvetica-Bold', fontSize=20, textColor=colors.black, leading=15)
    content.append(Paragraph('PROJECTOS ' + configs[1], titlestyle))

    pacotecomment = ParagraphStyle('paragraph', fontName='Helvetica', fontSize=12, spaceBefore=10, spaceAfter=30)
    content.append(Paragraph(f'ao abrigo do(s) pacote(s) de hora(s) adquirido(s) a partir de {configs[2]}', pacotecomment))


    # Read summary.txt of client
    config = open('/Users/Lúcia/Desktop/' + company + '/summary.txt', encoding='utf-8', mode='r')
    lines = config.readlines()
    config.close()

    timesbefore = []
    for i in lines:
        i = i.split('-')
        timesbefore.append(i[1][1:-1])

    day = int(enddateT[8:10]) + 1

    data = [['Período a que diz respeito', f'{startdateT} a {enddateT}              '],
            [catTimes[0][0][:-1]],
            [f'Total de horas no pacote adquirido ({configs[3]})', catTimes[0][1][1:-1] + ' horas'],
            [f'Total de horas a {startdateT}', timesbefore[0]],
            [f'Horas despendidas no período de {startdateT} a {enddateT}', 'calculo'], ###falta calculo por categoria e diferença de horas
            [f'Total de horas disponíveis a {day}/{enddateT[5:-1]}', 'total h - h despendidas'],
            [catTimes[1][0][:-1]],
            [f'Total de horas no pacote adquirido ({configs[3]})', catTimes[1][1][1:-1] + ' horas'],
            [f'Total de horas a {startdateT}', timesbefore[1]],
            [f'Horas despendidas no período de {startdateT} a {enddateT}', 'calculo'],###falta calculo por categoria e diferença de horas
            [f'Total de horas disponíveis a {day}/{enddateT[5:-1]}', 'total h - h despendidas']]


    table = Table(data)
    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black), #linhas dentro
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),#tipo de letra
                        ('FONTSIZE', (0, 0), (-1, -1), 10),#tamanho de letra
                        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 6), (0, 6), 'Helvetica-Bold'),
                        ('FONTNAME', (1, 3), (1, 5), 'Helvetica-Bold'),
                        ('FONTNAME', (1, 8), (1, -1), 'Helvetica-Bold')])

    # Alternate color BACKGROUND table
    rowNumb = len(data)
    for i in range(1, rowNumb):
        if i % 2 == 0:
            bc = colors.white
        else:
            bc = colors.cornflowerblue

        ts = TableStyle([('BACKGROUND', (0, i), (-1, i), bc)])
        table.setStyle(ts)

    table.setStyle(style)
    content.append(table)

    # Comments
    comments_cover(content)


    ####################Second Page################################
    titleStyle = ParagraphStyle('heading3', fontName='Helvetica-Bold', fontSize=14, textColor=colors.black,
                                leading=70, alignment=TA_CENTER)
    content.append(Paragraph('2. Resumo de Consumo de Horas Mensal', titleStyle))


    # Tempo em média e total somando as horas dos meses de summary.txt
    n = 0
    timepormes = {}
    for line in lines[2:]:
        x = line.split('-')
        count = x[1].replace(',', '.')
        count2 = x[2].replace(',', '.')
        timepormes[x[0]] = [float(count), float(count2)]
        n += 1

    total = 0
    total2 = 0
    for mounth, time in timepormes.items():
        total += time[0] #total do digital
        total2 += time[1] #total do brand design

    media = round(total/n, 2)#digital
    media2 = round(total2/n, 2)#brand design


    data2 = [['  Horas Mensais', '                   Reais', '          Contratadas', '            Diferença'],
            [f'Janeiro {year} '],
            [f'Fevereiro {year} '],
            [f'Março {year} '],
            [f'Abril {year} '],
            [f'Maio {year} '],
            [f'Junho {year} '],
            [f'Julho {year} '],
            [f'Agosto {year} '],
            [f'Setembro {year} '],
            [f'Outubro {year} '],
            [f'Novembro {year} '],
            [f'Dezembro {year} ']]


    horas_contratadas = round(float(catTimes[0][1][1:-1])/12, 2)
    diferencaTotal = round(float(catTimes[0][1])-total, 2)
    diferencaTotal = str(diferencaTotal).replace('.', ',')

    # ADD Information of Times in Table
    #https://pt.stackoverflow.com/questions/415967/array-atribu%c3%adda-em-outra-n%c3%a3o-mant%c3%a9m-o-mesmo-valor
    table_category1 = copy.deepcopy(data2)
    for row in table_category1:
        for i in row:
            for mounth, times in timepormes.items():
                if i == mounth:
                    row.append(str(times[0]).replace('.',','))
                    row.append(str(horas_contratadas).replace('.',','))
                    if times[0] > horas_contratadas:
                        diferenca = round(times[0] - horas_contratadas, 2)
                        diferenca = str(diferenca).replace('.',',')
                        row.append(f'+ {diferenca}')
                    else:
                        diferenca = round(horas_contratadas - times[0], 2)
                        diferenca = str(diferenca).replace('.',',')
                        row.append(diferenca)

    table_category1.append(['TOTAL', str(total).replace('.', ','),  catTimes[0][1][1:-1], diferencaTotal])
    table_category1.append(['Média Mensal', str(media).replace('.', ',')])

    width = 369
    titleTable = Table([[f'RESUMO DE CONSUMO DE HORAS {catTimes[0][0][:-1]}']], width)
    table = Table(table_category1)


    horas_contratadas2 = round(int(catTimes[1][1][1:-1])/12, 2)
    diferencaTotal2 = round(float(catTimes[1][1])-total, 2)
    diferencaTotal2 = str(diferencaTotal2).replace('.', ',')

    table_category2 = copy.deepcopy(data2)
    for row in table_category2:
        for i in row:
            for mounth, times in timepormes.items():
                if i == mounth:
                    row.append(str(times[1]).replace('.',','))
                    row.append(str(horas_contratadas2).replace('.',','))
                    if times[1] > horas_contratadas2:
                        diferenca = round(times[1] - horas_contratadas2, 2)
                        diferenca = str(diferenca).replace('.',',')
                        row.append(f'+ {diferenca}')
                    else:
                        diferenca = round(horas_contratadas2 - times[1], 2)
                        diferenca = str(diferenca).replace('.',',')
                        row.append(diferenca)


    table_category2.append(['TOTAL', str(total2).replace('.', ','),  catTimes[1][1][1:-1], diferencaTotal2])
    table_category2.append(['Média Mensal', str(media2).replace('.',',')])

    titleTable2 = Table([[f'RESUMO DE CONSUMO DE HORAS {catTimes[1][0][:-1]}']], width)
    table2 = Table(table_category2)

    #Style Tables creat
    titleStyle = TableStyle([('BACKGROUND', (0, 0), ( 0, 0), colors.blueviolet),
                             ('BOX', (0, 0), (0, 0), 0.15, colors.black),
                             ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold')])

    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black), #linhas dentro
                        ('ALIGN',(0, 0), (-1, -1), 'RIGHT'), #alinhamento a esquerda
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),#tipo de letra
                        ('FONTSIZE', (0, 0), (-1, -1), 10),#tamanho de letra
                        ('FONTNAME', (0, 0), (3, 0), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 13), (0, -1), 'Helvetica-Bold'),
                        ])

    titleTable.setStyle(titleStyle)
    content.append(titleTable)
    table.setStyle(style)
    content.append(table)

    espacoStyle = ParagraphStyle('space', spaceBefore=15)
    content.append(Paragraph(' ', espacoStyle))

    titleTable2.setStyle(titleStyle)
    content.append(titleTable2)
    table2.setStyle(style)
    content.append(table2)

    content.append(PageBreak())

    return content


def export_pdf(caracteristicas):
    type = read_config(client)[0]
    if 'mensal' in type:
        coverMensalVersion(caracteristicas, client, reportDate, startdateTable, enddateTable)
        secondPageMensalVersion(caracteristicas, client, atualYear)
    else:
        coverAnualVersion(taskslist, reportDate, caracteristicas, client, startdateTable, enddateTable, atualYear)

    pdf = SimpleDocTemplate(f'KOBU_{client}.pdf', pagesize=A4)
    pdf.build(caracteristicas)

    return pdf


if __name__ == '__main__':
    connect(teamlogin, client)
    #export_csv(taskslist, client, debuglog, date)
    #export_pdf(elements)

    if len(debuglog) >= 1:
        export_csv(taskslist, client, debuglog, date)
        exportNow = input('\nQuando o CSV estiver pronto para elaborar um PDF corretamente escreve yes:\n').upper()
        if exportNow == 'YES':
            import report_CSV

    else:
        export_csv(taskslist, client, debuglog, date)
        export_pdf(elements)
