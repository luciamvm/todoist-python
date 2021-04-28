# Import libraries
import todoist, re, csv
import datetime, sys
import locale
import json
from operator import itemgetter
from collections import defaultdict

import os
import reportlab
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, Table, TableStyle, PageBreak
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.styles import getSampleStyleSheet



taskslist = []
debuglog = []
elements = []


# Try the variables passed as arguments or use the variables defined by default
try:
    client = sys.argv[1]
    startdate = sys.argv[2] + 'T00:00'
    enddate = sys.argv[3] + 'T23:59'
except:
    client = 'tivoli'
    startdate = "2021-03-22T00:00"
    enddate = "2021-04-22T23:59"


# Define Team
teamlogin = {'Lúcia Moita': 'c85809e301b6d32847d6ea4a345d225a5b343673',
             'Nuno Tenazinha': '2557865df9c85bba8feb9bf086a1f25371bdac44',
             'Karolina Szmit': 'ffba60cd1d7377810f2f07502dd1c98b56c6271e',
             'Marta Gouveia': '8bac2c52ac124a9dbcf1dbe5baa3c118437b0ad8',
             'Gonçalo Cevadinha': 'fab0edf44f51a39ffa318fd21f3f76dd250dbee5'}


# Dates for use in PDF
locale.setlocale(locale.LC_ALL, 'pt_PT')
date = datetime.datetime.now()
atualYear = date.strftime('%Y')
reportDate = date.strftime('%B %Y').upper()
startdateTable = f'{startdate[8:10]}/{startdate[5:7]}/{startdate[:4]}'
enddateTable = f'{enddate[8:10]}/{enddate[5:7]}/{enddate[:4]}'


doc = BaseDocTemplate(f'KOBU_{client}.pdf', pagesize=A4)


# Intall KOBU FONTS for reports
folder = os.path.dirname(reportlab.__file__) + os.sep + 'fonts'

headline = os.path.join(folder, 'KobuHeadlineBdCn.ttf')
bold = os.path.join(folder, 'KOBUTextSansSerif-Bold.ttf')
regular = os.path.join(folder, 'KOBUTextSansSerif-Regular.ttf')
regular_italic = os.path.join(folder, 'KOBUTextSansSerif-RegularItalic.ttf')
extralight = os.path.join(folder, 'KOBUTextSansSerif-ExtraLight.ttf')

pdfmetrics.registerFont(TTFont("KOBU-Headline", headline))
pdfmetrics.registerFont(TTFont("KOBU-Bold", bold))
pdfmetrics.registerFont(TTFont("KOBU-Regular", regular))
pdfmetrics.registerFont(TTFont("KOBU-RegularItalic", regular_italic))
pdfmetrics.registerFont(TTFont("KOBU-Extralight", extralight))



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

            taskslist.append([username, data, taskdate, project, taskdescription.capitalize(), tasktime, taskcategory])

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
    api = todoist.TodoistAPI('2557865df9c85bba8feb9bf086a1f25371bdac44')

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

    # Tasks agrupadas pro categoria
    by_category = defaultdict(list)
    for task in tasksinformation:
        by_category[task[6]].append([task[0], task[1], task[2], task[3], task[4], task[5], task[6]])

    # Tasks ordenadas por data
    orderedtasks = []
    for category, task in by_category.items():
        ordered = sorted(task, key=itemgetter(0))
        orderedtasks.append(ordered)

    return orderedtasks



# Export csv file with data of tasks and log file with errors that have occurred in data processing
def export_csv(company, errors, report_date):

    ordered = organize(taskslist)

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



# Read Config File
def read_config(company):
    with open(f'/Users/Lúcia/Desktop/{company}/config.json', encoding='utf-8', mode='r') as configfile:
        data = json.load(configfile)

    return data



# Read Summary File
def read_summary(company):
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='r') as summary:
        data = json.load(summary)
        summary.close()

    return data



def alternate_color(dados, tabela):
    # Alternate color BACKGROUND table cover
    rowNumb = len(dados)
    for i in range(1, rowNumb):
        if i % 2 == 0:
            bc = colors.white
        else:
            bc = colors.cornflowerblue

        ts = TableStyle([('BACKGROUND', (0, i), (-1, i), bc)])
        tabela.setStyle(ts)

    return tabela



def comments_cover(content):
    titlecomment = ParagraphStyle('paragraph',
                                  fontName='KOBU-Bold',
                                  fontSize=12,
                                  spaceBefore=40,
                                  spaceAfter=10)

    content.append(Paragraph('Este relatório contém:', titlecomment))


    commentstyle = ParagraphStyle('paragraph',
                                  fontName='KOBU-Regular',
                                  fontSize=10)

    content.append(Paragraph('1.    Resumo de Campanhas Facebook e Instagram Ads', commentstyle))
    content.append(Paragraph('2.    Resumo de Consumo de Horas Mensal Digital e Brand Design', commentstyle))
    content.append(Paragraph('3.    Relação de Projetos/Horas Digital e Brand Design', commentstyle))
    content.append(PageBreak())

    return content


#########################Version 2 == 1 contrato############################
def calcules_version2(company, report_date):

    tasks_ordered = organize(taskslist)

    # Dict with minutes for project
    min_subproject = defaultdict(list)
    for task in tasks_ordered:
        for i in task:
            if i[5] != 'Error': #mudar
                min_subproject[i[3]].append(int(i[5]))

    summary = read_summary(client)

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

    for projects, info in summary['projects'].items():
        info[report_date.capitalize()] = projecthoursmonth

    summary = {'projects': {projects: info}}

    # Update Summary Hours
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='w') as summaryfile:
        json.dump(summary, summaryfile, indent=2)
        summaryfile.close()

    return projecthoursmonth, hour_subproject, min_subproject



# PDF Variables, Calcules and Export
def cover_version2(content, firstdate, startdateT, enddateT):

    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='KOBU-Extralight',
                                    fontSize=14,
                                    textColor=colors.blue,
                                    spaceAfter=5)

    content.append(Paragraph(f'RELATÓRIO {firstdate}', firstlinestyle))


    titlestyle = ParagraphStyle('cover_title',
                                fontName='KOBU-Headline',
                                fontSize=30,
                                textColor=colors.black,
                                leading=30)


    config = read_config(client)
    title = config['report_title'].upper()

    content.append(Paragraph(title.replace('\n', '<br/>'), titlestyle))
    content.append(Paragraph(' ', ParagraphStyle('space', spaceBefore=50)))

    variables = calcules_version2(client, reportDate)
    hour = variables[0]
    sub_hour = variables[1]

    # First table (capa)  https://www.youtube.com/watch?v=B3OCXBL4Hxs&t=309s
    data = [['Período a que diz respeito', f'{startdateT} a {enddateT}              '],
           [' ', ' '],
           [f'Horas despendidas no período de {startdateT} a {enddateT}', hour],
           ['Por rubrica:', ' ']]


    # ADD Projects and Hours Spended of Data
    for projects, info in config['projects'].items():
        times = '0,00'
        for project, hours in sub_hour.items():
            if project in projects:
                times = str(hours).replace('.', ',')
        else:
            data.append([info['name'] + ' (horas)', times])


    table = Table(data, colWidths=[280, 150])

    style = TableStyle([('BOX', (0,0), (-1,-1), 0.15, colors.black),
                        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                        ('ALIGN',(0,3),(0,-1),'RIGHT'),
                        ('FONTNAME', (0,0), (-1,-1), 'KOBU-Regular'),
                        ('FONTSIZE', (0,0), (-1,-1), 10),
                        ('BACKGROUND', (0,1), (2,1), colors.cornflowerblue),
                        ('FONTNAME', (1,2), (1,2), 'KOBU-Bold')])

    # Alternate color BACKGROUND table
    alternate_color(data, table)

    table.setStyle(style)
    content.append(table)

    # Comments
    comments_cover(content)

    return content



def summary_version2(content, year):

    #Estilo e titulo da página
    titlesecond = ParagraphStyle('heading3',
                                 fontName='KOBU-Headline',
                                 fontSize=14,
                                 textColor=colors.black,
                                 leading=70,
                                 alignment=TA_CENTER)

    content.append(Paragraph('Resumo de Consumo de Horas Mensal', titlesecond))


    # Calculos a partir da leitura do summary
    summary = read_summary(client)

    for projects, summarys in summary.items():
        n = 0
        for project, times in summarys.items():
            n += len(times)
            total = round(sum(times.values()), 2)

    media = round((total/n), 2)

    # Dados da segunda tabela
    data = [['Horas Mensais', 'Reais', 'Contratadas', 'Diferença'],
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
    for row in data:
        for mes in row:
            for mounth, time in times.items():
                if mounth in mes:
                    row.append(str(time).replace('.', ','))
                    row.append(str(time).replace('.', ','))
                    row.append('0,00')


    table = Table(data, colWidths=[90, 90, 90, 90])

    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                         ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black), #linhas dentro
                         ('ALIGN',(0, 0), (-1, -1), 'RIGHT'), #alinhamento a esquerda
                         ('FONTNAME', (0, 0), (-1, -1), 'KOBU-Regular'),#tipo de letra
                         ('FONTSIZE', (0, 0), (-1, -1), 10),#tamanho de letra
                         ('BACKGROUND', (0, 0), (0, -1), colors.gray),#cor de fundo
                         ('FONTNAME', (0, 0), (3, 0), 'KOBU-Bold'),
                         ('FONTNAME', (0, 13), (0, -1), 'KOBU-Bold'),
                         ('LINEBEFORE', (1, 0), (1, -1), 1, colors.black)])

    table.setStyle(style)
    content.append(table)
    content.append(PageBreak())

    return content



def body_version2(content, pdf):
    titleStyle = ParagraphStyle('heading',
                                fontName='KOBU-Headline',
                                fontSize=14,
                                textColor=colors.black,
                                leading=50,
                                alignment=TA_CENTER)

    content.append(Paragraph('Relação de Projetos/Horas', titleStyle))

    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]
    styleN.fontName = "KOBU-Regular"

    tasks = organize(taskslist)
    config = read_config(client)

    times = calcules_version2(client, reportDate)
    hours_total = times[0]
    times_hours = times[1]
    times_min = times[2]

    # Creat a table to project
    extra_style = defaultdict(list)
    tabelas = {}
    for project, info in config['projects'].items():
        id = 0
        data = {project: [[info['name']],
                          ['ID', 'Nome']]}

        # Add categories to tables
        for projecto, informacao in data.items():
            for category in info['categories']:
                id += 1
                informacao.append([id, category])

                #Add tasks info to tables
                exist_tasks = False
                for task_array in tasks:
                    for task in task_array:
                        description = Paragraph(task[4], styleN)
                        if task[6] == category and task[3] == projecto:
                            exist_tasks = True
                            informacao.append([task[2], description, task[5], '(min)'])
                if exist_tasks == False:
                    informacao.append(['-', '-', '', '(min)'])

            #Add total times to table
            minutes = False
            for area, min in times_min.items():
                if area == project:
                    minutes = True
                    informacao.append(['', 'Total', min, '(min)'])
            if minutes == False:
                informacao.append(['', 'Total', '0', '(min)'])
                informacao.append(['', '', '0,00', '(h)'])

            for area, hour in times_hours.items():
                if area == project:
                    informacao.append(['', '', hour, '(h)'])

            informacao.append([''])
            tabelas[projecto] = informacao

        for category in info['categories']:
            for row, values in enumerate(informacao):
                for indice, info in enumerate(values):
                    if category == info:
                        extra_style[projecto].append(('FONTNAME', (indice-1, row), (indice, row), 'KOBU-Bold'))


    #Estilo comum a todas as tabelas
    style = ([('BOX', (0, 0), (-1, -2), 0.15, colors.black),
              ('INNERGRID', (0, 2), (-1, -2), 0.15, colors.black),
              ('SCAN', (0,0), (0,0)),
              ('FONTNAME', (1,-1), (1, -1), 'KOBU-Bold'),
              ('FONTNAME', (0, 0), (1, 2), 'KOBU-Bold'),
              ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
              ('FONTSIZE', (0, 0), (-1, -1), 10),
              ('BACKGROUND', (0, 0), (3, 0), colors.dodgerblue),
              ('BACKGROUND', (0, 1), (-1, 1), colors.deepskyblue),
              ('BACKGROUND', (1,-3), (1, -3), colors.deepskyblue),
              ('BACKGROUND', (2,-3),(2,-2), colors.deepskyblue),
              ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
              ('VALIGN', (0, 0), (-1, -1), 'TOP'),
              ('ALIGN', (1,-3), (1, -3), 'RIGHT'),
              ('FONTNAME', (1,-3), (1, -3), 'KOBU-Bold'),
              ('FONTNAME', (2,-2), (2, -2), 'KOBU-Bold')])

    #https://www.javaer101.com/pt/article/25871033.html
    #caso não seja possivel ficarem duas tabelas completas numa página, a proxima tabela passa para a proxima página
    available_height = pdf.height
    for project, tabela in tabelas.items():
        table = Table(tabela, repeatRows=2, colWidths=[70, 345, 50, 45])
        # Add style to tables
        for area, extra in extra_style.items():
            for estilo in extra:
                if area == project:
                    style.append(estilo)

        table.setStyle(TableStyle(style))


        table_height = table.wrap(0, available_height)[1]
        if available_height < table_height:
            content.extend([PageBreak(), table])
            if table_height < pdf.height:
                available_height = pdf.height - table_height
            else:
                available_height = table_height % pdf.height
        else:
            content.append(table)
            available_height = available_height - table_height


    total_data = [['','TOTAL MÊS', hours_total, '(h)']]
    total_table = Table(total_data, colWidths=[70, 345, 50, 45])

    total_style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                              ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black),
                              ('BACKGROUND', (1,0), (1,0), colors.gray),
                              ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
                              ('FONTNAME', (0,0), (-2,-1), 'KOBU-Bold'),
                              ('TEXTCOLOR', (1,0), (1,0), colors.white)])

    total_table.setStyle(total_style)
    content.append(total_table)

    return content



####################Version 1 >= 2 contratos#########################################
def calcules_version1(company, reportdate):

    ordered = organize(taskslist)
    config = read_config(client)

    categories = {}
    #Categorias de cada projecto
    for project, info in config['projects'].items():
        categories[project] = info['categories']

    min_category = defaultdict(list)
    for task in ordered:
        for i in task:
            if i[6] != 'no-category' and i[5] != 'Error': #mudar mais tarde
                min_category[i[6]].append(i[5])


    times_category = {}
    # projecto com os tempos gastos este mês em minutos
    for category, min in min_category.items():
        for project, categoria in categories.items():
            if category in categoria:
                times_category[project] = min
            else:
                times_category[project] = [0]

    transform_int = defaultdict(list)
    # tempos transformados em números inteiros
    for category, times in times_category.items():
        for time in times:
            transform_int[category].append(int(time))

    total_time_hours = {}
    total_time_min = {}
    #tempo total gasto por categoria
    for category, time in transform_int.items():
        total_time_min[category] = sum(time)
        for project, min in total_time_min.items():
            total_time_hours[project] = round(min / 60, 2)

    # Read summary for update
    summary = read_summary(client)

    # Horas disponiveis até ao report do mês passado
    avaliable_hours = {}
    for project, category in summary['projects'].items():
        avaliable_hours[project] = category['avaliable']

    # Horas disponiveis até ao report de agora
    avaliable_hours_update = {}
    for category, time in avaliable_hours.items():
        for project, hours in total_time_hours.items():
            if project == category:
                avaliable_hours_update[project] = round(time-hours, 2)

    for project, info in summary['projects'].items():
        for category, hour in avaliable_hours_update.items():
            if category == project:
                info['avaliable'] = hour

        for cat, time in total_time_hours.items():
            if project == cat:
                info['hours_month'][reportdate.capitalize()] = time

    # Update summary
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='w') as summaryfile:
        json.dump(summary, summaryfile, indent=2)
        summaryfile.close()

    return avaliable_hours, total_time_hours, total_time_min



def cover_version1(content, reportdate, startdateT, enddateT):

    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='KOBU-Extralight',
                                    fontSize=14,
                                    textColor=colors.blue,
                                    leading=15)

    content.append(Paragraph(f'RELATÓRIO {reportdate}', firstlinestyle))

    # Informações da leitura do config
    configs = read_config(client)

    titlestyle = ParagraphStyle('headline',
                                fontName='KOBU-Headline',
                                fontSize=30,
                                textColor=colors.black,
                                leading=23)

    content.append(Paragraph(f'PROJECTOS {configs["report_title"].upper()}', titlestyle))

    pacoteStyle = ParagraphStyle('paragraph',
                                 fontName='KOBU-Regular',
                                 fontSize=12.5,
                                 spaceBefore=10,
                                 spaceAfter=30)

    content.append(Paragraph(f'ao abrigo do(s) pacote(s) de hora(s) adquirido(s) a partir de '
                             f'{configs["contract_start_date"]}', pacoteStyle))


    times = calcules_version1(client, reportDate)
    times_last_month = times[0]
    time_month = times[1]


    day = int(enddateT[:2]) + 1
    contract_start_date = f'(até {configs["contract_end_date"]})'

    #Table for cover
    tables =[['Período a que diz respeito', f'{startdateT} a {enddateT}']]
    for category, info in configs['projects'].items():
        for area, time in times_last_month.items():
            for tema, hours in time_month.items():

                if category == area and category == tema:
                    tables.append([category.upper()])
                    tables.append([f'Total de horas no pacote adquirido {contract_start_date}', f'{info["total"]} horas'])
                    tables.append([f'Total de horas a {startdateT}', time])
                    tables.append([f'Horas despendidas no período de {startdateT} a {enddateT}', hours])
                    tables.append([f'Total de horas disponíveis a {day}/{enddateT[3:]}', round(time-hours, 2)])


    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                        ('FONTNAME', (0, 0), (-1, -1), 'KOBU-Regular'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 1), (0, 1), 'KOBU-Bold'),
                        ('FONTNAME', (0, 6), (0, 6), 'KOBU-Bold'),
                        ('FONTNAME', (1, 3), (1, 5), 'KOBU-Bold'),
                        ('FONTNAME', (1, 8), (1, -1), 'KOBU-Bold'),
                        ('BACKGROUND', (2, -1), (3, -1), colors.aquamarine)])


    table = Table(tables, colWidths=[290, 150])
    alternate_color(tables, table)
    table.setStyle(style)
    content.append(table)

    # Comments
    comments_cover(content)

    return content


def summary_version1(content, year):
    titleStyle = ParagraphStyle('heading3',
                                fontName='KOBU-Headline',
                                fontSize=18,
                                textColor=colors.black,
                                spaceAfter=40,
                                alignment=TA_CENTER)

    content.append(Paragraph('Resumo de Consumo de Horas Mensal', titleStyle))


    configs = read_config(client)
    summary = read_summary(client)

    total = {}
    n = 0
    hours_month = {}
    for projects, summarys in summary.items():
        for project, info in summarys.items():
            n += len(info)
            x = info['hours_month']
            hours_month[project] = info['hours_month']
            total[project] = round(sum(x.values()), 2)

    media = {}
    for category, time in total.items():
        media[category] = round((time/n), 2)

    tables = []
    time_avaliable_month = {}
    for project, info in configs['projects'].items():
        contract_time = info['total']
        time_avaliable_month[project] = int(contract_time/12)
        for category, time in total.items():
            for area, media_time in media.items():
                if project == category:
                    data = [[f'RESUMO DE CONSUMO DE HORAS {project.upper()}'],
                            ['Horas Mensais', 'Reais', 'Contratadas', 'Diferença'],
                            [f'Janeiro {year}'],
                            [f'Fevereiro {year}'],
                            [f'Março {year}'],
                            [f'Abril {year}'],
                            [f'Maio {year}'],
                            [f'Junho {year}'],
                            [f'Julho {year}'],
                            [f'Agosto {year}'],
                            [f'Setembro {year}'],
                            [f'Outubro {year}'],
                            [f'Novembro {year}'],
                            [f'Dezembro {year}'],
                            ['TOTAL', time, contract_time, round(contract_time-time, 2)],
                            ['Média Mensal', media_time]]

        for row in data:
            for value in row:
                for topic, avaliable_month in time_avaliable_month.items():
                    for tema, times in hours_month.items():
                        for month, time in times.items():

                            if project == tema and month == value and project == topic:
                                row.append(time)
                                row.append(avaliable_month)
                                row.append(avaliable_month-time)

        data.append([])
        data.append([])
        tables.append(data)


    style = TableStyle([('BOX', (0, 0), (-1, -3), 0.15, colors.black),
                        ('INNERGRID', (0, 1), (-1, -3), 0.15, colors.black),
                        ('SCAN', (0,0), (0,0)),
                        ('ALIGN',(0, 1), (-1, -1), 'RIGHT'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 2), (-1, -5), 'KOBU-Regular'),
                        ('FONTNAME', (0, 0), (0, 0), 'KOBU-Headline'),
                        ('FONTNAME', (0, 1), (-1, 1), 'KOBU-Bold'),

                        ('FONTNAME', (0, -4), (0, -1), 'KOBU-Bold'),
                        ('BACKGROUND', (0,0), (-1,0), colors.coral)
                        ])

    for table in tables:
        table = Table(table, colWidths=[90, 90, 90, 90])
        table.setStyle(style)
        content.append(table)

    return content


def body_version1(content):
    # Title
    titleStyle = ParagraphStyle('heading',
                                fontName='KOBU-Headline',
                                fontSize=18,
                                textColor=colors.black,
                                spaceAfter=40,
                                alignment=TA_CENTER)

    content.append(Paragraph('Relação de Projetos/Horas', titleStyle))

    # Config info
    config = read_config(client)


    # Style for the text stay in her column (\n)
    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]
    styleN.fontName = "KOBU-Regular"

    # Create tables
    id = 0
    data = {}
    possible_categories = {}
    for project, info in config['projects'].items():
        id += 1
        possible_categories[project] = info['categories']

        data[project] = [[config['report_title']],
                         ['ID', 'Nome'],
                         [id, info["tasks_table_title"]]]

    # Tasks
    tasks = organize(taskslist)

    # Times
    times = calcules_version1(client, reportDate)
    times_hours = times[1]
    times_min = times[2]


    for area, table in data.items():
        # Add tasks information to tables
        for project, categories in possible_categories.items():
            for category in categories:
                if project == area:

                    for list_tasks in tasks:
                        for task in list_tasks:
                            description = Paragraph(task[4], styleN)
                            if task[6] == category:
                                table.append([task[2], description, task[5], '(min)'])

        # Add times for category
        for tema, time in times_min.items():
            if area == tema:
                table.append(['', '', time, '(min)'])

        for tema, time in times_hours.items():
            if area == tema:
                table.append(['', '', time, '(h)'])
                table.append([])
                table.append(['', 'TOTAL MÊS', time, '(h)'])


    # Style to tables
    tableStyle = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                             ('INNERGRID', (0, 2), (-1, -1), 0.15, colors.black),
                             ('SCAN', (0,0), (0,0)),
                             ('FONTSIZE', (0, 0), (-1, -1), 10),
                             ('FONTNAME', (0,0), (0,0), 'KOBU-Headline'),
                             ('FONTNAME', (0, 1), (1, 2), 'KOBU-Bold'),
                             ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
                             ('TEXTCOLOR', (1,-1), (1, -1), colors.white),
                             ('FONTNAME', (1,-1), (1, -1), 'KOBU-Bold'),
                             ('ALIGN',(0, 3), (0, -1), 'RIGHT'),
                             ('ALIGN',(2, 0), (-1, -1), 'LEFT'),
                             ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                             ('ALIGN', (1,-1), (1, -1), 'RIGHT'),
                             ('VALIGN', (0, 0), (-1, -1), 'TOP'),

                             ('BACKGROUND', (0, 0), (3, 0), colors.dodgerblue),
                             ('BACKGROUND', (0, 1), (-1, 1), colors.deepskyblue),
                             ('BACKGROUND', (1,-1), (1, -1), colors.gray),
                             ('BACKGROUND', (2,-4),(2,-3), colors.deepskyblue),
                             ('BACKGROUND', (-2,-1),(-2,-1), colors.deepskyblue)])

    # Add tables to report
    for project, table in data.items():
        tabela = Table(table, repeatRows=3, colWidths=[70, 345, 50, 45])
        tabela.setStyle(tableStyle)
        content.append(tabela)
        content.append(PageBreak())


    return content


#https://stackoverflow.com/questions/8827871/a-multilineparagraph-footer-and-header-in-reportlab
def footer(canvas, pdf):

    style = ParagraphStyle('paragraph',
                           fontName='KOBU-Regular',
                           fontSize=8,
                           alignment=TA_CENTER)

    empresa = Paragraph('KOBU Agência Criativa Digital, Lda', style)
    morada = Paragraph('Rua do Pé da Cruz, nº 24, 3º Esq e Dir, 8000-404 Faro', style)
    site = '<link href=https://kobu.agency/><u> kobu.agency </u></link>'
    mail = '<a href="mailto:hello@kobu.pt"><u> hello@kobu.pt </u></a>'
    add_footer = Paragraph(site + ' | ' + mail, style)

    w, h = empresa.wrap(pdf.width, pdf.bottomMargin)
    empresa.drawOn(canvas, pdf.leftMargin, h+28)
    w, h = morada.wrap(pdf.width, pdf.bottomMargin)
    morada.drawOn(canvas, pdf.leftMargin, h+15)
    w, h = add_footer.wrap(pdf.width, pdf.bottomMargin)
    add_footer.drawOn(canvas, pdf.leftMargin, h+2)



def export_pdf(caracteristicas, pdf):
    type = read_config(client)['number_contracts']
    if type == 1:
        cover_version2(caracteristicas, reportDate, startdateTable, enddateTable)
        summary_version2(caracteristicas, atualYear)
        body_version2(caracteristicas, doc)
    else:
        cover_version1(caracteristicas, reportDate, startdateTable, enddateTable)
        summary_version1(caracteristicas, atualYear)
        body_version1(caracteristicas)

    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='footer', frames=frame, onPage=footer)
    pdf.addPageTemplates([template])

    pdf.build(caracteristicas)
    print('PDF EXPORTED!')

    return pdf


if __name__ == '__main__':
    connect(teamlogin, client)
    export_csv(client, debuglog, date)
    export_pdf(elements, doc)

    # if len(debuglog) >= 1:
    #     import report_CSV
    #     export_csv(client, debuglog, date)
    #     exportNow = input('\nQuando o CSV estiver pronto para elaborar um PDF corretamente escreve yes:\n').upper()
    #     if exportNow == 'YES':
    #         report_CSV.ler_csv(client, taskslist)
    #         report_CSV.export_pdf(elements, doc)
    # else:
    #     export_csv(client, debuglog, date)
    #     export_pdf(elements, doc)
