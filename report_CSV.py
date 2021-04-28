# Import libraries
import todoist_project
import csv, json
from collections import defaultdict
from operator import itemgetter

from reportlab.platypus import Paragraph, PageBreak, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle

from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate



client = todoist_project.client
startdate = todoist_project.startdateTable
enddate = todoist_project.enddateTable

reportDate = todoist_project.reportDate
atualYear = todoist_project.atualYear

taskslist = []
elementsPDF = []

doc = BaseDocTemplate(f'KOBU_{client}_csv.pdf', pagesize=A4)


# Read CSV and Save info tasks in taskslist
def read_csv(company, tasksinformation):
    csv_file = '/Users/Lúcia/Desktop/report_' + company.replace(' ', '_')+'_csv' + '.csv'
    with open(csv_file) as output:
        csv_reader = csv.reader(output)
        csv_reader.__next__()

        for row in csv_reader:
            tasksinformation.append(row[2:])

    return tasksinformation



def organize_tasks(taskinformation):

    # Dicionario com tasks agrupadas por categoria
    by_category = defaultdict(list)
    for info in taskinformation:
        by_category[info[4]].append([info[0], info[1], info[2], info[3], info[4]])


    # Tasks ordenadas pro data
    ordered_tasks = []
    for category, task in by_category.items():
        ordered = sorted(task, key=itemgetter(0))
        ordered_tasks.append(ordered)

    return ordered_tasks



######################Version 2 == 1 contrato ##############################
def calcules_version2(company, report_date):
    ordered = organize_tasks(taskslist)

    # Dict with minutes for project
    min_subproject = defaultdict(list)
    for task in ordered:
        for i in task:
            min_subproject[i[1]].append(int(i[3]))

    summary = todoist_project.read_summary(client)

    # Calculate the sum of times in minutes per project and convert to hours, saving to a new dictionary
    total = 0
    hour_subproject = {}
    for project, minutes in min_subproject.items():
        min_subproject[project] = sum(minutes)
        hour_subproject[project] = round(sum(minutes) / 60, 2)
        for min in minutes:
            total += min

    # Is the total time in hours spended
    projecthoursmonth = round(total/60, 2)
    for projects, info in summary['projects'].items():
        info[report_date.capitalize()] = projecthoursmonth

    summary = {'projects': {projects: info}}

    # Update Summary Hours
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='w') as summaryfile:
        json.dump(summary, summaryfile, indent=2)
        summaryfile.close()

    return projecthoursmonth, hour_subproject, min_subproject



def cover_version2(content, firstdate, startdateT, enddateT):

    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='KOBU-Extralight',
                                    fontSize=12,
                                    textColor=colors.blue)

    content.append(Paragraph(f'RELATÓRIO {firstdate}', firstlinestyle))

    titlestyle = ParagraphStyle('heading1',
                                fontName='KOBU-Headline',
                                fontSize=20,
                                textColor=colors.black,
                                leading=34)

    content.append(Paragraph('PROJECTOS', titlestyle))

    config = todoist_project.read_config(client)
    title = config['report_title'].upper()

    content.append(Paragraph(title.replace('\n','<br/>'), titlestyle))#segunda linha do relatório
    content.append(Paragraph(' ', ParagraphStyle('space', spaceBefore=70)))

    # Total times in min and hour
    variables = calcules_version2(client, reportDate)
    hour = variables[0]
    sub_hour = variables[1]

    # First table (capa)  https://www.youtube.com/watch?v=B3OCXBL4Hxs&t=309s
    data = [['Período a que diz respeito', f'{startdateT} a {enddateT}              '],
            [' ', ' '],
            [f'Horas despendidas no período de {startdateT} a {enddateT}', hour],
            ['Por rubrica:', ' ']]


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
    todoist_project.alternate_color(data, table)

    table.setStyle(style)
    content.append(table)

    # Comments
    todoist_project.comments_cover(content)

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

    tasks = organize_tasks(taskslist)
    config = todoist_project.read_config(client)

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
                        description = Paragraph(task[2], styleN)
                        if task[4] == category and task[1] == projecto:
                            exist_tasks = True
                            informacao.append([task[0], description, task[3], '(min)'])
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
            informacao.append([''])
            tabelas[projecto] = informacao

        for category in info['categories']:
            for row, values in enumerate(informacao):
                for indice, info in enumerate(values):
                    if category == info:
                        extra_style[projecto].append(('FONTNAME', (indice-1, row), (indice, row), 'KOBU-Bold'))


    #Estilo comum a todas as tabelas
    style = ([('BOX', (0, 0), (-1, -3), 0.15, colors.black),
              ('INNERGRID', (0, 2), (-1, -3), 0.15, colors.black),
              ('SCAN', (0,0), (0,0)),
              ('FONTNAME', (1,-1), (1, -1), 'KOBU-Bold'),
              ('FONTNAME', (0, 0), (1, 2), 'KOBU-Bold'),
              ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
              ('FONTSIZE', (0, 0), (-1, -1), 10),
              ('BACKGROUND', (0, 0), (3, 0), colors.dodgerblue),
              ('BACKGROUND', (0, 1), (-1, 1), colors.deepskyblue),
              ('BACKGROUND', (1,-4), (1, -4), colors.deepskyblue),
              ('BACKGROUND', (2,-4),(2,-3), colors.deepskyblue),
              ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
              ('VALIGN', (0, 0), (-1, -1), 'TOP'),
              ('ALIGN', (1,-4), (1, -4), 'RIGHT'),
              ('FONTNAME', (1,-4), (1, -4), 'KOBU-Bold'),
              ('FONTNAME', (2,-3), (2, -3), 'KOBU-Bold')])


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



####################Version 1 = <1contrato#############################
def calcules_version1(company, reportdate):

    ordered = organize_tasks(taskslist)
    config = todoist_project.read_config(client)

    categories = {}
    #Categorias de cada projecto
    for project, info in config['projects'].items():
        categories[project] = info['categories']

    # Todas as categorias encontradas
    min_category = defaultdict(list)
    for task in ordered:
        for i in task:
            min_category[i[4]].append(i[3])


    # Times divididos pelas duas categorias principais
    times_category = defaultdict(list)
    for category, min in min_category.items():
        for project, categoria in categories.items():
           if category in categoria:
               times_category[project].append(min)


    transform_int = defaultdict(list)
    # Tempos transformados em números inteiros
    for category, list_times in times_category.items():
        for times in list_times:
            for time in times:
                transform_int[category].append(int(time))

    total_time_min = {}
    total_time_hours = {}
    # Tempo total gasto por categoria
    for category, time in transform_int.items():
        total_time_min[category] = sum(time)
        for project, hours in total_time_min.items():
            total_time_hours[project] = round(hours / 60, 2)


    # Read summary for update
    summary = todoist_project.read_summary(client)

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



def report_cover_version1(content, reportdate, startdateT, enddateT):

    titleStyle = ParagraphStyle('heading2',
                                fontName='KOBU-Extralight',
                                fontSize=14,
                                textColor=colors.blue,
                                leading=15)

    content.append(Paragraph(f'RELATÓRIO {reportdate}', titleStyle))

    # Função para ler os config.txt
    configs = todoist_project.read_config(client)

    subtitleStyle = ParagraphStyle('headline',
                                   fontName='KOBU-Headline',
                                   fontSize=30,
                                   textColor=colors.black,
                                   leading=23)

    content.append(Paragraph(f'PROJECTOS {configs["report_title"].upper()}', subtitleStyle))

    infoStyle = ParagraphStyle('paragraph',
                               fontName='KOBU-Regular',
                               fontSize=12.5,
                               spaceBefore=10,
                               spaceAfter=30)

    content.append(Paragraph(f'ao abrigo do(s) pacote(s) de hora(s) adquirido(s) a partir de'
                             f'{configs["contract_start_date"]}', infoStyle))


    times = calcules_version1(client, reportDate)
    times_last_month = times[0]
    time_month = times[1]

    day = int(enddate[:2]) + 1
    contract_start_date = f'(até {configs["contract_end_date"]})'

    #Table for cover
    tables =[['Período a que diz respeito', f'{startdateT} a {enddateT}']]
    for category, info in configs['projects'].items():
        for area, time in times_last_month.items():
            for tema, hours in time_month.items():

                if category == tema and category == area:
                    tables.append([category.upper()])
                    tables.append([f'Total de horas no pacote adquirido {contract_start_date}', f'{info["total"]} horas'])
                    tables.append([f'Total de horas a {startdate}', time])
                    tables.append([f'Horas despendidas no período de {startdate} a {enddate}', hours])
                    tables.append([f'Total de horas disponíveis a {day}/{enddate[3:]}', round(time-hours, 2)])


    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                        ('FONTNAME', (0, 0), (-1, -1), 'KOBU-Regular'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 1), (0, 1), 'KOBU-Bold'),
                        ('FONTNAME', (0, 6), (0, 6), 'KOBU-Bold'),
                        ('FONTNAME', (1, 3), (1, 5), 'KOBU-Bold'),
                        ('FONTNAME', (1, 8), (1, -1), 'KOBU-Bold'),
                        ('BACKGROUND', (2, -1), (3, -1), colors.aquamarine)])


    table = Table(tables, colWidths=[280, 150])
    todoist_project.alternate_color(tables, table)
    table.setStyle(style)
    content.append(table)

    # Comments
    todoist_project.comments_cover(content)

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
    config = todoist_project.read_config(client)


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
    tasks = organize_tasks(taskslist)

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
                            description = Paragraph(task[2], styleN)
                            if task[4] == category:
                                table.append([task[0], description, task[3], '(min)'])

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
                             ('FONTNAME', (0,0), (0,0), 'KOBU-Headline'),
                             ('FONTNAME', (0, 1), (1, 2), 'KOBU-Bold'),
                             ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
                             ('FONTSIZE', (0, 0), (-1, -1), 10),
                             ('FONTNAME', (1,-1), (1, -1), 'KOBU-Bold'),
                             ('TEXTCOLOR', (1,-1), (1, -1), colors.white),
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



def export_pdf(caracteristicas, pdf):
    type = todoist_project.read_config(client)["number_contracts"]
    if type == 1:
        cover_version2(caracteristicas, reportDate, startdate, enddate)
        todoist_project.summary_version2(caracteristicas, atualYear)
        body_version2(caracteristicas, pdf)
    else:
        report_cover_version1(caracteristicas, reportDate, startdate, enddate)
        todoist_project.summary_version1(caracteristicas, atualYear)
        body_version1(caracteristicas)

    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='footer', frames=frame, onPage=todoist_project.footer)
    pdf.addPageTemplates([template])

    pdf.build(caracteristicas)
    print('PDF EXPORTED!')

    return pdf


if __name__ == '__main__':
    read_csv(client, taskslist)
    export_pdf(elementsPDF, doc)
