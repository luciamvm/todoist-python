##
# \file export_pdf.py
# Para a criação dos relatórios foi usada a bliblioteca ‘reportlab’ do python
# (https://www.reportlab.com/docs/reportlab-userguide.pdf).
# De acordo com o tipo de acordo estipulado pelo cliente com a agência é criado um relatório com uma capa que contém os
# dados do contrato, as horas gastas e as horas ainda disponíveis.\n
# A segunda página é composta por um sumário do tempo despendido pela agência ao longo dos meses, incluído os dados média
# e total.\n
# O corpo do relatório é composto por tabelas que para a versão 1, existe uma tabela destinada a cada contrato que
# contêm a data da tarefa, a sua descrição e o tempo despendido em minutos. Para a versão 2 existe uma tabela para
# o projeto, ou uma para cada subprojecto, que está dividida por categorias nas quais são encaixadas as tarefas, com a
# sua data, descrição e tempo em minutos.
##


# Import libraries
import json
import os
import reportlab
from collections import defaultdict

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, Table, TableStyle, PageBreak
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

import matplotlib.pyplot as plt
from io import BytesIO
from svglib.svglib import svg2rlg

# Import variables from data_process.py
from data_process import client, startdate, enddate, date, date_folder, logging



# Dates for use in reports
date = date
atualYear = date.strftime('%Y')
reportDate = date.strftime('%B %Y')
startdateTable = f'{startdate[8:10]}/{startdate[5:7]}/{startdate[:4]}'
enddateTable = f'{enddate[8:10]}/{enddate[5:7]}/{enddate[:4]}'


# Elements to pdf and file
elements = []
doc = BaseDocTemplate(f'/Users/Lúcia/Desktop/{client}/{date_folder}/KOBU_{client.replace(" ", "_")}.pdf', pagesize=A4)


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


logo = 'C:/Users/Lúcia/PycharmProjects/estagio/KOBU-logo2.jpg'



##
# Read file with configurations about client
#
# \param company : name of client to which will be extracted report that is the folder name
# \return dicionary with information extracted of file
def read_config(company):
    with open(f'/Users/Lúcia/Desktop/{company}/config.json', encoding='utf-8', mode='r') as configfile:
        data = json.load(configfile)

    return data


##
# Read summary file with hours spent each month
#
# \param company : name of client to which will be extracted report that is the folder name
# \return dictionary with information extracted of file
def read_summary(company):
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='r') as summary:
        data = json.load(summary)
        summary.close()

    return data


##
# Alternate color BACKGROUND table cover
#
# \param dados : array with an array for each line of table
# \param tabela : to convert array for table
# \return the table with new alternate color style
def alternate_color(dados, tabela):
    rowNumb = len(dados)
    for i in range(1, rowNumb):
        if i % 2 == 0:
            bc = colors.white
        else:
            bc = colors.HexColor('#E6E6E6')

        ts = TableStyle([('BACKGROUND', (0, i), (-1, i), bc)])
        tabela.setStyle(ts)

    return tabela


##
# Area for write comments on cover report
#
# \param content : array with the elements that to compose the report
# \return array updated with comments information
def comments_cover(content):
    titlecomment = ParagraphStyle('paragraph',
                                  fontName='KOBU-Bold',
                                  fontSize=12,
                                  textColor=colors.HexColor('#3A3C4B'),
                                  spaceBefore=40,
                                  spaceAfter=10)
    content.append(Paragraph('Este relatório contém:', titlecomment))

    commentstyle = ParagraphStyle('paragraph',
                                  fontName='KOBU-Regular',
                                  fontSize=10,
                                  textColor=colors.HexColor('#3A3C4B'))
    content.append(Paragraph('1.\tResumo de Campanhas Facebook e Instagram Ads', commentstyle))
    content.append(Paragraph('2.\tResumo de Consumo de Horas Mensal Digital e Brand Design', commentstyle))
    content.append(Paragraph('3.\tRelação de Projetos/Horas Digital e Brand Design', commentstyle))


    return content


##
# Define footer - https://stackoverflow.com/questions/8827871/a-multilineparagraph-footer-and-header-in-reportlab
#
# \param canvas : property of reportlab for position of footer
# \param pdf : define the methods for manipulate the data in the report
def footer(canvas, pdf):

    style = ParagraphStyle('paragraph',
                           fontName='KOBU-Regular',
                           fontSize=8,
                           textColor=colors.HexColor('#3A3C4B'),
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

    # Add logo
    logo = Image('C:/Users/Lúcia/PycharmProjects/estagio/KOBU-logo.jpg', 4*cm, 0.571*cm)
    w, h = logo.wrap(pdf.width, pdf.topMargin)
    logo.drawOn(canvas, pdf.leftMargin, pdf.height + pdf.topMargin - h + 45)




## Version 2 is used when client have just 1 contract with
#
# Calculate spend times in all project and for each sub project and update summary file with hours spent this month
#
# \param company : name of client to which will be extracted report that is the folder name
# \param report_date : present month and year
# \param info_tasks : dictionary with tasks organized
# \return total time spent this month in hours
# \return dicitionary with name child projects and the total time spent in hours
# \return dicitionary with name child projects and the total time spent in minutes
#
def calcules_version2(company, report_date, info_tasks, logs):
    # Organized tasks
    tasks_ordered = info_tasks

    # Dict with minutes for project
    min_subproject = defaultdict(list)
    for category, tasks in tasks_ordered.items():
        for task in tasks:
            min_subproject[task[4]].append(task[6])

    # Transform times in int
    transform_int = defaultdict(list)
    for category, list_times in min_subproject.items():
        for times in list_times:
            transform_int[category].append(int(times))

    # Calculate the sum of times in minutes per project and convert to hours, saving to a new dictionary
    total = 0
    min_subproject = {}
    hour_subproject = {}
    for project, minutes in transform_int.items():
        min_subproject[project] = sum(minutes)
        hour_subproject[project] = round(sum(minutes) / 60, 2)
        for min in minutes:
            total += min

    # The total time in hours spent
    projecthoursmonth = round(total/60, 2)
    summary = read_summary(client)

    for projects, info in summary['projects'].items():
        info[report_date.capitalize()] = projecthoursmonth

    summary = {'projects': {projects: info}}

    # Update Summary Hours
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='w') as summaryfile:
        json.dump(summary, summaryfile, indent=2)
        summaryfile.close()

    return projecthoursmonth, hour_subproject, min_subproject



##
# Defines the composition of the report cover
#
# \param content : array with the elements that to compose the report
# \param firstdate : present month and year
# \param startdatT : date inserted at the beginning of the project for define the point of start extract tasks
# \param enddateT : date inserted at the beginning of the project for define the point of end extract tasks
# \param info_tasks : dictionary with tasks organized
# \return array updated with elements that compose the cover reports
#
def cover_version2(content, firstdate, startdateT, enddateT, info_tasks):

    print('Starting to doing the report . . .')

    # Titles cover
    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='KOBU-Extralight',
                                    fontSize=14,
                                    textColor=colors.HexColor('#E53E44'),
                                    spaceAfter=5)

    content.append(Paragraph(f'Relatório {firstdate}', firstlinestyle))


    titlestyle = ParagraphStyle('cover_title',
                                fontName='KOBU-Headline',
                                fontSize=30,
                                textColor=colors.HexColor('#3A3C4B'),
                                leading=30)


    config = read_config(client)
    title = config['report_title'].upper()

    # Case title has \n
    content.append(Paragraph(title.replace('\n', '<br/>'), titlestyle))
    content.append(Paragraph(' ', ParagraphStyle('space', spaceBefore=50)))

    variables = calcules_version2(client, reportDate, info_tasks, logging)
    hour = variables[0]
    sub_hour = variables[1]

    # First table  https://www.youtube.com/watch?v=B3OCXBL4Hxs&t=309s
    data = [['Período a que diz respeito', f'{startdateT} a {enddateT}              '],
            [' ', ' '],
            [f'Horas despendidas no período de {startdateT} a {enddateT}', str(hour).replace('.', ',')],
            ['Por rubrica:', ' ']]


    # Add Projects and Hours Spent of data
    for projects, info in config['projects'].items():
        times = '0,00'
        for project, hours in sub_hour.items():
            if project in projects:
                times = str(hours).replace('.', ',')
        else:
            data.append([info['name'] + ' (horas)', times])


    table = Table(data, colWidths=[280, 150])

    style = TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor('#D2D2D7')),
                        ('ALIGN',(0,3),(0,-1),'RIGHT'),
                        ('FONTNAME', (0,0), (-1,-1), 'KOBU-Regular'),
                        ('FONTSIZE', (0,0), (-1,-1), 10),
                        ('FONTNAME', (1,2), (1,2), 'KOBU-Bold'),
                        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#3A3C4B'))])

    # Alternate color BACKGROUND table
    alternate_color(data, table)

    table.setStyle(style)
    content.append(table)


    # Comments
    comments_cover(content)
    content.append(PageBreak())

    print("Report cover created")

    return content



def plot_version2(content, company, foldername):
    print("Graph in prodution . . .")

    fig = plt.figure(figsize=(5.6, 3))
    plt.style.use('bmh')
    plt.grid(alpha=0)

    summary = read_summary(client)
    x1 = []
    y1 = []
    for projects, info in summary.items():
        for project, resume in info.items():
            for month, hours in resume.items():
                x1.append(month[0:3])
                y1.append(hours)


    plt.plot(x1, y1, color='#E53E44', label=client)
    plt.ylim(ymin=0)

    plt.xlabel('Mês', fontname='KOBU-Bold', size=14)
    plt.ylabel('Horas', fontname='KOBU-Bold', size=14)
    plt.tight_layout()
    plt.legend(frameon=True, fontsize=8, framealpha=0.5, facecolor='#D2D2D7', bbox_to_anchor=(0.6, 1.05))

    fig.savefig(f'C:/Users/Lúcia/Desktop/{company}/{foldername}/graph.jpg')

    img = BytesIO()
    fig.savefig(img, format='svg')
    img.seek(0)

    drawing = svg2rlg(img)
    drawing.shift(-60, -50)

    content.append(drawing)

    return content

##
# Defines the composition of the summary page of report
#
# \param content : array with the elements that to compose the report
# \param year : present year
# \return array updated with elements that compose the second page of report
def summary_version2(content, year):

    # Style and title of page
    titlesecond = ParagraphStyle('heading3',
                                 fontName='KOBU-Headline',
                                 fontSize=14,
                                 textColor=colors.HexColor('#3A3C4B'),
                                 leading=70,
                                 alignment=TA_CENTER)

    content.append(Paragraph('Resumo de Consumo de Horas Mensal'.upper(), titlesecond))


    # Calculate the total and average hours spent
    summary = read_summary(client)
    for projects, summarys in summary.items():
        n = 0
        for project, times in summarys.items():

            n += len(times)
            total = round(sum(times.values()), 2)

    media = round((total/n), 2)
    diferenca = round(total-total, 2)

    # Data of summary table
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
            ['TOTAL', str(total).replace('.', ','), str(total).replace('.', ','), '-'+str(diferenca).replace('.',',')],
            ['Média Mensal', str(media).replace('.', ',')]]

    # Add data of table after calcules
    for row in data:
        for mes in row:

            for mounth, time in times.items():
                if mounth in mes:
                    row.append(str(time).replace('.', ','))
                    row.append(str(time).replace('.', ','))
                    row.append('0,00')

        if len(row) == 1:
            row.append('')
            row.append('')
            row.append('0,00')


    table = Table(data, colWidths=[90, 90, 90, 90])

    # Style table
    style = TableStyle([
                        ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.HexColor('#D2D2D7')),
                        ('ALIGN',(0, 0), (-1, -1), 'RIGHT'),
                        ('FONTNAME', (0, 0), (-1, -1), 'KOBU-Regular'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#E6E6E6')),
                        ('FONTNAME', (0, 0), (3, 0), 'KOBU-Bold'),
                        ('FONTNAME', (0, 13), (0, -1), 'KOBU-Bold'),
                        ('LINEBEFORE', (1, 0), (1, -1), 1, colors.HexColor('#3A3C4B')),
                        ('TEXTCOLOR', (0,0),(-1,-1), colors.HexColor('#3A3C4B'))
                    ])

    # Add table to report
    table.setStyle(style)
    content.append(table)

    plot_version2(content, client, date_folder)

    content.append(PageBreak())

    print('Hour consumption summary created')

    return content


##
# Defines the composition of the body to report, one table for child project with date task, description task and time in minutes
#
# \param content : array with the elements that to compose the report
# \param info_tasks : dictionary with tasks organized
# \return array updated with elements that compose the body report
def body_version2(content, info_tasks):
    # Body title
    titleStyle = ParagraphStyle('heading',
                                fontName='KOBU-Headline',
                                fontSize=14,
                                textColor=colors.HexColor('#3A3C4B'),
                                leading=50,
                                alignment=TA_CENTER)

    content.append(Paragraph('Relação de Projetos/Horas'.upper(), titleStyle))

    # Extra style for description task
    styles = getSampleStyleSheet()
    style_tasks_description = styles["Normal"]
    style_tasks_description.fontName = "KOBU-Regular"
    style_tasks_description.textColor = colors.HexColor('#3A3C4B')

    # Extra style to bold id
    style_id_bold = styles["BodyText"]
    style_id_bold.fontName = "KOBU-Bold"
    style_id_bold.textColor = colors.HexColor('#3A3C4B')

    tasks = info_tasks
    config = read_config(client)

    times = calcules_version2(client, reportDate, info_tasks, logging)
    times_hours = times[1]
    times_min = times[2]


    print('Body tables are in process . . .')

    # Creat a table for each project
    tabelas = {}
    for project, info in config['projects'].items():
        id = 0
        data = {project: [[info['name'].upper()],
                          ['ID', 'Nome']]}

        # Add categories to tables
        for projecto, informacao in data.items():
            for title, category in info['categories'].items():
                id += 1
                id_table = Paragraph(str(id), style_id_bold)
                category_table = Paragraph(title, style_id_bold)
                informacao.append([id_table, category_table])

                #Add tasks info to tables
                exist_tasks = False
                for key, list_tasks in tasks.items():
                    for task in list_tasks:
                        description = Paragraph(task[5].replace('""', '"'), style_tasks_description)
                        if task[7] in category and task[4] == projecto:
                            exist_tasks = True
                            informacao.append([task[3], description, task[6], '(min)'])

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
                    informacao.append(['', '', str(hour).replace('.', ','), '(h)'])

            informacao.append([''])
            tabelas[projecto] = informacao


    # Tables Style
    style = ([
              ('INNERGRID', (0, 2), (-1, -2), 0.15, colors.HexColor('#D2D2D7')),
              ('SCAN', (0,0), (0,0)),
              ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
              ('FONTNAME', (0, 1), (1, 1), 'KOBU-Bold'),
              ('FONTSIZE', (0, 0), (-1, -1), 10),
              ('TEXTCOLOR', (0,0),(-1,0), colors.HexColor('#ffffff')),
              ('TEXTCOLOR', (0,1),(-1,-1), colors.HexColor('#3A3C4B')),
              ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D2D2D7')), #TABLE TITLES
              ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E6E6E6')),
              ('BACKGROUND', (1,-3), (1, -3), colors.HexColor('#E6E6E6')),
              ('BACKGROUND', (2,-3),(2,-2), colors.HexColor('#E6E6E6')),
              ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
              ('VALIGN', (0, 0), (-1, -1), 'TOP'),
              ('ALIGN', (1,-3), (1, -3), 'RIGHT'),
              ('FONTNAME', (1,-3), (1, -3), 'KOBU-Bold'),
              ('FONTNAME', (2,-2), (2, -2), 'KOBU-Bold')
              ])

    # To assign an id to table
    project_id = 0
    for project, tabela in tabelas.items():
        project_id += 1
        table = Table(tabela, repeatRows=2, colWidths=[70, 345, 50, 45])
        table.setStyle(TableStyle(style))
        content.append(table)

        # When arrives to last table does't change to page
        if project_id != len(tabelas):
            content.append(PageBreak())


    return content


## Vension 1 is for clients with 2 contracts or more
#
# Calculate spend times for contract and update summary file with hours spent this month
#
# \param reportdate : atual month and year
# \param info_tasks : dictionary with tasks organized
# \param company : name of client to which will be extracted report that is the folder name
# \return dictionary with total avaliable hours for spend this month
# \return dicitionary with total time spent in hours this month for contract
# \return dicitionary with total time spent in minutes this month for contract
# \return dictionary with total avaliable hours at the moment
#
def calcules_version1(reportdate, info_tasks, logs, company):
    # Organized tasks
    ordered = info_tasks
    config = read_config(client)

    # Categories of each project
    categories_arrays = defaultdict(list)
    for project, info in config['projects'].items():
        categories_arrays[project].append(info['categories'].values())

    categories = defaultdict(list)
    for area, lists in categories_arrays.items():
        for array in lists:
            for i in array:
                for category in i:
                    categories[area].append(category)


    # Category with yours times
    min_category = defaultdict(list)
    for category, list_tasks in ordered.items():
        for task in list_tasks:
            min_category[category].append(task[6])


    # Times for project in minutes
    times_category = defaultdict(list)
    for category, min in min_category.items():
        for project, categoria in categories.items():
            if category in categoria:
                times_category[project].append(min)
            else:
                times_category[project].append(['0'])

    # To transform times for int
    transform_int = defaultdict(list)
    for category, list_times in times_category.items():
        for times in list_times:
            for time in times:
                transform_int[category].append(int(time))

    # Total time in minutes and hours for project
    total_time_hours = {}
    total_time_min = {}
    for category, time in transform_int.items():
        total_time_min[category] = sum(time)
        for project, min in total_time_min.items():
            total_time_hours[project] = round(min / 60, 2)


    # Info summary file
    summary = read_summary(client)

    # Avaliable hours at report past month
    avaliable_hours = {}
    for project, category in summary['projects'].items():
        avaliable_hours[project] = category['avalaible']


    # Avaliable hours at the moment
    avaliable_hours_update = {}
    for category, time in avaliable_hours.items():
        for project, hours in total_time_hours.items():
            if project == category:
                avaliable_hours_update[project] = round(time-hours, 2)

    for dict, projects in summary['projects'].items():
        for category, hour in avaliable_hours_update.items():
            if category == dict:
                projects['avalaible'] = hour

        for cat, time in total_time_hours.items():
            if dict == cat:
                projects['hours_month'][reportdate.capitalize()] = time

    # Update summary
    with open(f'/Users/Lúcia/Desktop/{company}/summary.json', encoding='UTF-8', mode='w') as summaryfile:
        json.dump(summary, summaryfile, indent=2)
        summaryfile.close()

    return avaliable_hours, total_time_hours, total_time_min, avaliable_hours_update


##
# Defines the composition of the report cover
#
# \param content : array with the elements that to compose the report
# \param reportdate : present month and year
# \param startdatT : date inserted at the beginning of the project for define the point of start extract tasks
# \param enddateT : date inserted at the beginning of the project for define the point of end extract tasks
# \param info_tasks : dictionary with tasks organized
# \return array updated with elements that compose the cover reports
#
def cover_version1(content, reportdate, startdateT, enddateT, info_tasks):

    print('Starting to doing the report . . .')

    # First line cover
    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='KOBU-Extralight',
                                    fontSize=14,
                                    textColor=colors.HexColor('#E53E44'),
                                    leading=15)

    content.append(Paragraph(f'Relatório {reportdate}', firstlinestyle))

    # Config file info
    configs = read_config(client)

    # Titles cover
    titlestyle = ParagraphStyle('headline',
                                fontName='KOBU-Headline',
                                fontSize=30,
                                textColor=colors.HexColor('#3A3C4B'),
                                leading=23)

    content.append(Paragraph(f'PROJECTOS {configs["report_title"].upper()}', titlestyle))

    pacoteStyle = ParagraphStyle('paragraph',
                                 fontName='KOBU-Regular',
                                 fontSize=12.5,
                                 textColor=colors.HexColor('#3A3C4B'),
                                 spaceBefore=10,
                                 spaceAfter=30)

    content.append(Paragraph(f'ao abrigo do(s) pacote(s) de hora(s) adquirido(s) a partir de '
                             f'{configs["contract_start_date"]}', pacoteStyle))

    # Times about avaliable hours last month and now
    times = calcules_version1(reportDate, info_tasks, logging, client)
    times_last_month = times[0]
    time_month = times[1]

    # Tommorrow day
    day = int(enddateT[:2]) + 1
    # Date that contrat end
    contract_start_date = f'(até {configs["contract_end_date"]})'

    # Define table for cover
    tables =[['Período a que diz respeito', f'{startdateT} a {enddateT}']]
    for category, info in configs['projects'].items():
        for area, time in times_last_month.items():
            for tema, hours in time_month.items():

                if category == area and category == tema:
                    tables.append([category.upper()])
                    tables.append([f'Total de horas no pacote adquirido {contract_start_date}', f'{info["total"]} horas'])
                    tables.append([f'Total de horas a {startdateT}', str(time).replace('.', ',')])
                    tables.append([f'Horas despendidas no período de {startdateT} a {enddateT}', str(hours).replace('.', ',')])
                    tables.append([f'Total de horas disponíveis a {day}/{enddateT[3:]}', str(round(time-hours, 2)).replace('.', ',')])

    # Style table
    style = TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#D2D2D7')),
                        ('FONTNAME', (0, 0), (-1, -1), 'KOBU-Regular'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 1), (0, 1), 'KOBU-Bold'),
                        ('FONTNAME', (0, 6), (0, 6), 'KOBU-Bold'),
                        ('FONTNAME', (1, 3), (1, 5), 'KOBU-Bold'),
                        ('FONTNAME', (1, 8), (1, -1), 'KOBU-Bold'),
                        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#3A3C4B')),
                        ])

    # Add table to cover
    table = Table(tables, colWidths=[290, 150])
    alternate_color(tables, table)
    table.setStyle(style)
    content.append(table)

    # Comments
    comments_cover(content)

    content.append(PageBreak())

    print('Report cover created')

    return content



def plot_version1(content, company, foldername):
    print("Graph in prodution . . .")
    fig = plt.figure(figsize=(5.6,3))
    plt.style.use('bmh')
    plt.grid(alpha=0)

    summary = read_summary(client)
    y = {}
    for projects, info in summary.items():
        for project, resume in info.items():
            y[project] = resume['hours_month']

    for project, info in y.items():
        x = []
        y = []
        for month, hours in info.items():
            x.append(month[0:3])
            y.append(hours)

        plt.plot(x,y, label=project)

    plt.ylim(ymin=0)

    plt.xlabel('Mês', fontname='KOBU-Bold', size=14)
    plt.ylabel('Horas', fontname='KOBU-Bold', size=14)
    plt.tight_layout()
    plt.legend(frameon=True, fontsize=8, framealpha=0.5, facecolor='#D2D2D7', bbox_to_anchor=(1.1, 1.05))

    fig.savefig(f'C:/Users/Lúcia/Desktop/{company}/{foldername}/graph.jpg')

    img = BytesIO()
    fig.savefig(img, format='svg')
    img.seek(0)

    drawing = svg2rlg(img)
    drawing.shift(-60, -50)
    content.append(drawing)

    return content


##
# Defines the composition of the summary page of report with a table to each contract
#
# \param content : array with the elements that to compose the report
# \param year : present year
# \return array updated with elements that compose the second page of report
def summary_version1(content, year):
    # Summary title
    titleStyle = ParagraphStyle('heading3',
                                fontName='KOBU-Headline',
                                fontSize=18,
                                textColor=colors.HexColor('#3A3C4B'),
                                spaceAfter=40,
                                alignment=TA_CENTER)

    content.append(Paragraph('Resumo de Consumo de Horas Mensal'.upper(), titleStyle))

    # Info summary file
    summary = read_summary(client)

    # Sum times of months in summary
    total = {}
    n = 0
    hours_month = {}
    for projects, summarys in summary.items():
        for project, info in summarys.items():
            x = info['hours_month']
            n = len(x)
            hours_month[project] = info['hours_month']
            total[project] = round(sum(x.values()), 2)

    # Calculate media of sum
    media = {}
    for category, time in total.items():
        media[category] = round((time/n), 2)


    # Creat a table for project and add totals and media
    tables = []
    time_avaliable_month = {}
    for project, info in summary['projects'].items():
        contract_time = info['total']
        time_avaliable_month[project] = int(contract_time/12)
        for category, time in total.items():
            for area, media_time in media.items():
                if project == category and project == area:
                    diferenca = round(contract_time-time, 2)
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
                            ['TOTAL', str(time).replace('.', ','), contract_time, str(diferenca).replace('.', ',')],
                            ['Média Mensal', str(media_time).replace('.', ',')]]

        # Add spended times for month to table
        for row in data:
            for value in row:
                for topic, avaliable_month in time_avaliable_month.items():
                    for tema, times in hours_month.items():
                        for month, time in times.items():

                            if project == tema and month == value and project == topic:
                                row.append(str(time).replace('.', ','))
                                row.append(avaliable_month)
                                row.append(str(round(avaliable_month-time, 2)).replace('.', ','))


                    # Case month doesn't have spended hours, add contract hours
                    if len(row) == 1 and project == topic and len(value) < 20:
                        row.append('')
                        row.append(avaliable_month)
                        row.append(str(round(float(avaliable_month), 2)).replace('.',','))

        data.append([])
        data.append([])
        tables.append(data)

    # Table style
    style = TableStyle([
                        ('INNERGRID', (0, 1), (-1, -3), 0.15, colors.HexColor('#D2D2D7')),
                        ('SCAN', (0,0), (0,0)),
                        ('ALIGN',(0, 1), (-1, -1), 'RIGHT'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 2), (-1, -5), 'KOBU-Regular'),
                        ('FONTNAME', (0, 0), (0, 0), 'KOBU-Headline'),
                        ('FONTNAME', (0, 1), (-1, 1), 'KOBU-Bold'),
                        ('FONTNAME', (0, -4), (0, -1), 'KOBU-Bold'),
                        ('TEXTCOLOR', (0,0),(-1,-1), colors.HexColor('#3A3C4B')),
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E6E6E6'))
                        ])

    # Add tables to report
    for table in tables:
        table = Table(table, colWidths=[90, 90, 90, 90])
        table.setStyle(style)
        content.append(table)

    plot_version1(content, client, date_folder)
    content.append(PageBreak())

    print('Hour consumption summary created')

    return content


##
# Defines the composition of the body to report, with a table for contract, each table has a date task, description and spent minutes
#
# \param content : array with the elements that to compose the report
# \param info_tasks : dictionary with tasks organized
# \return array updated with elements that compose the body report
def body_version1(content, info_tasks, logs):
    # Title Body
    titleStyle = ParagraphStyle('heading',
                                fontName='KOBU-Headline',
                                fontSize=18,
                                textColor=colors.HexColor('#3A3C4B'),
                                spaceAfter=40,
                                alignment=TA_CENTER)

    content.append(Paragraph('Relação de Projetos/Horas'.upper(), titleStyle))


    print('Body tables are in process . . .')

    # Extra style for description task
    styles = getSampleStyleSheet()
    style_tasks_description = styles["Normal"]
    style_tasks_description.fontName = "KOBU-Regular"
    style_tasks_description.textColor = colors.HexColor('#3A3C4B')

    # Extra style to bold id
    style_id_bold = styles["BodyText"]
    style_id_bold.fontName = "KOBU-Bold"
    style_id_bold.textColor = colors.HexColor('#3A3C4B')

    tasks = info_tasks
    config = read_config(client)

    times = calcules_version1(reportDate, info_tasks, logging, client)
    times_hours = times[1]
    times_min = times[2]

    # Create table for project
    tabelas = {}
    for project, info in config['projects'].items():
        id = 0
        data = {project: [[info['name'].upper()],
                          ['ID', 'Nome']]}

        # Add categories to tables
        for projecto, informacao in data.items():
            for title, category in info['categories'].items():
                id += 1
                id_table = Paragraph(str(id), style_id_bold)
                category_table = Paragraph(title, style_id_bold)
                informacao.append([id_table, category_table])

                #Add tasks info to tables
                exist_tasks = False
                for key, list_tasks in tasks.items():
                    for task in list_tasks:
                        description = Paragraph(task[5].replace('""', '"'), style_tasks_description)
                        if task[7] in category:
                            exist_tasks = True
                            informacao.append([task[3], description, task[6], '(min)'])

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
                    informacao.append(['', '', str(hour).replace('.', ','), '(h)'])

            informacao.append([''])
            tabelas[projecto] = informacao


    # Tables Style
    style = ([
        ('INNERGRID', (0, 2), (-1, -2), 0.15, colors.HexColor('#D2D2D7')),
        ('SCAN', (0,0), (0,0)),
        ('FONTNAME', (0, 3), (-1, -1), 'KOBU-Regular'),
        ('FONTNAME', (0, 1), (1, 1), 'KOBU-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0,0),(-1,0), colors.HexColor('#ffffff')),
        ('TEXTCOLOR', (0,1),(-1,-1), colors.HexColor('#3A3C4B')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D2D2D7')), #TABLE TITLES
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E6E6E6')),
        ('BACKGROUND', (1,-3), (1, -3), colors.HexColor('#E6E6E6')),
        ('BACKGROUND', (2,-3),(2,-2), colors.HexColor('#E6E6E6')),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1,-3), (1, -3), 'RIGHT'),
        ('FONTNAME', (1,-3), (1, -3), 'KOBU-Bold'),
        ('FONTNAME', (2,-2), (2, -2), 'KOBU-Bold')
    ])

    # To assign an id to table
    project_id = 0
    for project, tabela in tabelas.items():
        project_id += 1
        table = Table(tabela, repeatRows=2, colWidths=[70, 345, 50, 45])
        table.setStyle(TableStyle(style))
        content.append(table)

        # When arrives to last table does't change to page
        if project_id != len(tabelas):
            content.append(PageBreak())

    return content


##
# Checks what is the version of report will be extracted and extract the pdf document
#
# \param caracteristicas : array with the elements that to compose the report
# \param pdf : define the methods for manipulate the data in the report
# \param info_tasks : dictionary with tasks organized
#
def export_pdf(caracteristicas, pdf, info_tasks):
    # Differentiat how much contracts have and call function for your version
    type = read_config(client)['number_contracts']
    if type == 1:
        cover_version2(caracteristicas, reportDate, startdateTable, enddateTable, info_tasks)
        summary_version2(caracteristicas, atualYear)
        body_version2(caracteristicas, info_tasks)
    else:
        cover_version1(caracteristicas, reportDate, startdateTable, enddateTable, info_tasks)
        summary_version1(caracteristicas, atualYear)
        body_version1(caracteristicas, info_tasks, logging)

    # Add footer to report
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='footer', frames=frame, onPage=footer)
    pdf.addPageTemplates([template])

    # Add elements to the report
    pdf.build(caracteristicas)

    return pdf
