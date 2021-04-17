import todoist_project
import csv, copy
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


# Read CSV and Save info tasks in taskslist
def ler_csv(company, tasksinformation):
    csv_file = '/Users/Lúcia/Desktop/report_' + company.replace(' ', '_') + '.csv'
    with open(csv_file) as output:
        csv_reader = csv.reader(output)
        csv_reader.__next__()

        for row in csv_reader:
            tasksinformation.append(row[2:])

    return tasksinformation


def organize_tasks(taskinformation): #anual
    #Dicionario com tasks agrupadas por categoria e dict com os tempos por categoria
    by_category = defaultdict(list)
    time = defaultdict(list)
    for info in taskinformation:
        by_category[info[4]].append([info[0], info[1], info[2], info[3], info[4]])
        time[info[4]].append(float(info[3]))


    # Tempo total gasto por categoria em horas
    time_min = {}
    for category, times in time.items():
        time[category] = round(sum(times)/60, 2)
        time_min[category] = round(sum(times), 2)


    #Depois de agrupadas as tasks, ordenei por data
    by_date = defaultdict(list)
    for category, task in by_category.items():
        ordered = sorted(task, key=itemgetter(0))
        by_date[category].append(ordered)

    return time, time_min, by_date


####################Anual Version#############################
def prepare_PDF_anual(company, taskinformation):
    read_summary = todoist_project.read_summary_anual(company)
    hours_disponiveis = read_summary[0]#hrs que ficaram no mês passado

    transform = []
    for time in hours_disponiveis:
        transform.append(float(time.replace(',','.')))

    # Categorias do projecto
    times = organize_tasks(taskinformation)[0]
    justtimes = list(times.values())#horas despendidas este mês -  indice[0] = digital, indice[1] = brand design

    avaliable = [] #horas disponiveis a partir de hoje - indice[0] = digital, indice[1] = brand design
    for minutes, horas in zip(transform,justtimes):
        avaliable.append(minutes-horas)

    #Atualização de summary.txt com os tempos em avaliable
    get_names = read_summary[6][:2]
    names = []
    for name in get_names:
        spar = name.split('-')
        names.append(spar[0])

    safe = read_summary[6][2:]

    file = '/Users/Lúcia/Desktop/' + company + '/summary.txt'
    with open(file, 'w') as file:
        file.write(f'{names[0]}- {avaliable[0]}\n')
        file.write(f'{names[1]}- {avaliable[1]}\n')
        for i in safe:
            file.write(i)
        file.write(f'{reportDate.capitalize()} - {justtimes[0]} - {justtimes[1]} \n')


    return hours_disponiveis, justtimes, avaliable


def cover_versao_anual(content, company, taskinformation):

    # Função para ler os config.txt
    configuracoes = todoist_project.read_config(company)

    titleStyle = ParagraphStyle('heading2',
                                fontName='Helvetica',
                                fontSize=12,
                                textColor=colors.blue,
                                leading=15)

    content.append(Paragraph(f'RELATÓRIO {todoist_project.reportDate}', titleStyle))

    subtitleStyle = ParagraphStyle('heading1',
                                   fontName='Helvetica-Bold',
                                   fontSize=20,
                                   textColor=colors.black,
                                   leading=15)

    content.append(Paragraph(f'PROJECTOS {configuracoes[1]}', subtitleStyle))

    infoStyle = ParagraphStyle('paragraph',
                               fontName='Helvetica',
                               fontSize=12,
                               spaceBefore=10,
                               spaceAfter=30)

    content.append(Paragraph(f'ao abrigo do(s) pacote(s) de hora(s) adquirido(s) a partir de {configuracoes[2]}', infoStyle))

    times = prepare_PDF_anual(company, taskinformation)
    hours_disponiveis = times[0]
    justtimes = times[1]
    avaliable = times[2]

    #categorias do projectos encontradas no config
    categorys = configuracoes[4] + configuracoes[5]

    day = int(enddate[:2]) + 1
    dataTable = [['Perído a que diz respeito', f'{startdate} a {enddate}             '],
                 [categorys[0]],
                 [f'Total de horas no pacote adquirido ({configuracoes[3]})', categorys[1][1:-1] + ' horas'],
                 [f'Total de horas a {startdate}', str(hours_disponiveis[0].replace('.', ','))],
                 [f'Horas despendidas no período de {startdate} a {enddate}', str(justtimes[0]).replace('.',',')],
                 [f'Total de horas disponíveis a {day}/{enddate[3:]}', str(avaliable[0]).replace('.', ',')],
                 [categorys[2]],
                 [f'Total de horas no pacote adquirido ({configuracoes[3]})', categorys[3][1:-1] + ' horas'],
                 [f'Total de horas a {startdate}', str(hours_disponiveis[1].replace('.', ','))],
                 [f'Horas despendidas no período de {startdate} a {enddate}', str(justtimes[1]).replace('.',',')],
                 [f'Total de horas disponíveis a {day}/{enddate[3:]}', str(avaliable[1]).replace('.', ',')]]


    table = Table(dataTable)
    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black), #linhas dentro
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),#tipo de letra
                        ('FONTSIZE', (0, 0), (-1, -1), 10),#tamanho de letra
                        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 6), (0, 6), 'Helvetica-Bold'),
                        ('FONTNAME', (1, 3), (1, 5), 'Helvetica-Bold'),
                        ('FONTNAME', (1, 8), (1, -1), 'Helvetica-Bold')])

    # Alternate color BACKGROUND table
    todoist_project.alternate_color(dataTable, table)

    table.setStyle(style)
    content.append(table)

    # Comments
    todoist_project.comments_cover(content)

    return content



def bodyAnual(content, company, taskinformation):
    titleStyle = ParagraphStyle('heading',
                                fontName='Helvetica-Bold',
                                fontSize=14,
                                textColor=colors.black,
                                leading=70,
                                alignment=TA_CENTER)

    content.append(Paragraph('3. Relação de Projetos/Horas', titleStyle))

    titles = todoist_project.read_config(company)


    #Este estilo foi criado para o texto não passar a largura da tabela e fazer \n
    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]

    titleTable = titles[6][:-1]
    title1Table = titles[7][0][:-1]

    data = [[titleTable],
            ['ID', 'Nome'],
            ['1', title1Table]]

    title2Table = titles[7][2][:-1]
    data1 = [[titleTable],
            ['ID', 'Nome'],
            ['2' , title2Table]]

    categorys = titles[4] + titles[5]
    category1 = categorys[0][:-1]#digital
    category2 = categorys[2][:-1]#brand design

    # Adicionar as tasks á tabela data
    tableContent = organize_tasks(taskinformation)

    tasks = tableContent[2]
    time_hours = tableContent[0]
    time_min = tableContent[1]

    #Adiciona à tabela, data, task, tempo em min por task
    for category, task in tasks.items():
        for lista in task:
            for task in lista:
                descripition = Paragraph(task[2], styleN)
                if category == category1.lower():
                    data.append([task[0], descripition, task[3], '(min)'])

                elif category == category2.lower():
                    data1.append([task[0], descripition, task[3], '(min)'])


    #Adiciona à tabela tempo total em min na primeira categoria
    for category, time in time_min.items():
        if category == category1.lower():
            data.append(['', '', str(time).replace('.', ','), '(min)'])
        elif category == category2.lower():
            data1.append(['', '', str(time).replace('.', ','), '(min)'])

    #Adiciona o tempo em horas gasto com a primeira categoria
    for category, time in time_hours.items():
        if category == category1.lower():
            data.append(['', '', str(time).replace('.', ','), '(h)'])
            data.append([])
            data.append(['', 'TOTAL MÊS', str(time).replace('.', ','), '(h)'])
        elif category == category2.lower():
            data1.append(['', '', str(time).replace('.', ','), '(h)'])
            data1.append([])
            data1.append(['', 'TOTAL MÊS', str(time).replace('.', ','), '(h)'])


    table = Table(data, repeatRows=3, colWidths=[70, 345, 50, 45])#largura de cada célula
    table1 = Table(data1, repeatRows=3, colWidths=[70, 345, 50, 45])

    tableStyle = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                             ('INNERGRID', (0, 2), (-1, -1), 0.15, colors.black), #linhas dentro
                             ('SCAN', (0,0), (0,0)),
                             ('FONTNAME', (1,-1), (1, -1), 'Helvetica-Bold'), #TOTAL MÊS <--
                             ('FONTNAME', (0, 0), (1, 2), 'Helvetica-Bold'),
                             ('FONTNAME', (0, 3), (-1, -1), 'Helvetica'),
                             ('FONTSIZE', (0, 0), (-1, -1), 10), #tamanho de letra
                             ('TEXTCOLOR', (1,-1), (1, -1), colors.white),
                             ('ALIGN',(0, 3), (0, -1), 'RIGHT'), #alinhamento a esquerda
                             ('ALIGN',(2, 0), (-1, -1), 'LEFT'),
                             ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                             ('ALIGN', (1,-1), (1, -1), 'RIGHT'),
                             ('VALIGN', (0, 1), (0, -1), 'TOP'),
                             ('VALIGN', (2, 1), (3, -1), 'TOP'),
                             ('BACKGROUND', (0, 0), (3, 0), colors.dodgerblue),
                             ('BACKGROUND', (0, 1), (-1, 1), colors.deepskyblue),
                             ('BACKGROUND', (1,-1), (1, -1), colors.gray),
                             ('BACKGROUND', (2,-4),(2,-3), colors.deepskyblue),
                             ('BACKGROUND', (-2,-1),(-2,-1), colors.deepskyblue)])

    table.setStyle(tableStyle)
    table1.setStyle(tableStyle)

    content.append(table)
    content.append(PageBreak())
    content.append(table1)

    return content


######################Mensal Version##############################
def prepare_PDF_mensal(taskinformation, firstdate, company):
    ordered = organize_tasks(taskinformation)[2]

    time_min_subproject = defaultdict(list)#horas por subprojecto
    time_min_categoria = defaultdict(list)#min por categoria existente no csv
    for category, task in ordered.items():
        for task in task:
            for i in task:
                time_min_categoria[category].append(int(i[3]))
                time_min_subproject[i[1]].append(int(i[3]))

    total = 0 #total de horas gastas no projeto
    time_hours_suproject = {}#horas por subprojecto
    for subproject, min in time_min_subproject.items():
        time_min_subproject[subproject] = sum(min)
        time_hours_suproject[subproject] = round(sum(min) / 60, 2)
        for min in min:
            total += min

    total_hours = str(round(total/60, 2)).replace('.', ',')

    # Update Summary Hours
    summary = '/Users/Lúcia/Desktop/' + company + '/summary.txt'
    with open(summary, 'a') as file:
        file.write(firstdate.capitalize() + ' - ' + total_hours + '-' + '\n')

    return total_hours, time_hours_suproject


def cover_versao_mensal(content, company, firstdate):
    times = prepare_PDF_mensal(taskslist, reportDate, company)

    firstlinestyle = ParagraphStyle('heading2',
                                    fontName='Helvetica',
                                    fontSize=12,
                                    textColor=colors.blue)

    content.append(Paragraph(f'RELATÓRIO {firstdate}', firstlinestyle))

    config = todoist_project.read_config(company)[0]


    titlestyle = ParagraphStyle('heading1',
                                fontName='Helvetica-Bold',
                                fontSize=20,
                                textColor=colors.black,
                                leading=20)

    content.append(Paragraph('PROJECTOS', titlestyle))
    content.append(Paragraph(config[1], titlestyle))#segunda linha do relatório

    titlestyle2 = ParagraphStyle('heading1',
                                 fontName='Helvetica-Bold',
                                 fontSize=20,
                                 textColor=colors.black,
                                 leading=70)

    content.append(Paragraph(config[2], titlestyle2))

    # First table (capa)  https://www.youtube.com/watch?v=B3OCXBL4Hxs&t=309s
    data = [['Período a que diz respeito', f'{startdate} a {enddate}              '],
            [' ', ' '],
            [f'Horas despendidas no período de {startdate} a {enddate}', times[0]],
            ['Por rubrica:', ' ']]

    sub_hour = times[1]
    for pro in config[3:]:
        times_capa = '0,00'
        enumerate = pro.split('-')
        for z, m in sub_hour.items():
            if z in enumerate[0]:
                enumerate.append(str(m))
                times_capa = enumerate[2].replace('.', ',')
        else:
            data.append([enumerate[1] + (' (horas)'), times_capa])


    table = Table(data)
    style = TableStyle([('BOX', (0,0), (-1,-1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black), #linhas dentro
                        ('ALIGN',(0,3),(0,-1),'RIGHT'), #alinhamento a esquerda na primeira coluna na 3 celula até à ultima
                        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),#tipo de letra
                        ('FONTSIZE', (0,0), (-1,-1), 10),#tamanho de letra
                        ('BACKGROUND', (0,1), (2,1), colors.cornflowerblue),#cor de fundo
                        ('FONTNAME', (1,2), (1,2), 'Helvetica-Bold')])

    # Alternate color BACKGROUND table
    todoist_project.alternate_color(data, table)

    table.setStyle(style)
    content.append(table)

    # Comments
    todoist_project.comments_cover(content)

    return content


def bodyMensal(content, company, taskinformation):
    titleStyle = ParagraphStyle('heading',
                                fontName='Helvetica-Bold',
                                fontSize=14,
                                textColor=colors.black,
                                leading=70,
                                alignment=TA_CENTER)

    content.append(Paragraph('Relação de Projetos/Horas', titleStyle))

    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]


    tasks = organize_tasks(taskinformation)[2]

    config = todoist_project.read_config(company)[1]

    titles = {} #key = nome do suprojecto no csv, value = nome que terá no relatório
    categorys = defaultdict(list)
    for i in config:
        subproject = i[0].split('-')
        title = i[1].split('_')
        titles[subproject[0]] = title[0]

        category = title[1].split(',')
        categorys[subproject[0]].append(category)



    style = TableStyle([('BOX', (0, 0), (-1, -1), 0.15, colors.black), #linhas de fora
                        ('INNERGRID', (0, 2), (-1, -1), 0.15, colors.black), #linhas dentro
                        ('SCAN', (0,0), (0,0)),
                        ('FONTNAME', (1,-1), (1, -1), 'Helvetica-Bold'), #TOTAL MÊS <--
                        ('FONTNAME', (0, 0), (1, 2), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 3), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10), #tamanho de letra
                        ('TEXTCOLOR', (1,-1), (1, -1), colors.white),
                        ('ALIGN',(0, 3), (0, -1), 'RIGHT'), #alinhamento a esquerda
                        ('ALIGN',(2, 0), (-1, -1), 'LEFT'),
                        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                        ('ALIGN', (1,-1), (1, -1), 'RIGHT'),
                        ('VALIGN', (0, 1), (0, -1), 'TOP'),
                        ('VALIGN', (2, 1), (3, -1), 'TOP'),
                        ('BACKGROUND', (0, 0), (3, 0), colors.dodgerblue),
                        ('BACKGROUND', (0, 1), (-1, 1), colors.deepskyblue),
                        ('BACKGROUND', (1,-1), (1, -1), colors.gray),
                        ('BACKGROUND', (2,-4),(2,-3), colors.deepskyblue),
                        ('BACKGROUND', (-2,-1),(-2,-1), colors.deepskyblue)])


    tabelas = []

    for project, title in titles.items():
        data = [[title],
                ['ID', 'Nome']]

        for pro, category in categorys.items():
            for categoria in category:
                id = 0
                for i in categoria:
                    if project == pro:
                        id +=1
                        data.append([id ,i])

                        for tema, lista in tasks.items():
                            for task in lista:
                                for taskl in task:
                                    if tema in i and taskl[1] in project:
                                        data.append([taskl[0], taskl[2], taskl[3], '(min)'])

        data.append([' ', ' ', '-', '(min)'])
        tabelas.append(data)


    for tabela in tabelas:
        table = Table(tabela, colWidths=[70, 345, 50, 45])
        table.setStyle(style)
        content.append(table)

    return content



def export_pdf(caracteristicas):
    type = todoist_project.read_config(client)[0]
    if 'anual' in type:
        cover_versao_anual(caracteristicas, client, taskslist)
        todoist_project.secondPageAnual(caracteristicas, client, atualYear)
        bodyAnual(caracteristicas, client, taskslist)
    else:
        cover_versao_mensal(caracteristicas, client, reportDate)
        todoist_project.secondPageMensalVersion(caracteristicas,client, atualYear)
        bodyMensal(caracteristicas, client, taskslist)

    pdf = BaseDocTemplate(f'KOBU_{client}_csv.pdf', pagesize=A4)
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='footer', frames=frame, onPage=todoist_project.footer)
    pdf.addPageTemplates([template])

    pdf.build(caracteristicas)
    print('PDF EXPORTED!')

    return pdf


if __name__ == '__main__':
    ler_csv(client, taskslist)
    export_pdf(elementsPDF)
