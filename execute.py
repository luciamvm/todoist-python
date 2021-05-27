##
# \mainpage Automatização da criação de relatórios para os diversos clientes da KOBU
# \section intro_sec Introdution
# Os colaboradores da agência usam o Todoist, uma aplicação para gerir projetos e tarefas,
# na qual cada projeto é um cliente e inserido no projeto eles têm registadas as suas tarefas
# com a data a que foi realizada e o tempo que demoraram a concluí-la. O CEO todos os meses precisa
# de entregar relatórios aos clientes da agência com as tarefas desenvolvidas e o tempo que foi despendido.
# O objetivo do software é automatizar todo o processo, desde a extração de tarefas da aplicação como a
# criação do dito relatório em PDF. Estes relatórios têm duas versões sendo que a versão 1 é destinada a
# clientes que tenham estipulado dois ou mais contratos com a agência, incluindo um pacote de horas.
# E a versão 2 é destinada a clientes que só tenham um contrato estipulado.
#
# \subsection title Repository Github
# https://github.com/luciamvm/todoist-python.git
##

##
# \file execute.py
# É aqui que todos os ficheiros e funções principais definidos anteriormente se encontram e é criada uma interação
# com o utilizador do software. Informando-o daquilo que está a acontecer e validando qual o momento certo para extrair
# o documento .pdf, sendo que sempre que sejam detetados erros no conteúdo das tarefas não será permitido fazê-lo.
##


import data_process
import export_pdf
import os, csv
from collections import defaultdict
from operator import itemgetter


client = data_process.client
teamlogin = data_process.teamlogin
logging = data_process.logging
date = data_process.date
taskslist = data_process.taskslist

elements = export_pdf.elements
doc = export_pdf.doc

process_errors = []


##
# Checks for errors that may exist on the task date, time or category
#
# \param company : name of client to which will be extracted report
# \param tasks : dictionary with tasks organized
# \return array with the errors that occurred in the validation
def validations(tasks):
    print('A validar campos das tarefas')
    config = export_pdf.read_config(client)
    validation_tasks = tasks

    # Categories of each project
    categories = []
    for project, info in config['projects'].items():
        if type(info['categories']) == list:
            categories.append(info['categories'])
        elif type(info['categories']) == dict:
            for title, arraycategories in info['categories'].items():
                categories.append(arraycategories)

    # Validations that mustn't export pdf
    array_all_categories = []
    for array in categories:
        for i in array:
            array_all_categories.append(i)

    position = 0
    errors = []
    for task in validation_tasks:
        task[3].replace(' ', '')
        date = task[3].find('-')

        if task[7] not in array_all_categories:
            validation_tasks[position] = ['error', task[1], task[2], task[3], task[4], task[5], task[6], task[7]]
            errors.append(f'{task[5]} - Task does not have category association')

        if date != 4 or task[3] == 'Error':
            validation_tasks[position] = ['error', task[1], task[2], task[3], task[4], task[5], task[6], task[7]]
            errors.append(f'{task[5]} - Task date have a format error')

        if task[5] == 'Error':
            validation_tasks[position] = ['error', task[1], task[2], task[3], task[4], task[5], task[6], task[7]]
            errors.append(f'{task[5]} - Task time is "Error"')

        position += 1


    return validation_tasks, errors



##
# Agroup tasks by category and order by date
#
# \param tasksinformation : array with an array with information about each task
# \return dictionary where the key is the category and the values are an array with tasks that belong category
def organize(tasksinformation):
    print('A organizar as tarefas')

    tasks = validations(tasksinformation)[0]

    # Agroup tasks by category
    by_category = defaultdict(list)
    for task in tasks:
        by_category[task[7]].append([task[0], task[1], task[2], task[3], task[4], task[5].replace('__', ''), task[6], task[7].strip()])


    # Order tasks by date
    orderedtasks = defaultdict(list)
    for category, tasks in by_category.items():
        ordered = sorted(tasks, key=itemgetter(3))
        for task_info in ordered:
            orderedtasks[category].append(task_info)

    return orderedtasks



##
# Export csv file with data of tasks
#
# \param company : name of project where went tasks extracted
# \param folder : name of folder where csv file will be created
# \return the csv file
def export_csv(company, folder, errors):
    validation = validations(taskslist)
    validation_infotasks = validation[0]
    errors_validation = validation[1]
    for error in errors_validation:
        errors.append(error)

    ordered = organize(validation_infotasks)

    print('Is exporting csv doc...')

    # Add Titles to Generate Report List to Export
    tasksreportlist = [['error', 'user', 'closing date', 'task date', 'project name', 'task description', 'task time', 'category']]
    for category, tasks in ordered.items():
        tasksreportlist.extend(tasks)

    # Generate CSV File
    csvfile = f'/Users/Lúcia/Desktop/{company}/{folder}/tasksinfo_{company.replace(" ", "_")}.csv'
    with open(csvfile, "w") as output:
        writer = csv.writer(output, lineterminator='\n')
        writer.writerows(tasksreportlist)

    print(f'CSV for {company} finished.')

    return ordered



##
# Read the csv file of the project for extract the tasks
#
# \param company : name of client/project where is for operate
# \param folder : name of folder where are the csv file for read
# \return array with an array with information about each task present in csv file
def read_csv(company, folder, errors):

    print('Reading the CSV file after your corrected')

    info_tasks = []

    csv_file = f'/Users/Lúcia/Desktop/{company}/{folder}/tasksinfo_{company.replace(" " , "_")}.csv'
    with open(csv_file) as output:
        csv_reader = csv.reader(output, quoting=csv.QUOTE_NONE)
        csv_reader.__next__()
        for row in csv_reader:
            info_tasks.append([row[0], row[1].replace('"', ''), row[2], row[3], row[4],
                               row[5].replace('""""', '" "'), row[6], row[7].replace('"', '')])


    validation = validations(info_tasks)
    validation_infotasks = validation[0]
    ordered = organize(validation_infotasks)

    errors_validation = validation[1]
    for error in errors_validation:
        errors.append(error)

    return ordered



##
# Execute all files and verify if there are errors, interacting with user of software for correct it in csv file
#
# \param login : dictionary with name and token each user
# \param company : name of client to which will be extracted report
# \param tasks_todoist : array with an array with information about each task, extracted directly of todoist
# \param data_process_errors : array with errors that ocurred in validation of information tasks (dates)
# \param pdf : define the methods for manipulate the data in the report
# \param content : array with the elements that to compose the report
# \return exportation report
def execute(login, company, pdf, content, errors):

    data_process.connect(login, company, data_process.logging)

    # Export csv
    tasks = export_csv(company, data_process.date_folder, process_errors)

    # Logging file is empty or not
    loggingsize = os.path.getsize(f'/Users/Lúcia/Desktop/{client}/{data_process.date_folder}/debug.log')

    if loggingsize > 0:
        print('Sorry, failed operation. Check debug.log please!')

    else:
        # Checks if there are error in data exported to todoist, if there is request user for correct csv file
        if len(process_errors) > 0:
            print("\nImpossible to create the report, there are errors in the data")
            print('The following error wer found:')
            for error in process_errors:
                print('\t', '> ', error)

            report_now = input('\nWhen you have a csv ready write "ok":\n').upper()
            if report_now == 'OK':
                errors.clear()
                csv_tasks = read_csv(company, data_process.date_folder, process_errors)
                while len(process_errors) != 0:
                    errors.clear()
                    csv_tasks = read_csv(company, data_process.date_folder, process_errors)
                    print('The following errors were found:')
                    for erro in process_errors:
                        print('\t', '> ', erro)
                        csv_ready = input('\n If you have a csv with correction ready write "ok":\n')

                else:
                    export_pdf.export_pdf(content, pdf, csv_tasks)
                    print('Report was extracted!')

        else:
            export_pdf.export_pdf(content, pdf, tasks)
            print('Report was extracted!')

    return pdf


if __name__ == '__main__':
    execute(teamlogin, client, doc, elements, process_errors)
