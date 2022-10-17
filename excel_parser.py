import csv
import sys
import time
from progress.bar import IncrementalBar
import win32com.client as win32
import pathlib
from pathlib import Path
from win32com.client import constants as c


def xls_to_csv(xls_file_input, csv_file_output):
    dir_path = pathlib.Path.cwd()
    file_name = Path(dir_path, xls_file_input)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(file_name)
    except:
        print(f"\tФайл {xls_file_input} не обнаружен. Завершение работы через 3 секунды.\n")
        time.sleep(3)
        sys.exit()
    result_path = str(dir_path) + "\\" + csv_file_output
    wb.SaveAs(result_path, c.xlCSV)
    wb.Close()


def templates_search(job: str, dep: str, matrix, template_set: set):
    xtemp = ''
    for z in range(1, len(matrix)):
        if job != "None" and str(matrix[z][2]).upper() in job.upper() and str(matrix[z][3]).upper() in dep.upper():
            xtemp = str(matrix[z][5]).upper()
            template_set.add(xtemp)
            break
        else:
            xtemp = "НЕТ ПОЗИЦИИ"
    return xtemp


def set_and_result(set_template: set, template_app: str):
    set_template.discard('NONE')
    if '' in set_template and len(set_template) == 1:
        set_template.discard('')
        result_template = "НЕ ПОЛОЖЕНО"
        return result_template, "False"
    elif '' in set_template and len(set_template) > 1:
        set_template.discard('')
    if '##' and '##+##' in set_template:
        set_template.discard('##')
    if '##' in set_template and '##_SSO' in set_template:
        set_template.discard('##_SSO')
    if '##_SSO' in set_template and len(set_template) > 1:
        set_template.discard('##_SSO')
        set_template.add('##')
    if '###########_SSO' in set_template and len(set_template) > 1:
        set_template.clear()
        set_template.add('###########_SSO')
    if '##' in set_template and len(set_template) > 1:
        set_template.clear()
        set_template.add('##')
    result_template = '+'.join(sorted(set_template))

    if result_template == "":
        result_template = "ДОСТУП НЕ НАЙДЕН"

    return result_template, "True" if result_template == template_app else "False"


def with_open_file_to_list(file_to_list: csv):
    try:
        with open(file_to_list) as file_csv:
            file_reader = csv.reader(file_csv, delimiter=";")
            to_list = list(list(file_reader))
        file_csv.close()
        return to_list
    except FileNotFoundError:
        print(f"\tФайл {file_to_list} не обнаружен. Завершение работы через 3 секунды.")
        time.sleep(3)
        sys.exit()


if __name__ == '__main__':
    try:
        print("\n\tДобрый день!\n\n\tПроверяю наличие файлов:\n\t- app_users.csv"
              "\n\t- matrix.csv\n\t- job_combination.XLS\n")
        start_time = time.time()
        xls_to_csv('job_combination.XLS', 'job_combination.csv')
        job_list = with_open_file_to_list('job_combination.csv')
        matrix_list = with_open_file_to_list('matrix.csv')
        app_list = with_open_file_to_list('app_users.csv')

        print("\tФайлы на месте, начинается обработка.\n")

        csv_file = open('app_parser_result.csv', 'w', newline='')  # путь к файлу csv
        with csv_file as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(
                [['APP_LOGIN;CLIENT;TEMPLATE;RESULT;IF-ELSE;TTH;LOCATION;JOB1;DEP1;JOB2;DEP2;JOB3;DEP3;JOB4;DEP4;JOB5;DEP5'
                  ';TEMP1;TEMP2;TEMP3;TEMP4;TEMP5']]
            )
            bar = IncrementalBar('\tProgress', max=len(app_list) - 1)

            for x in range(1, len(app_list)):
                low = 1
                high = len(job_list)
                while (high - low) > 1:
                    mid = (low + high) // 2
                    if app_list[x][6] > job_list[mid][0]:
                        low = mid
                    elif app_list[x][6] < job_list[mid][0]:
                        high = mid
                    else:
                        break
                    mid = (low + high) // 2
                if job_list[mid][4] in ['####', '####', '####', '###', '####'] or job_list[mid][8] in \
                        {
                            '############################',
                            '#####################',
                            '#########################',
                            '##################################',
                            '#######################',
                            '############',
                            '#######################',
                            '######################',
                            '###########'
                        }:
                    bar.next()
                    continue
                app_login = str(app_list[x][1])
                client = str(app_list[x][0])
                template = str(app_list[x][5])
                TTH = str(app_list[x][6])
                location = str(job_list[mid][4])

                [job1, job2, job3, job4, job5] = str(job_list[mid][7]), \
                                                 str(job_list[mid][11]), \
                                                 str(job_list[mid][15]), \
                                                 str(job_list[mid][19]), \
                                                 str(job_list[mid][23])

                [dep1, dep2, dep3, dep4, dep5] = str(job_list[mid][8]), \
                                                 str(job_list[mid][12]), \
                                                 str(job_list[mid][16]), \
                                                 str(job_list[mid][20]), \
                                                 str(job_list[mid][24])

                temp_set = set()
                temp1 = templates_search(job1, dep1, matrix_list, temp_set)
                temp2 = templates_search(job2, dep2, matrix_list, temp_set)
                temp3 = templates_search(job3, dep3, matrix_list, temp_set)
                temp4 = templates_search(job4, dep4, matrix_list, temp_set)
                temp5 = templates_search(job5, dep5, matrix_list, temp_set)

                result, if_else = set_and_result(temp_set, template)

                writer.writerows(
                    [[
                        app_login + ";" +  # app
                        client + ";" +  # CLIENT
                        template + ";" +  # TEMPLATE
                        result + ";" +
                        if_else + ";" +
                        TTH + ";" +  # TTH
                        location + ";" +  # location
                        job1 + ";" +  # JOB1
                        dep1 + ";" +  # DEP1
                        job2 + ";" +  # JOB2
                        dep2 + ";" +  # DEP2
                        job3 + ";" +  # JOB3
                        dep3 + ";" +  # DEP3
                        job4 + ";" +  # JOB4
                        dep4 + ";" +  # DEP4
                        job5 + ";" +  # JOB5
                        dep5 + ";" +  # DEP5
                        temp1 + ";" + temp2 + ";" + temp3 + ";" + temp4 + ";" + temp5
                    ]]
                )
                bar.next()
        file.close()
        print("\n\n\t--- %s seconds, yo ---" % round(time.time() - start_time))
        print('\n\tГотово!\n')
    except KeyboardInterrupt:
        print("\n\n\tВызвано прерывание программы. Завершение работы через 3 секунды.\n")
        time.sleep(3)
        sys.exit()
