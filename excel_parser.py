import csv
import time
from progress.bar import IncrementalBar
import win32com.client as win32
import pathlib
from pathlib import Path
from win32com.client import constants as c


def xls_to_csv(xls_file, csv_file):
    dir_path = pathlib.Path.cwd()
    file_name = Path(dir_path, xls_file)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(file_name)
    except:
        input(f"\n\tФайла {xls_file} не обнаружено. Нажмите Enter, чтобы выйти.")
        exit(1)
    result_path = str(dir_path) + "\\" + csv_file
    wb.SaveAs(result_path, c.xlCSV)
    wb.Close()


def templates_search(job: str, dep: str, matrix: csv, template_set: set):
    for z in range(1, len(matrix)):
        if job != "None" and str(matrix[z][2]).upper() in job.upper() and str(matrix[z][3]).upper() in dep.upper():
            xtemp = str(matrix[z][5]).upper()
            template_set.add(xtemp)
            break
        else:
            xtemp = "НЕТ ПОЗИЦИИ"
    return xtemp


def set_and_result(set_template: set, template_mms: str):
    set_template.discard('NONE')
    if '' in set_template and len(set_template) == 1:
        set_template.discard('')
        result_template = "НЕ ПОЛОЖЕНО"
        return result_template, "False"
    elif '' in set_template and len(set_template) > 1:
        set_template.discard('')
    if '##' and '##+##' in set_template:
        set_template.discard('##')
    if '##' in set_template and '##_###' in set_template:
        set_template.discard('##_###')
    if '##_###' in set_template and len(set_template) > 1:
        set_template.discard('##_###')
        set_template.add('##')
    if '########_####_###' in set_template and len(set_template) > 1:
        set_template.clear()
        set_template.add('########_####_###')
    if '##' in set_template and len(set_template) > 1:
        set_template.clear()
        set_template.add('##')
    result_template = '+'.join(sorted(set_template))

    if result_template == "":
        result_template = "ДОСТУП НЕ НАЙДЕН"

    return result_template, "True" if result_template == template_mms else "False"


if __name__ == '__main__':
    try:
        input("\n\tДобрый день!\n\n\tНеобходимо добавить файлы в папке со скриптом:\n\t- app_users.csv"
              "\n\t- matrix.csv\n\t- job_combination.XLS\n\n\tИ нажать Enter.")
        start_time = time.time()
        xls_to_csv('job_combination.XLS', 'job_combination.csv')
        
        with open('job_combination.csv') as jobfile:
            reader = csv.reader(jobfile, delimiter=';')
            joblst = list(list(reader))

        with open('matrix.csv') as matrix:
            reader = csv.reader(matrix, delimiter=';')
            matrixlst = list(list(reader))

        with open('app_users.csv') as appfile:
            reader = csv.reader(appfile, delimiter=';')
            applst = list(list(reader))

        print('\n\tНачинается выполнение...\n')
        csv_file = open('###_parser_result.csv', 'w', newline='')
        with csv_file as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(
                [['APP_LOGIN;CLIENT;TEMPLATE;RESULT;IF-ELSE;TTH;LOCATION;JOB1;DEP1;JOB2;DEP2;JOB3;DEP3;JOB4;DEP4;JOB5;DEP5'
                  ';TEMP1;TEMP2;TEMP3;TEMP4;TEMP5']]
            )
            bar = IncrementalBar('Progress', max=len(applst)-1)

            for x in range(1, len(applst)):
                low = 1
                high = len(joblst)
                while (high - low) > 1:
                    mid = (low + high) // 2
                    if applst[x][6] > joblst[mid][0]:
                        low = mid
                    elif applst[x][6] < joblst[mid][0]:
                        high = mid
                    else:
                        break
                    mid = (low + high) // 2
                if joblst[mid][4] in ['####', '####', '####', '###'] or joblst[mid][8] in \
                        {
                            '##########',
                            '################',
                            '#######',
                            '####################################',
                            '######################',
                            '################',
                            '#########################',
                            '####################',
                            '#############'
                        }:
                    bar.next()
                    continue
                app_login = str(applst[x][1])
                client = str(applst[x][0])
                template = str(applst[x][5])
                TTH = str(applst[x][6])
                location = str(joblst[mid][4])

                [job1, job2, job3, job4, job5] = str(joblst[mid][7]), \
                                                 str(joblst[mid][11]), \
                                                 str(joblst[mid][15]), \
                                                 str(joblst[mid][19]), \
                                                 str(joblst[mid][23])

                [dep1, dep2, dep3, dep4, dep5] = str(joblst[mid][8]), \
                                                 str(joblst[mid][12]), \
                                                 str(joblst[mid][16]), \
                                                 str(joblst[mid][20]), \
                                                 str(joblst[mid][24])

                temp_set = set()
                temp1 = templates_search(job1, dep1, matrixlst, temp_set)
                temp2 = templates_search(job2, dep2, matrixlst, temp_set)
                temp3 = templates_search(job3, dep3, matrixlst, temp_set)
                temp4 = templates_search(job4, dep4, matrixlst, temp_set)
                temp5 = templates_search(job5, dep5, matrixlst, temp_set)

                result, if_else = set_and_result(temp_set, template)

                writer.writerows(
                    [[
                        app_login + ";" +  # mms
                        client + ";" +  # CLIENT
                        template + ";" +  # TEMPLATE
                        result + ";" +
                        if_else + ";" +
                        TTH + ";" +  # SAP
                        location + ";" +  # STORE
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
        jobfile.close()
        appfile.close()
        file.close()
        matrix.close()
        print("\n\n\t--- %s seconds, yo ---" % round(time.time() - start_time))
        input('\n\tГотово! Намжите Enter, чтобы выйти.\n')
    except KeyboardInterrupt:
        input("\n\n\tВызвано прерывание программы. Нажмите Enter, чтобы выйти.\n")
