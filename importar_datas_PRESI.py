"""
https://stackoverflow.com/questions/2600775/how-to-get-week-number-in-python
https://stackoverflow.com/questions/24370385/how-to-format-cell-with-datetime-object-of-the-form-yyyy-mm-dd-hhmmss-in-exc
"""

import datetime
import logging
import os
import subprocess
import sys
import time
from copy import copy

import openpyxl
from openpyxl.styles import NamedStyle
from openpyxl.styles import Alignment

main_logger = logging.getLogger(__name__)
main_logger.setLevel(logging.DEBUG)

file_formatter = logging.Formatter('%(asctime)s||%(levelname)s||%(name)s||%(message)s')
stream_formatter = logging.Formatter('%(message)s')

file_handler = logging.FileHandler(f'{os.path.join("Logs",os.path.splitext(os.path.basename(__file__))[0])}.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(file_formatter)

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(stream_formatter)

main_logger.addHandler(file_handler)
main_logger.addHandler(stream_handler)


def main():
    if sys.platform in ('linux', 'darwin'):
        excel = '/mnt/c/Users/joao.caldeira.ext/SPMS - Serviços Partilhados do Ministério da Saúde, EPE/SPMS DSI RDIS - RIS Corporate - RIS2020 Cadastro/RIS2020 - Cadastro.xlsx'
        excel2 = '/mnt/c/Users/joao.caldeira.ext/OneDrive - Portugal Telecom/Trabalho/COS/SPMS/PRESI/Calendarização-quattro.xlsx'
    elif sys.platform == 'win32':
        excel = r"C:\Users\joao.caldeira.ext\SPMS - Serviços Partilhados do Ministério da Saúde, EPE\SPMS DSI RDIS - RIS Corporate - RIS2020 Cadastro\RIS2020 - Cadastro.xlsx"
        excel2 = r"C:\Users\joao.caldeira.ext\OneDrive - Portugal Telecom\Trabalho\COS\SPMS\PRESI\Calendarização-quattro.xlsx"

    excelFileName = os.path.basename(excel)
    pathToExcelFile = os.path.dirname(excel)

    excelFileName2 = os.path.basename(excel2)
    pathToExcelFile2 = os.path.dirname(excel2)

    chars_a_remover = {
                        'á': 'a',
                        'Á': 'A',
                        'à': 'a',
                        'À': 'A',
                        'ã': 'a',
                        'Ã': 'A',
                        'â': 'a',
                        'Â': 'A',
                        'é': 'e',
                        'É': 'E',
                        'è': 'e',
                        'È': 'E',
                        'ê': 'e',
                        'Ê': 'E',
                        'í': 'i',
                        'Í': 'I',
                        'ó': 'o',
                        'Ó': 'O',
                        'õ': 'o',
                        'Õ': 'O',
                        'ô': 'o',
                        'Ô': 'O',
                        'ú': 'u',
                        'Ú': 'U',
                        'û': 'u',
                        'Û': 'U',
                        'ù': 'u',
                        'Ù': 'U',
                        'ç': 'c',
                        's/n': ''
    }

    # Open Excel
    try:
        wb = openpyxl.load_workbook(excel, data_only=True)
        delTempFileFlag = False
    except PermissionError as e:
        main_logger.info(f'File is in use, so cannot be accessed:\n{e.filename}\n')
        createTempFile(pathToExcelFile, os.getcwd(), excelFileName)
        main_logger.info('Temporary copy of the file created to work with\n')
        wb = openpyxl.load_workbook(excel, data_only=True)
        delTempFileFlag = True
    except FileNotFoundError as e:
        sys.exit(f'File not found: {e.filename}')


    # Open Excel 2
    try:
        wb2 = openpyxl.load_workbook(excel2, data_only=True)
        delTempFileFlag2 = False
    except PermissionError as e:
        main_logger.info(f'File is in use, so cannot be accessed:\n{e.filename}\n')
        createTempFile(pathToExcelFile2, os.getcwd(), excelFileName2)
        main_logger.info('Temporary copy of the file created to work with\n')
        wb2 = openpyxl.load_workbook(excel2, data_only=True)
        delTempFileFlag2 = True
    except FileNotFoundError as e:
        sys.exit(f'File not found: {e.filename}')


    sheet = wb['RIS2020']
    sheet2 = wb2[f"Week {datetime.date.today().strftime('%V')}"]

    # date_style = NamedStyle(name='date', number_format='DD/MM/YYYY')
    date_style = NamedStyle(
        name='date',
        number_format='dd-mmmm',
        alignment=Alignment(
            horizontal='center',
            vertical='center',
        ),
    )

    lista_sites = []

    # Get all sites with schedule for this week
    for row in sheet2.iter_rows(min_row=2, values_only=True):
        if type(row[5]) == datetime.datetime and str(row[10]).upper() == 'CONFIRMADO':
            site_info = {
                'site_id': str(row[0]).zfill(4),
                'date_presi': row[5].date(),
            }

            lista_sites.append(site_info)
    for site in lista_sites:
        for row_index, row2 in enumerate(sheet.iter_rows(min_row=3, values_only=True), 3):
            # if site_info['site_id'] == row2[0] and row2[27] == '':
            if site['site_id'] == row2[0] and row2[27] == None:
                sheet[f'AA{row_index}'].value = site['date_presi']
                sheet[f'AA{row_index}'].style = date_style


    wb.save(excel)

    wb.close()
    wb2.close()
    if delTempFileFlag == True:
        os.remove(excelFileName)
    if delTempFileFlag2 == True:
        os.remove(excelFileName2)

    return len(lista_sites)



def cls():
    if sys.platform in ("linux", "darwin"):
        subprocess.run(["clear"])
    elif sys.platform == "win32":
        subprocess.run(["cls"], shell=True)



def createTempFile(pathToExcelFile, cwd, excelFileName):
    if sys.platform in ('linux', 'darwin'):
        subprocess.run(['cp', pathToExcelFile, cwd, excelFileName])
    elif sys.platform == 'win32':
        subprocess.run(['robocopy', pathToExcelFile, cwd, excelFileName], shell=True, stdout=subprocess.DEVNULL)

    # tempFile = f'{os.path.splitext(excelFileName)[0]}.temp{os.path.splitext(excelFileName)[1]}'
    # os.rename(excelFileName, f'temp-{tempFile}')




if __name__ == '__main__':
    cls()
    time_start = datetime.datetime.now()
    time_start = time_start.strftime('%H:%M:%S %d/%m/%Y')
    perf_counter_start = time.perf_counter()

    sites = main()

    perf_counter_stop = time.perf_counter()
    time_stop = datetime.datetime.now()
    time_stop = time_stop.strftime('%H:%M:%S %d/%m/%Y')

    main_logger.info(f'\nNumber of sites: {sites}')
    main_logger.info(f'Finished in {(perf_counter_stop - perf_counter_start):.3f} second(s)')
    main_logger.info(f'Started at {time_start} and finished at {time_stop}\n')
