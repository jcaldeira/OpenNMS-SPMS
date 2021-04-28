"""Script criado por Rui Pereira"""

import os
import xlrd
import re
import os.path
import socket
from openpyxl import load_workbook
from os import path

if socket.gethostname() == 'LP013150':
    excel = 'C:\\Users\\rui.pereira.ext\\OneDrive - SPMS - Serviços Partilhados do Ministério da Saúde, EPE\\RIS2020 Cadastro\\RIS2020 - Cadastro.xlsx'
else:
    excel = 'D:\\SPMS OneDrive\\SPMS - Serviços Partilhados do Ministério da Saúde, EPE\\SPMS DSI RDIS - RIS Corporate - RIS2020 Cadastro\\RIS2020 - Cadastro.xlsx'

if not os.path.exists(excel):
    print ('Sem acesso ao ficheiro excel!')

# To open Workbook
open_excel = xlrd.open_workbook(excel)
ris_cadastro_sheet = open_excel.sheet_by_name('RIS2020')

for i in range(1, ris_cadastro_sheet.nrows):
    i += 1
ultima_coluna = i

sites_com_dois_routers = []
for i in range(2, ultima_coluna):
    hostname = str(ris_cadastro_sheet.cell_value(i, 23))
    ip = str(ris_cadastro_sheet.cell_value(i, 24))
    aces = str(ris_cadastro_sheet.cell_value(i, 28))
    site_id = hostname[0:4]
    chars_a_remover = ['á','à','ã','â','é','è','ê','í','ó','õ','ô','ú','û','ù','ç','s/n']
    chars_a_inserir = ['a','a','a','a','e','e','e','i','o','o','o','u','u','u','c','']
    if ip != '' and hostname == '2014-B01-SW01':
        if 'B01-SW01' not in hostname:
            for j in range(2, ultima_coluna):
                if str(ris_cadastro_sheet.cell_value(j, 0)) == site_id:
                    entidade = str(ris_cadastro_sheet.cell_value(j, 1))
                    entidade = entidade.replace(" ","_")
                    entidade_upper = entidade.upper()
                    morada = str(ris_cadastro_sheet.cell_value(j, 4))
                    cp = str(ris_cadastro_sheet.cell_value(j, 5))
                    localidade = str(ris_cadastro_sheet.cell_value(j, 6))
                    latitude = str(ris_cadastro_sheet.cell_value(j, 7))
                    longitude = str(ris_cadastro_sheet.cell_value(j, 8))
                    local = str(ris_cadastro_sheet.cell_value(j, 3))
                    for k in range(0,len(chars_a_remover)):
                        morada = morada.replace(chars_a_remover[k],chars_a_inserir[k])
                        localidade = localidade.replace(chars_a_remover[k], chars_a_inserir[k])
        else:
            entidade = str(ris_cadastro_sheet.cell_value(i, 1))
            entidade = entidade.replace(" ", "_")
            entidade_upper = entidade.upper()
            morada = str(ris_cadastro_sheet.cell_value(i, 4))
            cp = str(ris_cadastro_sheet.cell_value(i, 5))
            localidade = str(ris_cadastro_sheet.cell_value(i, 6))
            local = str(ris_cadastro_sheet.cell_value(i, 3))
            latitude = str(ris_cadastro_sheet.cell_value(i, 7))
            longitude = str(ris_cadastro_sheet.cell_value(i, 8))
            for k in range(0, len(chars_a_remover)):
                morada = morada.replace(chars_a_remover[k], chars_a_inserir[k])
                localidade = localidade.replace(chars_a_remover[k], chars_a_inserir[k])

        modelo = str(ris_cadastro_sheet.cell_value(i, 20))
        serial = str(ris_cadastro_sheet.cell_value(i, 21))
        aces = str(ris_cadastro_sheet.cell_value(i, 28))
        ip = str(ris_cadastro_sheet.cell_value(i, 24))
        node_id = str(int(ris_cadastro_sheet.cell_value(i, 27)))

        print ("/opt/opennms/bin/provision.pl service add '" + aces + "' " + node_id + " " + ip + " ICMP")
        print ("/opt/opennms/bin/provision.pl service add '" + aces + "' " + node_id + " " + ip + " SNMP")
        print ("/opt/opennms/bin/provision.pl category remove '" + aces + "' " + node_id + " '" + entidade + "'")
        print("/opt/opennms/bin/provision.pl category add '" + aces + "' " + node_id + " '" + str(entidade_upper) + "'")
        print ("/opt/opennms/bin/provision.pl category add '" + aces + "' " + node_id + " Production")
        print ("/opt/opennms/bin/provision.pl category add '" + aces + "' " + node_id + " Switches")
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " address1 '" + morada + "'")
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " zip " + cp)
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " city '" + localidade + "'")
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " modelNumber " + modelo)
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " serialNumber " + serial)
        print ("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " latitude " + latitude)
        print ("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " longitude " + longitude)
        print("/opt/opennms/bin/provision.pl asset add '" + aces + "' " + node_id + " country Portugal")
        print ("/opt/opennms/bin/provision.pl requisition import '" + aces + "'\n")
