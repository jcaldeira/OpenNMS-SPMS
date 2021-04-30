import ipaddress
import sys
import datetime
from typing import Counter
import openpyxl
from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import _validate_archive


def main():
    excel = "C:\\Users\\joao.caldeira.ext\\SPMS - Serviços Partilhados do Ministério da Saúde, EPE\\SPMS DSI RDIS - RIS Corporate - RIS2020 Cadastro\\RIS2020 - Cadastro.xlsx"
    chars_a_remover = ['á','à','ã','â','é','è','ê','í','ó','õ','ô','ú','û','ù','ç','s/n']
    chars_a_inserir = ['a','a','a','a','e','e','e','i','o','o','o','u','u','u','c','']

    try:
        wb = openpyxl.load_workbook(excel, data_only=True)
    except Exception as e:
        sys.exit(f"{e}\nImpossible to open excel file: {excel}")
    else:
        sheet = wb["RIS2020"]

    counter = 0
    excluded_ids = ['0417']
    excluded_ips = ['10.34.6.241', '10.11.20.241', '10.13.200.241', '10.13.200.242']
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if type(row[26]) == datetime.datetime:
        # if row[26] is not None:
            data_presi = row[26].date()
            ip = str(row[24])
            hostname = str(row[23])
            site_id = hostname[0:4]
            if validate_ip(ip) and validate_date_inter(data_presi) and site_id not in excluded_ids and ip not in excluded_ips:
            # if validate_ip(ip) and hostname == '0073-B01-SW01':
                if str(row[0]) == site_id:
                    entidade, morada, cp, localidade, latitude, longitude, requisition, modelo, serial, node_id = site_info(chars_a_remover, chars_a_inserir, row)

                    print(f"/opt/opennms/bin/provision.pl service add '{requisition}' {node_id} {ip} ICMP")
                    print(f"/opt/opennms/bin/provision.pl service add '{requisition}' {node_id} {ip} SNMP")
                    print(f"/opt/opennms/bin/provision.pl category remove '{requisition}' {node_id} '{entidade}'")
                    print(f"/opt/opennms/bin/provision.pl category add '{requisition}' {node_id} '{entidade.upper()}'")
                    print(f"/opt/opennms/bin/provision.pl category add '{requisition}' {node_id} 'Production'")
                    print(f"/opt/opennms/bin/provision.pl category add '{requisition}' {node_id} 'Switches'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} address1 '{morada}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} zip '{cp}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} city '{localidade}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} modelNumber '{modelo}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} serialNumber '{serial}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} latitude '{latitude}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} longitude '{longitude}'")
                    print(f"/opt/opennms/bin/provision.pl asset add '{requisition}' {node_id} country Portugal")
                    print(f"/opt/opennms/bin/provision.pl requisition import {requisition}\n\n")

                    counter += 1

                else:
                    print(f"IDs do not match {site_id} (on switch name)")

            else:
                continue
                print(f"Invalid IP: {ip}")

    print(f'Generated {counter} configurations')


def site_info(chars_a_remover, chars_a_inserir, row):
    entidade = str(row[1]).replace(" ","_")
    morada = str(row[4])
    cp = str(row[5])
    localidade = str(row[6])
    latitude = str(row[7])
    longitude = str(row[8])
    requisition = str(row[28])
    modelo = str(row[20])
    serial = str(row[21])
    node_id = str(row[27])
    for i in range(0,len(chars_a_remover)):
        morada = morada.replace(chars_a_remover[i],chars_a_inserir[i])
        localidade = localidade.replace(chars_a_remover[i], chars_a_inserir[i])
    return entidade, morada, cp, localidade, latitude, longitude, requisition, modelo, serial, node_id



def validate_date_inter(date):
    low_date_lim = datetime.date(2021, 4, 27)
    high_date_lim = datetime.date(2021, 4, 29)
    return low_date_lim <= date <= high_date_lim



def validate_ip(ip):
    try:
        ipaddress.ip_address(ip)
    except:
        return False
    else:
        return True



if __name__ == "__main__":
    main()
