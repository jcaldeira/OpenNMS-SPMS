import sys
import ipaddress
import openpyxl


def validate_ip(ip):
    try:
        ipaddress.ip_address(ip)
    except:
        return False
    else:
        return True


def prep_env():
    excel = "C:\\\Users\\joao.caldeira.ext\\SPMS - Serviços Partilhados do Ministério da Saúde, EPE\\SPMS DSI RDIS - RIS Corporate - RIS2020 Cadastro\\RIS2020 - Cadastro.xlsx"
    chars_a_remover = ['á','à','ã','â','é','è','ê','í','ó','õ','ô','ú','û','ù','ç','s/n']
    chars_a_inserir = ['a','a','a','a','e','e','e','i','o','o','o','u','u','u','c','']

    try:
        wb = openpyxl.load_workbook(excel, data_only=True)
    except Exception as e:
        sys.exit(f"{e}\nImpossible to open excel file: {excel}")
    else:
        sheet = wb["RIS2020"]

    for row in sheet.iter_rows(min_row=3, values_only=True):
        ip = str(row[24])
        if validate_ip(ip):
            hostname = str(row[23])
            aces = str(row[28])
            site_id = hostname[0:4]
            if 'B01-SW01' not in hostname:
                if str(row[0]) == site_id:
                    entidade = str(row[1]).replace(" ","_")
                    entidade_upper = entidade.upper()
                    morada = str(row[4])
                    cp = str(row[5])
                    localidade = str(row[6])
                    latitude = str(row[7])
                    longitude = str(row[8])
                    local = str(row[3])

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

            print(f"/opt/opennms/bin/provision.pl service add '{aces}' {node_id} {ip} ICMP")
            print(f"/opt/opennms/bin/provision.pl service add '{aces}' {node_id} {ip} SNMP")
            print(f"/opt/opennms/bin/provision.pl service add '{aces}' {node_id} '{entidade}'")
            print(f"/opt/opennms/bin/provision.pl service add '{aces}' {node_id} '{entidade_upper}'")



if __name__ == "__main__":
    prep_env()
