from os import write
import netmiko
import re

from paramiko.ssh_exception import BadHostKeyException

device = {
    'device_type': 'cisco_ios',
    'host': '10.21.38.242',
    'username': 'jcaldeira',
    'password': 'CoN9m2%wVn^G',
    'ssh_config_file': '~/.ssh/config'
}

net_connect = netmiko.ConnectHandler(**device)
output = net_connect.send_command("show version")
# output = """0073-B01-SW01#sh ver
# Cisco IOS Software, C2960L Software (C2960L-UNIVERSALK9-M), Version 15.2(7)E2, RELEASE SOFTWARE (fc3)
# Technical Support: http://www.cisco.com/techsupport
# Copyright (c) 1986-2020 by Cisco Systems, Inc.
# Compiled Sat 14-Mar-20 15:38 by prod_rel_team

# ROM: Bootstrap program is C2960L boot loader
# BOOTLDR: C2960L Boot Loader (C2960L-HBOOT-M) Version 15.2(6r)E1, RELEASE SOFTWARE (fc1)

# 0073-B01-SW01 uptime is 1 week, 3 days, 21 hours, 12 minutes
# System returned to ROM by power-on
# System restarted at 13:29:48 PT Thu Apr 29 2021
# System image file is "flash:/c2960l-universalk9-mz.152-7.E2/c2960l-universalk9-mz.152-7.E2.bin"
# Last reload reason: power-on



# This product contains cryptographic features and is subject to United
# States and local country laws governing import, export, transfer and
# use. Delivery of Cisco cryptographic products does not imply
# third-party authority to import, export, distribute or use encryption.
# Importers, exporters, distributors and users are responsible for
# compliance with U.S. and local country laws. By using this product you
# agree to comply with applicable laws and regulations. If you are unable
# to comply with U.S. and local laws, return this product immediately.

# A summary of U.S. laws governing Cisco cryptographic products may be found at:
# http://www.cisco.com/wwl/export/crypto/tool/stqrg.html

# If you require further assistance please contact us by sending email to
# export@cisco.com.

# cisco WS-C2960L-48PS-LL (Marvell PJ4B (584) v7 (Rev 2)) processor (revision L0) with 524288K bytes of memory.
# Processor board ID FOC2450L9AK
# Last reset from Reload
# 2 Virtual Ethernet interfaces
# 52 Gigabit Ethernet interfaces
# The password-recovery mechanism is enabled.

# 512K bytes of flash-simulated non-volatile configuration memory.
# Base ethernet MAC Address       : EC:CE:13:36:B4:00
# Motherboard assembly number     : U58O342T01
# Power supply part number        : 341-0528-02
# Motherboard serial number       : FOC24471RRK
# Power supply serial number      : LIT24362MA7
# Model revision number           : L0
# Motherboard revision number     : 27
# Model number                    : WS-C2960L-48PS-LL
# Daughterboard assembly number   : 95.1642T01
# System serial number            : FOC2450L9AK
# Top Assembly Part Number        : 74-105858-01
# Top Assembly Revision Number    : L0
# Version ID                      : V01
# CLEI Code Number                : CMMWM00ARA
# Daughterboard revision number   : 08
# Hardware Board Revision Number  : 0x02


# Switch Ports Model                     SW Version            SW Image
# ------ ----- -----                     ----------            ----------
# *    1 52    WS-C2960L-48PS-LL         15.2(7)E2             C2960L-UNIVERSALK9-M


# Configuration register is 0xF

# 0073-B01-SW01#
# """
# print(f'################################\n{output}################################\n')

re_model = r'(Model number.*)'
re_serial_number = r'(System serial number.*)'
re_ios = r'(System image file is.*)'
hostname = f'{net_connect.find_prompt()[:-1].strip()}'

re_model_res = re.search(re_model, output)
re_model_res = re_model_res.group()
model = re_model_res.partition(":")[2].strip()

re_serial_number_res = re.search(re_serial_number, output)
re_serial_number_res = re_serial_number_res.group()
serial_number = re_serial_number_res.partition(":")[2].strip()

re_ios = re.search(re_ios, output)
re_ios = re_ios.group()
pos_ios = re_ios.partition(":")[2].find("/", 1)
ios = re_ios.partition(":")[2][pos_ios+1:-1].replace("/", "").strip()


# print('Before:')
# print(f're_modelo_res: {re_model_res}')
# print(f're_serial_number_res: {re_serial_number_res}')
# print(f're_ios: {re_ios}')

# print('\nAfter:')
print(f'Modelo: {model}')
print(f'Serial Number: {serial_number}')
print(f'IOS: {ios}')
print(f'hostname: {hostname}')

# with open('REPL.txt', 'wt') as f:
#     f.write(f'{model}\t{serial_number}\t{ios}\t{hostname}')
