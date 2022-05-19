#! /usr/bin/env python2.6
import sys, os
sys.path.insert(0,(os.path.expanduser('~')+'/.local/lib/python2.7/site-packages/'))
import paramiko
import signal, re, getpass, xlsxwriter, time
from optparse import OptionParser

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def connection_establishment(USER, PASS, host):
    try:
        print('Processing HOST: ' + host)
        client = paramiko.SSHClient()
        client.load_system_host_keys()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, 22, username=USER, password=PASS)
        channel = client.invoke_shell()
        while not channel.recv_ready():
            time.sleep(0.5)

        output = channel.recv(8192)
    except paramiko.AuthenticationException as error:
        print ('Authentication Error on host: ' + host)
        exit()
    except IOError as error:
        print (error)
        exit()

    return (channel, client)

def execute_command(command, channel, host):
    cbuffer = []
    data = ''

    channel.send(command)
    while True:
        if channel.recv_ready():
            data = channel.recv(1024)
            cbuffer.append(data)

        time.sleep(0.01)
        data_no_trails = data.strip()

        if len(data_no_trails) > 0: #and
            if data_no_trails.upper().endswith(host+'#'):
                break

    if channel.recv_ready():
        data = channel.recv(1000)
        cbuffer.append(data)

    rbuffer = ''.join(cbuffer)
    return rbuffer

def connection_teardown(client):
    client.close()

def get_user_password():
    sys.stdin = open('/dev/tty')
    USER = raw_input("Username:")
    PASS = getpass.getpass(prompt='Enter user password: ')
    return USER, PASS

def open_xls_for_write(filname):
    book = xlsxwriter.Workbook(filname)
    sheet = book.add_worksheet('Interfaces')
    sheet.write(0, 0, 'Interface')
    sheet.write(0, 1, 'Admin status')
    sheet.write(0, 2, 'Operational status')
    sheet.write(0, 3, 'Circuit reference')
    return book, sheet

def write_xls(sheet, interface, admin_status, oper_status, tl, row):
    sheet.write(row, 0, interface)
    sheet.write(row, 1, admin_status)
    sheet.write(row, 2, oper_status)
    sheet.write(row, 3, tl)

def close_xls_book(book):
    book.close()

def main():
    tl = re.compile(r'(?<=tl\=).+?(?=\:)')

    #create command line options menu
    usage = 'usage: %prog options [arg]'
    parser = OptionParser(usage)
    parser.add_option('-d', '--device', dest='device',
                            help='Specify device name')

    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
        parser.print_help()
        exit()

    username, password = get_user_password()
    channel, client = connection_establishment(username, password, options.device)
    execute_command('environment no more\n', channel, options.device.upper())
    ports = execute_command('show port | match ^[0-9] expression\n', channel, options.device.upper())
    ports_list = ports.split('\n')
    row = 1
    book, sheet = open_xls_for_write(options.device + '_interface_audit.xlsx')
    for i in range(len(ports_list) - 2):
        if re.search(r'^\d\/.{2}\d$', ports_list[i+1].split()[0]) != None:
            interface = ports_list[i+1].split()[0]
            print('Processing interface ' + interface)

            admin_state = ports_list[i+1].split()[1]
            port_state = ports_list[i+1].split()[3]

            description = execute_command('show port ' + interface + ' description | match ^[0-9] expression\n', channel, options.device.upper())
            m = tl.search(description)
            if m != None:
                write_xls(sheet, interface, admin_state, port_state, m.group(), row)
            else:
                write_xls(sheet, interface, admin_state, port_state, 'None', row)

            row += 1

    connection_teardown(client)
    close_xls_book(book)

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()
