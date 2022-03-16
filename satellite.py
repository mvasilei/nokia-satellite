#! /usr/bin/env python2.6
import sys, os, signal, re, getpass, time, xlrd, xlsxwriter
from optparse import OptionParser
from zipfile import ZipFile
from optparse import OptionParser
sys.path.insert(0,(os.path.expanduser('~')+'/.local/lib/python2.6/site-packages/'))
import paramiko

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def progress(count, total, status=''):
    bar_len = 60
    filled_len = int(round(bar_len * (count+1) / float(total)))

    percents = round(100.0 * (count+1) / float(total), 0)
    bar = '=' * filled_len + '-' * (bar_len - filled_len)

    sys.stdout.write('[%s] %s%s ...%s\r' % (bar, percents, '%', status))
    sys.stdout.flush()

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

def connection_teardown(client):
    client.close()

def execute_command(command, channel, host):
    cbuffer = []
    data = ''

    channel.send(command)
    while True:
        if channel.recv_ready():
            data = channel.recv(1000)
            cbuffer.append(data)

        time.sleep(0.02)
        data_no_trails = data.strip()

        if len(data_no_trails) > 0: #and
            if data_no_trails.upper().endswith(host+'#'):
                break
    if channel.recv_ready():
        data = channel.recv(1000)
        cbuffer.append(data)

    rbuffer = ''.join(cbuffer)
    return rbuffer

def get_user_password():
    sys.stdin = open('/dev/tty')
    USER = raw_input("Username:")
    PASS = getpass.getpass(prompt='Enter user password: ')
    return USER, PASS

def read_from_mirgation_book(filename, column):
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_name('optical')

    return sheet.col_slice(column,1)

def read_from_intemediate_book(filename, index):
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(index)

    return sheet.col(0), sheet.col(1), sheet.col(2)

def open_xls_for_write(filname):
    book = xlsxwriter.Workbook(filname)
    optical = book.add_worksheet('Optical')
    vprn = book.add_worksheet('VPRN')
    l2 = book.add_worksheet('L2')
    return book, optical, vprn, l2

def close_xls_book(book):
    book.close()

def write_source(worksheet, value1, value2, value3, row):
    worksheet.write(row, 0, value1)
    worksheet.write(row, 1, value2)
    worksheet.write(row, 2, value3)

def write_final_optical_values(worksheet, srcif, srcstatus, srclight, dstif, dststatus, dstlight, delta, count):
    row = count + 1
    worksheet.write(row, 0, srcif)
    worksheet.write(row, 1, srcstatus)
    worksheet.write(row, 2, srclight)
    worksheet.write(row, 3, dstif)
    worksheet.write(row, 4, dststatus)
    worksheet.write(row, 5, dstlight)
    srccell = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
    dstcell = xlsxwriter.utility.xl_rowcol_to_cell(row, 4)
    worksheet.write(row, 6,'=IF('+srccell+'='+dstcell+',"OK","FAILED")')
    srccell = xlsxwriter.utility.xl_rowcol_to_cell(row, 2)
    dstcell = xlsxwriter.utility.xl_rowcol_to_cell(row, 5)
    worksheet.write(row, 7,'=IF(AND(VALUE('+srccell+')-2<=VALUE('+dstcell+'),VALUE('+srccell+')+2>=VALUE('+dstcell+')),"OK","FAILED")')
    worksheet.write(row, 8, delta)


def write_final_service_values(worksheet, src_svc, src_sap, src_value, dst_svc, dst_sap, dts_value, count):
    row = count + 1
    worksheet.write(row, 0, src_svc)
    worksheet.write(row, 1, src_sap)
    worksheet.write(row, 2, src_value)
    worksheet.write(row, 3, dst_svc)
    worksheet.write(row, 4, dst_sap)
    worksheet.write(row, 5, dts_value)
    srccell = xlsxwriter.utility.xl_rowcol_to_cell(row, 2)
    dstcell = xlsxwriter.utility.xl_rowcol_to_cell(row, 5)
    worksheet.write(row, 6,'=IF('+srccell+'='+dstcell+',"OK","ATTENTION")')

def write_final_header(worksheet, value, column):
    worksheet.write(0, column, value)

def get_int_values(interfaces, host, channel):
    count = 0
    status = []
    light = []
    oper_status = re.compile(r'(?<=:).[A-Z,a-z]{2,4}')
    rx_light = re.compile(r'-*\d{1,2}.\d{1,2}')

    execute_command('environment no more\n', channel, host)

    for interface in interfaces:
        if interface.value != '':
            print('Collecting information for interface ' + interface.value)
            output = execute_command('show port ' + interface.value + ' | match "Oper State"\n', channel, host)
            m = oper_status.search(output.split('\n',3)[1])
            status.append(m.group())

            output = execute_command('show port ' + interface.value + ' optical | match "Rx Optical"\n', channel, host)
            m = rx_light.search(output.split('\n',3)[1])
            light.append(m.group())
            count += 1
    return status, light,

def get_service_values(interfaces, host, channel):
    sap_list = []
    svc_list = []
    arp = []
    svc_status = []
    svc_type = []
    service_type = re.compile(r'(?<=:).[A-Z,a-z]{1,5}')

    print('Collecting SAP details')
    saps = execute_command('show service sap-using\n', channel, host)
    sap_lines = saps.split('\n')
    count = 0
    for interface in interfaces:
        if interface.value != '':
            progress(count, len(interfaces))
            for i in range(1, len(sap_lines)-1):
                if interface.value+':' in sap_lines[i]:
                    sap = sap_lines[i].split()[0]
                    sap_list.append(sap)
                    svc_id = sap_lines[i].split()[1]
                    svc_list.append(svc_id)
                    output = execute_command('show service id ' + svc_id + ' base | match "Service Type|Oper State" expression\n', channel, host)
                    m = service_type.search(output.split('\n',4)[1])
                    svc_type.append(m.group())
                    if m.group() == ' VPRN':
                        output = execute_command('show service id ' + svc_id + ' arp sap ' + sap + '| match Dynamic\n', channel, host)
                        # sample command output == 100.125.85.18   70:df:2f:c3:42:50 Dynamic 00h44m44s VODART_70140_014* 3/1/2:11*
                        if len(output.split('\n')) > 2:
                            arp.append(output.split('\n',1)[1].split()[0])
                        else:
                            arp.append('No arp entry found')
                    else:
                            output = execute_command('show service id ' + svc_id + ' sap ' + sap + ' | match "Oper State"\n', channel, host)
                            svc_status.append(output.split('\n',3)[1].split(':')[2])
            count += 1

    return svc_type, svc_list, sap_list, arp, svc_status

def get_int_traffic(interfaces, host, channel):
    traffic_in = []
    traffic_out = []

    for interface in interfaces:
        if interface.value != '':
            output = execute_command('show port ' + interface.value + ' statistics | match ' + interface.value + '\n',
                                     channel, host)
            traffic_in.append(output.split('\n', 3)[1].split()[1])
            traffic_out.append(output.split('\n', 3)[1].split()[3])

    return traffic_in, traffic_out

def main():
    status = []
    light = []
    usage = 'usage: %prog options <device name>'
    parser = OptionParser(usage)

    parser.add_option('-f', '--file', dest='file',
                      help='XLS file name to load data from')
    parser.add_option('-d', '--device', dest='device',
                      help='Device name to connect to')
    parser.add_option('-c', '--check', action='store_true', dest='check',
                      help='Check interface and service status pre-migration')
    parser.add_option('-p', '--post', action='store_true', dest='post',
                      help='Check interface and service status post-migration')

    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
        parser.print_help()
        exit()

    if not (options.post or options.check):
        print('Specify either check or post operation')
        exit()
    elif options.post and not (options.file and options.device):
        print('Post checks require the use of all -f/-d/-p switches')
        exit()
    elif options.check and not (options.device and options.check):
        print('Pre checks require the use of all -f/-d/-c switches')
        exit()

    if options.check:
        if not os.path.exists(options.file):
            print('The file you specified doesn''t exist')
            exit()

        source = read_from_mirgation_book(options.file, 0)
        if len(source) > 0:
            book, optical, vprn, l2 = open_xls_for_write(options.device.upper() + '.xlsx')
            user, password = get_user_password()
            channel, client = connection_establishment(user, password, options.device)
            int_status, int_light = get_int_values(source, options.device, channel)
            svc_type, svc_id, sap, arp, svc_status = get_service_values(source, options.device, channel)
            connection_teardown(client)
            for i in range(len(int_status)):
                write_source(optical, source[i].value, int_status[i], int_light[i], i)

            vprn_count = 0
            l2_count = 0
            for i in range(len(svc_type)):
                if "VPRN" in svc_type[i]:
                    write_source(vprn, svc_id[i], sap[i], arp[vprn_count], vprn_count)
                    vprn_count += 1
                else:
                    write_source(l2, svc_id[i], sap[i], svc_status[l2_count], l2_count)
                    l2_count += 1
            close_xls_book(book)
            print('\nResults written in ' + options.device.upper() + '.xlsx')
            print('Please do NOT delete this file until after your run the post checks')
    elif options.post:
        destination = read_from_mirgation_book(options.file, 1) #CHANGE 0 to 1 when finish
        if len(destination) > 0:
            book, optical, vprn, l2 = open_xls_for_write(options.device.upper() + '_POST_MIGRATION.xlsx')
            user, password = get_user_password()
            channel, client = connection_establishment(user, password, options.device)
            int_status, int_light = get_int_values(destination, options.device, channel)
            dst_svc_type, dst_svc_id, dst_sap, dst_arp, dst_svc_status = get_service_values(destination, options.device, channel)

            dst_traffic_init_in, dst_traffic_init_out = get_int_traffic(destination, options.device, channel)
            print ('\nPausing for 60 seconds to calculate L2 svc traffic delta...')
            dst_traffic_final_in, dst_traffic_final_out = get_int_traffic(destination, options.device, channel)

            connection_teardown(client)
            source_int, source_status, source_light = read_from_intemediate_book(options.device.upper() + '.xlsx',0)
            src_vprn_id, src_vprn_sap, src_vprn_arp = read_from_intemediate_book(options.device.upper() + '.xlsx', 1)
            src_l2_id, src_l2_sap, src_l2_status = read_from_intemediate_book(options.device.upper() + '.xlsx', 2)
            column = 0
            for value in ['Src i/f','Src i/f status', 'Src Rx Level', 'Dst i/f', 'Dst i/f status', 'Dst Rx Level', 'I/f status check', 'Rx Level Check', 'Traffic passing']:
                write_final_header(optical, value, column)
                column += 1

            column = 0
            for value in ['Src Svc ID','Src SAP', 'Src ARP', 'Dst Svc ID', 'Dst SAP', 'Dst ARP', 'ARP check']:
                write_final_header(vprn, value, column)
                column += 1

            column = 0
            for value in ['Src Svc ID','Src SAP', 'Src Svc status', 'Dst Svc ID', 'Dst SAP', 'Dst Svc status', 'Svc Status']:
                write_final_header(l2, value, column)
                column += 1

            for i in range(len(int_status)):
                delta = 'ATTENTION'
                if ((float(dst_traffic_final_in[i]) - float(dst_traffic_init_in[i])) > 5) and ((float(dst_traffic_final_out[i]) - float(dst_traffic_init_out[i])) > 5):
                    delta = 'OK'
                write_final_optical_values(optical, source_int[i].value, source_status[i].value, source_light[i].value, destination[i].value,
                                           int_status[i], int_light[i], delta, i)

            vprn_count = 0
            l2_count = 0
            for i in range(len(dst_svc_type)):
                if "VPRN" in dst_svc_type[i]:
                    write_final_service_values(vprn, src_vprn_id[vprn_count].value, src_vprn_sap[vprn_count].value, src_vprn_arp[vprn_count].value,
                                                dst_svc_id[i], dst_sap[i], dst_arp[vprn_count], vprn_count)
                    vprn_count += 1
                else:
                    write_final_service_values(l2, src_l2_id[l2_count].value, src_l2_sap[l2_count].value, src_l2_status[l2_count].value,
                                               dst_svc_id[i], dst_sap[i], dst_svc_status[l2_count], l2_count)
                    l2_count += 1
            close_xls_book(book)
            print('Results written in ' + options.device.upper() + '_POST_MIGRATION.xlsx')

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()
