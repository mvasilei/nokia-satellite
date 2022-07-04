#! /usr/bin/env python
import sys, signal, os, xlrd, subprocess, re
from optparse import OptionParser

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def read_from_mirgation_book(filename):
    book = xlrd.open_workbook(filename)
    optical = book.sheet_by_name('optical')
    electrical = book.sheet_by_name('electrical')

    return optical.col_slice(0), electrical.col_slice(0), \
           optical.col_slice(1), electrical.col_slice(1)

def replace_mda(src, dst, cfg, device):
    # read device config file and replace the mda interface with the esat equivalent from the migration spreadsheet
    mda = set()
    not_migrated = 0

    sf = open(cfg, 'r')
    contents = sf.read()
    sf.close()

    # replace mda port from migration xls to its esat equivalent
    for i in range(1, len(src)):
        source = src[i].value
        destination = dst[i].value
        if source != '':
            contents = re.sub(re.escape(source)+r'\b', destination, contents)
        mda.update(source.split('/', 1)[0])  # return a set with MDAs listed as src interfaces

    # shutdown the migrated mda
    for slot in mda:
        regex = re.compile(r'(?<=\s{4}card\s' + re.escape(str(slot)) + r')[\s\S]+?(?=^\s{4}exit)',
                           re.MULTILINE)
        contents = re.sub(regex, r'\n        shutdown\n', contents)

    tmp = open('temp_' + device + '.cfg', 'w+')
    tmp.write(contents)
    tmp.close()

    return mda

def delete_unused_ports(mda, device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

            for slot in mda:
                # delete ports from config that are of the removed mda and aren't migrated
                regex = re.compile(r'\s{4}port\s' + re.escape(str(slot)) + r'\/.+\/.+[\s\S]*?^\s{4}exit',
                                   re.MULTILINE)
                contents = re.sub(regex, '', contents)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def fix_bfd(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()
            # append type cpm-np to interface bfd config
            regex = re.compile(r'(.*bfd.*receive.*multiplier.\d+(?!.))')
            contents = re.sub(regex, r'\g<1> type cpm-np', contents)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def slope(port):
    regex = re.compile(r'^\s{8}access[\s\S]+?^\s{8}exit', re.MULTILINE)
    no_slope = re.sub(regex, '', port.group())
    return no_slope

def mtu(port):
    correct_mtu = re.sub(r'mtu 9212', 'mtu 9208', port.group())
    return correct_mtu

def fix_slope_mtu(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()
            # extract port configuration and remove slop from esat interfaces
            regex = re.compile(r'(^\s{4}port\sesat[\s\S]+?^\s{4}exit)', re.MULTILINE)
            contents = re.sub(regex, slope, contents)
            contents = re.sub(regex, mtu, contents)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def queuing(sap):
    no_queuing = re.sub(r'shared\-queuing', '', sap.group())
    print no_queuing
    return no_queuing

def fix_share_queueing(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()
            # extract port configuration and remove slop from esat interfaces
            regex = re.compile(r'(^\s{12,16}sap\sesat[\s\S]+?\s{12,16}exit)', re.MULTILINE)
            contents = re.sub(regex, queuing, contents)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def local_user (device):
    account = '            user "nokia"\n\
                password "$2y$10$6iQeEVoIn1LukUrojZZ5A.wAgiiewH3/we.p/1hftpm5bOvpKxBfG"\n\
                access console ftp\n\
                console\n\
                    member "default"\n\
                    member "administrative"\n\
                exit\n\
            exit\n'
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

            # add local user
            contents = re.sub(r'(\s{12}user.*)', account + r'\g<1>', contents, 1)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def remove_service_name(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

            # remove service name references from the config
            contents = re.sub(r'.*service-name.*', '', contents, )

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def esat_synce(device, master):
    satellite_set = set()
    book = xlrd.open_workbook(master)
    satellite_uplinks = book.sheet_by_name('Circuit References')

    for i in range(satellite_uplinks.nrows):
        if satellite_uplinks.cell_value(i, 0).upper() == device.upper():
            satellite_set.update(satellite_uplinks.cell_value(i, 2).split('-')[1])

    config = 'echo "System Satellite phase 2 Configuration"\n\
#--------------------------------------------------\n\
    system\n\
        satellite\n'

    for satellite in satellite_set:
        config +='            eth-sat ' + satellite + '\n\
                sync-e\n\
                port-map esat-' + satellite + '/1/1 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/2 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/3 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/4 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/5 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/6 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/7 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/8 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/9 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/10 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/11 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/12 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/13 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/14 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/15 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/16 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/17 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/18 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/19 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/20 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/21 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/22 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/23 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/24 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/25 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/26 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/27 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/28 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/29 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/30 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/31 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/32 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/33 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/34 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/35 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/36 primary esat-' + satellite + '/1/u1 secondary esat-' + satellite + '/1/u2\n\
                port-map esat-' + satellite + '/1/37 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/38 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/39 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/40 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/41 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/42 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/43 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/44 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/45 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/46 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/47 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
                port-map esat-' + satellite + '/1/48 primary esat-' + satellite + '/1/u2 secondary esat-' + satellite + '/1/u1\n\
            exit\n'

    config += '        exit\n\
    exit\n\
#--------------------------------------------------\n'

    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

        contents = re.sub(r'(echo \"Port Configuration\"\n)', config + r'\g<1>', contents)
        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def esat_uplinks(device, master):
    book = xlrd.open_workbook(master)
    satellite_uplinks = book.sheet_by_name('Circuit References')

    config = 'echo "System Port-Topology Configuration"\n\
#--------------------------------------------------\n\
    system\n\
        port-topology\n'

    for i in range(satellite_uplinks.nrows):
        if satellite_uplinks.cell_value(i, 0).upper() == device.upper():
            mdaport = str(satellite_uplinks.cell_value(i, 1)).lower()
            esat_ulink = str(satellite_uplinks.cell_value(i, 3)).lower()
            config = config + '            ' + mdaport + ' to ' + esat_ulink + ' create\n'

    config = config + '        exit\n\
    exit\n\
#--------------------------------------------------\n'

    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

        contents = re.sub(r'(echo "System Satellite phase 2 Configuration"\n)', config + r'\g<1>', contents)
        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def esat_init(device, master):
    book = xlrd.open_workbook(master)
    satellite_list = ['1', '2']
    mac = '8c:83:df:72:a3:9e'

    print 'FIX ME esat_init()'

    config = 'echo "System Satellite phase 1 Configuration"\n\
#--------------------------------------------------\n\
    system\n\
        satellite\n'

    for satellite in satellite_list:
        config += '            eth-sat ' + satellite + ' create\n\
                description "Ethernet Satellite"\n\
                mac-address ' + mac +'\n\
                sat-type "es48-1gb-sfp"\n\
                no shutdown\n\
            exit\n '

    config += '        exit\n\
    exit\n\
# --------------------------------------------------\n'

    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

        contents = re.sub(r'(echo \"System Security Configuration\"\n)', config + r'\g<1>', contents)
        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def sfm(device, master):
    book = xlrd.open_workbook(master)
    epe = book.sheet_by_name('EPE_SlotReport16072021')

    for i in range(epe.nrows):
        if epe.cell_value(i, 1).upper() == device.upper():
            if epe.cell_value(i, 15) == 'SFM6':
                try:
                    with open('temp_' + device + '.cfg', 'r+') as sf:
                        contents = sf.read()

                    contents = re.sub(r'sfm\-type m\-sfm5\-12', 'sfm-type m-sfm6-7/12', contents)
                    with open('temp_' + device + '.cfg', 'w') as df:
                        df.write(contents)
                except IOError as e:
                    print('Operation failed:' + e.strerror)
                    exit()

def main():
    usage = 'usage: %prog options'
    parser = OptionParser(usage)

    parser.add_option('-f', '--file', dest='file',
                      help='migration XLS file to load data from')
    parser.add_option('-d', '--device', dest='device',
                      help='Device name to connect to')
    parser.add_option('-m', '--master', dest='master',
                      help='Device name to connect to')


    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
       parser.print_help()
       exit()

    if not options.device or not options.file or not options.master:
        print('All options are compulsory')
        parser.print_help()
        exit()

    print('Working on creating configuration file please hold...')
    optical_src, electrical_src, optical_dst, electrical_dst = read_from_mirgation_book(options.file)
    original_cfg = '/curr/' + options.device.lower() + '.cfg'

    mda = replace_mda(optical_src, optical_dst, original_cfg, options.device)
    mda.update(replace_mda(electrical_src, electrical_dst, 'temp_' + options.device + '.cfg', options.device))
    delete_unused_ports(mda, options.device)
    fix_bfd(options.device)
    fix_slope_mtu(options.device)
    fix_share_queueing(options.device)
    local_user(options.device)
    sfm(options.device, options.master)
    remove_service_name(options.device)
    esat_init(options.device, options.master)
    esat_synce(options.device, options.master)
    esat_uplinks(options.device, options.master)


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()
