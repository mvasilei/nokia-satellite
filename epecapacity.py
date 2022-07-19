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

def replace_mda(master, src, dst, cfg, device):
    # read device config file and replace the mda interface with the esat equivalent from the migration spreadsheet
    card = set()
    removed_cards = set()
    swapped_cards = set()

    sf = open(cfg, 'r')
    contents = sf.read()
    sf.close()

    # replace mda port from migration xls to its esat equivalent
    for i in range(1, len(src)):
        source = src[i].value
        destination = dst[i].value
        if source != '':
            contents = re.sub(re.escape(source)+r'\b', destination, contents)
        card.update(source.split('/', 1)[0])  # return a set with card listed as src interfaces

    book = xlrd.open_workbook(master)
    epe = book.sheet_by_name('EPE_SlotReport16072021')

    card_config = '\n        card-type imm-2pac-fp3\n\
        mda 1\n\
            mda-type p10-10g-sfp\n\
            sync-e\n\
            no shutdown\n\
        exit\n\
        mda 2\n\
            mda-type p10-10g-sfp\n\
            sync-e\n\
            no shutdown\n\
        exit\n\
        fp 1\n\
            ingress\n\
                mcast-path-management\n\
                    no shutdown\n\
                exit\n\
            exit\n\
        exit\n\
        no shutdown\n'

    # Go through the lines of master spreadsheet and if the whole card is removed then shut it and add it to the list of
    # cards for which you can remove reference to port configuration.
    # If the card is swapped with 2-PAC FP3 IMM then add the configuration to provision the new HW
    for i in range(epe.nrows):
        if epe.cell_value(i, 1).upper() == device.upper():
            for slot in card:
                if epe.cell_value(i, 16) != 'Yes' and int(epe.cell_value(i, 4)) == int(slot) and \
                        ('Daughter' not in epe.cell_value(i, 5) and 'SFM' not in epe.cell_value(i, 5)) and \
                        epe.cell_value(i, 17) != '':
                    regex = re.compile(r'(?<=\s{4}card\s' + re.escape(str(slot)) + r')[\s\S]+?(?=^\s{4}exit)',
                                       re.MULTILINE)
                    contents = re.sub(regex, r'\n        shutdown\n', contents)
                    removed_cards.update(slot)
                elif epe.cell_value(i, 16) == 'Yes' and int(epe.cell_value(i, 4)) == int(slot) and \
                        '2-PAC FP3 IMM' in epe.cell_value(i, 15):
                    regex = re.compile(r'(?<=\s{4}card\s' + re.escape(str(slot)) + r')[\s\S]+?(?=^\s{4}exit)',
                                       re.MULTILINE)
                    contents = re.sub(regex, card_config, contents)
                elif epe.cell_value(i, 16) != 'Yes' and int(epe.cell_value(i, 4)) == int(slot) and \
                        'Daughter' in epe.cell_value(i, 5) and epe.cell_value(i, 17) != '':
                    removed_cards.add(str(epe.cell_value(i, 5).split()[4]))
                elif epe.cell_value(i, 16) == 'Yes' and int(epe.cell_value(i, 4)) == int(slot) and \
                        'Daughter' in epe.cell_value(i, 5):
                    # if swapped and daughtercard
                    mda = epe.cell_value(i, 5).split()[4]
                    regex = re.compile(r'\s{4}port\s' + re.escape(str(mda)) + r'\/([1][1-9]|[2-9][0-9])[\s\S]+?\s{4}exit',
                                       re.MULTILINE) # match interfaces that are over x/x/11 and delete the config
                    contents = re.sub(regex, '', contents)
                    regex = re.compile(r'.*port\s' + re.escape(str(mda)) + r'\/([1][1-9]|[2-9][0-9])+?')
                    contents = re.sub(regex, '', contents)

    tmp = open('temp_' + device + '.cfg', 'w+')
    tmp.write(contents)
    tmp.close()

    return removed_cards

def delete_unused_ports(card, device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

            for slot in card:
                # delete ports from config that are of the removed card and aren't migrated
                if '/' in slot:
                    regex = re.compile(r'\s{4}port\s' + re.escape(str(slot)) + r'\/.+[\s\S]*?^\s{4}exit',
                                       re.MULTILINE)
                else:
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
    regex = re.compile(r'^\s{8}(access|network)[\s\S]+?^\s{8}exit', re.MULTILINE)
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

def interfaces(interface):
    return re.sub(r'(.*authentication\-key).*', r'\g<1> juniper', interface.group())

def fix_authentication_key(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()
            # extract port configuration and remove slop from esat interfaces
            regex = re.compile(r'(?<=interface\s\"port\-esat\-)[\s\S]+?(?=exit)', re.MULTILINE)
            contents = re.sub(regex, interfaces, contents)

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

def add_soft_repo(device):
    config = '        snmp\n\
            streaming\n\
                no shutdown\n\
            exit\n\
            packet-size 9216\n\
        exit\n\
        software-repository "7210-SAS-Sx-TiMOS-20.9.R3" create\n\
            description "7210-SAS-Sx-2009R3-Images"\n\
            primary-location "cf3:/7210-SAS-Sx-TiMOS-20.9.R3"\n\
        exit'

    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

        regex = re.compile(r'(^\s{8}snmp\n[\s\S]+?^\s{8}exit)', re.MULTILINE)
        contents = re.sub(regex, config, contents)

        config = '        lldp\n\
            no shutdown\n\
        exit'
        regex = re.compile(r'(^\s{8}lldp\n[\s\S]+?^\s{8}exit)', re.MULTILINE)
        contents = re.sub(regex, config, contents)

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
    port_config = ''
    book = xlrd.open_workbook(master)
    satellite_uplinks = book.sheet_by_name('Circuit References')

    config = 'echo "System Port-Topology Configuration"\n\
#--------------------------------------------------\n\
    system\n\
        port-topology\n'
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()

        for i in range(satellite_uplinks.nrows):
            if satellite_uplinks.cell_value(i, 0).upper() == device.upper():
                mdaport = str(satellite_uplinks.cell_value(i, 1)).lower()
                esat_ulink = str(satellite_uplinks.cell_value(i, 3)).lower()
                config = config + '            ' + mdaport + ' to ' + esat_ulink + ' create\n'
                port_config += '    ' + str(satellite_uplinks.cell_value(i, 1)).lower() + '\n\
        description "vf=4445:dt=bb:bw=10G:ph=10GE:st=act:tl=#VF#' + str(satellite_uplinks.cell_value(i, 4)).lower() + ':di='\
                           + device + '-' + satellite_uplinks.cell_value(i, 2) +'#' + satellite_uplinks.cell_value(i, 3) + '"\n\
        ethernet\n\
            dot1x\n\
                tunneling\n\
            exit\n\
            mode hybrid\n\
            encap-type dot1q\n\
            ssm\n\
                no shutdown\n\
            exit\n\
        exit\n\
        no shutdown\n\
    exit\n'
                # If the mda port that connects to uplink already exist remove the configuration
                regex = re.compile(r'(^\s{4}' + re.escape(mdaport) + r'\n[\s\S]+?^\s{4}exit)', re.MULTILINE)
                contents = re.sub(regex, '', contents)

        config = config + '        exit\n\
    exit\n'

        contents = re.sub(r'(echo "Port Configuration"\n#\-.*\n)', r'\g<1>' + port_config, contents)
        contents = re.sub(r'(echo "System Satellite phase 2 Configuration"\n)', config + r'\g<1>', contents)
        with open(device + '_R20.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def esat_init(device, master):
    book = xlrd.open_workbook(master)
    sheet = book.sheet_by_name('PE List')
    count = 1

    config = 'echo "System Satellite phase 1 Configuration"\n\
#--------------------------------------------------\n\
    system\n\
        satellite\n'
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 0).upper() == device.upper():
            for j in range(13,18):
                if sheet.cell_value(i, j) != '':
                    config += '            eth-sat ' + str(count) + ' create\n\
                description "Ethernet Satellite"\n\
                mac-address ' + str(sheet.cell_value(i, j)) +'\n\
                sat-type "es48-1gb-sfp"\n\
                software-repository "7210-SAS-Sx-TiMOS-20.9.R3"\n\
                no shutdown\n\
            exit\n '
                    count += 1

    config += '        exit\n\
    exit\n\
#--------------------------------------------------\n'

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

    mda = replace_mda(options.master, optical_src, optical_dst, original_cfg, options.device)
    mda.update(replace_mda(options.master, electrical_src, electrical_dst, 'temp_' + options.device + '.cfg', options.device))
    delete_unused_ports(mda, options.device)
    add_soft_repo(options.device)
    fix_bfd(options.device)
    fix_slope_mtu(options.device)
    fix_share_queueing(options.device)
    fix_authentication_key(options.device)
    local_user(options.device)
    sfm(options.device, options.master)
    remove_service_name(options.device)
    esat_init(options.device, options.master)
    esat_synce(options.device, options.master)
    esat_uplinks(options.device, options.master)

    #result = subprocess.call(['scp temp_' + options.device + '.cfg ' + options.device+':cf3:\\' + options.device + '_R20.cfg'],
    #    stdout=subprocess.PIPE,
    #    shell=True)


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()
