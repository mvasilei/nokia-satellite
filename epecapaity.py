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

        # delete ports from config that are of the removed mda and aren't migrated
        regex = re.compile(r'\s{4}port\s' + re.escape(str(slot)) + r'\/.+\/.+[\s\S]*?^\s{4}exit',
                       re.MULTILINE)
        contents = re.sub(regex, '', contents)

    tmp = open('temp_' + device + '.cfg', 'w+')
    tmp.write(contents)
    tmp.close()

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
    no_slope = re.sub(r'.*slope.*', '', port.group())
    return no_slope

def fix_slope(device):
    try:
        with open('temp_' + device + '.cfg', 'r+') as sf:
            contents = sf.read()
            # extract port configuration and remove slop from esat interfaces
            regex = re.compile(r'(^\s{4}port\sesat[\s\S]+?^\s{4}exit)', re.MULTILINE)
            contents = re.sub(regex, slope, contents)

        with open('temp_' + device + '.cfg', 'w') as df:
            df.write(contents)
    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def esat_sync():
    pass

def esat_uplinks():
    pass

def esat_init():
    pass

def sfm():
    pass

def tacacs_user (device):
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

            # add user
            contents = re.sub(r'(\s{12}user.*)', account + r'\g<1>', contents, 1)
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


    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
       parser.print_help()
       exit()

    if not options.device:
        print('You need to specify a device (-d)')
        exit()
    elif not options.file:
        print('You need to specify migration file (-f)')
        exit()

    print('Working on creating configuration files please hold...')
    optical_src, electrical_src, optical_dst, electrical_dst = read_from_mirgation_book(options.file)
    original_cfg = '/curr/' + options.device.lower() + '.cfg'

    replace_mda(optical_src, optical_dst, original_cfg, options.device)
    replace_mda(electrical_src, electrical_dst, 'temp_' + options.device + '.cfg', options.device)
    fix_bfd(options.device)
    fix_slope(options.device)
    tacacs_user(options.device)

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()
