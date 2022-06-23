#! /usr/bin/env python
import sys, signal, os, xlrd, subprocess, re
from optparse import OptionParser

# This script calls mdarepace script to create the removal and apply configuration. Apply can be used as rollback file
# It then creates an Apply_esat file with the configuration to be used after the migration.
# It takes as input the migration xls file which should consist of two sheets called optical and electrical (all lower)
# for the optical and copper ports respectively

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def call_mdareplace(mda, device):
    result = subprocess.call(
        ['mdareplace ' + device + ' ' + str(mda)],
        stdout=subprocess.PIPE,
        shell=True)

def read_from_mirgation_book(filename):
    book = xlrd.open_workbook(filename)
    optical = book.sheet_by_name('optical')
    electrical = book.sheet_by_name('electrical')

    return optical.col_slice(0), electrical.col_slice(0), \
           optical.col_slice(1), electrical.col_slice(1)

def replace_mda(mda, esat, src, dst, mda_file, esat_file):
    try:
        with open(mda_file, 'r') as sf:
            with open(esat_file, 'w+') as df:
                contents = sf.read()
                for i in range(1, len(src)):
                    source = src[i].value
                    destination = dst[i].value
                    contents = re.sub(re.escape(source)+r'\b', destination, contents)

                # For those that don't have connections/configuration
                contents = re.sub(re.escape(mda)+r'\/\d', esat+'/1', contents)
                df.write(contents)

    except IOError as e:
        print('Operation failed:' + e.strerror)
        exit()

def main():
    usage = 'usage: %prog options'
    parser = OptionParser(usage)

    parser.add_option('-f', '--file', dest='file',
                      help='XLS file name to load data from')
    parser.add_option('-d', '--device', dest='device',
                      help='Device name to connect to')
    parser.add_option('-e', '--esat',  dest='esat',
                      help='esat you migrate to')
    parser.add_option('-m', '--mda',  dest='mda',
                      help='mda you migrate from')

    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
       parser.print_help()
       exit()

    if not (options.esat and options.mda):
        print('You need to specify both mda (-m) and esat (-e)')
        exit()
    elif not options.device:
        print('You need to specify a device (-d)')
        exit()
    elif not options.file:
        print('You need to specify migration file (-f)')
        exit()

    print('Working on creating configuration files please hold...')
    #call_mdareplace(options.mda, options.device)
    optical_src, electrical_src, optical_dst, electrical_dst = read_from_mirgation_book(options.file)
    apply_file = '/usr/local/scripts/datafiles/MDACardReplace/' + options.device.lower() + '.' + options.mda.replace('/','_') + '.Apply'
    esat_file = os.path.expanduser('~') + '/' + options.device.lower() + '.' + options.esat + '.Apply'
    remove_file = '/usr/local/scripts/datafiles/MDACardReplace/' + options.device.lower() + '.' + options.mda.replace('/','_') + '.Remove'
    remove_esat_file = os.path.expanduser('~') + '/' + options.device.lower() + '.' + options.esat + '.Remove'

    replace_mda(options.mda, options.esat, optical_src, optical_dst, apply_file, esat_file)
    replace_mda(options.mda, options.esat, electrical_src, electrical_dst, apply_file, esat_file)
    replace_mda(options.mda, options.esat, optical_src, optical_dst, remove_file, remove_esat_file)
    replace_mda(options.mda, options.esat, electrical_src, electrical_dst, remove_file, remove_esat_file)

    print('The following files were created:')
    print(apply_file + '\n' + esat_file + '\n' + remove_file + '\n' + remove_esat_file + '\n')

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  #catch ctrl-c and call handler to terminate the script
    main()