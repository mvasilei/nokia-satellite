#! /usr/bin/env python
import sys, signal, os, xlsxwriter, subprocess, re
from optparse import OptionParser

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def vprn_lookup(vprn_id, worksheet):
    policy_regex = re.compile(r'vrf-[ei][mx]port.*', re.MULTILINE)
    row = 0
    result = subprocess.Popen(
        ['rlist alcatel'],
        stdout=subprocess.PIPE,
        shell=True)
    output = result.stdout.read().split('\n')
    vprn_regex = re.compile(r'(^\s{8}vprn\s' + re.escape(vprn_id) + '.*(\n|.)*?exit)', re.MULTILINE) #(^\s{8}vprn\s' + re.escape(vprn_id) + '.*(\n|.)*?\s{8}exit)
    for line in output:
        device = line.split(':')[0].strip().lower()
        if os.path.exists('/curr/' + device.lower() + '.cfg'):
            result = subprocess.Popen(
                ['egrep "vprn ' + vprn_id + ' " /curr/' + device + '.cfg'],
                stdout=subprocess.PIPE,
                shell=True)
            file = result.stdout.read()
            if file:
                with open('/curr/' + device.lower() + '.cfg', 'r') as cfg:
                    config = cfg.read()
                    if vprn_id in config:
                        for match in vprn_regex.finditer(config):
                            policy = re.findall(policy_regex, match.group(1))
                            if policy:
                                row += 1
                                worksheet.write(row, 0, vprn_id)
                                worksheet.write(row, 1, device)
                                for i in range(len(policy)):
                                    if 'import' in policy[i].split()[0]:
                                        worksheet.write(row, i + 2, policy[i].split()[1])
                                    else:
                                        if i == 1:
                                            worksheet.write(row, i + 2, policy[i].split()[1])
                                        else:
                                            worksheet.write(row, i + 3, policy[i].split()[1])

def open_xls_to_write(svc):
    book = xlsxwriter.Workbook('EPE_SVC_'+svc+'_Policies.xlsx')
    svc_sheet = book.add_worksheet('SVC')

    return book, svc_sheet

def main():
    usage = 'usage: %prog options'
    parser = OptionParser(usage)

    parser.add_option('-s', '--svc', dest='svc',
                      help='Service ID (VPRN#)')

    (options, args) = parser.parse_args()

    if not len(sys.argv) > 1:
       parser.print_help()
       exit()

    try:
        val = int(options.svc)
        if val < 0:
            print("The Service ID provided doesn't have the correct format")
            sys.exit()
    except ValueError:
        print("The Service ID provided doesn't have the correct format")

    book, svc_sheet = open_xls_to_write(options.svc)
    svc_sheet.write(0, 0, 'Service ID')
    svc_sheet.write(0, 1, 'Device')
    svc_sheet.write(0, 2, 'Import Policy')
    svc_sheet.write(0, 3, 'Export Policy')
    vprn_lookup(options.svc, svc_sheet)
    book.close()
    print ('Results file ' + 'EPE_SVC_'+options.svc+'_Policies.xlsx')


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
