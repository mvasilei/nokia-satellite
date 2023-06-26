#! /usr/bin/env python
import sys, signal, os, xlsxwriter, subprocess, re
from optparse import OptionParser

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def vprn_lookup(vprn_id, worksheet):
    policy_regex = re.compile(r'vrf-[ei][mx]port.*', re.MULTILINE) #match import export policies
    vprn_regex = re.compile(r'(^\s{8}vprn\s' + re.escape(vprn_id) + '\s(\n|.)*?^\s{8}exit)', re.MULTILINE) #match entire vprn configuration
    bgp_group_regex = re.compile(r'^\s{12}bgp(\n|.)*?^\s{12}exit', re.MULTILINE) #match entire bgp configuration
    row = 1

    result = subprocess.Popen(
        ['rlist alcatel'],
        stdout=subprocess.PIPE,
        shell=True)
    output = result.stdout.read().split('\n')

    for line in output: #for each entry returned from the rlist command
        device = line.split(':')[0].strip().lower() #extract host name
        if os.path.exists('/curr/' + device.lower() + '.cfg'): #check config file exists then grep with vprn to see if relevant
            result = subprocess.Popen(
                ['egrep "vprn ' + vprn_id + ' " /curr/' + device + '.cfg'],
                stdout=subprocess.PIPE,
                shell=True)
            file = result.stdout.read()
            if file: #if vprn exists in the config file
                with open('/curr/' + device.lower() + '.cfg', 'r') as cfg:
                    config = cfg.read()
                    for match in vprn_regex.finditer(config):
                        policies = re.findall(policy_regex, match.group(1))

                        if policies:
                            worksheet.write(row, 0, vprn_id)
                            worksheet.write(row, 1, device)

                            mxroutes = re.findall(r'maximum-routes.+|mc-maximum-routes.+', match.group(1))

                            for mxroute in mxroutes:
                                if 'mc' in mxroute:
                                    worksheet.write(row, 5, mxroute.split()[1] + '+' + mxroute.split()[3])
                                else:
                                    worksheet.write(row, 4, mxroute.split()[1] + '+' + mxroute.split()[4])

                            for bgp in bgp_group_regex.finditer(match.group(1)):
                                for bgp_group in re.finditer(r'group.*', bgp.group()):
                                    worksheet.write(row, 6, bgp_group.group().split()[1])

                            for policy in policies:
                                if 'import' in policy:
                                    worksheet.write(row, 2, policy.split()[1])
                                else:
                                    worksheet.write(row, 3, policy.split()[1])
                            row += 1

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
    svc_sheet.write(0, 4, 'Max Routes')
    svc_sheet.write(0, 5, 'Max Multicast Routes')
    svc_sheet.write(0, 6, 'BGP Group')
    vprn_lookup(options.svc, svc_sheet)
    book.close()
    print ('Results file ' + 'EPE_SVC_'+options.svc+'_Policies.xlsx')

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
