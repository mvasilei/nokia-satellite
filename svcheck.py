#! /usr/bin/env python
import sys, signal, os, csv, subprocess, re
from optparse import OptionParser

def signal_handler(sig, frame):
    print('Exiting gracefully Ctrl-C detected...')
    sys.exit()

def vprn_lookup(vprn_id, csvwriter, rlist_group):
    policy_regex = re.compile(r'vrf-[ei][mx]port.*', re.MULTILINE) #match import export policies
    vprn_regex = re.compile(r'(^\s{8}vprn\s' + re.escape(vprn_id) + '\s(\n|.)*?^\s{8}exit)', re.MULTILINE) #match entire vprn configuration
    bgp_group_regex = re.compile(r'^\s{12}bgp(\n|.)*?^\s{12}exit', re.MULTILINE) #match entire bgp configuration
    row = 1

    result = subprocess.Popen(
        ['rlist ' + rlist_group],
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
                            mc_mx_routes =''
                            mx_routes = ''
                            bgp_policy = ''
                            import_policy = ''
                            export_policy = ''

                            mxroutes = re.findall(r'maximum-routes.+|mc-maximum-routes.+', match.group(1))

                            for mxroute in mxroutes:
                                if 'mc' in mxroute:
                                    mc_mx_routes = mxroute.split()[1] + '+' + mxroute.split()[3]
                                else:
                                    mx_routes =  mxroute.split()[1] + '+' + mxroute.split()[4]

                            for bgp in bgp_group_regex.finditer(match.group(1)):
                                for bgp_group in re.finditer(r'group.*', bgp.group()):
                                    bgp_policy = bgp_group.group().split(' ',1)[-1]

                            for policy in policies:
                                if 'import' in policy:
                                    import_policy = policy.split(' ',1)[-1]
                                else:
                                    export_policy = policy.split(' ',1)[-1]

                            csvwriter.writerow([vprn_id, device, import_policy, export_policy, mx_routes, mc_mx_routes, bgp_policy])

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

    with open('EPE_RA_SVC_'+options.svc+'_Policies.csv', 'w') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter=',',
                               quotechar='|')
        csvwriter.writerow(['Service ID', 'Device', 'Import Policy', 'Export Policy', 'Max Routes',
                            'Max Multicast Routes', 'BGP Group'])
        vprn_lookup(options.svc, csvwriter,'alcatel')
        vprn_lookup(options.svc, csvwriter,'UKI-RA')
        print ('Results file ' + 'EPE_RA_SVC_'+options.svc+'_Policies.csv')

if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)  # catch ctrl-c and call handler to terminate the script
    main()
