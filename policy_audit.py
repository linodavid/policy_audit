import os
import glob
import xlsxwriter
from ciscoconfparse import CiscoConfParse


def ws_ingress_create():

    # create the ingress policies worksheet
    ws_ingress = wb.add_worksheet('Ingress Policies')

    # format columns and rows for ingress policies
    ws_ingress.set_column('A:A', 12, a_center)  # for hostname
    ws_ingress.set_column('B:B', 6, a_center)   # for IOS XR version
    ws_ingress.set_column('C:C', 28, a_left)    # for UNI
    ws_ingress.set_column('D:D', 50, a_left)    # for description
    ws_ingress.set_column('E:E', 50, a_left)    # for policy name
    ws_ingress.set_column('F:F', 20, a_left)    # for CIR/ CBS
    ws_ingress.set_column('G:G', 12, a_center)  # for COS
    ws_ingress.set_column('H:H', 12, a_center)  # for circuit type
    ws_ingress.set_column('I:I', 35, a_left)    # for remarks
    ws_ingress.set_row(0, 23, header)           # for header

    # write the column headers
    ws_ingress.write_row('A1', ['Hostname', 'IOS XR', 'Port', 'Description', 'Policy Name',
                                'CIR/ CBS', 'COS Setting', 'Circuit Type', 'Remarks'])

    return ws_ingress


def ws_egress_create():

    # create the parent policies worksheet
    ws_egress = wb.add_worksheet('Egress Policies')

    # Format columns and rows for parent policies
    ws_egress.set_column('A:A', 12, a_center)   # for hostname
    ws_egress.set_column('B:B', 6, a_center)    # for IOS XR version
    ws_egress.set_column('C:C', 18, a_left)     # for BE port
    ws_egress.set_column('D:D', 50, a_left)     # for description
    ws_egress.set_column('E:E', 20, a_left)   # for parent policy
    ws_egress.set_column('F:F', 20, a_left)   # for child policy
    ws_egress.set_column('G:G', 30, a_left)     # for remarks
    ws_egress.set_row(0, 23, header)            # for header

    # write the column headers
    ws_egress.write_row('A1', ['Hostname', 'IOS XR', 'BE Port', 'Description', 
                               'Parent Policy', 'Child Policy', 'Remarks'])

    return ws_egress


def ws_c_map_create():

    # create the child policies worksheet
    ws_c_map = wb.add_worksheet('Class Maps')

    # Format columns and rows for child policies
    ws_c_map.set_column('A:A', 12, a_center)    # for hostname
    ws_c_map.set_column('B:B', 6, a_center)     # for IOS XR version
    ws_c_map.set_column('C:C', 20, a_left)      # for child policy
    ws_c_map.set_column('D:D', 18, a_left)      # for class map
    ws_c_map.set_column('E:E', 35, a_left)      # for CIR/ CBS
    ws_c_map.set_column('F:F', 35, a_left)      # for remarks
    ws_c_map.set_row(0, 23, header)             # for header

    # Write some data headers.
    ws_c_map.write_row('A1', ['Hostname', 'IOS XR', 'Policy Map',
                              'Class Map', 'CIR/ CBS', 'Remarks'])

    return ws_c_map


def parse_ingress(ws):

    row = 1

    # open all .txt files in the current directory
    for filename in glob.glob('*.txt'):

        config = CiscoConfParse(filename, factory=True, syntax='ios')
        hostname = config.find_objects_dna(r'Hostname')[0].hostname
        ios_ver = config.find_lines('IOS')[0][len('!! IOS XR Configuration '):]
        print('... %s (IOS XR %s)' % (hostname, ios_ver))

        for ports in config.find_objects_w_child('^interface.+l2transport', 'service-policy input'):

            # write the hostname in column A
            ws.write(row, 0, hostname)
            # write the IOS XR version in column B
            ws.write(row, 1, ios_ver)
            # write the ingress port in column C
            port = ports.text[len('interface '):-len(' l2transport')]
            ws.write(row, 2, port)

            for line in ports.re_search_children(r'^ description '):
                description = line.text[len('description '):]
                ws.write(row, 3, description)

            for maps in ports.re_search_children('service-policy'):

                # write the ingress policy map in column D
                policy_map = maps.text[len(' service-policy input '):]
                ws.write(row, 4, policy_map)
                policy_maps = config.find_all_children(policy_map)
                policy_rate = [i for i in policy_maps if 'police rate' in i][0].strip()
                cos = [i for i in policy_maps if 'set cos' in i][0].strip()
                ws.write(row, 5, policy_rate)
                # write the COS setting in column F
                ws.write(row, 6, cos)

                # determine the circuit type based on the policy names
                if 'aovl' in policy_map.lower():
                    circuit_type = 'Adva Overlay'
                elif ('test' in policy_map.lower()) or ('mef' in policy_map.lower()):
                    circuit_type = 'Test'
                elif ('nid' in policy_map.lower()) or ('dcn' in policy_map.lower()):
                    circuit_type = 'DCN'
                else:
                    circuit_type = 'Customer'

                # write the circuit type in column G
                ws.write(row, 7, circuit_type)

                if circuit_type == 'Test':
                    if 'cos 4' in cos:
                        ws.write(row, 8, 'Ok')
                    else:
                        ws.write(row, 8,
                                 'Matched keyword test and/ or MEF.'
                                 'Please confirm if this is a test policy and delete. '
                                 'If not, change/ add set cos 4.')

                elif circuit_type == 'Adva Overlay':
                    if 'cos 6' in cos:
                        ws.write(row, 8, 'Ok')
                    else:
                        ws.write(row, 8, 'Change to set cos 6.')

                elif circuit_type == 'DCN':
                    if ('cos 4' in cos) or ('cos 7' in cos):
                        ws.write(row, 8, 'Ok')
                    else:
                        ws.write(row, 8, 'DCN circuit - change set to cos 7.')

                elif circuit_type == 'Customer':
                    if 'cos 4' in cos:
                        ws.write(row, 8, 'Ok')
                    else:
                        ws.write(row, 8, 'Change/ add set cos 4.')

                cos = ''
                row += 1

    print('... Completed %s ingress policies.' % row)

    # set the autofilter
    ws.autofilter(0, 0, row, 8)


def parse_egress(ws):

    row = 1

    # open all .txt files in the current directory
    for filename in glob.glob('*.txt'):

        config = CiscoConfParse(filename, factory=True, syntax='ios')
        hostname = config.find_objects_dna(r'Hostname')[0].hostname
        ios_ver = config.find_lines('IOS')[0][len('!! IOS XR Configuration '):]
        print('... %s (IOS XR %s)' % (hostname, ios_ver))

        # for be in config.find_objects_w_child(r'^interface.+Bundle[^.]+$', 'service-policy output'):
        for be in config.find_objects(r'^interface.+Bundle[^.]+$'):

            # write the hostname in column A
            ws.write(row, 0, hostname)
            # write the IOS XR version in column B
            ws.write(row, 1, ios_ver)
            # write the BE port in column C
            be_port = be.text[len('interface '):]
            ws.write(row, 2, be_port)

            # write the description in column D
            for line in be.re_search_children(r'^ description '):
                description = line.text[len('description '):]
                ws.write(row, 3, description)

            # pre-fill remarks for be ports without parent policies
            ws.write(row, 6, 'Missing parent and/ or child policy.')

            for parent in be.re_search_children('service-policy'):

                # write the parent policy in column E
                parent_policy = parent.text[len(' service-policy output '):]
                parent_policy1 = parent.text[len(' service-policy input '):]
                ws.write(row, 4, parent.text.strip())

                # write the child policy in column F
                try:
                    p_map = config.find_all_children(r'^policy-map ' + parent_policy + '$')
                    child_policy = [i for i in p_map if 'service-policy' in i][0][len('  service-policy '):]
                    ws.write(row, 5, child_policy)
                except Exception:
                    p_map = config.find_all_children(r'^policy-map ' + parent_policy1 + '$')
                    child_policy = [i for i in p_map if 'service-policy' in i][0][len('  service-policy '):]
                    ws.write(row, 5, child_policy)

                if 'input' in parent.text:
                    ws.write(row, 6, 'Policy incorrectly applied on ingress instead of egress.')
                elif be_port[len('Bundle-Ether'):] == parent_policy[len('policy_port_BE'):] \
                        and parent_policy[len('policy_port_BE'):] == child_policy[len('policy_BVID_BE'):]:
                    ws.write(row, 6, 'Ok')
                else:
                    ws.write(row, 6, 'Incorrect parent and/ or child policy.')

            row += 1

    print('... Completed %s egress policies.' % row)

    # set the autofilter
    ws.autofilter(0, 0, row, 6)


def parse_c_map(ws):

    row = 1

    # open all .txt files in the current directory
    for filename in glob.glob('*.txt'):

        config = CiscoConfParse(filename, factory=True, syntax='ios')
        hostname = config.find_objects_dna(r'Hostname')[0].hostname
        ios_ver = config.find_lines('IOS')[0][len('!! IOS XR Configuration '):]
        print('... %s (IOS XR %s)' % (hostname, ios_ver))

        for be in config.find_objects_w_child(r'^interface.+Bundle[^.]+$', 'service-policy output'):
            for parent in be.re_search_children('service-policy'):
                parent_policy = parent.text[len(' service-policy output '):]
                p_map = config.find_all_children(r'^policy-map ' + parent_policy + '$')
                child_policy = [i for i in p_map if 'service-policy' in i][0][len('  service-policy '):]
                for c_policy in config.find_objects('policy-map ' + child_policy + '$'):
                    for c_map in c_policy.re_search_children('class'):
                        # write the hostname in column A
                        ws.write(row, 0, hostname)
                        # write the IOS XR version in column B
                        ws.write(row, 1, ios_ver)
                        # write child policy in column C
                        ws.write(row, 2, child_policy)
                        # write class maps in column D
                        ws.write(row, 3, c_map.text[len(' class '):])
                        for police_rate in c_map.re_search_children('police'):
                            # write police rate in column E
                            ws.write(row, 4, police_rate.text.strip())
                        if 'vlan' in c_map.text.lower() or 'default' in c_map.text:
                            ws.write(row, 5, 'Ok')
                        else:
                            ws.write(row, 5, 'Incorrect class map.')
                        row += 1

    print('... Completed %s class maps.' % row)

    # set the autofilter
    ws.autofilter(0, 0, row, 5)


if __name__ == "__main__":

    regions = ['India', 'International']
    for region in regions:
        print('\nStarting audit of %s ASR running-configuration.' % region)
        path = 'config_files/' + region
        os.chdir(path)

        # create a workbook and add a worksheet
        wbName = 'Policy Audit - ' + region + '.xlsx'
        wb = xlsxwriter.Workbook(wbName)

        # add cell formats
        header = wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                                'align': 'center', 'font': 'Tahoma', 'font_size': 8})
        a_center = wb.add_format({'text_wrap': True, 'valign': 'vcenter',
                                  'align': 'center', 'font': 'Tahoma', 'font_size': 8})
        a_left = wb.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'left',
                                'indent': 1, 'font': 'Tahoma', 'font_size': 8})

        print('\nParsing %s ASR ingress policies.' % region)
        parse_ingress(ws_ingress_create())

        print('\nParsing %s ASR egress policies.' % region)
        parse_egress(ws_egress_create())

        print('\nParsing %s ASR class maps.' % region)
        parse_c_map(ws_c_map_create())

        wb.close()
        os.chdir('../..')

    for region in regions:
        os.startfile(os.getcwd() + '/config_files/' + region + '/Policy Audit - ' + region + '.xlsx')
