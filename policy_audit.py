import os
import glob
import xlsxwriter

# set the path for either Intl or India config files
path_dir = 'Intl'
path = 'config_files/' + path_dir
os.chdir(path)

# create a workbook and add a worksheet
wbName = 'Policy Audit - ' + path_dir + '.xlsx'
wb = xlsxwriter.Workbook(wbName)

# add cell formats
header = wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                        'align': 'center', 'font': 'Tahoma', 'font_size': 8})
a_center = wb.add_format({'text_wrap': True, 'valign': 'vcenter',
                          'align': 'center', 'font': 'Tahoma', 'font_size': 8})
a_left = wb.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'left',
                        'indent': 1, 'font': 'Tahoma', 'font_size': 8})


def ws_ingress_create():

    # create the ingress policies worksheet
    ws_ingress = wb.add_worksheet('Ingress Policies')

    # format columns and rows for ingress policies
    ws_ingress.set_column('A:A', 12, a_center)   # for hostname
    ws_ingress.set_column('B:B', 40, a_left)     # for policy name
    ws_ingress.set_column('C:C', 18, a_center)   # for CIR/ CBS
    ws_ingress.set_column('D:D', 12, a_center)   # for COS
    ws_ingress.set_column('E:E', 12, a_center)   # for circuit type
    ws_ingress.set_column('F:F', 35, a_left)     # for remarks
    ws_ingress.set_row(0, 23, header)            # for header

    # write the column headers
    ws_ingress.write_row('A1', ['Hostname', 'Policy Name',
                                'CIR/ CBS', 'COS Setting', 'Circuit Type', 'Remarks'])

    return ws_ingress


def ws_parent_create():

    # create the parent policies worksheet
    ws_parent = wb.add_worksheet('Parent Policies')

    # Format columns and rows for parent policies
    ws_parent.set_column('A:A', 12, a_center)   # for hostname
    ws_parent.set_column('B:B', 20, a_left)     # for bundle-ether port
    ws_parent.set_column('C:C', 20, a_center)   # for policy name
    ws_parent.set_column('D:D', 30, a_left)   # for remarks
    ws_parent.set_row(0, 23, header)            # for header

    # write the column headers
    ws_parent.write_row('A1', ['Hostname', 'BE Port', 'Port Policy', 'Remarks'])

    return ws_parent


def ws_child_create():

    # create the child policies worksheet
    ws_child = wb.add_worksheet('Child Policies')

    # Format columns and rows for child policies
    ws_child.set_column('A:A', 12, a_center)   # for hostname
    ws_child.set_column('B:B', 20, a_left)     # for policy name
    ws_child.set_column('C:C', 13, a_center)   # for class map
    ws_child.set_column('D:D', 18, a_center)   # for CIR/ CBS
    ws_child.set_column('E:E', 35, a_left)     # for remarks
    ws_child.set_row(0, 23, header)            # for header

    # Write some data headers.
    ws_child.write_row('A1', ['Hostname', 'Policy Name',
                              'Class Map', 'CIR/ CBS', 'Remarks'])
    
    return ws_child


def parse_ingress(ws):
    
    row = 1
    
    # 0pen all .txt files in the current directory
    for filename in glob.glob('*.txt'):

        # read the contents of the text file
        with open(filename, 'r') as f:

            # each file uses the ASR hostname in the filename
            hostname = filename.split('.')

            # parse each line
            for line in f:

                # exclude the egress policy maps
                if ('policy_port' in line.lower()) \
                        or ('policy_bvid' in line.lower()) \
                        or ('policy_svid' in line.lower()) \
                        or ('policy_child' in line.lower()) \
                        or ('policy_parent' in line.lower()) \
                        or ('egress' in line.lower()):

                    for line in f:
                        
                        # skip all config lines of the egress policy maps
                        if 'end-policy-map' not in line:
                            next(iter(line))
                        else:
                            break

                else:

                    # parse the ingress policy maps
                    if 'policy-map ' in line:

                        # write the hostname in column A
                        ws.write(row, 0, hostname[0])
                        # write the policy names in column B
                        ws.write(row, 1, line[11:])

                        # determine the circuit type based on the policy names
                        if 'aovl' in line.lower():
                            circuit_type = 'Adva Overlay'
                        elif ('test' in line.lower()) or ('mef' in line.lower()):
                            circuit_type = 'Test'
                        elif ('nid' in line.lower()) or ('dcn' in line.lower()):
                            circuit_type = 'DCN'
                        else:
                            circuit_type = 'Customer'

                        # write the circuit type in column E
                        ws.write(row, 4, circuit_type)

                    # Determine and write the current CIR and CBS in column C
                    elif 'police rate' in line:
                        ws.write(row, 2, line.strip())

                    # Determine and write the current CoS setting in column D
                    elif 'set cos' in line:
                        cos = line.strip()
                        ws.write(row, 3, cos)

                    # at end of policy map, check if settings are per the guidelines
                    elif 'end-policy-map' in line:

                        if circuit_type == 'Test':
                            if 'cos 4' in cos:
                                ws.write(row, 5, 'Ok')
                            else:
                                ws.write(row, 5,
                                         'Matched keyword test and/ or MEF.'
                                         'Please confirm if this is a test policy and delete. '
                                         'If not, change/ add set cos 4.')

                        elif circuit_type == 'Adva Overlay':
                            if 'cos 6' in cos:
                                ws.write(row, 5, 'Ok')
                            else:
                                ws.write(row, 5, 'Change to set cos 6.')

                        elif circuit_type == 'DCN':
                            if ('cos 4' in cos) or ('cos 7' in cos):
                                ws.write(row, 5, 'Ok')
                            else:
                                ws.write(row, 5, 'DCN circuit - change set to cos 7.')

                        elif circuit_type == 'Customer':
                            if 'cos 4' in cos:
                                ws.write(row, 5, 'Ok')
                            else:
                                ws.write(row, 5, 'Change/ add set cos 4.')

                        cos = ''
                        row += 1

    # set the autofilter
    ws.autofilter(0, 0, row, 5)


def parse_parent(ws):

    row = 1
    port_policy_start = 'service-policy'

    # 0pen all .txt files in the current directory
    for filename in glob.glob('*.txt'):

        # read the contents of the text file
        with open(filename, 'r') as f:

            # each file uses the ASR hostname in the filename
            hostname = filename.split('.')

            # parse each line
            for line in f:

                # parse the parent policy maps for each BE port
                if ('interface Bundle-Ether' in line) and ('.' not in line):

                    # write the hostname in column A
                    ws.write(row, 0, hostname[0])
                    # write the BE port in column B
                    be_port = line[10:].strip()

                    if 'face' in be_port:
                        continue
                    else:
                        ws.write(row, 1, be_port)

                        for line in f:

                            if 'instance' in line:
                                break
                            elif port_policy_start in line:
                                port_policy = line.strip()
                                ws.write(row, 2, port_policy)
                                if 'input' in port_policy:
                                    ws.write(row, 3, 'Input policy applied instead of output policy.')
                                elif be_port[12:] == port_policy[36:]:
                                    ws.write(row, 3, 'Ok')
                                else:
                                    ws.write(row, 3, 'Incorrect parent policy.')
                                row += 1
                                break
                            elif '!' in line:
                                ws.write(row, 2, 'NA')
                                ws.write(row, 3, 'Missing port policy.')
                                row += 1
                                break
                            else:
                                continue

                else:
                    continue

    # set the autofilter
    ws.autofilter(0, 0, row, 3)


if __name__ == "__main__":

    parse_ingress(ws_ingress_create())
    parse_parent(ws_parent_create())

    # # Set the autofilter
    #
    # ws_parent.autofilter(0, 0, row_parent, 5)
    # ws_child.autofilter(0, 0, row_child, 4)

    wb.close()
    os.startfile(wbName)
