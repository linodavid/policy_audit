import os
import glob
import xlsxwriter

# Create a workbook and add a worksheet
# wbName = input('Enter new file name without extension: ')
# wbName = wbName + '.xlsx'
wbName = 'test.xlsx'
wb = xlsxwriter.Workbook(wbName)
ws_ingress = wb.add_worksheet('Ingress Policies')
ws_parent = wb.add_worksheet('Parent Policies')
ws_child = wb.add_worksheet('Child Policies')

# Add cell formats
header = wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                        'align': 'center', 'font': 'Tahoma', 'font_size': 8})
a_center = wb.add_format({'text_wrap': True, 'valign': 'vcenter',
                          'align': 'center', 'font': 'Tahoma', 'font_size': 8})
a_left = wb.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'left',
                        'indent': 1, 'font': 'Tahoma', 'font_size': 8})

# Format columns and rows for ingress policies
ws_ingress.set_column('A:A', 12, a_center)   # for hostname
ws_ingress.set_column('B:B', 40, a_left)     # for policy name
ws_ingress.set_column('C:C', 18, a_center)   # for CIR/ CBS
ws_ingress.set_column('D:D', 12, a_center)   # for COS
ws_ingress.set_column('E:E', 12, a_center)   # for circuit type
ws_ingress.set_column('F:F', 35, a_left)     # for remarks
ws_ingress.set_row(0, 23, header)            # for header

# Format columns and rows for parent policies
ws_parent.set_column('A:A', 12, a_center)   # for hostname
ws_parent.set_column('B:B', 20, a_left)     # for policy name
ws_parent.set_column('C:C', 15, a_center)   # for class
ws_parent.set_column('D:D', 20, a_center)   # for child policies
ws_parent.set_column('E:E', 12, a_center)   # for shaper
ws_parent.set_column('F:F', 35, a_left)     # for remarks
ws_parent.set_row(0, 23, header)            # for header

# Format columns and rows for child policies
ws_child.set_column('A:A', 12, a_center)   # for hostname
ws_child.set_column('B:B', 20, a_left)     # for policy name
ws_child.set_column('C:C', 13, a_center)   # for class map
ws_child.set_column('D:D', 18, a_center)   # for CIR/ CBS
ws_child.set_column('E:E', 35, a_left)     # for remarks
ws_child.set_row(0, 23, header)            # for header

# Write some data headers.
ws_ingress.write_row('A1', ['Hostname', 'Policy Name',
                            'CIR/ CBS', 'COS Setting', 'Circuit Type', 'Remarks'])
ws_parent.write_row('A1', ['Hostname', 'Policy Name',
                           'Class', 'Child Policy', 'Shaper', 'Remarks'])
ws_child.write_row('A1', ['Hostname', 'Policy Name',
                          'Class Map', 'CIR/ CBS', 'Remarks'])

# Start from the first cell below headers
row_ingress = 1
row_parent = 1
row_child = 1
# col = 0

# Open all .txt files in the current directory
for filename in glob.glob('*.txt'):

    # Read the contents of the text file
    with open(filename, 'r') as f:

        # Each file uses the ASR hostname in the filename
        hostname = filename.split('.')

        # Parse each line
        for line in f:

            # Parse the parent policy maps
            if ('policy_port' in line.lower()) or ('policy_parent' in line.lower()):

                # Write the hostname in column A
                ws_parent.write(row_parent, 0, hostname[0])
                # Write the policy names in column B
                policy = line[11:].strip()
                ws_parent.write(row_parent, 1, policy)

                col = 2

                for line in f:

                    if not line.strip() or ('!' in line):
                        continue
                    elif 'end-policy-map' not in line:
                        ws_parent.write(row_parent, col, line.strip())
                    else:
                        break

                    col += 1

                if ('temp' in policy.lower()) or ('test' in policy.lower()):
                    ws_parent.write(row_parent, 5,
                                    'Temporary or test policy - please delete.')
                elif 'parent' in policy.lower():
                    ws_parent.write(row_parent, 5,
                                    'Old parent policy - please delete.')
                else:
                    ws_parent.write(row_parent, 5, 'Ok')

                row_parent += 1

            elif ('policy_bvid' in line.lower()) \
                    or ('policy_svid' in line.lower()) \
                    or ('policy_child' in line.lower()) \
                    or ('egress' in line.lower()):

                policy = line[11:].strip()
                col = 2

                for line in f:

                    if not line.strip():
                        continue
                    elif 'end-policy-map' in line:
                        row_child -= 1
                        break
                    elif '!' not in line:

                        if 'class' in line:

                            # Write the hostname in column A
                            ws_child.write(row_child, 0, hostname[0])
                            # Write the policy names in column B
                            ws_child.write(row_child, 1, policy)
                            ws_child.write(row_child, col, line.strip())
                            col += 1

                            if ('temp' in policy.lower()) or ('test' in policy.lower()):
                                ws_child.write(row_child, 4,
                                               'Temporary or test policy - please delete.')
                            elif 'child' in policy.lower():
                                ws_child.write(row_child, 4,
                                               'Old child policy - please delete.')
                            elif ('bvid' in policy.lower()) or ('svid' in policy.lower()):
                                ws_child.write(row_child, 4, 'Ok')
                            else:
                                ws_child.write(row_child, 4,
                                               'Unknown policy - please delete.')

                        else:
                            ws_child.write(row_child, col, line.strip())
                            col = 2
                            row_child += 1

            else:

                # Parse the ingress policy maps
                if 'policy-map ' in line:

                    # Write the hostname in column A
                    ws_ingress.write(row_ingress, 0, hostname[0])
                    # Write the policy names in column B
                    ws_ingress.write(row_ingress, 1, line[11:].strip())

                    # Determine the circuit type based on the policy names
                    if 'aovl' in line.lower():
                        circuitType = 'Adva Overlay'
                    elif ('test' in line.lower()) or ('mef' in line.lower()):
                        circuitType = 'Test'
                    elif ('nid' in line.lower()) or ('dcn' in line.lower()):
                        circuitType = 'DCN'
                    else:
                        circuitType = 'Customer'

                    # Write the circuit type in column E
                    ws_ingress.write(row_ingress, 4, circuitType)

                # Determine and write the current CIR and CBS in column C
                elif 'police rate' in line:
                    ws_ingress.write(row_ingress, 2, line.strip())

                # Determine and write the current CoS setting in column D
                elif 'set cos' in line:
                    cos = line.strip()
                    ws_ingress.write(row_ingress, 3, cos)

                # end of policy map, check if settings are per the guidelines
                elif 'end-policy-map' in line:

                    if circuitType == 'Test':
                        if 'cos 4' in cos:
                            ws_ingress.write(row_ingress, 5, 'Ok')
                        else:
                            ws_ingress.write(row_ingress, 5,
                                             'Matched keyword test and/ or MEF. '
                                             'Confirm if test policy and delete. '
                                             'If not, change/ add set cos 4.')

                    elif circuitType == 'Adva Overlay':
                        if 'cos 6' in cos:
                            ws_ingress.write(row_ingress, 5, 'Ok')
                        else:
                            ws_ingress.write(row_ingress, 5, 'Change to set cos 6.')

                    elif circuitType == 'DCN':
                        if ('cos 4' in cos) or ('cos 7' in cos):
                            ws_ingress.write(row_ingress, 5, 'Ok')
                        else:
                            ws_ingress.write(row_ingress, 5,
                                             'DCN circuit - change set to cos 7.')

                    elif circuitType == 'Customer':
                        if 'cos 4' in cos:
                            ws_ingress.write(row_ingress, 5, 'Ok')
                        else:
                            ws_ingress.write(row_ingress, 5, 'Change/ add set cos 4.')

                    cos = ''
                    row_ingress += 1

# Set the autofilter
ws_ingress.autofilter(0, 0, row_ingress, 5)
ws_parent.autofilter(0, 0, row_parent, 5)
ws_child.autofilter(0, 0, row_child, 4)

wb.close()

os.startfile(wbName)
