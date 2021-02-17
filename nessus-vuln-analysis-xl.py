from lxml import etree # used for parsing .nessus files
import datetime # used for formatting scan date and accurately comparing datetime values
import os # used for file pathing and file backups

# Takes string instructions and a flag for whether or not an extension is expected and exits if too many failed attempts occur or some other issue occurs
def _Inp_Path (instruct, opt):
    count = 0 # counts errors
    while True:
        try:
            p = str(input(instruct+': '))
        except:
            "\nUnexpected error. Try again...\n"

        if count == 3:
            _Err_Exit('\nYou seem to be having trouble. Confirm your desired path and come back later.\nExiting...')
        elif opt == 'f': # f option indicates path directly to file, checks simply for valid path
            if path.isfile(p):
                break
            else:
                count+=1
                print('\nInvalid path to file. Try again...\n')
                continue
        elif opt == 'd': # d option indicates path to a directory, checks simply for valid path
            if path.isdir(p):
                break
            else:
                count+=1
                print('\nInvalid path to directory. Try again...\n')
                continue
        elif opt == 'x': # x option indicates an excel file that has been generated by this Python program
            print(p.split('.')[-1])
            if p.split('.')[-1] == "xlsx" and path.isdir('\\'.join(p.split('\\')[:-2])): # isolate/check the extension and check validity of target dir
                wb = load_workbook(p)
                sheetlist = []
                for sheet in wb.sheetnames:
                    sheetlist.append(sheet)
                if "statuses" in sheetlist and "targets" in sheetlist and "columns" in sheetlist: # determines if the excel file has been generated by this program by checking for the existance of specific sheetnames
                    wb.close()
                    break
                else:
                    wb.close()
                    count+=1
                    print("\nThis does not appear to be a compatible workbook. Please provide a workbook that has been generated by this program.\n")
                    continue
            else:
                count+=1
                print('\nInvalid path for writing. Make sure the full path is correct, that the extension is .xlsx, and that the target is a workbook generated originally by this Python tool...\n')
                continue
    return p

def _Parse_Nessus(report_path):
    client = ""
    report_dict = dict()
    host_params = ["HOST_START",
                   "mac-address",
                   "netbios-name",
                   "operating-system",
                   "host-ip"]
    vuln_params = ["pluginName",
                   "pluginID",
                   "port",
                   "svc_name",
                   "severity",
                   "synopsis",
                   "solution",
                   "plugin_output"]

    f = open(report_path, 'r')
    xml_content = f.read()
    f.close()

    parz = etree.XMLParser(huge_tree=True)
    root = etree.fromstring(text=xml_content, parser=parz)
    for block in root:
        if block.tag == "Report":
            client = block.attrib['name'].split(" ", 1)[0] # Grabs the client acronym from the scan name
            for ReportHost in block:
                props_dict = dict() # dict for holding host properties
                for ReportItem in ReportHost:
                    vuln_dict = dict() # dict for holding individual vulnerability details; to be attached to the host props dict
                    if ReportItem.tag == "HostProperties":
                        for tagg in ReportItem:
                            if tagg.attrib['name'] in host_params:
                                #if tagg.attrib['name'] == "mac-address" and "virtual-mac-address" in tagg.attrib.values(): WORKING ON MAC SORTING WHEN THERE'S MORE THAN ONE
                                props_dict[tagg.attrib['name']] = tagg.text
                    else:
                        for attr in ReportItem.attrib:
                            if attr in vuln_params:
                                vuln_dict[attr] = ReportItem.attrib[attr]
                        for param in ReportItem:
                            if param.tag in vuln_params:
                                vuln_dict[param.tag] = param.text
                        props_dict['vulns'] = vuln_dict
                report_dict[ReportItem.attrib['name']] = props_dict


def main ():
    report_path = _Inp_Path("Enter a filepath to your .nessus file", 'f') # Provide path to .nessus report file for importing
    report_dict = _Parse_Nessus(report_path) # parse .nessus file for report dicts, sheet (client) name, and scan date
