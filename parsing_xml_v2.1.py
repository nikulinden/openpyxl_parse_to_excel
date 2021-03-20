import xml.etree.ElementTree as ET
import os
import os.path
from utils import *


working_dir = os.getcwd()
#name of the excel to dump the xml values
spreadsheet_name = 'works.xlsx'


if __name__ == '__main__':
    #get the names of all xml files and put them in a list
    sheet_list = get_xml_filenames(working_dir)
    #each script run a new excel is created
    if os.path.isfile(spreadsheet_name):
        os.remove(os.path.join(working_dir,spreadsheet_name))
        create_workbook(spreadsheet_name)
    else:
        create_workbook(spreadsheet_name)
    #iter over xml files and call appropriate function
    for i in sheet_list:
        #name the sheets the same as xmls
        sheet_name = xml_name = i
        tree = ET.parse(xml_name)
        root = tree.getroot()
        print (root.tag, root.attrib)

        if 'weighttables' in xml_name:
            feature_weights = parse_xml_to_dict(root)
            dict_to_excel(feature_weights, spreadsheet_name, sheet_name)
        elif ('freegames_reelstrips' or 'sjdspin_reelstrips') in xml_name:
            reels = get_xml_reels(root)
            list_to_excel(reels, spreadsheet_name, sheet_name)
        elif 'main_reelstrips' in xml_name:
            reels = get_xml_reels_with_weights(root)
            list_to_excel(reels, spreadsheet_name, sheet_name)
        elif 'slot' in xml_name:
            main_slot = get_main_slot_combo_xml(root)
            main_dict_to_excel(main_slot, spreadsheet_name, sheet_name)
