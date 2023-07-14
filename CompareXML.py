import datetime
from tkinter import messagebox
# import xml.etree.ElementTree as ET
from lxml import etree
import os
import openpyxl as XL
from openpyxl.styles import Font
from openpyxl.styles import Border, Side
from difflib import SequenceMatcher
from collections import OrderedDict
import codecs

red = '\033[91m'
green = '\033[92m'
blue = '\033[94m'
bold = '\u001b[1m'
italics = '\033[3m'
underline = '\033[4m'
end = '\u001b[0m'

global row_count, resultfile, resultsheet


def compare_xml_files(backupFilePath, updateFilepath):
    if backupFilePath.get(1.0, "end-1c") and updateFilepath.get(1.0, "end-1c"):
        global row_count, resultfile, resultsheet
        filePath1 = backupFilePath.get(1.0, "end")
        filePath1 = filePath1[:-1]
        filePath2 = updateFilepath.get(1.0, "end")
        filePath2 = filePath2[:-1]
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current_time = current_time.replace(" ","_")
        current_time = current_time.replace(":", "-")
        output_file_path = os.path.join(os.getcwd(),'CompareResults', f'xmlCompare_output_{str(current_time)}' + ".xlsx")
        row_count = 1
        border_style = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
        resultfile = XL.Workbook()
        resultsheet = resultfile.active
        resultsheet[f'A{row_count}'].value = "Label"
        resultsheet.column_dimensions['A'].width = 20
        resultsheet.column_dimensions['B'].width = 30
        resultsheet.column_dimensions['C'].width = 30
        resultsheet.column_dimensions['D'].width = 30
        resultsheet.column_dimensions['E'].width = 30
        resultsheet.column_dimensions['F'].width = 30
        resultsheet[f'B{row_count}'].value = "Text Content Difference"
        resultsheet[f'C{row_count}'].value = "Attributes & Text Changes of Backup"
        resultsheet[f'C{row_count}'].alignment = resultsheet[f'C{row_count}'].alignment.copy(wrapText=True)
        resultsheet[f'D{row_count}'].value = "Attributes & Text Changes of Updated"
        resultsheet[f'D{row_count}'].alignment = resultsheet[f'D{row_count}'].alignment.copy(wrapText=True)
        resultsheet[f'E{row_count}'].value = "Extra Nodes in Backup"
        resultsheet[f'E{row_count}'].alignment = resultsheet[f'E{row_count}'].alignment.copy(wrapText=True)
        resultsheet[f'F{row_count}'].value = "Extra Nodes in Updated"
        resultsheet[f'F{row_count}'].alignment = resultsheet[f'F{row_count}'].alignment.copy(wrapText=True)
        resultsheet[f'G{row_count}'].value = "CompareResults A != B"
        resultsheet.column_dimensions['G'].width = 400
        resultsheet[f'A{row_count}'].font = Font(bold=True)
        resultsheet[f'B{row_count}'].font = Font(bold=True)
        resultsheet[f'C{row_count}'].font = Font(bold=True)
        resultsheet[f'D{row_count}'].font = Font(bold=True)
        resultsheet[f'E{row_count}'].font = Font(bold=True)
        resultsheet[f'F{row_count}'].font = Font(bold=True)
        resultsheet[f'G{row_count}'].font = Font(bold=True)
        row_count += 2
        # with open(output_file_path, 'w') as f:
        #     f.write('')

        with open(filePath1, encoding='utf8') as file1, open(filePath2, encoding='utf8') as file2:
            try:
                tree1 = etree.parse(file1)
                tree2 = etree.parse(file2)
                root1 = tree1.getroot()
                root2 = tree2.getroot()

                # with open(output_file_path, 'w') as f:
                mycompare(root1, root2, resultsheet)
                for row in resultsheet.iter_rows():
                    for cell in row:
                        cell.border = border_style

                resultfile.save(output_file_path)
                messagebox.showinfo("Task Completed",
                                    "Compare successful! Please Check Location.")

            except etree.ParseError as e:
                messagebox.showerror("Error", f"Error parsing XML: {e}")

    else:
        messagebox.showwarning("Warning", "Please select XML files!")


def mycompare(ele1, ele2, output_fil):
    global row_count, resultfile, resultsheet

    for chil1, chil2 in zip(ele1, ele2):
        processflag1 = False
        processflag2 = False
        if chil1.tag == 'Process' and chil2.tag == 'Process' and chil1.attrib.get('Name') == chil2.attrib.get('Name'):
            compare_xml_elements(chil1, chil2, output_fil)

        else:
            if chil1.tag != 'Process' and chil2.tag != 'Process':
                compare_xml_elements(chil1, chil2, output_fil)

            if chil1.tag == 'Process':
                for chi2, t1 in zip(ele2, ele1):
                    if chil1.attrib.get('Name') == chi2.attrib.get('Name'):
                        processflag1 = True
                        compare_xml_elements(chil1, chi2, output_fil)
                        break

                if processflag1 == False:
                    resultsheet[f'E{row_count}'].value = f"Process - {chil1.attrib.get('Label')}"
                    resultsheet[f'E{row_count}'].alignment = resultsheet[f'E{row_count}'].alignment.copy(
                        wrapText=True)
                    xmlstr = etree.tostring(chil1, encoding='unicode')
                    resultsheet[f'C{row_count}'].value = f"{xmlstr}"
                    resultsheet[f'C{row_count}'].alignment = resultsheet[f'C{row_count}'].alignment.copy(
                        wrapText=True)

                row_count += 1

            if chil2.tag == 'Process':
                for chi1, t2 in zip(ele1, ele2):
                    if chil2.attrib.get('Name') == chi1.attrib.get('Name'):
                        processflag2 = True
                        compare_xml_elements(chi1, chil2, output_fil)
                        break

                if processflag2 == False:
                    resultsheet[f'F{row_count}'].value = f"{chil2.attrib.get('Label')}"
                    resultsheet[f'F{row_count}'].alignment = resultsheet[f'F{row_count}'].alignment.copy(
                        wrapText=True)
                    xmlstr = etree.tostring(chil2, encoding='unicode')
                    resultsheet[f'D{row_count}'].value = f"{xmlstr}"
                    resultsheet[f'D{row_count}'].alignment = resultsheet[f'D{row_count}'].alignment.copy(
                        wrapText=True)

                row_count += 1
        #
        # if len(elem1) != len(elem2):
        #     resultsheet[
        #         f'G{row_count}'].value = f"Child element count mismatch for tag '{elem1.tag}': {len(elem1)} != {len(elem2)}\n"
        #     if elem1.attrib.get("Label") == elem2.attrib.get("Label") and elem1.attrib.get("Label") != None:
        #         resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label')}"
        #
        #     elif elem1.attrib.get("Label") != elem2.attrib.get("Label") and elem1.attrib.get(
        #             "Label") != None and elem2.attrib.get("Label") != None:
        #         resultsheet[
        #             f'A{row_count}'].value = f"{elem1.attrib.get('Label'), elem2.attrib.get('Label')}"
        #
        #     else:
        #         Parent_node = elem1.getparent()
        #         while Parent_node != None and Parent_node.attrib.get('Label') == None:
        #             Parent_node = Parent_node.getparent()
        #
        #         if Parent_node != None and Parent_node.attrib.get('Label') != None:
        #             resultsheet[f'A{row_count}'].value = f"{Parent_node.attrib.get('Label')}"
        #
        #     if len(elem1) > len(elem2):
        #         resultsheet[f'E{row_count}'].value = f"{find_extra_beads(elem1)}"
        #         resultsheet[f'E{row_count}'].alignment = resultsheet[f'E{row_count}'].alignment.copy(
        #             wrapText=True)
        #         xmlstr = etree.tostring(elem1, encoding='unicode')
        #         resultsheet[f'C{row_count}'].value = f"{xmlstr}"
        #         resultsheet[f'C{row_count}'].alignment = resultsheet[f'C{row_count}'].alignment.copy(
        #             wrapText=True)
        #
        #     else:
        #         resultsheet[f'F{row_count}'].value = f"{find_extra_beads(elem2)}"
        #         resultsheet[f'F{row_count}'].alignment = resultsheet[f'F{row_count}'].alignment.copy(
        #             wrapText=True)
        #         xmlstr = etree.tostring(elem2, encoding='unicode')
        #         resultsheet[f'D{row_count}'].value = f"{xmlstr}"
        #         resultsheet[f'D{row_count}'].alignment = resultsheet[f'D{row_count}'].alignment.copy(
        #             wrapText=True)


def compare_xml_elements(elem1, elem2, output_file):
    global row_count, resultfile, resultsheet

    if elem1.tag != elem2.tag:
        resultsheet[f'G{row_count}'].value = f"Tag mismatch: {elem1.tag} != {elem2.tag}\n"
        if elem1.attrib.get("Label") == elem2.attrib.get("Label") and elem1.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label')}"

        elif elem1.attrib.get("Label") != elem2.attrib.get("Label") and elem1.attrib.get("Label") != None and elem2.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label'), elem2.attrib.get('Label')}"

        else:
            Parent_node = elem1.getparent()
            while Parent_node != None and Parent_node.attrib.get('Label') == None:
                Parent_node = Parent_node.getparent()

            if Parent_node != None and Parent_node.attrib.get('Label') != None:
                resultsheet[f'A{row_count}'].value = f"{Parent_node.attrib.get('Label')}"
        row_count += 1

    if compare_attributes(elem1.attrib,elem2.attrib, 'Visible') != 0:
        resultsheet[f'G{row_count}'].value = f"Attribute mismatch for tag '{elem1.tag}': {elem1.attrib} != {elem2.attrib}\n"
        if elem1.attrib.get("Label") == elem2.attrib.get("Label") and elem1.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label')}"

        elif elem1.attrib.get("Label") != elem2.attrib.get("Label") and elem1.attrib.get("Label") != None and elem2.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label'), elem2.attrib.get('Label')}"

        else:
            Parent_node = elem1.getparent()
            while Parent_node != None and Parent_node.attrib.get('Label') == None:
                Parent_node = Parent_node.getparent()

            if Parent_node != None and Parent_node.attrib.get('Label') != None:
                resultsheet[f'A{row_count}'].value = f"{Parent_node.attrib.get('Label')}"

        keys_to_compare1 = [key for key in elem1.attrib.keys() if key != 'Visible']
        keys_to_compare2 = [key for key in elem2.attrib.keys() if key != 'Visible']
        dict1 = {}
        dict2 = {}
        for key in keys_to_compare1:
            if elem1.attrib.get(key) != elem2.attrib.get(key):
                dict1[key] = elem1.attrib.get(key)

        for key in keys_to_compare2:
            if elem1.attrib.get(key) != elem2.attrib.get(key):
                dict2[key] = elem2.attrib.get(key)
        resultsheet[f'C{row_count}'].value = f"{dict1}"
        resultsheet[f'C{row_count}'].alignment = resultsheet[f'C{row_count}'].alignment.copy(wrapText=True)
        resultsheet[f'D{row_count}'].value = f"{dict2}"
        resultsheet[f'D{row_count}'].alignment = resultsheet[f'D{row_count}'].alignment.copy(wrapText=True)

        row_count += 1

    if elem1.text != elem2.text and elem1.text != None and elem1.text != "" and elem2.text != None and elem2.text != "":
        if elem1.attrib.get("Label") == elem2.attrib.get("Label") and elem1.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label')}"

        elif elem1.attrib.get("Label") != elem2.attrib.get("Label") and elem1.attrib.get("Label") != None and elem2.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label'), elem2.attrib.get('Label')}"

        else:
            Parent_node = elem1.getparent()
            while Parent_node != None and Parent_node.attrib.get('Label') == None:
                Parent_node = Parent_node.getparent()

            if Parent_node != None and Parent_node.attrib.get('Label') != None:
                resultsheet[f'A{row_count}'].value = f"{Parent_node.attrib.get('Label')}"

        resultsheet[f'G{row_count}'].value = f"Text content mismatch for tag '{elem1.tag}': {elem1.text} != {elem2.text}\n"
        difference = get_string_difference(elem1.text, elem2.text)
        resultsheet[f'B{row_count}'].value = f"{difference}"
        resultsheet[f'B{row_count}'].alignment = resultsheet[f'B{row_count}'].alignment.copy(wrapText=True)
        row_count += 1

    if len(elem1) != len(elem2):
        resultsheet[f'G{row_count}'].value = f"Child element count mismatch for tag '{elem1.tag}': {len(elem1)} != {len(elem2)}\n"
        if elem1.attrib.get("Label") == elem2.attrib.get("Label") and elem1.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label')}"

        elif elem1.attrib.get("Label") != elem2.attrib.get("Label") and elem1.attrib.get("Label") != None and elem2.attrib.get("Label") != None:
            resultsheet[f'A{row_count}'].value = f"{elem1.attrib.get('Label'), elem2.attrib.get('Label')}"

        else:
            Parent_node = elem1.getparent()
            while Parent_node != None and Parent_node.attrib.get('Label') == None:
                Parent_node = Parent_node.getparent()

            if Parent_node != None and Parent_node.attrib.get('Label') != None:
                resultsheet[f'A{row_count}'].value = f"{Parent_node.attrib.get('Label')}"

        if len(elem1) > len(elem2):
            resultsheet[f'E{row_count}'].value = f"{find_extra_beads(elem1)}"
            resultsheet[f'E{row_count}'].alignment = resultsheet[f'E{row_count}'].alignment.copy(wrapText=True)
            xmlstr = etree.tostring(elem1, encoding='unicode')
            resultsheet[f'C{row_count}'].value = f"{xmlstr}"
            resultsheet[f'C{row_count}'].alignment = resultsheet[f'C{row_count}'].alignment.copy(wrapText=True)

        else:
            resultsheet[f'F{row_count}'].value = f"{find_extra_beads(elem2)}"
            resultsheet[f'F{row_count}'].alignment = resultsheet[f'F{row_count}'].alignment.copy(wrapText=True)
            xmlstr = etree.tostring(elem2, encoding='unicode')
            resultsheet[f'D{row_count}'].value = f"{xmlstr}"
            resultsheet[f'D{row_count}'].alignment = resultsheet[f'D{row_count}'].alignment.copy(wrapText=True)

        row_count += 1

    for child1, child2 in zip(elem1, elem2):
        compare_xml_elements(child1, child2, output_file)


def compare_attributes(dict1, dict2, ignorekey1):
    keys_to_compare = [key for key in dict1.keys() if key != ignorekey1]
    count = 0
    for key in keys_to_compare:
        if dict1.get(key) != dict2.get(key):
            count += 1
    return count


def get_string_difference(string1, string2):
    differences = " "
    if string1 != "" and string2 != "" and string1 != None and string2 != None:
        matcher = SequenceMatcher(None, string1, string2)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'replace':
                differences += f'Replace: {string1[i1:i2]} With {string2[j1:j2]}\n'
            elif tag == 'delete':
                differences += f'Delete: {string1[i1:i2]}\n'
            elif tag == 'insert':
                differences += f'Insert: {string2[j1:j2]}\n'
    return differences


def find_extra_beads(elem):
    stack = [elem]
    d = [OrderedDict()]
    while stack:
        element = stack.pop()
        if element.attrib.get('Label') != None and element.tag != None:
            d.append({element.tag: element.attrib.get('Label')})
        stack.extend(reversed(list(element)))

    labels = ""
    i = 1
    while i < len(d):
        labels += f"{d[i]}\n"
        i += 1
    return labels

# '''IF first check process name then 	check process by their 	lable numbers and Name
#   Match comparison
#
# Else
#     IF loop on all process and check if Name is matching with the process
#  	Comparison
#     Else
# 	show in backup or updated as a extra process
#
#
#
# check also tags names of process
#
#
#   for inside tags check with label numbers
# '''