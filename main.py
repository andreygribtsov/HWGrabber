#%%
import re
import os
import xlwt
# from xlwt import Workbook
from openpyxl import Workbook


def parse_file(filepath):
    
    finded_lines = list()

    with open(filepath) as fp:
        for cnt, line in enumerate(fp):
            # print("Line {} : {}".format(cnt, line))
            # print(type(line))
            if "RACK" in line or "DPADDRESS" in line:                
                if "6" in line or "SIWAREX" in line:
                    if not ".GSD" in line:
                        if not "PPO" in line:
                            if not "Base" in line:
                                if not "Bytes" in line:
                                    if not "MODULE" in line:
                                        print(line)
                                        # print(re.findall(r'\"(.+?)\"', line))

def parse_file_v2(filepath):
    
    finded_lines = list()

    with open(filepath) as fp:
        for cnt, line in enumerate(fp):
            # print("Line {} : {}".format(cnt, line))
            # print(type(line))
            if "RACK" in line or "DPADDRESS" in line:                
                if "6" in line or "SIWAREX" in line:
                    if not ".GSD" in line:
                        if not "PPO" in line:
                            if not "Base" in line:
                                if not "Bytes" in line:
                                    if not "MODULE" in line:
                                        if not "_S7" in line:
                                            if not '-->' in line:
                                                if not 'Universal' in line:
                                                    if not 'Voltage' in line:
                                                        if not 'Current' in line:
                                                            if not 'Average' in line:
                                                                if not 'Freq' in line:
                                                                    if not 'Power' in line:
                                                                        if not 'Energy' in line:
                                                                            if not 'Input' in line:
                                                                                if not 'Geraet' in line:
                                                                                    if not 'MASTER DP' in line:
                                                                                        if not 'Words In' in line:
                                                                                            if not 'PORT1' in line:
                                                                                                if not '.GSE' in line:
                                                                                                    if not 'PMD Diagnostics' in line:
                                                                                                        if not '(read)' in line:
                                                                                                            if not '(POSITION)' in line:
                                                                                                                if not 'Basic type 2' in line:
                                                                                                                    if not 'MB_REG_READ_2W' in line:
                                                                                                                        if not 'Type 3:' in line:
                                                                                                                            if not 'Output 1 byte' in line:
                                                                                                                                if not 'Word In' in line:
                                                                                                                                    if not 'byte' in line:
                                                                                                                                        if not 'Byte' in line:
                                                                                                                                            if not 'TOTAL' in line:
                                                                                                                                                if not 'Words' in line:
                                                                                                                                                    if not 'Main Process Value' in line:
                                                                                                                                                # print(line)
                                                                                                                                                        # print(re.findall(r'\"(.+?)\"', line))
                                                                                                                                                            finded_lines.append(re.findall(r'\"(.+?)\"', line)) 
    return finded_lines

    # for item in finded_lines:
    #     print(item)

    # print("elems qty = {} \n".format(finded_lines.count()))

    # finded_lines.clear()

def main():

    files = os.listdir('./sources/')
    # files = ['UTIC02.cfg']
    # wb = Workbook()
    # sheet = wb.add_sheet('Sheet 1')

    workbook = Workbook()
    sheet = workbook.active

    row_number = 1

    for _file in files:
        result = parse_file_v2('./sources/' + _file)
        area_name = re.sub(r'\.cfg$','', _file)
        for _line in result:
            if len(_line) == 2:
                _line.append('')
                _line.append(area_name.upper())
            else:
                _line.append(area_name.upper())
            # print(_line)
            # sheet.write(row_number, 1, _line)
            # sheet.write(row_number, 2, )
            sheet.append(_line)
            # sheet['B' + str(row_number)] = str(_line)
            # sheet['C' + str(row_number)] = re.sub(r'\.cfg$','', _file)

            row_number += 1
    # wb.save("order_list.xls")
    workbook.save(filename='order_list.xlsx')

    
    

















if __name__ == "__main__":
    main()
