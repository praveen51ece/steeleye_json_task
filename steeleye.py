import xlrd
import os
import json

def read_data_and_store_injson(filepath, start, end):
    """
    This file is for Open and store data from Excel file to json file.
    here parameter will provide filepath, start and end point.That will open 
    xl file row number parameter(start to end).
    @Author: Praveen
    @Date: 16/04/2018
    """
    if os.path.exists(filepath):
        book = xlrd.open_workbook(filepath, "r")   
        MICs_List_by_CC = book.sheet_by_index(1)
        dict_key = MICs_List_by_CC.row_values(0)
        # print dict_key
        data_list = []   
        for row in xrange(start, end):
            try:
                dict_data = {}
                dict_data[dict_key[0]] = MICs_List_by_CC.cell_value(row, 0)  
                dict_data[dict_key[1]] = MICs_List_by_CC.cell_value(row, 1)
                dict_data[dict_key[2]] = MICs_List_by_CC.cell_value(row, 2)
                dict_data[dict_key[3]] = MICs_List_by_CC.cell_value(row, 3)
                dict_data[dict_key[4]] = MICs_List_by_CC.cell_value(row, 4)
                dict_data[dict_key[5]] = MICs_List_by_CC.cell_value(row, 5)
                dict_data[dict_key[6]] = MICs_List_by_CC.cell_value(row, 6)
                dict_data[dict_key[7]] = MICs_List_by_CC.cell_value(row, 7)
                dict_data[dict_key[8]] = MICs_List_by_CC.cell_value(row, 8)
                dict_data[dict_key[9]] = MICs_List_by_CC.cell_value(row, 9)
                dict_data[dict_key[10]] = MICs_List_by_CC.cell_value(row, 10)
                dict_data[dict_key[11]] = MICs_List_by_CC.cell_value(row, 11)
                dict_data[dict_key[12]] = MICs_List_by_CC.cell_value(row, 12)
                data_list.append(dict_data)
                
                with open('json_data.json', 'w') as outfile:
                    json.dump(data_list, outfile)

            except Exception as e:
                print e
        # print 'data_list %s' %data_list
    else:
        print "Filepath does not exist"

filepath = 'ISO10383_MIC.xls'
start = 1
end = 1617
read_data_and_store_injson(filepath, start, end) 