# -*- coding: utf-8 -*-
"""
Created on Thu Jul 20 07:51:36 2017

@author: shubham
"""

from pandas import read_csv
import difflib as dl
from openpyxl import Workbook
from openpyxl.styles import Font


#Global Varibles
group_counter=0
input_file_path='F:\SHUBHAM\MachineLearning\POC\Data\pankaj.csv'
output_file_path='F:\SHUBHAM\MachineLearning\POC\Data\pankaj1.xlsx'
header_names=['a','b']
data=read_csv(input_file_path,header=None,encoding='"ISO-8859-1"',dtype=str)
input_1=data.values.T.tolist()
temp=input_1[0]
original_list=temp
row_count=2
#==============================================================================
# #read the input file 
#encoding is used because input may have anytype data
#
#header_names =header name of any csv file .csv file does not contain header
#==============================================================================

def find_the_similar_data(input1):
    global temp
    global group_counter
    list_of_similiar_product=dl.get_close_matches(input1,temp,n=len(temp),cutoff=0.65)
    group_counter+=1
    return list_of_similiar_product,group_counter


def remove_the_word_from_list(input_list,group):
    global temp
    global original_list
    for i in input_list:
        insert_data_into_excel_Sheet(i,group)
        temp.remove(i)
       

def createExcelSheetforlog(path):
    global sheet1
    global wb
    wb=Workbook()
    sheet=wb.active
    sheet.title="logs"
    sheet1=wb.get_sheet_by_name("logs")
    header_font=Font(size=12,bold=True)
    sheet1['A1'].font,sheet1['B1'].font=header_font,header_font
    sheet1['A1']="Product Description"
    sheet1['B1']="Group Number"
    wb.save(path)
    
    

def insert_data_into_excel_Sheet(a,b):
    global row_count
    sheet1['A'+str(row_count)]=a
    sheet1['B'+str(row_count)]=b
    row_count+=1       
    wb.save(output_file_path)
    
def main():
    print("Started....")
    createExcelSheetforlog(output_file_path)
    while(len(original_list)!=0):
        searched_elements,group_nbr,=find_the_similar_data(original_list[0])
        Group_nbr="Group"+str(group_nbr)
        remove_the_word_from_list(searched_elements,Group_nbr)
    print("Complete...")
   
main()