import xlsxwriter
import xlrd
import os
import difflib

def filter(file_names):

    #통합할 파일
    wt_app_sheetnum=1
    wt_app_workbook=xlsxwriter.Workbook('removed_file.xlsx', {'strings_to_urls':False})
    wt_app_worksheet=wt_app_workbook.add_worksheet()

    #쓸 파일의 row, col
    wt_row=0
    wt_col=0
    
    #읽을 파일의 row, col
    rd_row=0
    rd_col=0
    
    
    
    for file_name in file_names:
        if 'integrated' in file_name:
            black_set=set([])
            #기사 선별
            ex_path=os.getcwd()+'\\'+file_name
            print(ex_path)
            #읽을 파일
            rd_workbook=xlrd.open_workbook(ex_path)
            rd_worksheet=rd_workbook.sheet_by_index(0)
            


            #읽은 엑셀 파일의 총 row와 col수
            tot_rows=rd_worksheet.nrows
                
            #읽을 엑셀 파일 row, col
            for row in range(0, tot_rows):
                if row in black_set:
                    continue
                compareing_cell=[]
                compareing_cell.append(rd_worksheet.cell_value(row, 0))
                compareing_cell.append(rd_worksheet.cell_value(row, 4))
                
                #compared_cell=[]
                #compared_cell.append(rd_worksheet.cell_value(com_row,0))
                #compared_cell.append(rd_worksheet.cell_value(com_row,4))
                
                for com_row in range(row, row+50):
                    
                    if com_row == row:
                        continue
                    if com_row in black_set:
                        continue
                    if com_row >= tot_rows:
                        break
                    
                    #print('com_row:',com_row)
                    #print(compareing_cell[0], '\n', compareing_cell[1])
                    matcher_0=difflib.SequenceMatcher(None, compareing_cell[0], rd_worksheet.cell_value(com_row,0))
                    matcher_4=difflib.SequenceMatcher(None, compareing_cell[1], rd_worksheet.cell_value(com_row,4))
                    if (float(0.8) < matcher_0.ratio() or float(0.8) < matcher_4.ratio()) and com_row != row:
                        print('Matching point row:', row, 'com_row:', com_row)
                        black_set.add(com_row)
                        
                        continue
                    else:
                        print('It is going well, row:',row, 'com_row:', com_row)
            for row in range(0, tot_rows):
                if row not in black_set:
                    value=rd_worksheet.row_values(row)
                    wt_app_worksheet.write_row(row, 0, value)

    print('Print black_set')
    print(black_set)            
    wt_app_workbook.close()
    
def main():
    file_names=os.listdir(os.getcwd())
    filter(file_names)

if __name__ == '__main__':
    main()