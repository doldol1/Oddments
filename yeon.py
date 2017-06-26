import xlsxwriter
import xlrd
import os
import difflib

def filter(file_names):

    #통합할 파일
    wt_app_sheetnum=1
    wt_app_workbook=xlsxwriter.Workbook('removed_yeon.xlsx', {'strings_to_urls':False})
    wt_app_worksheet=wt_app_workbook.add_worksheet()

    #쓸 파일의 row, col
    wt_row=0
    wt_col=0
    
    #읽을 파일의 row, col
    rd_row=0
    rd_col=0
    
    
    
    for file_name in file_names:
        if 'removed' in file_name:
            ex_path=os.getcwd()+'\\'+file_name
            print(ex_path)
            #읽을 파일
            rd_workbook=xlrd.open_workbook(ex_path)
            rd_worksheet=rd_workbook.sheet_by_index(0)
            


            #읽은 엑셀 파일의 총 row와 col수
            tot_rows=rd_worksheet.nrows
            for row in range(0, tot_rows):
                if rd_worksheet.cell_value(row, 4) != '해당 뉴스는 스포츠연애 부문 언론사의 본문이며, 자료를 가져올 수 없습니다.':
                    value=rd_worksheet.row_values(row)
                    wt_app_worksheet.write_row(row, 0, value)

    wt_app_workbook.close()
    
def main():
    file_names=os.listdir(os.getcwd())
    filter(file_names)

if __name__ == '__main__':
    main()