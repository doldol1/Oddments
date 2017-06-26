import xlwt
import xlrd
import os

def filter(file_name, keyword_list):
    
    #기사 선별
    ex_path=os.getcwd()+'\기사 선별\\'+file_name
    print(ex_path)
    
    #파일 읽기
    rd_workbook=xlrd.open_workbook(ex_path)
    rd_worksheet=rd_workbook.sheet_by_index(0)
    
    #파일 쓰기(매칭)
    wt_app_sheetnum=1
    wt_app_workbook=xlwt.Workbook(encoding='utf-8')
    wt_app_worksheet=wt_app_workbook.add_sheet('Sheet'+str(wt_app_sheetnum))
    
    #파일 쓰기(언매칭)
    wt_un_sheetnum=1
    wt_un_workbook=xlwt.Workbook(encoding='utf-8')
    wt_un_worksheet=wt_un_workbook.add_sheet('Sheet'+str(wt_un_sheetnum))
    
    #읽은 엑셀 파일의 총 row와 col수
    tot_rows=rd_worksheet.nrows
    tot_cols=rd_worksheet.ncols
        
    #읽을 엑셀 파일 row, col
    rd_row=0
    rd_col=0

    #쓸 파일의 row, col
    wt_row=0
    wt_col=0
    
    #match_count=0
    #unmatch_count=0
    total=0
    
    #매치 리스트(조건에 맞는 row의 번호)
    match_set=set([])
    unmatch_set=set([])
    
    while rd_row < tot_rows:
        cell=rd_worksheet.cell_value(rd_row, rd_col)
        for keyword in keyword_list:
            if keyword in cell:
                #print(rd_worksheet.row_values(rd_row))
                match_set.add(rd_row)
                #print('매칭',rd_row)
        rd_row=rd_row+1
    
    print(len(match_set))
    print(match_set)

    for row in range(0, tot_rows):
        for value in rd_worksheet.row_values(row):
            if row in match_set:
                wt_app_worksheet.write(row, wt_col, value)
                #match_count=match_count+1
            else:
                wt_un_worksheet.write(row, wt_col, value)
                #unmatch_count=unmatch_count+1
            wt_col=wt_col+1
        
        wt_col=0

    wt_app_workbook.save(ex_path.replace('.xls','')+'_match.xls')
    wt_un_workbook.save(ex_path.replace('.xls','')+'_unmatch.xls')

'''        
    print('Match Count:', match_count/5)
    print('Unmatch Count:', unmatch_count/5)
    print('Total Count:', tot_rows)

    if match_count/5+unmatch_count/5 is not tot_rows:
        print('Something Wrong')
'''
    
def main():

    
    file_names=os.listdir(os.getcwd()+'\기사 선별')
    
    

    keyword_list=('다문화', '다문화가족', '혼혈', '외국인근로자', '외국인 근로자', '외국인노동자', '외국인 노동자', '이주', '불법체류', '불법 체류', '국제결혼', '국제 결혼', '이민', '필리핀', '몽골', '베트남')
    
    #print(file_names)
    
    
    for file_name in file_names:
        filter(file_name, keyword_list)

if __name__ == '__main__':
    main()