from bs4 import BeautifulSoup
import urllib.request
import xlwt
import datetime
import re
#-*- coding: utf-8 -*-
 


#주소 추가
OR_URLS=['http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%C0%CF%C0%DA%B8%AE&sm=title.basic&pd=4&'\
'startDate=1960-01-01&endDate=2009-12-31', 'http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%C0%CF%C0%DA%B8%AE&sm=title.basic&'\
'pd=4&startDate=2010-01-01&endDate=2012-12-31','http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%C0%CF%C0%DA%B8%AE&sm=title.basic&pd=4&'\
'startDate=2013-01-01&endDate=2015-12-31', 'http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%C0%CF%C0%DA%B8%AE&sm=title.basic&pd=4&'\
'startDate=2016-01-01&endDate=2017-06-19', 'http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%BD%C7%BE%F7&sm=title.basic&pd=4&'\
'startDate=1960-01-01&endDate=2011-12-31', 'http://news.naver.com/main/search/search.nhn?'\
'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
'3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
'ie=MS949&detail=0&rcsection=&query=%BD%C7%BE%F7&sm=title.basic&pd=4&'\
'startDate=2012-01-01&endDate=2017-06-19']

# s='http://news.naver.com/main/search/search.nhn?'\
# 'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
# '3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
# 'ie=MS949&detail=0&rcsection=&query=%C7%D8%BE%E7%BD%C9%C3%FE%BC%F6&'\
# 'sm=all.basic&pd=4&startDate=2017-01-01&endDate=2017-06-28'
# OR_URLS=['http://news.naver.com/main/search/search.nhn?'\
# 'rcnews=exist%3A032%3A005%3A086%3A020%3A021%3A081%3A022%3A023%3A025%'\
# '3A028%3A038%3A469%3A&refresh=&so=rel.dsc&stPhoto=&stPaper=&stRelease=&'\
# 'ie=MS949&detail=0&rcsection=&query=%C7%D8%BE%E7%BD%C9%C3%FE%BC%F6&'\
# 'sm=all.basic&pd=4&startDate=2017-01-01&endDate=2017-06-28']


# 출력 파일 명
OUTPUT_FILE_NAME = ''
 
# 스크래핑 함수
def get_text(OR_URLS, OUTPUT_FILE_NAME):
    for OR_URL in OR_URLS:
        excel_file=xlwt.Workbook(encoding='utf-8')
        excel_sheet=excel_file.add_sheet('Sheet1')

        print(OR_URL)
        t=0
        source_stack=list()
        while(t <5):

            
            URL=OR_URL      #URL초기화
            text=''           #파일에 저장할 변수
            dup=''              #중복 페이지 여부 확인
            num=0            #row의 숫자를 맞춰주기 위한 변수. 
            t=t+1               #스크래핑할 col

            OUTPUT_FILE_NAME=set_stp(t=t)
            print(str(OUTPUT_FILE_NAME)+' started')
            
            i=1                 #스크래핑할 페이지의 수
            while i < 401:      #페이지 수 400개(최대 한계)로 설정
                source=''


                
            
                #URL 변경
                URL=chkURL(i=i, URL=URL)
            
                #URL 요철을 열 때 예외 발생 확률이 높으므로 try사용
                try:
                    if(t is 1):
                        source_code_from_URL = urllib.request.urlopen(URL)
                        soup = BeautifulSoup(source_code_from_URL, 'lxml', from_encoding='utf-8')
                        source_stack.append(soup)
                    else:
                        soup =source_stack[i-1]
                    
                except urllib.error.URLError as e:
                    print(e.reason)
                    break

                #soup에 해당 URL페이지의 내용이 담긴다.
                

                #종료 여부 확인
                tmp2=soup.find('div', {"class" : "paging"}).select("strong")
                #print(tmp2)
                if(dup==tmp2):
                    print('end')
                    break
                else:
                    dup=tmp2
            
                #term_extractor로 soup에 담긴 data를 추출한다.
                num=term_extractor(t=t, soup=soup, num=num, excel_sheet=excel_sheet)
                #term_extractor의 외부에 있기 때문에, 정확한 row의 개수를 확인할 때 num을 사용하면 안 됨
                if(t==5):
                    print(i,'번째 페이지 완료')

                i=i+1
        
            print('Process end. i='+str(i)+'\n')
        #저장할 파일 이름을 excel_file.save에 입력
        excel_name=OR_URL.split('&')
        excel_file.save(excel_name[-2]+'부터'+excel_name[-1]+'.xls')
        print(OR_URL)

#진행현황 네이밍
def set_stp(t):
    if(t==1):
        OUTPUT_FILE_NAME = 'Subject'
    elif(t==2):
        OUTPUT_FILE_NAME = 'Press'
    elif(t==3):
        OUTPUT_FILE_NAME = 'Time'
    elif(t==4):
        OUTPUT_FILE_NAME = 'News URL'
    elif(t==5):
        OUTPUT_FILE_NAME = 'Contents'
    else:
        OUTPUT_FILE_NAME = 'error'
        print('Try error: Filing')
    return OUTPUT_FILE_NAME
#항에서 data 추출
def term_extractor(t, soup, num, excel_sheet) :

    if(t==1):
        for item in soup.find_all('div', {"class" : "ct"}):
            tmp1=item.select(".tit")
            excel_sheet.write(num, 0, tmp1[0].text)
            num=num+1

    elif(t==2):
        for item in soup.find_all('div', {"class" : "info"}):
            tmp1=item.select(".press")
            excel_sheet.write(num, 1, tmp1[0].text)
            num=num+1
            
    elif(t==3):
        for item in soup.find_all('div', {"class" : "info"}):
            tmp1=item.select(".time")
            ref_time=chk_time(tmp1[0].text)
            excel_sheet.write(num, 2, ref_time)
            num=num+1

    elif(t==4):
        for item in soup.find_all('div', {'class':'ct'}):
            '''
            tmp1=str(item).split()
            tmp1[2]=tmp1[2].replace("href=", "")
            tmp1[2]=tmp1[2].strip('"')
            '''
            tmp1=str(item).split()
            splited_item=[s for s in tmp1 if "href" in s]
            splited_item[0]=splited_item[0].replace("href=", "")
            body_URL=splited_item[0].strip('"')
            body_URL=body_URL.replace("amp;",'')

            #둘 중 하나 사용할 것. 하이퍼링크는 복붙이 안되는 문제좀 존재.......
            excel_sheet.write(num, 3, body_URL)
            #excel_sheet.write(num, 3, xlwt.Formula('HYPERLINK("%s")' % body_URL))
            num=num+1

    elif(t==5):#일부 매우 마이너한 언론은 본문 안됨
        for item in soup.find_all('div', {'class':'info'}):
            try:
                tmp1=str(item).split()
                splited_item=[s for s in tmp1 if "href" in s]
                splited_item[0]=splited_item[0].replace("href=", "")
                body_URL=splited_item[0].strip('"')
                body_URL=body_URL.replace("amp;",'')
                
                news_body_URL=urllib.request.urlopen(body_URL)
                body= BeautifulSoup(news_body_URL, 'lxml', from_encoding='utf-8')
            except:
                print('본문을 가져올 URL에 문제가 있어 공란으로 처리합니다.')
            #일반 네이버뉴스
            try:
                if body_URL is '#' : 
                    body_st=""

                elif body.find('div', id="articleBodyContents"):
                    body_st=body.find('div', id="articleBodyContents").text
                    body_st=body_st.strip()
                    body_st=body_st.replace("// flash 오류를 우회하기 위한 함수 추가", "")
                    body_st=body_st.strip()
                    body_st=body_st.replace("function _flash_removeCallback() {}", "")
                    body_st=body_st.strip()

            #연애 부문(실제로 모두 연애 내용은 아님) 네이버뉴스
                elif body.find('div', id='articeBody'):
                    body_st=body.find('div', id="articeBody").text
                    body_st=body_st.strip()
                    #print(str(body.find('div', id='articleBody')))
                    #body_st=body.find('div', id='articleBody').text

                else:
                    body_st='해당 뉴스는 스포츠/연애 부문 언론사의 본문이며, 자료를 가져올 수 없습니다.'
                
                body_st=re.sub("[^().,가-힣0-9a-zA-Z\\s]","",body_st)
                content_length=len(body_st)//30000
                content_ed=len(body_st)+1
                while(content_length >= 0):
                    content_st=content_length*30000
                    excel_sheet.write(num, 4+content_length, body_st[content_st:content_ed])
                    content_ed=content_st
                    content_length=content_length-1
                        
            except:
                print("parsing exception")
                body_st='해당 뉴스는 자료를 가져올 수 없습니다.'
            
            #print(len(body_st))

            #if len(body_st)> 30000:
            #    excel_sheet.write(num, 4, body_st[:30001])
            #    excel_sheet.write(num, 5, body_st[30000:])

            #else:
            #    excel_sheet.write(num, 4, body_st)
            #    print('seperate')
            num=num+1


    else:
        print('Try error: Scrapping')

    return num
#시간 변환기
def chk_time(tmp1):
    
    tmp1=tmp1.strip()
    
    if tmp1.count('일전') is not 0:
        print(tmp1)
        tmp1='-'+tmp1.replace('일전', '')
        c_time=datetime.datetime.now()+datetime.timedelta(int(tmp1))
        time_list=str(c_time).split()
        ymd=str(time_list[0]).split('-')
        tmp1=ymd[0]+'.'+ymd[1]+'.'+ymd[2]
        

    elif tmp1.count('시간전') is not 0:
        print(tmp1)
        tmp1='-'+tmp1.replace('시간전', '')
        c_time=datetime.datetime.now()+datetime.timedelta(hours=int(tmp1))
        time_list=str(c_time).split()
        ymd=str(time_list[0]).split('-')
        tmp1=ymd[0]+'.'+ymd[1]+'.'+ymd[2]
        

    elif tmp1.count('분전') is not 0:
        print(tmp1)
        c_time=datetime.datetime.now()
        time_list=str(c_time).split()
        ymd=str(time_list[0]).split('-')
        tmp1=ymd[0]+'.'+ymd[1]+'.'+ymd[2]
        

    return tmp1
#URL을 자동으로 변경시킨다.
def chkURL(i, URL):
    #네이버 뉴스 특성상 URL의 맨 뒤에 page='페이지번호n'이 있으며 if-elif문은 URL의 '페이지번호n'을 바꿔주는 역할을 한다.
    
    if URL.count('page=') is 0:
        URL=URL+'&page=1'
    if i >= 1 and i < 10:
        URL=URL[:-1]+str(i)
    elif i >= 10 and i < 100:
        if i==10:
            URL=URL[:-1]+str(i)
        else:
            URL=URL[:-2]+str(i)
    elif i >= 100 and i < 1000:
        if i==100:
            URL=URL[:-2]+str(i)
        else:
            URL=URL[:-3]+str(i)
    elif i >= 1000 and i < 10000:
        if i==1000:
            URL=URL[:-3]+str(i)
        else:
            URL=URL[:-4]+str(i)
    elif i >= 10000 and i < 100000:
        if i==10000:
            URL=URL[:-4]+str(i)
        else:
            URL=URL[:-5]+str(i)
    elif i >= 100000 and i < 1000000:
        if i==100000:
            URL=URL[:-5]+str(i)
        else:
            URL=URL[:-6]+str(i)
    elif i >= 1000000 and i < 10000000:
        if i==1000000:
            URL=URL[:-6]+str(i)
        else:
            URL=URL[:-7]+str(i)

    return URL
# 메인 함수
def main():

    get_text(OR_URLS=OR_URLS, OUTPUT_FILE_NAME=OUTPUT_FILE_NAME)
    
if __name__ == '__main__':
    main()
