try:
    #! selenium의 webdriver를 사용하기 위한 import
    from selenium import webdriver

    #! selenium으로 무엇인가 입력하기 위한 import
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    #! ActionChains 를 사용하기 위해서.
    from selenium.webdriver import ActionChains


    #! 페이지 로딩을 기다리는데에 사용할 time 모듈 import
    import time
    import datetime as dt
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials

    import pandas as pd
    from pyasn1.codec.der import decoder as der_decoder
    import re

    import os
    from distutils.dir_util import copy_tree
    from distutils.dir_util import remove_tree

    import sys
    #current_dir = os.path.dirname(os.path.abspath(__file__))
    #parent_dir = os.path.dirname(current_dir)
    #sys.path.append(parent_dir)  # 상위 폴더의 경로를 추가
    from my_package._func import *
    from my_package._cert import *
    from my_package._chrom_control import *
    from my_package._google_sheet import *

    from smp_email.ex_mysql import *
except Exception as e:
    print(e)
    input('1:')

######################################################################
### start
######################################################################

#import chromedriver_autoinstaller
#! 크롬드라이버 버전확인
#chrome_ver = chromedriver_autoinstaller.get_chrome_version()
#print(chrome_ver)

#chromedriver_autoinstaller.install(True)
#Chrome_path = f'./{chrome_ver.split(".")[0]}/chromedriver.exe'
url = 'https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index.xml'

CAPITAL_LIST = list(range(ord('A'),ord('A')+26))
ISSUE_SUB_FIELD = ['발전량','SMP공급가액','SMP발행일','REC수량','REC공급가액','REC발행일']

#Chrome_path = './chromedriver_win32/chromedriver.exe'
Chrome_path = 'E:/hahn/python/chromedriver_win32/chromedriver.exe'
print( os.path.isfile(Chrome_path))

#!!! log file ###
date_now = dt.datetime.now()
str_dt = date_now.strftime('%y%m%d235959Z')
rsa_from_path = '//192.168.0.150/설악서버/※태양광/사후관리/공인인증서/NPKI'
rsa_to_path = 'G:/NPKI'
auto_path = f'//192.168.0.150/설악서버/※태양광/사후관리/공인인증서/auto/{date_now.year}-{date_now.month}/{date_now.month}-{date_now.day}'

os.makedirs(auto_path, exist_ok=True)

log_file_path = f'{auto_path}/_log_{date_now.month}_{date_now.day}.txt'


### 스프레드시트 문서 가져오기 #########################################
scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]
json_file_name = 'civil-flash-378808-055ecdc1d33b.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)


spreadsheet_url = 'https://docs.google.com/spreadsheets/d/146kU3SF9taLub8eEqWrWZGLQ2dSHJiyihrN5aR7Uvhw/edit?usp=sharing'

doc = gc.open_by_url(spreadsheet_url)


### 사업2팀 에너지공단 시트 ###############
#* 공인인증서 탭
spreadsheet_url3 = 'https://docs.google.com/spreadsheets/d/1MNri81wM8mhoEqo9IenoSKhpssGvNkakevWOirirKRE'
cGS2 = cGoogleSheet(json_file_name = json_file_name, spreadsheet_url = spreadsheet_url3)
sheetnames2 = ['공인인증서']
cGS2.open_sheets(sheetnames2)
sheet_data2 = {}
sheet = sheetnames2[0]
sheet_data2['field'] = cGS2.sheets[sheet].row_values(1)
sheet_data2['lastday'] = cGS2.sheets[sheet].col_values(1)
sheet_data2['store'] = cGS2.sheets[sheet].col_values(sheet_data2['field'].index('사업자등록번호')+1)

#!## 시트 값 가져오기 ################################

#* 4. 22/23년발행완료 탭
start_row = 2
sheetname_issue = ['22년발행완료', '23년발행완료']
issue_sheet = []
issue_sheet.append(doc.worksheet(sheetname_issue[0]))
issue_sheet.append(doc.worksheet(sheetname_issue[1]))
issue_ids = {}
issue_all = {}
for i in range(len(sheetname_issue)):
    issue_all[sheetname_issue[i]] = issue_sheet[i].get_all_values()
    issue_ids[sheetname_issue[i]] = {'field1':issue_all[sheetname_issue[i]][0],
                                     'field2':issue_all[sheetname_issue[i]][1],
                                     'id':issue_sheet[i].col_values(1)}

#* 1. 미발행 탭
sheetname_incom = f'미발행'
incom_sheet = doc.worksheet(sheetname_incom)
incom_df = get_sheet_df(incom_sheet, 1)
get_field_incom = incom_df.columns.to_list()

#* 2. 대행 탭
sheetname_customer = '대행'
costomer_sheet = doc.worksheet(sheetname_customer)
get_field_costomer = ['계약번호', '발전소ID', '사업자번호', '대행여부', '종사업장번호']
costomer_record_dic = get_recode_data(costomer_sheet, get_field_costomer, start_row, 0)

#* 3. 사업자 탭
sheetname_store = '사업자'
store_sheet = doc.worksheet(sheetname_store)
store_row_list = store_sheet.col_values(1)
store_col_list = store_sheet.row_values(2)
get_field_store = store_col_list[1:]
store_record_dic = get_recode_data_dic(store_sheet, start_row, '사업자번호')

#* 5. 한전지사 탭
sheetname_kepco_sub = '한전지사'
kepco_sub_sheet = doc.worksheet(sheetname_kepco_sub)
kepco_sub_row_list = store_sheet.col_values(1)
kepco_sub_col_list = store_sheet.row_values(1)
kepco_sub_record_dic = get_recode_data_dic(kepco_sub_sheet, 1, 'ID')

#* 6. 거래처 탭
sheetname_account = '거래처'
account_sheet = doc.worksheet(sheetname_account)
account_row_list = store_sheet.col_values(1)
account_col_list = store_sheet.row_values(1)
account_dict = get_recode_data_dic(account_sheet, 1, '상호')
'''
account_dict = {'한국전력공사'  :['120-82-00052','정승일','',''],
                '한국남부발전'  :['120-86-19165','이승우','rps@kospo.co.kr',''],
                '한국동서발전'  :['120-86-19199','김영문','ewprps@ewp.co.kr',''],
                '한국서부발전'  :['120-86-19205','박형덕','rps@iwest.co.kr',''],
                '한국수력원자력':['120-86-18943','정재훈','yeorimson@khnp.co.kr','ktax@khnp.co.kr'],
                '한국중부발전'  :['120-86-19170','박형구','rec@komipo.co.kr','']}
'''

#* 7. 승인번호 탭
sheet_name = '승인번호'
approval_number_sheet = doc.worksheet(sheetname_account)
approval_number_id_list = store_sheet.col_values(1)[2:]
approval_number_id_max = max(approval_number_id_list)

##############################################################################

#! MySQLConnector 클래스 사용 예시
connector = MySQLConnector(
    host="192.168.0.150",
    user="test",
    password="oW9_82AVYf",
    database="testDB",
    port=3307
)
'''

## 수정발급
//*[@id="sub_a_0104010200"]

##승인번호 입력
//*[@id="edtAprvNo1"]

##확인
//*[@id="grp1031"]

##수정
//*[@id="textbox9973"]

##삭제
//*[@id="textbox9948"]

##상호
//*[@id="edtDmnrTnmNmTop"]

##공급가액
//*[@id="edtSumSplCftTop"]

##발급하기
//*[@id="btnIsn"]

##확인
//*[@id="trigger20"]

##인증서

'''
#!## start #########################################
driver = None

save_log(log_file_path, f'\n\n************* SMP START {date_now.strftime("%F %T")} *************\n')

max_row = incom_df.shape[0]
#! 미발행건 검색 (발행일에 공란인 것)
condition = incom_df['발행일'] == ''
incom_df_sub = incom_df[condition]
#for idx, row in incom_df.iterrows():
for idx, row in incom_df_sub.iterrows():
    log_data = row["지역"] +' ' + row["고객명"]+' ' + row["계약번호"]
    try:
        print(idx, log_data)
        i=idx
        id_registration = False

        close_driver(driver)
        driver = None

        '''
        #! 1. 미 발행건 체크
        log_proc = '미 발행건 체크'
        issuance_day = row['발행일']
        if issuance_day != '':
            print(f'\t{issuance_day} 발행완료')
            continue
        '''

        #! 시트 데이터 체크
        log_proc = '시트 데이터 체크'
        action = row['구분']

        if action != '발행':
            continue

        if ((row['구분'] == '발행') | (row['구분'] == '취소')) == False:
            raise Exception(f'구분에 발행 or 취소 아님')
        #if ((row['구분'] == '자가') | (row['구분'] == '발전')) == False:
        #if ((row['구분'] == '자가') ) == False:
        #if ( (row['구분'] == '발전')) == False:
            #raise Exception(f'구분에 자가 or 발전 아님')
            #continue
        
        if row['세금계산서\n발행처']=='한전':
            s_field_id = 'SMP발행일'
            i_field_off = 0
            account_name = '한국전력공사'
            smp_rec = 'SMP'
        else:
            account_name = row['발전사']
            s_field_id = 'REC발행일'
            i_field_off = 3
            smp_rec = 'REC'
            
        row_idx = i+start_row+1
        issuance_price = row['공급가액']
        kepco_id = row['계약번호']

        #! 계약번호 체크
        if kepco_id not in costomer_record_dic:
            raise Exception(f'{kepco_id} 대행 탭에 없음')
            #print(kepco_id + ' 대행 탭에 없음')
            #save_log(log_file_path, '\n계약번호\t{i}\t{kepco_id}\t대행 탭에 없음')
            #continue

        kepco_sub_num = costomer_record_dic[kepco_id][get_field_costomer.index('종사업장번호')]

        #! 23년발행완료 탭의 해당 월 위치 및 데이터 가져오기
        [s_y, s_m] = row['연월'].split('-')
        if (len(s_y)>2):
            s_y = s_y[-2:]
        s_y2 = s_y+'년발행완료'
        s_m2 = str(int(s_m))+'월'
        issue_row = issue_ids[s_y2]['id'].index(costomer_record_dic[kepco_id][get_field_costomer.index('발전소ID')])
        issue_col = issue_ids[s_y2]['field1'].index(s_m2)
        
        #! 공급가액 0원 체크
        log_proc = '공급가액 0원 체크'
        incom_str = re.sub(r'[^0-9]', '', issuance_price)
        if(int(incom_str)==0):
            set_data = '0원'
            print(i, set_data)
            set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, set_data)
            if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                print('\terror: set_sheet_data1')
            row['발행일'] = set_data
            continue
        
        
        
        #! 2. 대행/무상 체크
        log_proc = '대행/무상 체크'
        agency = costomer_record_dic[kepco_id][get_field_costomer.index('대행여부')]
        if not((agency =='대행') | (agency =='무상')):
        #if not((agency =='대행') | (agency =='무상') | (agency =='에공포기') | (agency =='대기중') | (agency =='일반발전소') | (agency =='발전사미체결')): # | (agency =='계약전')
        #if ((agency =='미입금') | (agency =='직접')):
            if s_y=='23':
                cell_range = get_issue_cell_range(issue_ids[s_y2], int(s_m), costomer_record_dic[kepco_id][get_field_costomer.index('발전소ID')])
                if issue_all[s_y2][issue_row][issue_col]=='':
                    time.sleep(1)
                    print (row['고객명'])
                    issue_all[s_y2][issue_row][issue_col] = int(re.sub(r'[^0-9]', '', row['발전량']))
                    set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col, issue_row+1, issue_all[s_y2][issue_row][issue_col])
                if issue_all[s_y2][issue_row][issue_col+1]=='':
                    issue_all[s_y2][issue_row][issue_col+1] = int(re.sub(r'[^0-9]', '', row['공급가액']))
                    set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+1, issue_row+1, issue_all[s_y2][issue_row][issue_col+1])
            continue
        
        
        
        store_id = costomer_record_dic[kepco_id][get_field_costomer.index('사업자번호')]
        power_id = costomer_record_dic[kepco_id][get_field_costomer.index('발전소ID')]
        store_name = row["발전소명"]
        store_kW = row["용량"]
        owner_name = row["고객명"]
        rec_sell = row["발전사"]
        
        #! 2-5 발행완료탭에서 발행 되었는지 확인
        log_proc = '발행완료탭 체크'
        s_y2 = s_y+'년발행완료'
        s_m2 = str(int(s_m))+'월'
        cell_range = get_issue_cell_range(issue_ids[s_y2], int(s_m), costomer_record_dic[kepco_id][get_field_costomer.index('발전소ID')])
        issue_cells = issue_all[s_y2][issue_row][issue_col:issue_col+6]
        
        set_data = issue_cells[ISSUE_SUB_FIELD.index(s_field_id)]
        if set_data!='':
            # 이미 세금계산서 발행 됨
            print(f'세금계산서 이미 발행 됨: {owner_name} {store_name} {set_data}')
            if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                print('sheet row error')
            continue
        
        #! 3. 사업자번호로 인증서 검색
        log_proc = '사업자번호로 인증서 검색'
        if store_id not in store_record_dic:
            err_text = f'\n인증서 없음\t{owner_name}\t{store_name}\t{store_id}\t{kepco_id}\t{rec_sell}\n'
            print(err_text)
            #save_log(log_file_path, err_text)
            set_data = '인증서 없음'
            row['발행일'] = set_data
            continue
        
        rsa_state = store_record_dic[store_id][get_field_store.index('인증서 상태')]
        rsa_last_day = store_record_dic[store_id][get_field_store.index('인증서만료일')]
        rsa_pw = store_record_dic[store_id][get_field_store.index('인증서PW')]
        rsa_folder_name = store_record_dic[store_id][get_field_store.index('folder_name')]
        if rsa_folder_name=='':
            err_text = f'\n인증서 없음\t{owner_name}\t{store_name}\t{store_id}\t{kepco_id}\t{rec_sell}\n'
            print(err_text)
            continue
        [rsa_id, rsa_folder_path] = extract_rsa_id_path(rsa_folder_name)

        
        #! 4. 인증서 복사
        log_proc = '인증서 상태 확인'
        reset_cert_D(rsa_to_path)
        rsa_folder_path = rsa_folder_path.replace('\\','/')
        from_path = f'{rsa_from_path}{rsa_folder_path}/{rsa_folder_name}'
        to_path = f'{rsa_to_path}{rsa_folder_path}/{rsa_folder_name}'
        
        from_info = check_cert(from_path, str_dt)
        
        if (from_info[1]=='없음'):
            print (' - 인증서 없음')
            if (store_record_dic[store_id][get_field_store.index('인증서 상태')]!='없음'):
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='없음')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '없음'
            continue
        
        if (rsa_last_day == from_info[1]):
            #@ 인증서 만료일이 같으면
            if ((rsa_state=='만료') | (rsa_state=='폐지') | (rsa_state=='은행용') |
                (rsa_state=='휴업') | (rsa_state=='폐업') | (rsa_state=='PW오류') |
                (rsa_state=='없음') | (rsa_state=='no folder')):
                print(f'\terror: {owner_name} 인증서 {rsa_state}')
                continue 
        else:
            #@ 새로운 인증서
            #@ 만료일 입력
            rsa_last_day = from_info[1]
            set_sheet_data1(sheet=store_sheet,
                            col_idx=store_col_list.index('인증서만료일'),
                            row_idx=store_row_list.index(store_id)+1,
                            text=from_info[1])
            store_record_dic[store_id][get_field_store.index('인증서만료일')] = rsa_last_day
            if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '':
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = ''
            
        if (from_info[0]=='만료'):
            #@ 인증서 만료됨
            print (' - 인증서 만료')
            if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '만료':
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='만료')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '만료'
            continue
        
        
        log_proc = '인증서 복사'
        try:
            copy_tree(from_path, to_path)
        except Exception as e:
            print('\t인증서 없음 (복사 필요)')
            set_data = '인증서 없음'
            #save_log(log_file_path,set_data)

            row['발행일'] = set_data
            continue
        
        
        ### 홈택스 발행 시작 ########################################
        print(f'\n\n *** start {rsa_id}\t{i}/{max_row} ***')
        
        #! 크롬드라이버 실행  (경로 예: '/Users/Roy/Downloads/chromedriver')
        log_proc = '크롬드라이버 실행'
        driver = open_driver(Chrome_path)
        
        #! 크롬 드라이버에 url 주소 넣고 실행
        log_proc = '크롬 드라이버에 url 주소 넣고 실행'
        driver.get(url)
        driver.implicitly_wait(3)   # selenium 라이브러리 자체적으로 기다려주는 방법
        time.sleep(2)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        

        #! 로그인 버튼 클릭
        log_proc = '로그인 버튼 클릭'
        #xp = '//*[@id="textbox81212912"]'
        xp = '//*[@id="textbox915"]'
        if web_btn_click(driver, xp, 2)==False:
            cancel_routine(to_path, log_file_path, f'{rsa_id} - 로그인 버튼 클릭 실패 !!')
            continue


        #! iframe 이동
        #*** 참조 https://coding-kindergarten.tistory.com/168
        #*** iframe이란 inline frame의 약자로 쉽게 말해 페이지 안의 페이지입니다.
        #*** driver.switch_to.parent_frame()          #다시 상위 frame으로 전환하는 법
        web_iframe_switch(driver, 'txppIframe', 1)


        #! 공동 인증서 버튼 클릭
        log_proc = '공동 인증서 버튼 클릭'
        if web_btn_click(driver, '//*[@id="anchor22"]', 7)==False:
            cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 버튼 클릭 실패 !!')
            continue


        #! 공인인증서 로그인
        log_proc = '공인인증서 로그인'
        popup_text = login_RSA(driver, rsa_id, rsa_pw, rsa_last_day)
        
        if popup_text==False:
            cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 로그인 실패 !!')
            continue
        elif popup_text=='해당 인증서 목록 조회에 실패하였습니다.':
            #! 인증서 폐지
            cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 로그인 실패 !!')
            set_data = popup_text
            issue_all[s_y2][issue_row][issue_col+i_field_off+2] = set_data
            if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '폐지':
                set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, '폐지')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '폐지'
            
            #* <사업2팀> 시트 [공인인증서] 탭에 만료일 쓰기(폐지)
            try:
                lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
                lastrow_idx = sheet_data2['store'].index(store_id) +1
                cGS2.set_sheet_data('공인인증서', 0, lastrow_idx, 'x폐지\n'+lastday)
            except Exception as e:
                print(f' - error - 사업2팀 시트쓰기 {e}')
            
            continue
        elif popup_text.find('은행')!=-1:
            #! 은행용 인증서
            print('error 은행용 인증서')
            if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '은행용':
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='은행용')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '은행용'
            if issue_all[s_y2][issue_row][issue_col+i_field_off+2] != '은행용':
                issue_all[s_y2][issue_row][issue_col+i_field_off+2] = '은행용'
                set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, '은행용')
                
            #* <사업2팀> 시트 [공인인증서] 탭에 만료일 쓰기(은행용)
            try:
                lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
                lastrow_idx = sheet_data2['store'].index(store_id) +1
                cGS2.set_sheet_data('공인인증서', 0, lastrow_idx, 'x은행용\n'+lastday)
            except Exception as e:
                print(f' - error - 사업2팀 시트쓰기 {e}')
            continue
            
        
        #! Alert 창 확인 클릭 
        try:
            result = driver.switch_to.alert
            alert_text = result.text
            print(alert_text)
            if alert_text[0:9] == '홈택스에 등록된 인증서가 아닙니다.'[0:9]:
                id_registration = True
                result.accept()
            elif alert_text[0:9] == '선택하신 인증서는 폐지된 인증서입니다.'[0:9]:
                log_text = f'\n{rsa_id}\t{alert_text}'
                cancel_routine(to_path, log_file_path, log_text)
                if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '폐지':
                    store_record_dic[store_id][get_field_store.index('인증서 상태')] = '폐지'
                    set_sheet_data1(sheet=store_sheet,
                                    col_idx=store_col_list.index('인증서 상태'),
                                    row_idx=store_row_list.index(store_id)+1,
                                    text='폐지')
                
                #* <사업2팀> 시트 [공인인증서] 탭에 만료일 쓰기
                try:
                    lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
                    lastrow_idx = sheet_data2['store'].index(store_id) +1
                    cGS2.set_sheet_data('공인인증서', 0, lastrow_idx, 'x폐지\n'+lastday)
                except Exception as e:
                    print(f' - error - 사업2팀 시트쓰기 {e}')
                result.accept()
                continue
            elif alert_text[0:9] == '[ETICMZ0008]전자서명 검증에 실패하였습니다.'[0:9]:
                result.accept()
                continue
            else:
                log_text = f'\n{rsa_id}\t{alert_text}'
                cancel_routine(to_path, log_file_path, log_text)
                set_data = alert_text
                issue_all[s_y2][issue_row][issue_col+i_field_off] = int(re.sub(r'[^0-9]', '', row['발전량']))
                issue_all[s_y2][issue_row][issue_col+i_field_off+1] = int(re.sub(r'[^0-9]', '', row['공급가액']))
                issue_all[s_y2][issue_row][issue_col+i_field_off+2] = set_data
                set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, issue_all[s_y2][issue_row][issue_col+i_field_off+2])
                if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, alert_text)==False:
                    print('sheet row error')
                row['발행일'] = set_data
                result.accept()
                continue
        except:
            print('공동인증서 로그인 성공')
            
        
        #! <사업2팀 에너지공단> 시트 [공인인증서] 탭에 만료일자 확인
        try:
            lastrow_idx = sheet_data2['store'].index(store_id)
            lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
            if lastday != sheet_data2['lastday'][lastrow_idx]:
                lastday_day = make_datetime(lastday)
                cGS2.set_sheet_data('공인인증서', 0, lastrow_idx+1, lastday_day)
        except Exception as e:
            print(f' - error - 사업2팀 {e}')
        
        #! 암호오류 확인
        if check_pw(driver, rsa_id, log_file_path)==False:
            set_data = driver.find_element(By.XPATH,'//*[@id="alert_msg"]').text
            if ((set_data[:9] == '인증서 로그인에 실패하였습니다.'[:9]) | 
                (set_data[:9] == '홈택스 이용자 증가로 서비스 지연이 발생하고 있습니다.'[:9])):
                print(f"\terror: {set_data}")
            else:
                #if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                #    print('sheet row error')
                #row['발행일'] = set_data
                #issue_all[s_y2][issue_row][issue_col+i_field_off+2] = set_data
                if store_record_dic[store_id][get_field_store.index('인증서 상태')] != 'PW오류':
                    set_sheet_data1(sheet=store_sheet,
                                    col_idx=store_col_list.index('인증서 상태'),
                                    row_idx=store_row_list.index(store_id)+1,
                                    text='PW오류')
                    store_record_dic[store_id][get_field_store.index('인증서 상태')] = 'PW오류'
            cancel_routine(to_path, log_file_path, set_data)
            continue
        time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        ###------------------------------------------------
        
        #!### 공동인증서 등록 루틴 ###############
        if id_registration:
            log_proc = '공동인증서 등록 루틴'
            web_iframe_switch(driver, 'txppIframe', 1)
            
            #! 공동인증서 - 사업자번호 등록 ###########################################
            store_id_sub = store_id.split('-')
            driver.find_element(By.XPATH, '//*[@id="edtBsno1"]').send_keys(store_id_sub[0])
            driver.find_element(By.XPATH, '//*[@id="edtBsno2"]').send_keys(store_id_sub[1])
            driver.find_element(By.XPATH, '//*[@id="edtBsno3"]').send_keys(store_id_sub[2])
            
            if web_btn_click(driver, '//*[@id="btnRgtBman"]', 7)==False:
                cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 등록 버튼 클릭 실패1 {store_id}!!')
                continue
            
            #! 공인인증서 확인
            if login_RSA(driver, rsa_id, rsa_pw, rsa_last_day)==False:
                cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 등록 실패2 {store_id}!!')
                continue
            
            #! Alert 창 확인 클릭 
            try:
                result = driver.switch_to.alert
                alert_text = result.text
                print(alert_text)
                if alert_text[0:9] != '인증서가 정상적으로 등록되었습니다.'[0:9]:
                    log_text = f'\n{rsa_id}\t{alert_text}'
                    cancel_routine(to_path, log_file_path, log_text)
                    result.accept()
                    continue
                result.accept()
            except:
                print('인증서 등록 실패')
                continue
            time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
            
            #! 다시 로그인
            #! 로그인 버튼 클릭
            xp = '//*[@id="textbox915"]'
            if web_btn_click(driver, xp, 2)==False:
                cancel_routine(to_path, log_file_path, f'{rsa_id} - 로그인 버튼 클릭 실패 !!')
                continue


            #! iframe 이동
            # 참조 https://coding-kindergarten.tistory.com/168
            # iframe이란 inline frame의 약자로 쉽게 말해 페이지 안의 페이지입니다.
            # driver.switch_to.parent_frame()          #다시 상위 frame으로 전환하는 법
            web_iframe_switch(driver, 'txppIframe', 1)


            #! 공동 인증서 버튼 클릭
            if web_btn_click(driver, '//*[@id="anchor22"]', 7)==False:
                cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 버튼 클릭 실패 !!')
                continue


            #! 공인인증서 로그인
            log_proc = '공동인증서 로그인'
            if login_RSA(driver, rsa_id, rsa_pw, rsa_last_day)==False:
                cancel_routine(to_path, log_file_path, f'{rsa_id} 공동 인증서 로그인 실패 !!')
                continue
            
            
            #! Alert 창 확인 클릭 
            try:
                result = driver.switch_to.alert
                alert_text = result.text
                print(alert_text)
                if alert_text[0:9] != '홈택스에 등록된 인증서가 아닙니다.'[0:9]:
                    log_text = f'\n{rsa_id}\t{alert_text}'
                    cancel_routine(to_path, log_file_path, log_text)
                    set_data = log_text
                    if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                        print('sheet row error')
                    row['발행일'] = set_data
                    result.accept()
                    continue
                result.accept()
                id_registration = True
            except:
                print('공동인증서 로그인 성공')
            time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        #^### 공동인증서 등록 루틴 끝, 재로그인 완료 ###############
        ### ------------------------------------------------------
        
        homtax_id = rsa_id
        login_time = dt.datetime.now()
        print(f"{login_time.hour}:{login_time.minute} log in [{homtax_id}]")
        #save_log(log_file_path, f"\n{login_time.hour}:{login_time.minute} log in [{homtax_id}]")
        
        
        #! 조회/발급 클릭
        log_proc = '상단 메뉴 -> 전자(세금)계산서 클릭' #'조회/발급 클릭'
        #xp = '//*[@id="textbox81212923"]'
        xp = '//*[@id="hdTextbox548"]'
        if web_btn_click(driver, xp, 2)==False:
            if web_btn_click(driver, '//*[@id="hdTextbox544"]', 2)==False:
                cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 조회/발급 클릭 실패')
                continue

        #! iframe
        web_iframe_switch(driver, 'txppIframe', 1)
        
        '''
        #! 발급 클릭
        log_proc = '발급 클릭'
        if web_btn_click(driver, '//*[@id="sub_a_0104010000"]', 2)==False:
            cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 발급 클릭 실패')
            continue
        '''

        #! 건별 발급 클릭
        log_proc = '건별 발급 클릭'
        #xp = '//*[@id="sub_a_0104010100"]'
        xp = '//*[@id="group23231569"]'
        if web_btn_click(driver, xp, 3)==False:
            cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 건별 발급 클릭 실패')
            continue
        ### ------------------------------------------------------
        '''
        try:
            if web_iframe_switch(driver, 'UTEETZZD02_iframe', 1)==False:
                raise Exception('정상')
        
            message = driver.find_element(By.XPATH, '//*[@id="textbox18"]/span').text
            if (message[0:9] == '귀하께서 로그인한 ‘인증서(개인범용/금융·신용카드·보험거래용 등)’는 [전자(세금)계산서 발급 가능 인증서]가 아닙니다.'[0:9]):
                print(message)
                set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, message)
                if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, message)==False:
                    print('sheet row error')
                row['발행일'] = message
                cancel_routine(to_path, log_file_path, message)
                continue
        except:
            dbg=1
        ### ------------------------------------------------------
        '''
        
        #! iframe
        web_iframe_switch(driver, 'txppIframe', 1)

        log_proc = 'save data to store sheet'
        #* 사업자 등록번호
        store_number = ''
        xp = '//*[@id="edtSplrTxprNoDispTop"]'
        while True:
            store_number = driver.find_element(By.XPATH,xp).text
            if store_number != '':
                break
            time.sleep(1)

        #* 상호
        xp = '//*[@id="edtSplrTnmNmTop"]'
        store_name = driver.find_element(By.XPATH,xp).get_attribute('value')  #^ input box에 입력된 값 가져오기
        #* 성명 
        xp = '//*[@id="edtSplrRprsFnmTop"]'
        owner_name = driver.find_element(By.XPATH,xp).get_attribute('value')  #^ input box에 입력된 값 가져오기
        #* 사업장 
        xp = '//*[@id="edtSplrPfbAdrTop"]'
        store_address = driver.find_element(By.XPATH,xp).get_attribute('value')  #^ input box에 입력된 값 가져오기

        print (f'{store_number} / {store_name} / {owner_name} / {store_address}')
        
        if store_id != store_number:
            set_data = '사업자번호 오류'
            
            #set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, set_data)
            if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                print('sheet row error')
            raise Exception(f'{set_data} - 시트:{store_id} 홈택스:{store_id}')
            #row['발행일'] = set_data
            #cancel_routine(to_path, log_file_path, f'\nsheet - {store_id}\t 홈택스 - {store_number}')
            #continue
        
        #! save store information to sheet
        row_id = store_row_list.index(store_number)+1
        if (store_record_dic[store_id][get_field_store.index('상호')]!=store_name):
            set_sheet_data1(sheet=store_sheet, 
                            col_idx=store_col_list.index('상호'),
                            row_idx=row_id,
                            text=store_name)
            #update_sheet_data(store_sheet, store_col['상호']+str(row_id), store_name)
        if (store_record_dic[store_id][get_field_store.index('성명')]!=owner_name):
            set_sheet_data1(sheet=store_sheet, 
                            col_idx=store_col_list.index('성명'),
                            row_idx=row_id,
                            text=owner_name)
            #update_sheet_data(store_sheet, store_col['성명']+str(row_id), owner_name)
        if (store_record_dic[store_id][get_field_store.index('사업장주소')]!=store_address):
            set_sheet_data1(sheet=store_sheet, 
                            col_idx=store_col_list.index('사업장주소'),
                            row_idx=row_id,
                            text=store_address)
            #update_sheet_data(store_sheet, store_col['사업장주소']+str(row_id), store_address)
        #! mySQL 저장 추가
        ####---------------------------------------------------
            
            
        while True:
            #! 거래처 조회 클릭 
            log_proc = '거래처 조회 클릭'
            cnt1 = 5;
            
            while True:
                xp = '//*[@id="btnDmnrClplcInqrTop"]'
                web_btn_click(driver, xp, 3)
                
                #! iframe
                xp = 'clplcInqrPopup_iframe'
                if web_iframe_switch(driver, xp, 2):
                    break
                else:
                    cnt1 = cnt1-1
                    if cnt1<=0:
                        break
                    cancel_routine(to_path, log_file_path, f'\n{rsa_id}\terror: 거래처 조회 실패')
                    continue

            if cnt1<=0:
                print('\terror: 거래처 조회 실패')
                cancel_routine(to_path, log_file_path, f'\n{rsa_id}\terror: 거래처 조회 실패')
                continue
            ###-----------------------------------------------------------------------

            #! 거래처명 입력
            log_proc = '거래처명 입력'
            xp = '//*[@id="edtTxprNm"]'
            driver.find_element(By.XPATH, xp).send_keys(account_name)
            driver.implicitly_wait(1)   # selenium 라이브러리 자체적으로 기다려주는 방법
            ###-----------------------------------------------------------------------


            #! 조회하기 클릭
            log_proc = '조회하기 클릭'
            if web_btn_click(driver, '//*[@id="btnSearch"]', 1)==False:
                cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 한전 조회하기 클릭 실패')
                continue
            ###-----------------------------------------------------------------------


            #! 한전 선택
            log_proc = '한전 선택'
            error = ''
            if web_btn_click(driver, '//*[@id="G_grdResult___radio_chk_0"]', 1)==False:
                set_data = '거래처 조회 실패'
                
                #! --- 거래처 등록 루틴 --------------------------------
                log_proc = '거래처 등록'
                
                #! 닫기 클릭
                web_btn_click(driver, '//*[@id="btnClose"]', 1)
                
                #! 창 이동
                web_iframe_switch(driver, 'txppIframe', 1)
    
                #! 거래처 관리 클릭 
                web_btn_click(driver, '//*[@id="btnClplcInqrTop"]', 1)
                
                #! 건별 등록 클릭 
                web_btn_click(driver, '//*[@id="textbox1395"]', 1)
                
                #! 사업자번호 입력
                account_num = account_dict[account_name][0]
                web_send_key1(driver, '//*[@id="txtBsno1"]', account_num.replace('-', ''))
                
                #! 확인 클릭
                web_btn_click(driver, '//*[@id="btnValidCheck"]', 1)

                #if account_name == '한국전력공사' and kepco_sub_num == '':
                if account_name == '한국전력공사':
                    #! 종사업장 선택 창으로 이동
                    web_iframe_switch(driver, "ABTIBsnoUnitPopup2_iframe", 1)
                    

                    #! 종된사업장 일련번호 입력
                    xp = '//*[@id="iptMpbSn"]'
                    web_send_key1(driver, xp, kepco_sub_num)

                    #! 조회하기 클릭
                    xp = '//*[@id="btnSearch"]'
                    web_btn_click(driver, xp, 1)

                    #! 라디오 버튼 선택
                    xp = '//*[@id="grid1_cell_0_0"]/input'
                    web_btn_click(driver, xp, 1)

                    #! 선택 클릭
                    xp = '//*[@id="trigger66"]'
                    web_btn_click(driver, xp, 1)

                    #! 이메일 쓰기
                    account_dict[account_name][2] = kepco_sub_record_dic[kepco_sub_num][1]
                    account_dict[account_name][3] = kepco_sub_record_dic[kepco_sub_num][2]

                        
                    #! alert
                    try:
                        result = driver.switch_to.alert
                        print(result.text)
                        if result.text != '정상적인 사업자번호입니다.':
                            error = '거래처 등록 실패'
                            break
                        result.accept()
                        time.sleep(1)
                        #result.dismiss()
                    except:
                        error = '거래처 등록 실패'
                        break

                    #! 창이동
                    web_iframe_switch(driver, 'txppIframe', 1)

                    #cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 한전 선택 실패')
                    #if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                    #    print('sheet row error')
                    #row['발행일'] = set_data
                    #break
                    
                else:
                        
                    #! alert
                    try:
                        result = driver.switch_to.alert
                        print(result.text)
                        if result.text != '정상적인 사업자번호입니다.':
                            error = '거래처 등록 실패'
                            break
                        result.accept()
                        time.sleep(1)
                        #result.dismiss()
                    except:
                        error = '거래처 등록 실패'
                        break
                    
                #! 상호 입력
                web_send_key1(driver, '//*[@id="txtTnmNm"]', account_name)
                time.sleep(0.5)
                
                #! 대표자 입력
                account_ceo = account_dict[account_name][1]
                web_send_key1(driver, '//*[@id="txtRprs"]', account_ceo)
                time.sleep(0.5)
                
                #! 주담당자이메일 입력
                account_email = account_dict[account_name][2].split('@')
                web_send_key1(driver, '//*[@id="txtChrgEmlAdr1"]', account_email[0])
                time.sleep(0.5)
                web_send_key1(driver, '//*[@id="txtChrgEmlAdr2"]', account_email[1])
                time.sleep(0.5)
                
                if account_dict[account_name][3]!='':
                    #! 부담당자이메일 입력
                    account_email = account_dict[account_name][3].split('@')
                    web_send_key1(driver, '//*[@id="txtSchrgEmlAdr1"]', account_email[0])
                    time.sleep(0.5)
                    web_send_key1(driver, '//*[@id="txtSchrgEmlAdr2"]', account_email[1])
                    time.sleep(0.5)
                
                #! 등록하기 클릭
                web_btn_click(driver, '//*[@id="btnRgt"]', 1)
                
                #! alert
                try:
                    result = driver.switch_to.alert
                    print(result.text)
                    if result.text != '거래처 정보를 등록하시겠습니까?':
                        error = '거래처 등록 실패 1'
                        break
                    result.accept()
                    time.sleep(0.5)
                    #result.dismiss()
                except:
                    error = '거래처 등록 실패 2'
                    break
                
                #! alert
                try:
                    result = driver.switch_to.alert
                    print(result.text)
                    if result.text[:9] != '거래처 정보가 성공적으로 등록되었습니다.'[:9]:
                        error = '거래처 등록 실패 3'
                        break
                    #result.accept()
                    result.dismiss()
                    time.sleep(0.5)
                except:
                    error = '거래처 등록 실패 4'
                    break
                
                #! alert
                try:
                    result = driver.switch_to.alert
                    print(result.text)
                    if result.text[:9] != '거래처 담당자를 추가 등록하시겠습니까?'[:9]:
                        error = '거래처 등록 실패 5'
                        break
                    #result.accept()
                    result.dismiss()
                    time.sleep(0.5)
                except:
                    error = '거래처 등록 실패 6'
                    break
                
                #! 목록 클릭
                web_btn_click(driver, '//*[@id="btnList"]', 1)
                
                #! 건별발급 이동 클릭
                web_btn_click(driver, '//*[@id="btnIsnMove"]', 1)
                time.sleep(3)
                continue
            else:
                break
        ###-----------------------------------------------------------------------

        if error != '':
            if error[:-2] == '거래처 등록 실패':
                raise Exception(error)
        
        #! 확인 클릭 
        log_proc = '확인 클릭'
        if web_btn_click(driver, '//*[@id="btnProcess"]', 2)==False:
            cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 확인 클릭 실패')
            continue
        ###-----------------------------------------------------------------------


        #! Alert 창 확인 클릭 
        log_proc = 'Alert 창 확인 클릭'
        try:
            result = driver.switch_to.alert
            print(result.text)
            result.accept()
            #result.dismiss()
        except:
            print("한전 검색 실패")
        time.sleep(2)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        ###-----------------------------------------------------------------------


        #! iframe
        web_iframe_switch(driver, 'txppIframe', 1)
        #! 스크롤 높이 가져옴
        last_height = driver.execute_script("return document.body.scrollHeight")
        ###-----------------------------------------------------------------------




        #! 한전 종사업장번호 체크
        if account_name == '한국전력공사':
            log_proc = '한전 종사업장번호 체크'
            xp = '//*[@id="edtDmnrMpbNoTop"]'
            kepco_sub_num_homtax = driver.find_element(By.XPATH, xp).text
            if kepco_sub_num != kepco_sub_num_homtax:
                print (f'종사업장번호 오류: 시트={kepco_sub_num}, 홈텍스={kepco_sub_num_homtax}')
                error_data = []
                error_data.append(log_data)
                error_data.append(f'\terror: {log_proc}')
                error_data.append(f'\t종사업장번호 오류: 시트={kepco_sub_num}, 홈텍스={kepco_sub_num_homtax}\n')
                log_data_str = '\n'.join(error_data)  # 리스트의 각 요소를 개행 문자('\n')로 연결하여 하나의 문자열로 만듦
                save_log(log_file_path, log_data_str)
                continue

        #! 일자 입력
        log_proc = '계산서 내용 입력'
        date_now = dt.datetime.now()
        auto_path = f'//192.168.0.150/설악서버/※태양광/사후관리/공인인증서/auto/{date_now.year}-{date_now.month}/{date_now.month}-{date_now.day}'
        os.makedirs(auto_path, exist_ok=True)
        log_file_path = f'{auto_path}/_log_{date_now.month}_{date_now.day}.txt'

        driver.find_element(By.XPATH, '//*[@id="genEtxivLsatTop_0_edtLsatSplDdTop"]').send_keys(date_now.day)
        driver.implicitly_wait(1)   # selenium 라이브러리 자체적으로 기다려주는 방법

        #! 품목명 입력
        write_month2 = row['연월']
        item_name = f"{store_name}({store_kW}kW) - {write_month2}월분 {smp_rec}"
        driver.find_element(By.XPATH, '//*[@id="genEtxivLsatTop_0_edtLsatNmTop"]').send_keys(item_name)
        driver.implicitly_wait(1)   # selenium 라이브러리 자체적으로 기다려주는 방법

        #! [공급가액] 입력
        try:
            incom_value = incom_value[0:incom_value.index('.')]
        except:
            pass
        incom_value = re.sub(r'[^0-9]', '', issuance_price)
        driver.find_element(By.XPATH, '//*[@id="genEtxivLsatTop_0_edtLsatSplCftTop"]').send_keys(incom_value)
        driver.implicitly_wait(1)   # selenium 라이브러리 자체적으로 기다려주는 방법
        
        #! 공급밭는자 성명 입력
        if (smp_rec == 'SMP'):
            driver.find_element(By.XPATH, '//*[@id="edtDmnrRprsFnmTop"]').clear()
            driver.find_element(By.XPATH, '//*[@id="edtDmnrRprsFnmTop"]').send_keys('정승일')
            driver.implicitly_wait(1)   # selenium 라이브러리 자체적으로 기다려주는 방법
        
        ###---------------------------------------------------
        
            

        #! [발급하기] 버튼 클릭
        log_proc = '[발급하기] 버튼 클릭'
        #driver.find_element(By.ID,'btnIsn').click()
        driver.find_element(By.XPATH,'//*[@id="btnIsn"]').click()
        driver.implicitly_wait(3)   # selenium 라이브러리 자체적으로 기다려주는 방법
        time.sleep(2)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        #---------------------------------------------------
        
        
        #! Alert 창 확인 클릭 
        Alert_text = GetAlert(driver)
        if Alert_text == '':
            print ('[발급하기] 버튼 클릭 성공')
        else:
            if (set_data == '귀사업자는 휴업중인 사업자입니다. 발급하시겠습니까?'):
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='휴업')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '휴업'
                
                #! <사업2팀> 시트 [공인인증서] 탭에 만료일 쓰기
                try:
                    lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
                    lastrow_idx = sheet_data2['store'].index(store_id) +1
                    cGS2.set_sheet_data('공인인증서', 0, lastrow_idx, 'x휴업\n'+lastday)
                except Exception as e:
                    print(f' - error - 사업2팀 {e}')
                #continue
            elif (set_data[0:9] == '공급자(또는 수탁사업자)가 폐업상태입니다.'[0:9]):
                set_sheet_data1(sheet=store_sheet,
                                col_idx=store_col_list.index('인증서 상태'),
                                row_idx=store_row_list.index(store_id)+1,
                                text='폐업')
                store_record_dic[store_id][get_field_store.index('인증서 상태')] = '폐업'
                #! <사업2팀> 시트 [공인인증서] 탭에 만료일 쓰기
                try:
                    lastday = '20' + from_info[1][:2] + '-' + from_info[1][2:4] + '-' + from_info[1][4:6]
                    lastrow_idx = sheet_data2['store'].index(store_id) +1
                    cGS2.set_sheet_data('공인인증서', 0, lastrow_idx, 'x폐업\n'+lastday)
                except Exception as e:
                    print(f' - error - 사업2팀 {e}')
                #continue
            else:
                issue_all[s_y2][issue_row][issue_col+i_field_off+2] = set_data
                set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], 
                                issue_col+i_field_off+2, 
                                issue_row+1, 
                                issue_all[s_y2][issue_row][issue_col+i_field_off+2])
                if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
                    print('sheet row error')
                row['발행일'] = set_data
                #cancel_routine(to_path, log_file_path, result.text)
                #result.dismiss()
                #continue
            result.accept()
            raise Exception(set_data)
        time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        #---------------------------------------------------


        #! 전자세금계산서 발급 창 iframe 이동
        log_proc = '전자세금계산서 발급 확인 창'
        web_iframe_switch(driver, 'UTEETZZA89_iframe', 2)
        #---------------------------------------------------


        #! 발행금액 체크
        price = driver.find_element(By.XPATH,'//*[@id="textbox1012"]').text
        price = re.sub(r'[^0-9]', '', price)
        #print(f'발행 공급가액: {incom_value} / {price}')
        if incom_value != price:
            raise Exception(f'발행금액 오류: {incom_value} / {price}')
            #cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 발행금액 오류')
            #continue
        #---------------------------------------------------


        #! [확인] 버튼 클릭
        if web_btn_click(driver, '//*[@id="trigger20"]', 2)==False:
            raise Exception('전자세금계산서 발급 [확인] 클릭 실패')
            #cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 전자세금계산서 발급 [확인] 클릭 실패')
            #continue
        

        driver.execute_script("window.scrollTo(0,0)")
        driver.implicitly_wait(2)   # selenium 라이브러리 자체적으로 기다려주는 방법
        #---------------------------------------------------

        #! iframe
        web_iframe_switch(driver, 'txppIframe',2)
        driver.execute_script("window.scrollTo(0,0)")
        driver.implicitly_wait(2)   # selenium 라이브러리 자체적으로 기다려주는 방법
        #---------------------------------------------------
        
        #! 공인인증서 확인
        log_proc = '공인인증서 확인'
        last_day = login_RSA(driver, rsa_id, rsa_pw, rsa_last_day, 1)
        if last_day==False:
            raise Exception(f'{rsa_id} 공동 인증서 확인 실패 {store_id}!!')
        #---------------------------------------------------
        
        #! Alert 창 확인 클릭 
        try:
            result = driver.switch_to.Alert
            #print(result.text)
            #cancel_routine(to_path, log_file_path, result.text)
            result.accept()
            #result.dismiss()
            raise Exception(result.text)
            #continue
        except:
            print("공인인증서 확인 성공")
        time.sleep(2)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        #---------------------------------------------------



        driver.switch_to.parent_frame()
        time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        web_iframe_switch(driver, 'isnCmplPopup_iframe', 1)
        time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        
        #! 승인번호
        confirm_num = driver.find_element(By.XPATH, '//*[@id="txtMsg1"]/span').text
        confirm_num = confirm_num.replace(': ','')
        print(confirm_num)


        #! 시트에 세금계산서 발행일 작성
        log_proc = '시트에 세금계산서 발행일 작성'
        if store_record_dic[store_id][get_field_store.index('인증서 상태')] != '정상':
            set_sheet_data1(sheet=store_sheet,
                            col_idx=store_col_list.index('인증서 상태'),
                            row_idx=store_row_list.index(store_id)+1,
                            text='정상')
            store_record_dic[store_id][get_field_store.index('인증서 상태')] = '정상'
            
        set_data = f'{date_now.month}-{date_now.day}'
        if set_sheet_data1(incom_sheet, get_field_incom.index('발행일'), row_idx, set_data)==False:
            print('sheet row error')
        issue_all[s_y2][issue_row][issue_col+i_field_off+2] = set_data #make_datetime(set_data)
        set_sheet_data1(issue_sheet[sheetname_issue.index(s_y2)], issue_col+i_field_off+2, issue_row+1, issue_all[s_y2][issue_row][issue_col+i_field_off+2])
        


        #! 전자세금계산서 발급 [확인] 클릭
        log_proc = '전자세금계산서 발급 [확인] 클릭'
        if web_btn_click(driver, '//*[@id="btnClose1"]', 2)==False:
            raise Exception('전자세금계산서 발급 [확인] 클릭 실패')
            #print ('error web_btn_click')
            #cancel_routine(to_path, log_file_path, f'\n{rsa_id}\tcontinue: 전자세금계산서 발급 [확인] 클릭 실패')
            #continue

        row['발행일'] = set_data
        #save_log(log_file_path, f'\t{store_number}\t{store_name}\t{owner_name}\t발행 성공\t{price}원\t{confirm_num}')


        log_proc = '스크린샷'
        close_popup_window(driver)
        
        web_iframe_switch(driver, 'txppIframe', 1)
        time.sleep(.5)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법

        #! 출력 클릭 //*[@id="trigger29"]
        #driver.find_element(By.XPATH, '//*[@id="trigger29"]').send_keys(Keys.ENTER)
        driver.find_element(By.XPATH, '//*[@id="trigger29"]').click()
        #web_btn_click(driver, '//*[@id="trigger29"]', 1)
        time.sleep(2)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법

        parent_h = chang_popup_window(driver, 'Report')
        time.sleep(1)               # time 라이브러리에 코드를 직접적으로 쉬어주는 방법
        
        #! 스크린샷
        save_screenshot_path = f'{auto_path}/{owner_name} {item_name} {power_id} {confirm_num}.png'
        driver.save_screenshot(save_screenshot_path)
        
        #! remove
        remove_tree(to_path)
        log_data = log_data + f'\t: {smp_rec} {write_month2}월분 발행 완료\n'
        save_log(log_file_path, log_data)
        
        driver.quit()
        driver = None
        
        
    except Exception as e:
        error_data = []
        error_data.append(log_data)
        error_data.append(f'\terror: {log_proc}')
        error_data.append(f'\t{e}')
        log_data_str = '\n'.join(error_data)  # 리스트의 각 요소를 개행 문자('\n')로 연결하여 하나의 문자열로 만듦
        save_log(log_file_path, log_data_str)
#---------------------------------------------------


try:
    close_driver(driver)
except Exception as e:    # 모든 예외의 에러 메시지를 출력할 때는 Exception을 사용
    print('예외가 발생했습니다.', e)

print('END')
sys.exit()