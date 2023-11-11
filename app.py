import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# 웹 드라이버 경로
DRIVER_PATH = 'C:/Users/Gony/OneDrive/바탕 화면/code/chromedriver_win32/chromedriver.exe'

service = Service(DRIVER_PATH)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)  # 브라우저 창을 유지하도록 설정
driver = webdriver.Chrome(service=service, options=options)

# 웹사이트 접근
driver.get('https://www.foodsafetykorea.go.kr/portal/specialinfo/searchInfoCompany.do?menu_grp=MENU_NEW04&menu_no=2813')

# 초기 DataFrame 생성
df = pd.DataFrame()

# -----------------------------여기부터 남양주시 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_41")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("남양주시")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 남양주 데이터를 DataFrame으로 변환
df_namyangju = pd.DataFrame(data, columns=headers)

# 남양주 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_namyangju])

driver.refresh()

# ------------------------------ 여기까지 남양주시 -----------------------------
# ------------------------------ 여기부터 구리시 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_41")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("구리시")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 구리 데이터를 DataFrame으로 변환
df_guri = pd.DataFrame(data, columns=headers)

# 구리 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_guri])

driver.refresh()


# ------------------------------ 여기까지 구리시 -----------------------------

# ------------------------------ 여기부터 하남시 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_41")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("하남시")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 구리 데이터를 DataFrame으로 변환
df_hanam = pd.DataFrame(data, columns=headers)

# 구리 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_hanam])

driver.refresh()


# ------------------------------ 여기까지 하남시 -----------------------------

# ------------------------------ 여기부터 송파구 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_11")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("송파구")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 구리 데이터를 DataFrame으로 변환
df_Songpa = pd.DataFrame(data, columns=headers)

# 구리 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_Songpa])

driver.refresh()


# ------------------------------ 여기까지 송파구 -----------------------------

# ------------------------------ 여기부터 성동구 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_11")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("성동구")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 구리 데이터를 DataFrame으로 변환
df_Seongdong = pd.DataFrame(data, columns=headers)

# 구리 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_Seongdong])

driver.refresh()


# ------------------------------ 여기까지 성동구 -----------------------------

# ------------------------------ 여기부터 광진구 -----------------------------

# 체크박스 클릭
checkbox = driver.find_element(By.ID, "chk_sido1_11")
checkbox.click()

# '클릭' 부분
click_element = driver.find_element(By.CSS_SELECTOR, "#mode1 > div.ddCheck2c > div.dsL > ul > li:nth-child(4) > a")
click_element.click()

# 텍스트 입력
input_field = driver.find_element(By.CLASS_NAME, "w100")
input_field.send_keys("광진구")

# 드롭다운 메뉴 열기
driver.find_element(By.ID, "a_list_cnt").click()

# 값이 "50"인 항목 클릭
driver.find_element(By.CSS_SELECTOR, '[val="50"]').click()

# 헤더 데이터를 가져오는 부분
headers = []
header_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > thead > tr > th')))
for header in header_elements:
    headers.append(header.text)

# 데이터를 가져오는 부분
data = []
row_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#tbl_bsn_list > tbody > tr')))

for i in range(1, 51):  # 각 행에는 9개의 셀이 있습니다
    row_data = []
    for j in range(1, 10):
        try:
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'#tbl_bsn_list > tbody > tr:nth-child({i})')))
            if j == 3:  # '업소명'이 포함된 셀
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span > a').text
            else:
                cell_text = row_element.find_element(By.CSS_SELECTOR, f'td:nth-child({j}) > span.table_txt').text
            row_data.append(cell_text)
        except NoSuchElementException:
            row_data.append('')
    data.append(row_data)

# 구리 데이터를 DataFrame으로 변환
df_Gwangjin = pd.DataFrame(data, columns=headers)

# 구리 데이터를 전체 DataFrame에 추가
df = pd.concat([df, df_Gwangjin])


# ------------------------------ 여기까지 광진구 -----------------------------

# 데이터를 엑셀 파일로 저장
df.to_excel("data.xlsx", index=False)

# 브라우저 종료
driver.quit()























# 다음은 페이지 순회를 위한 코드입니다. 필요할 때 주석을 해제하고 사용하세요.
# try:
#     # 다음 페이지 버튼을 찾아 클릭합니다.
#     next_page_button = driver.find_element(By.CSS_SELECTOR, "#contents > main > section > div.board-footer > div > ul > li:nth-child(7) > a")
#     next_page_button.click()
#     # 페이지 로딩을 기다립니다.
#     time.sleep(5)
# except NoSuchElementException:
#     # 다음 페이지 버튼이 없다면, 마지막 페이지이므로 루프를 종료합니다.
#     break