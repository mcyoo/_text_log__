import openpyxl
from openpyxl.styles import PatternFill, Color, Font
import codecs
import os

# 사용자가 입력한 파일 불러오기
search_line_start = "↓↓↓ 찾으시는 단어 또는 문장을 한 줄씩 입력해주세요. (start) 예시 : 김치볶음밥 ↓↓↓"
search_line_end = "↑↑↑ 찾으시는 단어 또는 문장을 한 줄씩 입력해주세요. (end) ↑↑↑"

filter_line_start = "↓↓↓ 시작하는 단어 또는 문장 ~ 끝나는 단어 또는 문장을 입력해주세요. (start) 예시 : 김치볶음밥~간장밥 ↓↓↓ "
filter_line_end = "↑↑↑ 시작하는 단어 또는 문장 ~ 끝나는 단어 또는 문장을 입력해주세요. (end) ↑↑↑"

filter_search_line_start = "↓↓↓ 시작하는 단어 또는 문장 ~ 찾는 단어 또는 문장 ~ 끝나는 단어 또는 문장을 입력해주세요. (start) 예시 : 김치볶음밥~김치~간장밥 ↓↓↓ "
filter_search_line_end = "↑↑↑ 시작하는 단어 또는 문장 ~ 찾는 단어 또는 문장 ~ 끝나는 단어 또는 문장을 입력해주세요. (end) ↑↑↑"

user_file_name = 'save.txt'


def user_save_file1():
    with open(user_file_name, 'r') as file:
        lines = file.readlines()

    line_get = False
    user_search_lines = []

    for line in lines:
        if line == search_line_end:
            line_get = False

        if line_get == True:
            user_search_lines.append(line)

        if line == search_line_start:
            line_get = True

    # 공백 지우기
    user_search_lines = ' '.join(user_search_lines).split()
    return user_search_lines


def user_save_file2():
    with open(user_file_name, 'r') as file:
        lines = file.readlines()

    line_get = False
    user_filter_lines = []

    for line in lines:
        if line == filter_line_end:
            line_get = False

        if line_get == True:
            user_filter_lines.append(line)

        if line == filter_line_start:
            line_get = True

    # 공백 지우기
    user_filter_lines = ' '.join(user_filter_lines).split()
    return user_filter_lines


# 엑셀 문서를 만드는 함수
def excelOpen():
    try:
        wb = openpyxl.Workbook()
        return wb

    except Exception as ex:  # 에러 종류
        print('error', ex)
        return -1


def find_user_search(file, user_search_lines):
    log = []
    file.seek(0)
    lines = file.readlines()
    lines = ' '.join(lines).split()

    for line in lines:
        for user_line in user_search_lines:
            if line.find(user_line) >= 0:
                # 찻았따
                log.append(line)

    return log


def find_user_filter(file, user_filter_lines):
    log = []
    file.seek(0)
    lines = file.readlines()
    lines = ' '.join(lines).split()
    line_get = False

    for filter_line in user_filter_lines:
        start_line, end_line = filter_line.split('~')
        for line in lines:
            if line.find(start_line) >= 0:
                line_get = True

            if line_get == True:
                log.append(line)

            if line.find(end_line) >= 0:
                line_get = False

    return log

# pathnote 경로 안에 .txt .log 로 끝나는 파일을 열어서 리스트에 문자열을 하나씩 엑셀에 저장하는 함수


def main_crawling(wb, pathnote):

    # 사용자가 저장한 텍스트 파일 가져오기
    user_search_lines = user_save_file1()
    user_filter_lines = user_save_file2()

    ws_search = wb.create_sheet('수집 데이터 search')
    ws_filter = wb.create_sheet('수집 데이터 filter')

    for (path, dir, files) in os.walk(pathnote):
        for filename in files:
            ext = os.path.splitext(filename)[-1]
            if ext == '.txt' or ext == '.log':
                print("%s/%s" % (path, filename))

                f = codecs.open("%s/%s" % (path, filename),
                                'r', "utf-8-sig", errors='ignore')

                log1 = find_user_search(f, user_search_lines)
                log2 = find_user_filter(f, user_filter_lines)

                # log1 | find_user_search | user_search_lines | 수집 데이터 search
                excel_index = 1
                ws_search.cell(excel_index, 1).value = filename
                ws_search.cell(excel_index, 1).font = Font(size=15, bold=True)

                for line in log1:
                    excel_index += 1
                    ws_search.cell(excel_index, 1).value = line
                    excel_index += 1

                # log2 | find_user_filter | user_filter_lines | 수집 데이터 filter
                excel_index = 1
                ws_filter.cell(excel_index, 1).value = filename
                ws_filter.cell(excel_index, 1).font = Font(size=15, bold=True)

                for line in log2:
                    excel_index += 1
                    ws_filter.cell(excel_index, 1).value = line
                    excel_index += 1

                f.close()

# 엑셀 저장하는 함수


def wbsave(wb):
    try:
        wb.save("수집수집.xlsx")
        return wb
    except Exception as ex:  # 에러 종류
        print('엑셀 저장 에러(엑셀 프로그램을 모두 끄고 다시 실행해주세요.)', ex)
        return -1

# 엑셀 문서를 닫는 함수


def wbclose(wb):
    try:
        wb.close()
        return wb

    except Exception as ex:  # 에러 종류
        print('엑셀 프로세스 에러(엑셀 프로그램을 모두 끄고 다시 실행해주세요.)', ex)
        return -1


#!! MAIN!!
if __name__ == "__main__":
    user_input_txt = input("로그 파일 경로:")
    wb = excelOpen()
    main_crawling(wb, user_input_txt)
    wbsave(wb)
    wbclose(wb)
