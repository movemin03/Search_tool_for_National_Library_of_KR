import sys
import urllib.request
import json
import pandas as pd
import re

print("도서 검색 프로그램입니다")
print("made by 이동민 사회복무요원, 2022.11.17 ver")
print("위 프로그램은 국립중앙도서관 데이터를 활용하고 있습니다")

#저자명에서 초성을 추출하는 모듈
def kr_to_en(korean_word):
    r_lst = []
    for w in list(korean_word.strip()):
        # 영어인 경우 구분해서 작성함.
        if '가' <= w <= '힣':
            ## 588개 마다 초성이 바뀜.
            ch1 = (ord(w) - ord('가')) // 588
            ## 중성은 총 28가지 종류
            ch2 = ((ord(w) - ord('가')) - (588 * ch1)) // 28
            ch3 = (ord(w) - ord('가')) - (588 * ch1) - 28 * ch2

            # 초성 리스트. 00 ~ 18
            CHOSUNG_LIST = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ',
                            'ㅎ']
            # 중성 리스트. 00 ~ 20
            JUNGSUNG_LIST = ['ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ',
                             'ㅡ', 'ㅢ',
                             'ㅣ']
            # 종성 리스트. 00 ~ 27 + 1(1개 없음)
            JONGSUNG_LIST = [' ', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ',
                             'ㅄ', 'ㅅ',
                             'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']

            r_lst.append([CHOSUNG_LIST[ch1], JUNGSUNG_LIST[ch2], JONGSUNG_LIST[ch3]])
        else:
            r_lst.append([w])
    return r_lst

#추출된 초성을 저자기호로 변경해주는 모듈
def kr_to_num():
    k_c_1 = kr_to_en(class_num_3_re)[0][0]
    if "ㅊ" in k_c_1:
        dic = {'ㅏ': 2, 'ㅐ': 2, 'ㅑ': 2, 'ㅒ': 2, 'ㅓ': 3, 'ㅔ': 3, 'ㅕ': 3, 'ㅖ': 3, 'ㅗ': 4, 'ㅘ': 4, 'ㅙ': 4,
               'ㅚ': 4, 'ㅛ': 4,
               'ㅜ': 5, 'ㅝ': 5, 'ㅞ': 5, 'ㅟ': 5, 'ㅠ': 5, 'ㅡ': 5, 'ㅢ': 5, 'ㅣ': 6}
    else:
        dic = {'ㅏ': 2, 'ㅐ': 3, 'ㅑ': 3, 'ㅒ': 3, 'ㅓ': 4, 'ㅔ': 4, 'ㅕ': 4, 'ㅖ': 4, 'ㅗ': 5, 'ㅘ': 5, 'ㅙ': 5,
               'ㅚ': 5, 'ㅛ': 5,
               'ㅜ': 6,
               'ㅝ': 6, 'ㅞ': 6, 'ㅟ': 6, 'ㅠ': 6, 'ㅡ': 7, 'ㅢ': 7, 'ㅣ': 8}
    k_c_2 = kr_to_en(class_num_3_re)[1][1]
    convert1 = dic.get(str(k_c_2))

    # 2번째 글자 자음 저장
    dic2 = {'ㄱ': 1, "ㄲ": 1, "ㄴ": 19, "ㄷ": 2, "ㄸ": 2, "ㄹ": 29, "ㅁ": 3, "ㅂ": 4, "ㅃ": 4, "ㅅ": 5, "ㅆ": 5,
            "ㅇ": 6, "ㅈ": 7,
            "ㅉ": 7, "ㅊ": 8, "ㅋ": 87, "ㅌ": 88, "ㅍ": 89, "ㅎ": 9}
    k_c_3 = kr_to_en(class_num_3_re)[1][0]
    convert2 = dic2.get(str(k_c_3))
    author_num = (class_num_3_re[:1] + str(convert1) + str(convert2))
    d_7.append(author_num)

# ISBN 13 to ISBN 10 모듈
def ISBN_10(x):
    ISBN_before = x[3:12]
    ISBN_split = []

    for z in str(ISBN_before):
        ISBN_split.append(z)

    ISBN_int = [int(z) for z in ISBN_split]

    ISBN_int_cul = 11 - ISBN_int[0] * 10 - ISBN_int[1] * 9 - ISBN_int[2] * 8 - ISBN_int[3] * 7 - ISBN_int[4] * 6 - \
                   ISBN_int[5] * 5 - ISBN_int[6] * 4 - ISBN_int[7] * 3 - ISBN_int[8] * 2
    ISBN_int_cul_2 = str(ISBN_int_cul % 11).replace('10', 'X')
    ISBN_int.append(ISBN_int_cul_2)

    ISBN_10 = (" ".join(map(str, ISBN_int))).replace(' ', '')
    List.append(ISBN_10)

# 국립중앙도서관 API 검색 모듈
def find_book(x):
    progress.append("1")
    print(x + "에 대한 정보를 검색합니다")
    print("전체진행도: " + str(len(progress)) + "/" + str(len(List)))
    print("")
    kwd = x
    key = "83a6dcb07e447377b1b6b1fb79fce8fcafec2afcfc8adc768ee2d32a1124b573"
    url = (str("https://www.nl.go.kr/NL/search/openApi/search.do?key=") + key + str("&kwd=") + kwd + str(
        "&pageNum=1&pageSize=1&apiType=json"))
    res = urllib.request.urlopen(url)
    res_msg = res.read().decode('utf-8')
    data = json.loads(res_msg).get('result')
    # 책 명 검색 엑셀 파일 추가
    print("제목 정보 검색중(1/5)")
    try:
        for h in data:
            class_num_2 = h['titleInfo']
    except:
        pass

    d_2.append(class_num_2)

    # 책 저자 검색 및 엑셀 추가
    print("저자 정보 검색중(2/5)")
    try:
        for j in data:
            class_num_3 = str(j["authorInfo"])
            class_num_3_2 = re.sub(r"[^\uAC00-\uD7A30]", "", class_num_3)

            if class_num_3_2 != "":
                if (class_num_3_2.find('구성') != -1) or (class_num_3_2.find('저자') != -1):
                    class_num_3_re = class_num_3_2[2:]
                else:
                    if (class_num_3_2.find('글쓴이') != -1) or (class_num_3_2.find('엮은이') != -1) or (class_num_3_2.find('지은이') != -1):
                        class_num_3_re = class_num_3_2[3:]
                    else:
                        if '글' in class_num_3_2:
                            class_num_3_re = class_num_3_2.replace('글', '')
                        else:
                            class_num_3_re = class_num_3_2
            else:
                class_num_3_re = "-"

            d_3.append(class_num_3_re)
    
            if class_num_3_re != "-":
                kr_to_num()
            else:
                d_7.append("-")
                pass
    except:
        pass

    # 책 ISBN 엑셀 파일 추가
    d_1.append(kwd)

    # 책 듀이십진분류법 엑셀 파일 추가
    print("분류 번호 검색중(3/5)")
    for i in data:
        class_num = i['callNo']
        if class_num == "":
            blank_list.append(x)
            d_8.append("-")
            for i in data:
                class_num = i['kdcCode1s']
            if class_num == "":
                d_4.append("-")
            else:
                d_4.append(str(class_num) + "00")
        else:
            c_n_1 = class_num[:1]
            if '0' <= c_n_1 <= '9':
                d_4.append(class_num[0:3])
                d_8.append("")
            else:
                d_4.append(class_num[1:4])
                d_8.append(class_num[:1])

    # 출판사 엑셀 파일 추가
    print("출판사 정보 검색중(4/5)")
    for k in data:
        class_num_4 = k['pubInfo']
        if class_num_4 == "":
            blank_list.append(x)
            d_5.append("-")
        else:
            d_5.append(class_num_4)

    # 출판 연도 엑셀 파일 추가
    print("출판 연도 검색중(5/5)")

    for l in data:
        class_num_5 = l['pubYearInfo']
        if class_num_5 == "":
            blank_list.append(x)
            d_6.append("-")
        else:
            d_6.append(class_num_5[:4])


# 저장 공간
progress = []
count = []
recount = []
error_count = []
error_list = []
blank_list = []
d_1 = []
d_2 = []
d_3 = []
d_4 = []
d_5 = []
d_6 = []
d_7 = []
d_8 = []

# 실행 영역
while 1 == 1:
    print("저자기호 변환만 원한다면 숫자 0 을, 도서정보 검색을 하려면 아무키나 입력해주세요:")
    first = input()

    if first == "0":
        List = [x for x in input('\n검색할 저자명을 입력해주세요. 한글만 인식 가능합니다. 띄어쓰기로 구분합니다: \n').split()]
        print(str(len(List)) + '개의 항목에 대하여 검색을 시작합니다 \n')
        for x in List:
            if str(x.isalpha()) == "False":
                print ("저자명이 잘못되었습니다. 계속 진행시 프로그램이 강제종료 될 수 있습니다. 프로그램을 다시 시작해주세요")
                a = input()
            else:
                pass
            class_num_3_re = x
            kr_to_num()
        print(d_7)
    else:
    # 텍스트 파일 생성
        f_f = open('실패항목.txt', 'w')

    #실행
        List = [x for x in input('\n검색할 도서 isbn을 입력해주세요. 띄어쓰기로 구분합니다: \n').split()]
        print(str(len(List)) + '개의 항목에 대하여 검색을 시작합니다 \n')


        for x in List:
            if len(x) == 10 or len(x) == 13:
                try:
                    find_book(x)
                    count.append('1')
                except:
                    if len(x) == 10:
                        error_count.append('1')
                        error_list.append(x)
                        pass
                    else:
                        try:
                            recount.append('1')
                            ISBN_10(x)
                        except:
                            error_count.append('1')
            else:
                print('올바른 ISBN인지 확인해주세요. 10자리 혹은 13자리여야 합니다\n')
                error_count.append('1')
        pass

        print('\n검색 성공 개수: ' + str(len(count)))
        print('ISBN 변환 시도 개수: ' + str(len(recount)))
        print('오류 개수: ' + str(len(error_count)))
        print('오류 항목: ' + str(error_list))
        print('빈 결과 값 나온 항목' + str(blank_list))
        print('빈 결과 값이 나온 경우, 검색은 성공했으나 어린이 도서관 혹은 디지털 보관 자료입니다. 별도로 구분해야 합니다.')

# 파일 기록 및 저장
        f_f.write("오류 항목: ")
        f_f.write(str(error_list))
        f_f.write("\n")
        f_f.write("빈 결과 값 나온 항목: ")
        f_f.write("빈 결과 값 나온 항목: ")
        f_f.write(str(blank_list))
        f_f.write('\n빈 결과 값이 나온 경우, 검색은 성공했으나 어린이 도서관 혹은 디지털 보관 자료입니다. 별도로 구분해야 합니다.')

        f_f.close()

        data_frame = (pd.DataFrame([d_2, d_3, d_7, d_8, d_4, d_1, d_5, d_6]))
        data_frame_re = data_frame.transpose()
        data_frame_re.to_excel('성공목록.xlsx')

    pause = str(input("\n완료되었습니다. 숫자 0을 입력하면 종료되고, 아니라면 처음으로 돌아갑니다."))
    if pause == "0" :
        break
    else :
        continue
sys.exit()
