import tkinter as tk
from tkinter import ttk
import urllib.request
import json
import pandas as pd
import re
import tkinter.messagebox
import time

win = tk.Tk()
win.title("국립중앙도서관 대량 검색 프로그램")

w = 510  # width for the Tk root
h = 710  # height for the Tk root

# get screen width and height
ws = win.winfo_screenwidth()  # width of the screen
hs = win.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws / 2) - (w / 2)
y = (hs / 2) - (h / 2)

# set the dimensions of the screen
# and where it is placed
win.geometry('%dx%d+%d+%d' % (w, h, x, y))
win.resizable(False, True)

tabcontrol = ttk.Notebook(win)

tab1 = ttk.Frame()  # 2개의 탭을 나누는 것
tabcontrol.add(tab1, text="")
tabcontrol.pack(expand=0, fill="both")

# ==========================================================================
# 검색창 프레임
firstlabel = ttk.LabelFrame(tab1, text="검색어 입력")  # 첫번째 테이블에서 레이블 프레임을 만듬
firstlabel.grid(column=0, row=0, padx=6, pady=4)

# 검색박스 입력창
label1_textbox = tk.StringVar()
label1_textbox = ttk.Entry(firstlabel, width=30)
label1_textbox.grid(column=0, row=0)

# 필수 옵션 프레임
secondlabel = ttk.LabelFrame(tab1, text="옵션")
secondlabel.grid(column=0, row=5, sticky="w", padx=8, pady=4)

# 필수 옵션 오디오 버튼_1
option_1 = tk.IntVar()
rad1 = tk.Radiobutton(secondlabel, text="   전체항목 검색   ", value=0, variable=option_1)
rad1.grid(column=6, row=0)

# 필수 옵션 오디오 버튼_2
rad2 = tk.Radiobutton(secondlabel, text="   저자기호 검색   ", value=1, variable=option_1)
rad2.grid(column=6, row=1)

option_3 = tk.IntVar()
chk1 = tk.Checkbutton(secondlabel, text="오류 무시(선택)", variable=option_3)
chk1.grid(column=6, row=3, sticky="w")

# 파일 저장명 프레임
thirdlabel = ttk.LabelFrame(tab1, text="파일 저장명")
thirdlabel.grid(column=0, row=5, sticky="e", padx=8, pady=4)

# 성공 파일 저장명 라벨
label2 = ttk.Label(thirdlabel, text="성공목록 파일명:", width=13)
label2.grid(column=0, row=0)

# 성공 파일 저장명 입력창
label2_textbox = ttk.Entry(thirdlabel, width=35)
label2_textbox.insert(0, "성공목록")
label2_textbox.grid(column=1, row=0)

# 실패 파일 저장명 라벨
label3 = ttk.Label(thirdlabel, text="실패목록 파일명:", width=13)
label3.grid(column=0, row=1)

# 실패 파일 저장명 입력창
label3_textbox = ttk.Entry(thirdlabel, width=35)
label3_textbox.insert(0, "실패목록")
label3_textbox.grid(column=1, row=1)

ResultViewlabel_2 = ttk.LabelFrame(tab1, text="알림")
ResultViewlabel_2.grid(column=0, row=14, padx=8, pady=4)

ResultViewlabel_Scrollbar_2 = tk.Listbox(ResultViewlabel_2, selectmode="extended", width=80, height=9,
                                         font=('Normal', 9),
                                         yscrollcommand=tk.Scrollbar(ResultViewlabel_2).set)
ResultViewlabel_Scrollbar_2.grid(column=0, row=0)

# 실행 결과 스크롤 박스
ResultViewlabel = ttk.LabelFrame(tab1, text="실행내용")
ResultViewlabel.grid(column=0, row=3, padx=8, pady=4)

ResultViewlabel_Scrollbar = tk.Listbox(ResultViewlabel, selectmode="extended", width=80, height=20, font=('Normal', 9),
                                       yscrollcommand=tk.Scrollbar(ResultViewlabel).set)
ResultViewlabel_Scrollbar.grid(column=0, row=1)

ResultViewlabel_Scrollbar.insert(tk.END, "ver 2022-11-29")
ResultViewlabel_Scrollbar.insert(tk.END, "Developer : https://github.com/movemin03")
ResultViewlabel_Scrollbar.insert(tk.END, "필수 옵션을 먼저 변경한 후 자료를 입력해주세요")
ResultViewlabel_Scrollbar.insert(tk.END, "여러 개를 입력 시 띄어쓰기로 항목을 구분합니다")
ResultViewlabel_Scrollbar.insert(tk.END, "많은 양의 자료를 입력시 버벅거릴 수 있으나 기다려주세요")
ResultViewlabel_Scrollbar.insert(tk.END, "너무 버벅일 경우 불러오기 기능 사용을 권장합니다")
ResultViewlabel_Scrollbar.insert(tk.END, "자료가 800개 이상일 경우는 나눠서 하기를 권장 드립니다")

# progress_bar 타입 설정
pb_type = tk.DoubleVar()
pb_type_2 = tk.DoubleVar()

progress_bar = ttk.Progressbar(ResultViewlabel, orient="horizontal", length=400, mode="determinate", maximum=100,
                               variable=pb_type)

# ==========================================================================


count = 0  # 값 저장 공간
recount = 0  # isbn10 으로 변경한 횟수
error_count = 0  # 에러 개수
progress = 0  # 전체 progress 정도 측정
extend = 0  # 불러오기 클릭 횟수

error_list = []  # 에러 목록
blank_list = []  # 빈 항목 목록
d_1 = []  # 입력된 ISBN
d_2 = []  # 제목정보
d_3 = []  # 저자정보
d_4 = []  # 대분류번호
d_5 = []  # 출판사정보
d_6 = []  # 출판연도
d_7 = []  # 저자기호
d_8 = []  # 분류번호

global firstlabel_sub
firstlabel_sub = ttk.LabelFrame(firstlabel, text="불러올 경로 입력(텍스트 파일만)")  # 첫번째 테이블에서 레이블 프레임을 만듬
firstlabel_sub.grid(column=0, row=1, padx=6, pady=4)


def execute_3():
    global extend  # 불러오기 버튼 클릭 횟수 기록
    extend += 1
    if extend < 2:   # 불러오기 입력창 열기
        global label1_textbox_sub
        label1_textbox_sub = ttk.Entry(firstlabel_sub, width=30, textvariable=tk.StringVar())
        label1_textbox_sub.grid(column=0, row=1)
    else:
        pass
    global path  # 불러오기 입력창 내용 가져오기
    path = str(label1_textbox_sub.get())

    if len(path) < 1:
        tkinter.messagebox.showwarning(title="알림",
                                       message="경로를 지정하고 불러오기 버튼을 다시 눌러주세요. 텍스트 파일만 가능합니다. 확장자를 포함하여 입력합니다. 예를 들어 C:\\\\test\\test.txt")
    else:
        try:  # 텍스트 파일 불러오기
            isbn_txt = open(path, 'r')
            global List
            List = isbn_txt.readlines()
            isbn_txt.close
            tkinter.messagebox.showinfo(title="알림", message="데이터를 불러왔습니다. 바로 검색을 시작합니다")
            global path_load
            path_load = "true"  # 불러오기 여부 변경
            execute()
            path_load = "false"  # 불러오기 여부 변경
        except:
            tkinter.messagebox.showwarning(title="알림", message=path + " 경로에서 txt 파일을 찾지 못했습니다.")


# 검색 버튼 실행 모듈
def execute():
    # 걸린 시간 측정 시작
    time_start = time.time()
    # 실행 영역
    ResultViewlabel_Scrollbar_2.delete(0, 99999999)

    try:  # path_load 값 없을 것에 대비
        global List
        if path_load == "true":  # 텍스트 파일에서 내용을 불러와서 이미 List 가 만들어졌으므로
            pass
        else:
            List = label1_textbox.get().split()  # 엔트리 창에서 입력 내용 가져오기
    except:
        List = label1_textbox.get().split()

    if List:
        pass
    else:
        tkinter.messagebox.showwarning(title="알림", message="검색어를 입력해주세요")  # 경고 메시지 박스

    # 저자명 옵션 시 실행 프로세스
    if option_1.get() == 1:

        ResultViewlabel_Scrollbar.insert(tk.END, str(len(List)) + '개의 항목에 대하여 검색을 시작합니다')
        ResultViewlabel_Scrollbar.insert(tk.END, "")
        for x in List:
            # 한글, 영어 여부 확인
            k_count = 0
            e_count = 0
            c = x[:1]
            if ord('가') <= ord(c) <= ord('힣'):
                k_count += 1
            elif ord('a') <= ord(c.lower()) <= ord('z'):
                e_count += 1
            else:
                pass

            # 여부 확인 실행 영역
            if str(x.isalpha()) == "False":
                global msg_cancel
                msg_cancel = tkinter.messagebox.askokcancel(title="알림", message=str(
                    x) + "에 숫자가 포함됐습니다. 저자명이 잘못되었습니다. 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요")
                ResultViewlabel_Scrollbar.insert(tk.END, str(x) + "에 숫자가 포함됐습니다.. 저자명이 잘못되었습니다.")
                ResultViewlabel_Scrollbar.insert(tk.END, "")
                if str(msg_cancel) == "False":
                    break
                else:
                    pass
            elif k_count < e_count:
                msg_cancel = tkinter.messagebox.askokcancel(title="알림", message=str(
                    x) + "에 영어가 포함됐습니다. 저자명이 잘못되었습니다. 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요")
                ResultViewlabel_Scrollbar.insert(tk.END, str(x) + "에 영어가 포함됐습니다. 저자명이 잘못되었습니다.")
                ResultViewlabel_Scrollbar.insert(tk.END, "")
                if str(msg_cancel) == "False":
                    break
                else:
                    pass
            else:
                global count
                count += 1
                global class_num_3_re
                class_num_3_re = str(x)
                kr_to_num()
                ResultViewlabel_Scrollbar.see(tk.END)
                pass
        ResultViewlabel_Scrollbar.insert(tk.END, str(d_7))
        ResultViewlabel_Scrollbar.insert(tk.END, "")
        ResultViewlabel_Scrollbar.see(tk.END)
        clear()
        time_end = time.time()
        try:
            take_time = round(len(List) / (time_end - time_start), 4)
        except:
            take_time = "∞"

        ResultViewlabel_Scrollbar_2.delete(0, 99999999)
        ResultViewlabel_Scrollbar_2.insert(tk.END, '\n검색 성공 개수: ' + str(len(List)))
        ResultViewlabel_Scrollbar_2.insert(tk.END, '처리 속도: ' + str(take_time) + " unit/s (높을수록 좋습니다)")

    else:
        # 전체 검색 옵션 시 실행 프로세스
        # 텍스트 파일 생성
        f_f = open(label3_textbox.get() + '.txt', 'w')

        # 실행
        ResultViewlabel_Scrollbar.insert(tk.END, str(len(List)) + '개의 항목에 대하여 검색을 시작합니다')
        time_start_2 = time.time()

        for x in List:
            x = x[0:13]
            if len(x) == 10 or len(x) == 13:
                try:
                    find_book(x)
                    count += 1
                except:
                    if len(x) == 10:
                        global error_count
                        error_count += 1
                        error_list.append(x)
                        pass
                    else:
                        try:
                            global recount
                            recount += 1
                            ISBN_10(x)
                        except:
                            error_count += 1
            else:
                if option_3.get() == 0:
                    global msg_cancel_2
                    msg_cancel_2 = tkinter.messagebox.askokcancel(title="알림",
                                                                  message="올바른 ISBN인지 확인해주세요. 10자리 혹은 13자리 숫자여야 합니다. 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요")
                    ResultViewlabel_Scrollbar.insert(tk.END, '')
                    ResultViewlabel_Scrollbar.insert(tk.END, '올바른 ISBN 인지 확인해주세요. 10자리 혹은 13자리 숫자여야 합니다\n')
                    error_count += 1
                    if str(msg_cancel_2) == "False":
                        break
                    else:
                        pass
                else:
                    pass

        pass
        ResultViewlabel_Scrollbar_2.delete(0, 99999999)
        ResultViewlabel_Scrollbar_2.insert(tk.END, '\n검색 성공 개수: ' + str(count))
        ResultViewlabel_Scrollbar_2.insert(tk.END, 'ISBN 변환 시도 개수: ' + str(recount))
        ResultViewlabel_Scrollbar_2.insert(tk.END, '오류 개수: ' + str(error_count))
        ResultViewlabel_Scrollbar_2.insert(tk.END, '오류 항목: ' + str(error_list))
        ResultViewlabel_Scrollbar_2.insert(tk.END, '빈 결과 값 나온 항목' + str(blank_list))
        ResultViewlabel_Scrollbar_2.insert(tk.END, '빈 결과 값이 나온 경우, 검색은 성공했으나 어린이 도서관 혹은 디지털 보관 자료입니다. 별도로 구분해야 합니다.')
        ResultViewlabel_Scrollbar_2.insert(tk.END, '결과 값은 이 프로그램이 저장된 위치에 저장됩니다')
        time_end_2 = time.time()

        pb_type_2.set(len(List))
        progress_bar_2.update()
        pb_type.set(100)
        progress_bar.update()

        try:
            take_time_2 = round(len(List) / (time_end_2 - time_start_2), 4)
        except:
            take_time_2 = "∞"
        ResultViewlabel_Scrollbar_2.insert(tk.END, '처리 속도: ' + str(take_time_2) + ' unit/s (높을수록 좋습니다)')

        # 파일 기록 및 저장
        f_f.write("오류 항목: ")
        f_f.write(str(error_list))
        f_f.write("\n")
        f_f.write("빈 결과 값 나온 항목: ")
        f_f.write("빈 결과 값 나온 항목: ")
        f_f.write(str(blank_list))
        f_f.write('\n빈 결과 값이 나온 경우, 검색은 성공했으나 어린이 도서관 혹은 디지털 보관 자료입니다. 별도로 구분해야 합니다.\n')

        f_f.close()

        error_list.clear()
        blank_list.clear()
        count = 0
        recount = 0
        error_count = 0

        data_frame = (pd.DataFrame([d_2, d_3, d_7, d_8, d_4, d_1, d_5, d_6]))
        data_frame_re = data_frame.transpose()
        data_frame_re.to_excel(label2_textbox.get() + '.xlsx')

        try:
            if str(msg_cancel_2) == "False":
                tkinter.messagebox.showinfo(title="알림", message="오류로 인해 검색이 중지되었습니다")
                pass
            else:
                tkinter.messagebox.showinfo(title="알림", message="검색이 완료되었습니다")
        except:
            tkinter.messagebox.showinfo(title="알림", message="검색이 완료되었습니다")


def execute_2():
    ResultViewlabel_Scrollbar.delete(7, 99999999)
    ResultViewlabel_Scrollbar_2.delete(0, 99999999)
    pb_type.set(00)
    progress_bar.update()
    pb_type_2.set(00)
    progress_bar_2.update()
    clear()


def clear():
    try:
        d_1.clear()
    except:
        pass
    try:
        d_2.clear()
    except:
        pass
    try:
        d_3.clear()
    except:
        pass
    try:
        d_4.clear()
    except:
        pass
    try:
        d_5.clear()
    except:
        pass
    try:
        d_6.clear()
    except:
        pass
    try:
        d_7.clear()
    except:
        pass
    try:
        d_8.clear()
    except:
        pass
    try:
        error_list.clear()
    except:
        pass
    try:
        error_list.clear()
    except:
        pass
    try:
        blank_list.clear()
    except:
        pass
    try:
        progress == 0
    except:
        pass
    try:
        count == 0
    except:
        pass
    try:
        recount == 0
    except:
        pass
    try:
        error_count == 0
    except:
        pass
    try:
        firstlabel_sub == ""
    except:
        pass


# ==========================================================================

# 검색박스 내 확인 버튼
action = ttk.Button(firstlabel, text="확인", command=execute)
action.grid(column=5, row=0, sticky="W")

# 검색박스 내 초기화 버튼
reset = ttk.Button(firstlabel, text="초기화", command=execute_2)
reset.grid(column=6, row=0, sticky="W")

read_txt = ttk.Button(firstlabel, text="불러오기", command=execute_3)
read_txt.grid(column=7, row=0, sticky="W")


# ==========================================================================
def kr_to_en(korean_word):
    r_lst = []
    for w in list(korean_word.strip()):
        # 영어인 경우 구분해서 작성함.
        if '가' <= w <= '힣':
            # 588개 마다 초성이 바뀜.
            ch1 = (ord(w) - ord('가')) // 588
            # 중성은 총 28가지 종류
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


# 추출된 초성을 저자 기호로 변경 해주는 모듈
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
    before_isbn10 = str(x)[3:12]
    ISBN_int_cul = (11 - (int(x / 10 ** 9) % 10) * 10 - (int(x / 10 ** 8) % 10) * 9 - (int(x / 10 ** 7) % 10) * 8 - (
            int(x / 10 ** 6) % 10) * 7 - (int(x / 10 ** 5) % 10) * 6 - (int(x / 10 ** 4) % 10) * 5 - (
                            int(x / 10 ** 3) % 10) * 4 - (int(x / 10 ** 2) % 10) * 3 - (int(x / 10) % 10) * 2) % 11
    if ISBN_int_cul == 10:
        ISBN_int_cul = 'X'
    else:
        pass

    ISBN_10 = before_isbn10 + ISBN_int_cul + " "
    List.append(ISBN_10)


# 국립 중앙 도서관 API 검색 모듈
def find_book(x):
    global progress
    progress += 1
    ResultViewlabel_Scrollbar.insert(tk.END, "")
    ResultViewlabel_Scrollbar.insert(tk.END, x + "에 대한 정보를 검색합니다")
    ResultViewlabel_Scrollbar.insert(tk.END, "전체진행도: " + str(progress) + "/" + str(len(List)))
    # ============================================================

    # 진행도
    progress_bar.grid(column=0, row=2, sticky="e")

    global progress_bar_2
    progress_bar_2 = ttk.Progressbar(ResultViewlabel, orient="horizontal", length=400, mode="determinate",
                                     maximum=len(List), variable=pb_type_2)
    progress_bar_2.grid(column=0, row=3, sticky="e")

    pb_type_2.set(progress)
    progress_bar_2.update()

    # 현재 검색어 진행도 라벨
    label3 = ttk.Label(ResultViewlabel, text="단일 진행도:", width=11)
    label3.grid(column=0, row=2, sticky="W")

    label4 = ttk.Label(ResultViewlabel, text="전체 진행도:", width=11)
    label4.grid(column=0, row=3, sticky="W")

    # ===========================================================
    kwd = x
    key = "83a6dcb07e447377b1b6b1fb79fce8fcafec2afcfc8adc768ee2d32a1124b573"
    url = (str("https://www.nl.go.kr/NL/search/openApi/search.do?key=") + key + str("&kwd=") + kwd + str(
        "&pageNum=1&pageSize=1&apiType=json"))
    res = urllib.request.urlopen(url)
    res_msg = res.read().decode('utf-8')
    data = json.loads(res_msg).get('result')
    # 책 명 검색 엑셀 파일 추가

    ResultViewlabel_Scrollbar.insert(tk.END, "제목 정보 검색중(1/5)")
    ResultViewlabel_Scrollbar.see(tk.END)
    pb_type.set(16)
    progress_bar.update()

    try:
        for h in data:
            class_num_2 = h['titleInfo']
    except:
        pass

    d_2.append(class_num_2)

    # 책 저자 검색 및 엑셀 추가
    ResultViewlabel_Scrollbar.insert(tk.END, "저자 정보 검색중(2/5)")
    pb_type.set(32)
    progress_bar.update()

    try:
        for j in data:
            class_num_3 = str(j["authorInfo"])
            class_num_3_2 = re.sub(r"[^\uAC00-\uD7A30]", "", class_num_3)

            if class_num_3_2 != "":
                if (class_num_3_2.find('구성') != -1) or (class_num_3_2.find('저자') != -1):
                    class_num_3_re = class_num_3_2[2:]
                elif (class_num_3_2.find('글쓴이') != -1) or (class_num_3_2.find('엮은이') != -1) or (class_num_3_2.find('지은이') != -1):
                    class_num_3_re = class_num_3_2[3:]
                elif '글' in class_num_3_2:
                    class_num_3_re = class_num_3_2.replace('글', '')
                else:
                    class_num_3_re = class_num_3_2
            else:
                class_num_3_re = "저자정보없음"

            d_3.append(class_num_3_re)

            if class_num_3_re != "저자정보없음":
                kr_to_num()
            else:
                d_7.append("저자정보없음")
                pass
    except:
        pass

    # 책 ISBN 엑셀 파일 추가
    d_1.append(kwd)

    # 책 듀이십진분류법 엑셀 파일 추가
    ResultViewlabel_Scrollbar.insert(tk.END, "분류 번호 검색중(3/5)")
    pb_type.set(48)
    progress_bar.update()

    for i in data:
        class_num = i['callNo']
        if class_num == "":
            blank_list.append(x)
            d_8.append("-")
            for j in data:
                class_num = j['kdcCode1s']
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
    ResultViewlabel_Scrollbar.insert(tk.END, "출판사 정보 검색중(4/5)")
    pb_type.set(64)
    progress_bar.update()
    for k in data:
        class_num_4 = k['pubInfo']
        if class_num_4 == "":
            blank_list.append(x)
            d_5.append("-")
        else:
            d_5.append(class_num_4)

    # 출판 연도 엑셀 파일 추가
    ResultViewlabel_Scrollbar.insert(tk.END, "출판 연도 검색중(5/5)")
    pb_type.set(80)
    progress_bar.update()

    for l in data:
        class_num_5 = l['pubYearInfo']
        if class_num_5 == "":
            blank_list.append(x)
            d_6.append("출판연도정보없음")
        else:
            d_6.append(class_num_5[:4])
    pb_type.set(100)
    progress_bar.update()


# ==========================================================================

win.mainloop()
