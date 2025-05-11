import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import urllib.request
import json
import pandas as pd
import re
import time
import os
import shutil
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


class LibrarySearchApp:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_widgets()
        self.initialize_variables()

    def setup_window(self):
        """윈도우 기본 설정"""
        self.root.title("국립중앙도서관 대량 검색 프로그램")

        width, height = 510, 710
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_position = (screen_width / 2) - (width / 2)
        y_position = (screen_height / 2) - (height / 2)

        self.root.geometry(f"{width}x{height}+{int(x_position)}+{int(y_position)}")
        self.root.resizable(False, True)

    def create_widgets(self):
        """UI 위젯 생성"""
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=0, fill="both")

        # 검색창 프레임
        self.create_search_frame()

        # 옵션 프레임
        self.create_options_frame()

        # 파일 저장명 프레임
        self.create_file_name_frame()

        # 파일 저장 위치 프레임
        self.create_save_location_frame()

        # 알림 프레임
        self.create_notification_frame()

        # 실행 결과 프레임
        self.create_result_frame()

        # 버튼 프레임
        self.create_button_frame()

    def create_search_frame(self):
        """검색창 프레임 생성"""
        self.search_frame = ttk.LabelFrame(self.main_frame, text="검색어 입력")
        self.search_frame.pack(fill="x", padx=6, pady=4)

        # 검색어 입력창
        self.search_entry = ttk.Entry(self.search_frame, width=30)
        self.search_entry.grid(column=0, row=0, padx=5, pady=5)

        # 초기화, 불러오기 버튼
        self.reset_button = ttk.Button(self.search_frame, text="초기화", command=self.reset_search)
        self.reset_button.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        self.load_button = ttk.Button(self.search_frame, text="불러오기", command=self.load_from_file)
        self.load_button.grid(column=2, row=0, padx=5, pady=5, sticky="W")

        # 파일 불러오기 프레임 (초기에는 숨겨져 있음)
        self.file_load_frame = ttk.LabelFrame(self.search_frame, text="불러올 경로 입력(텍스트 파일만)")
        self.file_load_frame.grid(column=0, row=1, columnspan=3, padx=6, pady=4, sticky="ew")

    def create_options_frame(self):
        """옵션 프레임 생성"""
        self.options_frame = ttk.LabelFrame(self.main_frame, text="옵션")
        self.options_frame.pack(fill="x", padx=8, pady=4)

        # 내부 프레임 생성 (가로 배치용)
        inner_frame = ttk.Frame(self.options_frame)
        inner_frame.pack(fill="x", padx=5, pady=5)

        # 검색 옵션 라디오 버튼과 체크박스를 가로로 배치
        self.search_option = tk.IntVar()
        self.radio_all = tk.Radiobutton(inner_frame, text="전체항목 검색",
                                        value=0, variable=self.search_option)
        self.radio_all.pack(side=tk.LEFT, padx=10, pady=2)

        self.radio_author = tk.Radiobutton(inner_frame, text="저자기호 검색",
                                           value=1, variable=self.search_option)
        self.radio_author.pack(side=tk.LEFT, padx=10, pady=2)

        # 오류 무시 체크박스
        self.ignore_errors = tk.IntVar()
        self.check_ignore = tk.Checkbutton(inner_frame, text="오류 무시(선택)",
                                           variable=self.ignore_errors)
        self.check_ignore.pack(side=tk.LEFT, padx=10, pady=2)

    def create_file_name_frame(self):
        """파일 저장명 프레임 생성"""
        self.file_name_frame = ttk.LabelFrame(self.main_frame, text="파일 저장명")
        self.file_name_frame.pack(fill="x", padx=8, pady=4)

        # 성공 파일명
        self.success_label = ttk.Label(self.file_name_frame, text="성공목록 파일명:", width=13)
        self.success_label.grid(column=0, row=0, padx=5, pady=2)

        self.success_entry = ttk.Entry(self.file_name_frame, width=35)
        self.success_entry.insert(0, "성공목록")
        self.success_entry.grid(column=1, row=0, padx=5, pady=2)

        # 실패 파일명
        self.fail_label = ttk.Label(self.file_name_frame, text="실패목록 파일명:", width=13)
        self.fail_label.grid(column=0, row=1, padx=5, pady=2)

        self.fail_entry = ttk.Entry(self.file_name_frame, width=35)
        self.fail_entry.insert(0, "실패목록")
        self.fail_entry.grid(column=1, row=1, padx=5, pady=2)

    def create_save_location_frame(self):
        """파일 저장 위치 프레임 생성"""
        self.save_location_frame = ttk.LabelFrame(self.main_frame, text="파일 저장 위치")
        self.save_location_frame.pack(fill="x", padx=8, pady=4)

        # 기본 저장 경로 설정 (사용자 바탕화면/국립중앙도서관_ISBN검색)
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        default_save_path = os.path.join(desktop_path, "국립중앙도서관_ISBN검색")

        # 저장 경로 입력창
        self.save_path_entry = ttk.Entry(self.save_location_frame, width=50)
        self.save_path_entry.insert(0, default_save_path)
        self.save_path_entry.grid(column=0, row=0, padx=5, pady=5)

        # 경로 찾기 버튼
        self.browse_button = ttk.Button(self.save_location_frame, text="찾기", command=self.browse_save_location)
        self.browse_button.grid(column=1, row=0, padx=5, pady=5)

    def create_notification_frame(self):
        """알림 프레임 생성"""
        self.notification_frame = ttk.LabelFrame(self.main_frame, text="알림")
        self.notification_frame.pack(fill="x", padx=8, pady=4)

        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(self.notification_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.notification_listbox = tk.Listbox(
            self.notification_frame,
            selectmode="extended",
            width=80,
            height=9,
            font=('Normal', 9),
            yscrollcommand=scrollbar.set
        )
        self.notification_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        scrollbar.config(command=self.notification_listbox.yview)

        # 알림창 초기 메시지
        self.notification_listbox.insert(tk.END, "알림: 검색 결과 요약 정보가 여기에 표시됩니다")
        self.notification_listbox.insert(tk.END, "검색 성공 개수, 오류 개수, 처리 속도 등이 표시됩니다")

    def create_result_frame(self):
        """실행 결과 프레임 생성"""
        self.result_frame = ttk.LabelFrame(self.main_frame, text="실행내용")
        self.result_frame.pack(fill="x", padx=8, pady=4)

        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(self.result_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 실행 내용 로그창 크기를 알림창과 동일하게 조정 (height=9)
        self.result_listbox = tk.Listbox(
            self.result_frame,
            selectmode="extended",
            width=80,
            height=9,
            font=('Normal', 9),
            yscrollcommand=scrollbar.set
        )
        self.result_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        scrollbar.config(command=self.result_listbox.yview)

        # 초기 메시지 추가
        self.result_listbox.insert(tk.END, "ver 2025-05-12")
        self.result_listbox.insert(tk.END, "Developer : https://github.com/movemin03")
        self.result_listbox.insert(tk.END, "필수 옵션을 먼저 변경한 후 자료를 입력해주세요")
        self.result_listbox.insert(tk.END, "여러 개를 입력 시 띄어쓰기로 항목을 구분합니다")
        self.result_listbox.insert(tk.END, "많은 양의 자료를 입력시 버벅거릴 수 있으나 기다려주세요")
        self.result_listbox.insert(tk.END, "너무 버벅일 경우 불러오기 기능 사용을 권장합니다")
        self.result_listbox.insert(tk.END, "자료가 800개 이상일 경우는 나눠서 하기를 권장 드립니다")

        # 진행 상태바 프레임
        self.progress_frame = ttk.Frame(self.result_frame)
        self.progress_frame.pack(fill="x", padx=5, pady=5)

        self.progress_value = tk.DoubleVar()
        self.total_progress_value = tk.DoubleVar()

        # 단일 진행도
        self.single_progress_label = ttk.Label(self.progress_frame, text="단일 진행도:", width=11)
        self.single_progress_label.grid(column=0, row=0, sticky="W", padx=5, pady=2)

        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            length=400,
            mode="determinate",
            maximum=100,
            variable=self.progress_value
        )
        self.progress_bar.grid(column=1, row=0, sticky="e", padx=5, pady=2)

        # 전체 진행도
        self.total_progress_label = ttk.Label(self.progress_frame, text="전체 진행도:", width=11)
        self.total_progress_label.grid(column=0, row=1, sticky="W", padx=5, pady=2)

        self.total_progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            length=400,
            mode="determinate",
            maximum=100,
            variable=self.total_progress_value
        )
        self.total_progress_bar.grid(column=1, row=1, sticky="e", padx=5, pady=2)

    def create_button_frame(self):
        """버튼 프레임 생성"""
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill="x", padx=8, pady=4)

        # 검색 버튼
        self.search_button = ttk.Button(self.button_frame, text="검색", command=self.execute_search)
        self.search_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 중지 버튼
        self.stop_button = ttk.Button(self.button_frame, text="중지", command=self.stop_search)
        self.stop_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 저장된 폴더 열기 버튼
        self.open_folder_button = ttk.Button(self.button_frame, text="저장 폴더 열기", command=self.open_save_folder)
        self.open_folder_button.pack(side=tk.RIGHT, padx=5, pady=5)

    def initialize_variables(self):
        """변수 초기화"""
        self.search_results = SearchResults()
        self.file_load_count = 0
        self.is_file_loaded = False
        self.search_items = []
        self.api_key = "83a6dcb07e447377b1b6b1fb79fce8fcafec2afcfc8adc768ee2d32a1124b573"
        self.search_running = False

    def browse_save_location(self):
        """저장 위치 탐색 다이얼로그 열기"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.save_path_entry.delete(0, tk.END)
            self.save_path_entry.insert(0, folder_path)

    def open_save_folder(self):
        """저장 폴더 열기"""
        save_path = self.save_path_entry.get()
        if os.path.exists(save_path):
            if os.name == 'nt':  # Windows
                os.startfile(save_path)
            elif os.name == 'posix':  # macOS, Linux
                import subprocess
                subprocess.Popen(['open', save_path])
        else:
            messagebox.showinfo("알림", "저장 폴더가 아직 생성되지 않았습니다.")

    def prepare_save_directory(self):
        """저장 디렉토리 준비"""
        save_path = self.save_path_entry.get()

        # 이미 폴더가 존재하고 내용물이 있는지 확인
        if os.path.exists(save_path) and os.listdir(save_path):
            # 기존 폴더 이름 변경
            i = 1
            while True:
                new_path = f"{save_path}_old({i})"
                if not os.path.exists(new_path):
                    try:
                        os.rename(save_path, new_path)
                        break
                    except PermissionError:
                        # 폴더를 사용 중인 경우
                        result = messagebox.askretrycancel(
                            "오류",
                            f"'{save_path}' 폴더가 사용 중입니다. 폴더를 닫고 다시 시도하세요."
                        )
                        if result:  # 재시도
                            continue
                        else:  # 취소
                            return False
                i += 1

        # 새 폴더 생성
        os.makedirs(save_path, exist_ok=True)
        return True

    def load_from_file(self):
        """파일에서 검색어 불러오기"""
        self.file_load_count += 1

        # 첫 번째 클릭 시 입력창 표시
        if self.file_load_count < 2:
            self.file_path_entry = ttk.Entry(self.file_load_frame, width=30)
            self.file_path_entry.grid(column=0, row=1, padx=5, pady=5)

            self.file_browse_button = ttk.Button(
                self.file_load_frame,
                text="찾기",
                command=self.browse_file_to_load
            )
            self.file_browse_button.grid(column=1, row=1, padx=5, pady=5)
            return

        # 파일 경로 가져오기
        file_path = self.file_path_entry.get()

        if not file_path:
            messagebox.showwarning(
                title="알림",
                message="경로를 지정하고 불러오기 버튼을 다시 눌러주세요. 텍스트 파일만 가능합니다."
            )
            return

        # 경로에서 따옴표 제거
        file_path = file_path.replace('"', '')

        try:
            # 텍스트 파일 불러오기
            with open(file_path, 'r') as file:
                self.search_items = file.readlines()

            messagebox.showinfo(title="알림", message="데이터를 불러왔습니다. 바로 검색을 시작합니다")
            self.is_file_loaded = True
            self.execute_search()
            self.is_file_loaded = False
        except Exception:
            messagebox.showwarning(title="알림", message=f"{file_path} 경로에서 txt 파일을 찾지 못했습니다.")

    def browse_file_to_load(self):
        """불러올 파일 탐색 다이얼로그 열기"""
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)

    def stop_search(self):
        """검색 중지"""
        if self.search_running:
            self.search_running = False
            messagebox.showinfo("알림", "검색이 중지되었습니다.")

    def execute_search(self):
        """검색 실행"""
        # 검색 실행 중 표시
        self.search_running = True

        # 저장 디렉토리 준비
        if not self.prepare_save_directory():
            self.search_running = False
            return

        # 시작 시간 기록
        start_time = time.time()

        # 알림창 초기화
        self.notification_listbox.delete(0, tk.END)
        self.notification_listbox.insert(tk.END, "검색이 시작되었습니다...")

        # 검색어 가져오기
        if self.is_file_loaded:
            items = self.search_items
        else:
            items = self.search_entry.get().split()

        if not items:
            messagebox.showwarning(title="알림", message="검색어를 입력해주세요")
            self.search_running = False
            return

        # 검색 옵션에 따라 다른 처리
        if self.search_option.get() == 1:
            self.search_by_author(items)
        else:
            self.search_by_isbn(items, start_time)

    def search_by_author(self, items):
        """저자명으로 검색"""
        # 실행 로그 초기화 및 정보 표시
        self.result_listbox.delete(7, tk.END)
        self.result_listbox.insert(tk.END, f"{len(items)}개의 항목에 대하여 검색을 시작합니다")

        # 알림창에 진행 상황 표시
        self.notification_listbox.insert(tk.END, f"총 {len(items)}개 항목 처리 중...")

        # 진행 상태바 초기화
        self.total_progress_value.set(0)
        self.total_progress_bar.update()

        for idx, item in enumerate(items):
            # 검색 중지 확인
            if not self.search_running:
                break

            # 한글, 영어 여부 확인
            first_char = item[:1]
            is_korean = ord('가') <= ord(first_char) <= ord('힣')
            is_english = ord('a') <= ord(first_char.lower()) <= ord('z')

            # 진행 상태바 업데이트
            progress_percent = ((idx + 1) / len(items)) * 100
            self.total_progress_value.set(progress_percent)
            self.total_progress_bar.update()

            # 입력값 검증
            if not item.isalpha():
                msg = f"{item}에 숫자가 포함됐습니다. 저자명이 잘못되었습니다."
                should_continue = messagebox.askokcancel(
                    title="알림",
                    message=f"{msg} 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요"
                )
                self.result_listbox.insert(tk.END, msg)

                if not should_continue:
                    break
            elif not is_korean and is_english:
                msg = f"{item}에 영어가 포함됐습니다. 저자명이 잘못되었습니다."
                should_continue = messagebox.askokcancel(
                    title="알림",
                    message=f"{msg} 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요"
                )
                self.result_listbox.insert(tk.END, msg)

                if not should_continue:
                    break
            else:
                self.search_results.count += 1
                author_code = self.convert_korean_to_author_code(item)
                self.search_results.author_codes.append(author_code)

                # 단일 진행도 업데이트 (100%)
                self.progress_value.set(100)
                self.progress_bar.update()

                # 실행 로그 업데이트
                self.result_listbox.insert(tk.END, f"처리 중: {item} → {author_code}")
                self.result_listbox.see(tk.END)

        # 저장 경로
        save_path = self.save_path_entry.get()

        # 검색 결과 저장
        with open(os.path.join(save_path, f"{self.success_entry.get()}.txt"), 'w') as f:
            for code in self.search_results.author_codes:
                f.write(f"{code}\n")

        # 알림창에 결과 요약 표시
        self.notification_listbox.delete(0, tk.END)
        self.notification_listbox.insert(tk.END, f"저자 기호 변환 완료")
        self.notification_listbox.insert(tk.END, f"총 처리 항목: {len(items)}")
        self.notification_listbox.insert(tk.END, f"성공 항목: {self.search_results.count}")
        self.notification_listbox.insert(tk.END, f"결과 파일: {os.path.join(save_path, self.success_entry.get())}.txt")

        self.search_running = False
        self.clear_search_results()
        messagebox.showinfo("알림", "저자 기호 변환이 완료되었습니다.")

    def search_by_isbn(self, items, start_time):
        """ISBN으로 검색"""
        # 저장 경로
        save_path = self.save_path_entry.get()

        # 실행 로그 초기화 및 정보 표시
        self.result_listbox.delete(7, tk.END)
        self.result_listbox.insert(tk.END, f"{len(items)}개의 항목에 대하여 검색을 시작합니다")

        # 알림창에 진행 상황 표시
        self.notification_listbox.insert(tk.END, f"총 {len(items)}개 항목 처리 중...")

        # 진행 상태바 초기화
        self.total_progress_value.set(0)
        self.total_progress_bar.update()

        # 실패 목록 파일 생성
        with open(os.path.join(save_path, f"{self.fail_entry.get()}.txt"), 'w') as fail_file:
            for idx, isbn in enumerate(items):
                # 검색 중지 확인
                if not self.search_running:
                    break

                isbn = isbn[:13]  # ISBN은 최대 13자리

                # 진행 상태바 업데이트
                progress_percent = ((idx + 1) / len(items)) * 100
                self.total_progress_value.set(progress_percent)
                self.total_progress_bar.update()

                if len(isbn) in (10, 13):
                    try:
                        self.find_book_info(isbn)
                        self.search_results.count += 1
                    except Exception as e:
                        if len(isbn) == 10:
                            self.search_results.error_count += 1
                            self.search_results.error_list.append(isbn)
                            # 실행 로그에 오류 표시
                            self.result_listbox.insert(tk.END, f"오류: {isbn} - {str(e)}")
                        else:
                            try:
                                self.search_results.recount += 1
                                isbn_10 = self.convert_isbn13_to_isbn10(isbn)
                                self.search_items.append(isbn_10)
                                # 실행 로그에 변환 정보 표시
                                self.result_listbox.insert(tk.END, f"ISBN 변환: {isbn} → {isbn_10}")
                            except Exception:
                                self.search_results.error_count += 1
                                self.search_results.error_list.append(isbn)
                                # 실행 로그에 오류 표시
                                self.result_listbox.insert(tk.END, f"변환 오류: {isbn}")
                else:
                    if self.ignore_errors.get() == 0:
                        msg = "올바른 ISBN인지 확인해주세요. 10자리 혹은 13자리 숫자여야 합니다."
                        should_continue = messagebox.askokcancel(
                            title="알림",
                            message=f"{msg} 검색을 중단하려면 취소 버튼을, 계속하려면 확인 버튼을 눌러주세요"
                        )
                        self.result_listbox.insert(tk.END, f'잘못된 ISBN 형식: {isbn}')
                        self.search_results.error_count += 1

                        if not should_continue:
                            break

                # 실행 로그 스크롤
                self.result_listbox.see(tk.END)

            # 파일에 오류 정보 기록
            fail_file.write(f"오류 항목: {self.search_results.error_list}\n")
            fail_file.write(f"빈 결과 값 나온 항목: {self.search_results.blank_list}\n")
            fail_file.write('빈 결과 값이 나온 경우, 검색은 성공했으나 일부 항목에 있어 누락된 결과 값이 있을 수 있습니다\n')

        # 결과 데이터프레임 생성 및 저장
        if self.search_results.titles:  # 결과가 있는 경우에만 저장
            data = [
                self.search_results.titles,
                self.search_results.authors,
                self.search_results.author_codes,
                self.search_results.classification_codes,
                self.search_results.main_categories,
                self.search_results.isbns,
                self.search_results.publishers,
                self.search_results.publication_years
            ]

            df = pd.DataFrame(data).transpose()
            df.columns = ['제목', '저자', '저자기호', '분류기호', '주제분류', 'ISBN', '출판사', '출판년도']
            df.to_excel(os.path.join(save_path, f"{self.success_entry.get()}.xlsx"), index=False)

        # 처리 시간 계산
        end_time = time.time()
        try:
            processing_speed = round(len(items) / (end_time - start_time), 4)
        except ZeroDivisionError:
            processing_speed = "∞"

        # 알림창에 결과 요약 표시
        self.notification_listbox.delete(0, tk.END)
        self.notification_listbox.insert(tk.END, f"검색 완료")
        self.notification_listbox.insert(tk.END, f"검색 성공 개수: {self.search_results.count}")
        self.notification_listbox.insert(tk.END, f"ISBN 변환 시도 개수: {self.search_results.recount}")
        self.notification_listbox.insert(tk.END, f"오류 개수: {self.search_results.error_count}")
        self.notification_listbox.insert(tk.END, f"처리 속도: {processing_speed} unit/s (높을수록 좋습니다)")
        self.notification_listbox.insert(tk.END, f"결과 파일: {os.path.join(save_path, self.success_entry.get())}.xlsx")
        self.notification_listbox.insert(tk.END, f"오류 파일: {os.path.join(save_path, self.fail_entry.get())}.txt")

        # 검색 종료
        self.search_running = False

        # 검색 결과 초기화
        self.clear_search_results()

        # 완료 메시지
        messagebox.showinfo(title="알림", message="검색이 완료되었습니다")

    def reset_search(self):
        """검색 초기화"""
        self.result_listbox.delete(7, tk.END)
        self.notification_listbox.delete(0, tk.END)
        self.notification_listbox.insert(tk.END, "알림: 검색 결과 요약 정보가 여기에 표시됩니다")
        self.notification_listbox.insert(tk.END, "검색 성공 개수, 오류 개수, 처리 속도 등이 표시됩니다")
        self.progress_value.set(0)
        self.progress_bar.update()
        self.total_progress_value.set(0)
        self.total_progress_bar.update()
        self.clear_search_results()

    def clear_search_results(self):
        """검색 결과 초기화"""
        self.search_results = SearchResults()

    def find_book_info(self, isbn):
        """국립 중앙 도서관 API로 도서 정보 검색"""
        # 검색 중지 확인
        if not self.search_running:
            raise Exception("검색이 중지되었습니다")

        # 진행 상태 업데이트
        self.search_results.progress += 1

        # 실행 로그 업데이트 (간결하게)
        self.result_listbox.delete(7, tk.END)
        self.result_listbox.insert(tk.END, f"{isbn}에 대한 정보를 검색합니다")

        # API 요청
        url = f"https://www.nl.go.kr/NL/search/openApi/search.do?key={self.api_key}&kwd={isbn}&pageNum=1&pageSize=1&apiType=json"
        response = urllib.request.urlopen(url)
        response_data = response.read().decode('utf-8')
        data = json.loads(response_data).get('result')

        # 단일 진행도 초기화
        self.progress_value.set(0)
        self.progress_bar.update()

        # 제목 정보 검색
        self.result_listbox.insert(tk.END, "제목 정보 검색중(1/5)")
        self.result_listbox.see(tk.END)
        self.progress_value.set(20)
        self.progress_bar.update()

        try:
            title = data[0]['titleInfo']
            self.search_results.titles.append(title)
        except (IndexError, KeyError):
            self.search_results.error_list.append(isbn)
            raise Exception(f"ISBN {isbn}에 대한 정보를 찾을 수 없습니다")

        # 저자 정보 검색
        self.result_listbox.insert(tk.END, "저자 정보 검색중(2/5)")
        self.progress_value.set(40)
        self.progress_bar.update()

        try:
            author_info = str(data[0]["authorInfo"])
            author_name = re.sub(r"[^\uAC00-\uD7A30]", "", author_info)

            if author_name:
                # 저자 정보 정제
                if any(keyword in author_name for keyword in ['구성', '저자']):
                    author_name = author_name[2:]
                elif any(keyword in author_name for keyword in ['글쓴이', '엮은이', '지은이']):
                    author_name = author_name[3:]
                elif '글' in author_name:
                    author_name = author_name.replace('글', '')
            else:
                author_name = "저자정보없음"

            self.search_results.authors.append(author_name)

            # 저자 기호 생성
            if author_name != "저자정보없음":
                author_code = self.convert_korean_to_author_code(author_name)
                self.search_results.author_codes.append(author_code)
            else:
                self.search_results.author_codes.append("저자정보없음")
        except Exception:
            self.search_results.authors.append("저자정보없음")
            self.search_results.author_codes.append("저자정보없음")

        # ISBN 추가
        self.search_results.isbns.append(isbn)

        # 분류 번호 검색
        self.result_listbox.insert(tk.END, "분류 번호 검색중(3/5)")
        self.progress_value.set(60)
        self.progress_bar.update()

        try:
            call_no = data[0]['callNo']
            if not call_no:
                self.search_results.blank_list.append(isbn)
                self.search_results.classification_codes.append("-")

                kdc_code = data[0]['kdcCode1s']
                if not kdc_code:
                    self.search_results.main_categories.append("-")
                else:
                    self.search_results.main_categories.append(f"{kdc_code}00")
            else:
                first_char = call_no[:1]
                if '0' <= first_char <= '9':
                    self.search_results.main_categories.append(call_no[0:3])
                    self.search_results.classification_codes.append("")
                else:
                    self.search_results.main_categories.append(call_no[1:4])
                    self.search_results.classification_codes.append(call_no[:1])
        except Exception:
            self.search_results.classification_codes.append("-")
            self.search_results.main_categories.append("-")

        # 출판사 정보 검색
        self.result_listbox.insert(tk.END, "출판사 정보 검색중(4/5)")
        self.progress_value.set(80)
        self.progress_bar.update()

        try:
            publisher = data[0]['pubInfo']
            if not publisher:
                self.search_results.blank_list.append(isbn)
                self.search_results.publishers.append("-")
            else:
                self.search_results.publishers.append(publisher)
        except Exception:
            self.search_results.publishers.append("-")

        # 출판 연도 검색
        self.result_listbox.insert(tk.END, "출판 연도 검색중(5/5)")
        self.progress_value.set(100)
        self.progress_bar.update()

        try:
            pub_year = data[0]['pubYearInfo']
            if not pub_year:
                self.search_results.blank_list.append(isbn)
                self.search_results.publication_years.append("출판연도정보없음")
            else:
                self.search_results.publication_years.append(pub_year[:4])
        except Exception:
            self.search_results.publication_years.append("출판연도정보없음")

    def convert_korean_to_author_code(self, author_name):
        """한글 저자명을 저자 기호로 변환"""
        # 첫 글자의 초성 추출
        first_char_components = self.decompose_korean_char(author_name[0])
        first_consonant = first_char_components[0]

        # 두 번째 글자의 중성 추출 (있는 경우)
        if len(author_name) > 1:
            second_char_components = self.decompose_korean_char(author_name[1])
            second_vowel = second_char_components[1]
            second_consonant = second_char_components[0]

            # 중성을 숫자로 변환
            if first_consonant == 'ㅊ':
                vowel_to_num = {
                    'ㅏ': 2, 'ㅐ': 2, 'ㅑ': 2, 'ㅒ': 2, 'ㅓ': 3, 'ㅔ': 3, 'ㅕ': 3, 'ㅖ': 3,
                    'ㅗ': 4, 'ㅘ': 4, 'ㅙ': 4, 'ㅚ': 4, 'ㅛ': 4, 'ㅜ': 5, 'ㅝ': 5, 'ㅞ': 5,
                    'ㅟ': 5, 'ㅠ': 5, 'ㅡ': 5, 'ㅢ': 5, 'ㅣ': 6
                }
            else:
                vowel_to_num = {
                    'ㅏ': 2, 'ㅐ': 3, 'ㅑ': 3, 'ㅒ': 3, 'ㅓ': 4, 'ㅔ': 4, 'ㅕ': 4, 'ㅖ': 4,
                    'ㅗ': 5, 'ㅘ': 5, 'ㅙ': 5, 'ㅚ': 5, 'ㅛ': 5, 'ㅜ': 6, 'ㅝ': 6, 'ㅞ': 6,
                    'ㅟ': 6, 'ㅠ': 6, 'ㅡ': 7, 'ㅢ': 7, 'ㅣ': 8
                }

            vowel_num = vowel_to_num.get(second_vowel, 0)

            # 자음을 숫자로 변환
            consonant_to_num = {
                'ㄱ': 1, 'ㄲ': 1, 'ㄴ': 19, 'ㄷ': 2, 'ㄸ': 2, 'ㄹ': 29, 'ㅁ': 3, 'ㅂ': 4,
                'ㅃ': 4, 'ㅅ': 5, 'ㅆ': 5, 'ㅇ': 6, 'ㅈ': 7, 'ㅉ': 7, 'ㅊ': 8, 'ㅋ': 87,
                'ㅌ': 88, 'ㅍ': 89, 'ㅎ': 9
            }

            consonant_num = consonant_to_num.get(second_consonant, 0)

            # 저자 기호 생성
            return f"{author_name[0]}{vowel_num}{consonant_num}"
        else:
            # 한 글자인 경우
            return author_name

    def decompose_korean_char(self, char):
        """한글 문자를 초성, 중성, 종성으로 분해"""
        if '가' <= char <= '힣':
            char_code = ord(char) - ord('가')

            # 초성 (19개)
            consonants = [
                'ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ',
                'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'
            ]

            # 중성 (21개)
            vowels = [
                'ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ',
                'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ'
            ]

            # 종성 (28개, 공백 포함)
            finals = [
                ' ', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ',
                'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ',
                'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'
            ]

            consonant_index = char_code // 588
            vowel_index = (char_code - (588 * consonant_index)) // 28
            final_index = char_code - (588 * consonant_index) - (28 * vowel_index)

            return [consonants[consonant_index], vowels[vowel_index], finals[final_index]]
        else:
            return [char, '', '']

    def convert_isbn13_to_isbn10(self, isbn13):
        """ISBN-13을 ISBN-10으로 변환"""
        isbn10_base = isbn13[3:12]

        # 체크섬 계산
        sum_val = 0
        for i in range(9):
            sum_val += int(isbn10_base[i]) * (10 - i)

        check_digit = (11 - (sum_val % 11)) % 11
        if check_digit == 10:
            check_digit = 'X'

        return f"{isbn10_base}{check_digit}"


@dataclass
class SearchResults:
    """검색 결과를 저장하는 데이터 클래스"""
    count: int = 0
    recount: int = 0
    error_count: int = 0
    progress: int = 0

    titles: List[str] = field(default_factory=list)
    authors: List[str] = field(default_factory=list)
    author_codes: List[str] = field(default_factory=list)
    classification_codes: List[str] = field(default_factory=list)
    main_categories: List[str] = field(default_factory=list)
    isbns: List[str] = field(default_factory=list)
    publishers: List[str] = field(default_factory=list)
    publication_years: List[str] = field(default_factory=list)

    error_list: List[str] = field(default_factory=list)
    blank_list: List[str] = field(default_factory=list)


def main():
    root = tk.Tk()
    app = LibrarySearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
