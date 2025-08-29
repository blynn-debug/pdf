# court_downloader_gui_final.py
import os
import sys

# Tcl/Tk 경로 설정 (tkinter import 전에 반드시 설정)
os.environ['TCL_LIBRARY'] = r'C:\Users\청현\Python13\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\청현\Python13\tcl\tk8.6'

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import queue
import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime
import re
import json


class CourtNoticeDownloader:
    def __init__(self, base_dir="court_downloads", log_queue=None):
        self.base_url = "https://www.scourt.go.kr"
        self.list_url = f"{self.base_url}/portal/notice/realestate/RealNoticeList.work"
        self.view_url = f"{self.base_url}/portal/notice/realestate/RealNoticeView.work"
        self.download_url = "https://file.scourt.go.kr/AttachDownload"
        self.base_dir = base_dir
        self.session = requests.Session()
        self.log_queue = log_queue
        self.stop_flag = False

        # 헤더 설정
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Origin': 'https://www.scourt.go.kr',
            'Referer': 'https://www.scourt.go.kr/portal/notice/realestate/RealNoticeList.work'
        })

        # 다운로드 디렉토리 생성 (필요한 경우)
        if not os.path.exists(base_dir):
            os.makedirs(base_dir)

    def log(self, message, level="INFO"):
        if self.log_queue:
            self.log_queue.put(f"[{level}] {message}")
        else:
            print(f"[{level}] {message}")

    def search_notices(self, keyword="", pages=None):
        """빠른 검색을 위한 개선된 함수"""
        if pages is None:
            pages = [1]

        all_notices = []

        for page in pages:
            if self.stop_flag:
                self.log("사용자에 의해 중지됨")
                break

            self.log(f"페이지 {page} 스캔 중...")

            # 전체 목록 가져오기 (검색어 없이)
            form_data = {
                'pageIndex': str(page),
                'bub_cd': '',
                'searchWord': '',  # 빈 값으로 전체 목록 요청
                'searchOption': '',
                'pageSize': '10'
            }

            try:
                response = self.session.post(self.list_url, data=form_data, timeout=10)
                response.encoding = 'euc-kr'
                soup = BeautifulSoup(response.text, 'html.parser')

                # 테이블 찾기
                table = soup.find('table', {'class': 'tableHor'})
                if not table:
                    table = soup.find('table', {'summary': lambda x: x and '공고게시판' in x})

                if not table:
                    self.log(f"페이지 {page}: 테이블을 찾을 수 없음", "WARNING")
                    continue

                tbody = table.find('tbody')
                if not tbody:
                    self.log(f"페이지 {page}: tbody가 없습니다", "WARNING")
                    continue

                rows = tbody.find_all('tr')
                if not rows:
                    self.log(f"페이지 {page}: 데이터가 없습니다", "WARNING")
                    break  # 더 이상 페이지가 없음

                # 각 행을 빠르게 스캔
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) >= 5:
                        # 매각기관 텍스트 (세 번째 열)
                        agency_text = cells[2].text.strip()
                        # 제목 텍스트 (네 번째 열)
                        title_text = cells[3].text.strip()

                        # 키워드가 있으면 필터링
                        if keyword:
                            # 매각기관 또는 제목에 키워드가 포함되어 있는지 확인
                            search_text = f"{agency_text} {title_text}".lower()
                            if keyword.lower() not in search_text:
                                continue  # 키워드가 없으면 건너뛰기

                        # 키워드가 포함된 경우 또는 전체 검색인 경우
                        # 번호 추출 (첫 번째 열)
                        number = cells[0].text.strip()
                        court = cells[1].text.strip()

                        # 링크에서 seq_id 추출
                        link = cells[3].find('a')
                        if link and 'href' in link.attrs:
                            href = link['href']
                            seq_id_match = re.search(r'seq_id=(\d+)', href)

                            if seq_id_match:
                                seq_id = seq_id_match.group(1)
                                title = link.text.strip()
                                views = cells[4].text.strip()

                                notice = {
                                    'seq_id': seq_id,
                                    'number': number,
                                    'court': court,
                                    'agency': agency_text,
                                    'title': title,
                                    'views': views,
                                    'page': page,
                                    'keyword': keyword
                                }

                                all_notices.append(notice)

                                if keyword:
                                    self.log(f"✓ [{number}] {title} - 매각기관: {agency_text}")
                                else:
                                    self.log(f"  [{number}] {title}", "DEBUG")

                time.sleep(0.3)  # 서버 부하 방지

            except Exception as e:
                self.log(f"페이지 {page} 처리 오류: {str(e)}", "ERROR")
                continue

        if keyword:
            self.log(f"'{keyword}' 검색 결과: {len(all_notices)}개 발견")
        else:
            self.log(f"총 {len(all_notices)}개 공고 수집 완료")

        return all_notices

    def get_notice_detail(self, seq_id):
        """상세 페이지에서 첨부파일 정보 추출"""
        url = f"{self.view_url}?seq_id={seq_id}"

        try:
            response = self.session.get(url, timeout=10)
            response.encoding = 'euc-kr'
            soup = BeautifulSoup(response.text, 'html.parser')

            attachments = []

            # javascript:download 형식의 링크 찾기
            for link in soup.find_all('a'):
                href = link.get('href', '')
                onclick = link.get('onclick', '')

                # href 또는 onclick에서 download 함수 찾기
                download_text = href if 'download' in href else onclick

                if 'download' in download_text:
                    # download('파일ID','파일명') 형식 파싱
                    match = re.search(r"download\s*\(\s*'([^']+)'\s*,\s*'([^']+)'\s*\)", download_text)
                    if match:
                        file_id = match.group(1)
                        file_name = match.group(2)
                        attachments.append({
                            'file_id': file_id,
                            'file_name': file_name
                        })
                        self.log(f"  첨부파일 발견: {file_name}")

            if not attachments:
                self.log("  첨부파일 없음")

            return attachments, {}

        except Exception as e:
            self.log(f"상세 페이지 {seq_id} 조회 실패: {str(e)}", "ERROR")
            return [], {}

    def download_file(self, file_id, file_name, court, seq_id):
        """파일 다운로드 - 지정된 경로에 직접 저장"""
        # 파일명을 안전하게 변환
        safe_filename = f"{seq_id}_{file_name}"
        safe_filename = re.sub(r'[<>:"/\\|?*]', '_', safe_filename)

        # 지정된 경로에 직접 저장 (법원 폴더 생성하지 않음)
        file_path = os.path.join(self.base_dir, safe_filename)

        if os.path.exists(file_path):
            self.log(f"  이미 존재: {safe_filename}")
            return True

        data = {
            'file': file_id,
            'path': '011',
            'downFile': file_name
        }

        try:
            response = self.session.post(self.download_url, data=data, stream=True, timeout=30)

            if response.status_code == 200:
                with open(file_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

                self.log(f"  ✓ 다운로드 완료: {safe_filename}")
                return True
            else:
                self.log(f"  ✗ 다운로드 실패 ({response.status_code}): {file_name}", "ERROR")
                return False

        except Exception as e:
            self.log(f"  ✗ 다운로드 오류: {str(e)}", "ERROR")
            return False

    def process_notices(self, notices, delay=1):
        """수집된 공고들을 처리하고 파일 다운로드"""
        results = []
        total = len(notices)

        for idx, notice in enumerate(notices, 1):
            if self.stop_flag:
                self.log("사용자에 의해 중지됨")
                break

            self.log(f"\n[{idx}/{total}] 처리 중: [{notice['number']}] {notice['title']}")
            self.log(f"  매각기관: {notice['agency']}")
            self.log(f"  법원: {notice['court']}")

            attachments, detail_info = self.get_notice_detail(notice['seq_id'])

            if attachments:
                for attachment in attachments:
                    success = self.download_file(
                        attachment['file_id'],
                        attachment['file_name'],
                        notice['court'],
                        notice['seq_id']
                    )

                    result = notice.copy()
                    result['file_name'] = attachment['file_name']
                    result['download_status'] = '성공' if success else '실패'
                    results.append(result)
            else:
                result = notice.copy()
                result['file_name'] = ''
                result['download_status'] = '첨부파일 없음'
                results.append(result)

            time.sleep(delay)

        # 결과 요약
        success_count = sum(1 for r in results if r['download_status'] == '성공')
        failed_count = sum(1 for r in results if r['download_status'] == '실패')
        no_attach_count = sum(1 for r in results if r['download_status'] == '첨부파일 없음')

        self.log("\n" + "=" * 50)
        self.log(f"완료! 성공: {success_count}, 실패: {failed_count}, 첨부파일 없음: {no_attach_count}")
        self.log(f"다운로드 경로: {os.path.abspath(self.base_dir)}")

        return results

    def stop(self):
        """다운로드 중지"""
        self.stop_flag = True


class DownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("법원 공고 파일 다운로더 v3.0")
        self.root.geometry("950x750")

        self.log_queue = queue.Queue()
        self.downloading = False
        self.downloader = None
        self.download_path = "court_downloads"

        # 설정 로드
        self.load_settings()

        self.setup_ui()
        self.check_queue()

    def load_settings(self):
        """설정 파일 로드"""
        try:
            if os.path.exists('settings.json'):
                with open('settings.json', 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.download_path = settings.get('download_path', 'court_downloads')
        except:
            pass

    def save_settings(self):
        """설정 파일 저장"""
        try:
            settings = {
                'download_path': self.download_path
            }
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except:
            pass

    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 제목
        title_label = ttk.Label(main_frame, text="법원 회생·파산 자산매각 공고 다운로더",
                                font=('맑은 고딕', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # 다운로드 경로 설정
        path_frame = ttk.LabelFrame(main_frame, text="다운로드 경로", padding="10")
        path_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.path_label = ttk.Label(path_frame, text=self.download_path,
                                    relief=tk.SUNKEN, padding=5)
        self.path_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)

        ttk.Button(path_frame, text="경로 변경",
                   command=self.change_download_path).grid(row=0, column=1, padx=5)

        path_frame.columnconfigure(0, weight=1)

        # 검색 프레임
        search_frame = ttk.LabelFrame(main_frame, text="검색 옵션", padding="10")
        search_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # 키워드 입력
        ttk.Label(search_frame, text="검색 키워드:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.keyword_entry = ttk.Entry(search_frame, width=30)
        self.keyword_entry.grid(row=0, column=1, padx=5)

        # 설명 레이블
        info_label = ttk.Label(search_frame,
                               text="매각기관 또는 제목에서 검색 (비워두면 전체)",
                               font=('맑은 고딕', 9))
        info_label.grid(row=0, column=2, padx=5)

        # 예시
        example_label = ttk.Label(search_frame,
                                  text="예시: 공동선, 제일도어, 채권, 부동산",
                                  font=('맑은 고딕', 9), foreground='gray')
        example_label.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        # 페이지 선택
        mode_frame = ttk.LabelFrame(main_frame, text="페이지 선택", padding="10")
        mode_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.download_mode = tk.StringVar(value="specific")

        # 특정 페이지 모드
        specific_frame = ttk.Frame(mode_frame)
        specific_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))

        ttk.Radiobutton(specific_frame, text="특정 페이지",
                        variable=self.download_mode, value="specific",
                        command=self.toggle_mode).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(specific_frame, text="페이지:").grid(row=0, column=1, padx=(20, 5))
        self.page_entry = ttk.Entry(specific_frame, width=20)
        self.page_entry.grid(row=0, column=2, padx=5)
        self.page_entry.insert(0, "1")
        ttk.Label(specific_frame, text="(쉼표로 구분, 예: 1,2,3)",
                  font=('맑은 고딕', 9)).grid(row=0, column=3, padx=5)

        # 범위 모드
        range_frame = ttk.Frame(mode_frame)
        range_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Radiobutton(range_frame, text="페이지 범위",
                        variable=self.download_mode, value="range",
                        command=self.toggle_mode).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(range_frame, text="시작:").grid(row=0, column=1, padx=(20, 5))
        self.start_page = ttk.Spinbox(range_frame, from_=1, to=100, width=10, state='disabled')
        self.start_page.grid(row=0, column=2, padx=5)
        self.start_page.set(1)

        ttk.Label(range_frame, text="종료:").grid(row=0, column=3, padx=5)
        self.end_page = ttk.Spinbox(range_frame, from_=1, to=100, width=10, state='disabled')
        self.end_page.grid(row=0, column=4, padx=5)
        self.end_page.set(5)

        # 전체 페이지 모드 추가
        all_frame = ttk.Frame(mode_frame)
        all_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Radiobutton(all_frame, text="전체 페이지",
                        variable=self.download_mode, value="all",
                        command=self.toggle_mode).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(all_frame, text="(최대 44페이지까지 검색)",
                  font=('맑은 고딕', 9), foreground='gray').grid(row=0, column=1, padx=(20, 5))

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)

        self.search_btn = ttk.Button(button_frame, text="검색 및 다운로드",
                                     command=self.start_download)
        self.search_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = ttk.Button(button_frame, text="중지",
                                   command=self.stop_download,
                                   state='disabled')
        self.stop_btn.grid(row=0, column=1, padx=5)

        ttk.Button(button_frame, text="로그 지우기",
                   command=self.clear_log).grid(row=0, column=2, padx=5)

        # 로그 창
        log_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=100)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 프로그레스 바
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # 상태 표시
        self.status_label = ttk.Label(main_frame, text="준비됨", font=('맑은 고딕', 10))
        self.status_label.grid(row=7, column=0, columnspan=3)

        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def change_download_path(self):
        """다운로드 경로 변경"""
        new_path = filedialog.askdirectory(initialdir=self.download_path)
        if new_path:
            self.download_path = new_path
            self.path_label.config(text=self.download_path)
            self.save_settings()
            messagebox.showinfo("경로 변경", f"다운로드 경로가 변경되었습니다:\n{self.download_path}")

    def clear_log(self):
        """로그 창 지우기"""
        self.log_text.delete(1.0, tk.END)

    def toggle_mode(self):
        mode = self.download_mode.get()
        if mode == "specific":
            self.page_entry.config(state='normal')
            self.start_page.config(state='disabled')
            self.end_page.config(state='disabled')
        elif mode == "range":
            self.page_entry.config(state='disabled')
            self.start_page.config(state='normal')
            self.end_page.config(state='normal')
        else:  # all
            self.page_entry.config(state='disabled')
            self.start_page.config(state='disabled')
            self.end_page.config(state='disabled')

    def get_pages(self):
        """선택된 모드에 따라 페이지 목록 반환"""
        mode = self.download_mode.get()

        try:
            if mode == "specific":
                pages_text = self.page_entry.get()
                pages = []
                for p in pages_text.split(','):
                    p = p.strip()
                    if p.isdigit():
                        pages.append(int(p))
                return pages
            elif mode == "range":
                start = int(self.start_page.get())
                end = int(self.end_page.get())
                return list(range(start, end + 1))
            else:  # all
                return list(range(1, 45))  # 최대 44페이지
        except:
            return []

    def start_download(self):
        if self.downloading:
            messagebox.showwarning("경고", "이미 다운로드가 진행 중입니다.")
            return

        pages = self.get_pages()
        if not pages:
            messagebox.showerror("오류", "올바른 페이지를 입력해주세요.")
            return

        self.downloading = True
        self.search_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress.start()
        self.status_label.config(text="다운로드 중...")

        thread = threading.Thread(target=self.download_thread)
        thread.daemon = True
        thread.start()

    def stop_download(self):
        if self.downloader:
            self.downloader.stop()
        self.downloading = False
        self.status_label.config(text="중지됨")

    def download_thread(self):
        try:
            keyword = self.keyword_entry.get().strip()
            pages = self.get_pages()

            # 다운로드 경로 포함하여 downloader 생성
            self.downloader = CourtNoticeDownloader(
                base_dir=self.download_path,
                log_queue=self.log_queue
            )

            # 검색 수행
            self.log_queue.put(f"\n{'=' * 50}")
            self.log_queue.put(f"검색 시작: 키워드='{keyword}', 페이지={pages}")
            self.log_queue.put(f"다운로드 경로: {os.path.abspath(self.download_path)}")
            self.log_queue.put(f"{'=' * 50}\n")

            notices = self.downloader.search_notices(keyword=keyword, pages=pages)

            if not notices:
                self.log_queue.put("\n❌ 검색 결과가 없습니다.")
                if keyword:
                    self.log_queue.put(f"💡 '{keyword}'가 포함된 공고가 없습니다.")
                    self.log_queue.put("💡 다른 키워드를 시도하거나 페이지 범위를 넓혀보세요.")
                return

            # 다운로드 수행
            self.downloader.process_notices(notices)

            self.log_queue.put("\n✅ 작업 완료!")

        except Exception as e:
            self.log_queue.put(f"\n❌ 오류 발생: {str(e)}")
            import traceback
            self.log_queue.put(traceback.format_exc())

        finally:
            self.downloading = False
            self.root.after(0, self.download_complete)

    def download_complete(self):
        self.search_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress.stop()
        self.status_label.config(text="완료" if self.downloading else "준비됨")

        if not self.downloader.stop_flag:
            messagebox.showinfo("완료", "작업이 완료되었습니다.")

    def check_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                # DEBUG 메시지는 무시
                if not message.startswith("[DEBUG]"):
                    self.log_text.insert(tk.END, message + "\n")
                    self.log_text.see(tk.END)
        except queue.Empty:
            pass

        self.root.after(100, self.check_queue)


if __name__ == "__main__":
    root = tk.Tk()
    app = DownloaderGUI(root)
    root.mainloop()