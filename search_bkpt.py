# court_downloader_gui.py
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import queue
import requests
from bs4 import BeautifulSoup
import time
import os
import csv
from datetime import datetime
import logging


class CourtNoticeDownloader:
    def __init__(self, base_dir="court_downloads", log_queue=None):
        self.base_url = "https://www.scourt.go.kr"
        self.list_url = f"{self.base_url}/portal/notice/realestate/RealNoticeList.work"
        self.view_url = f"{self.base_url}/portal/notice/realestate/RealNoticeView.work"
        self.download_url = "https://file.scourt.go.kr/AttachDownload"
        self.base_dir = base_dir
        self.session = requests.Session()
        self.log_queue = log_queue

        os.makedirs(base_dir, exist_ok=True)

    def log(self, message, level="INFO"):
        if self.log_queue:
            self.log_queue.put(f"[{level}] {message}")
        else:
            print(f"[{level}] {message}")

    def get_all_notices(self, start_page=1, end_page=44):
        all_notices = []

        for page in range(start_page, end_page + 1):
            self.log(f"페이지 {page}/{end_page} 수집 중...")

            params = {
                'pageIndex': page,
                'bub_cd': '',
                'searchWord': '',
                'searchOption': ''
            }

            try:
                response = self.session.get(self.list_url, params=params)
                response.encoding = 'euc-kr'
                soup = BeautifulSoup(response.text, 'html.parser')

                rows = soup.select('table.tableHor tbody tr')

                for row in rows:
                    cells = row.select('td')
                    if len(cells) >= 5:
                        link = cells[3].select_one('a')
                        if link and 'href' in link.attrs:
                            href = link['href']
                            if 'seq_id=' in href:
                                seq_id = href.split('seq_id=')[-1]

                                notice = {
                                    'seq_id': seq_id,
                                    'number': cells[0].text.strip(),
                                    'court': cells[1].text.strip(),
                                    'agency': cells[2].text.strip(),
                                    'title': link.text.strip(),
                                    'views': cells[4].text.strip(),
                                    'page': page
                                }
                                all_notices.append(notice)

                time.sleep(0.5)

            except Exception as e:
                self.log(f"페이지 {page} 수집 실패: {str(e)}", "ERROR")
                continue

        self.log(f"총 {len(all_notices)}개 공고 수집 완료")
        return all_notices

    def get_notice_detail(self, seq_id):
        url = f"{self.view_url}?seq_id={seq_id}"

        try:
            response = self.session.get(url)
            response.encoding = 'euc-kr'
            soup = BeautifulSoup(response.text, 'html.parser')

            attachments = []
            for link in soup.select('a[href*="javascript:download"]'):
                onclick = link.get('href', '')
                if 'download(' in onclick:
                    import re
                    match = re.search(r"download\('([^']+)','([^']+)'\)", onclick)
                    if match:
                        file_id = match.group(1)
                        file_name = match.group(2)
                        attachments.append({
                            'file_id': file_id,
                            'file_name': file_name
                        })

            return attachments, {}

        except Exception as e:
            self.log(f"상세 페이지 {seq_id} 조회 실패: {str(e)}", "ERROR")
            return [], {}

    def download_file(self, file_id, file_name, court, seq_id):
        court_dir = os.path.join(self.base_dir, court.replace('/', '_'))
        os.makedirs(court_dir, exist_ok=True)

        safe_filename = f"{seq_id}_{file_name}".replace('/', '_').replace('\\', '_')
        file_path = os.path.join(court_dir, safe_filename)

        if os.path.exists(file_path):
            self.log(f"이미 존재: {safe_filename}")
            return True

        data = {
            'file': file_id,
            'path': '011',
            'downFile': file_name
        }

        try:
            response = self.session.post(self.download_url, data=data, stream=True)

            if response.status_code == 200:
                with open(file_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                self.log(f"다운로드 완료: {safe_filename}")
                return True
            else:
                self.log(f"다운로드 실패 ({response.status_code}): {file_name}", "ERROR")
                return False

        except Exception as e:
            self.log(f"다운로드 오류: {str(e)}", "ERROR")
            return False

    def download_specific_pages(self, pages, delay=1):
        if isinstance(pages, int):
            pages = [pages]

        all_notices = []

        for page in pages:
            self.log(f"페이지 {page} 수집 중...")
            notices = self.get_all_notices(start_page=page, end_page=page)
            all_notices.extend(notices)

        results = []
        total = len(all_notices)

        for idx, notice in enumerate(all_notices, 1):
            self.log(f"처리 중 [{idx}/{total}]: {notice['title']}")

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
                    result['download_status'] = 'success' if success else 'failed'
                    results.append(result)
            else:
                result = notice.copy()
                result['file_name'] = ''
                result['download_status'] = 'no_attachment'
                results.append(result)

            time.sleep(delay)

        success_count = sum(1 for r in results if r['download_status'] == 'success')
        failed_count = sum(1 for r in results if r['download_status'] == 'failed')
        no_attach_count = sum(1 for r in results if r['download_status'] == 'no_attachment')

        self.log("=" * 50)
        self.log(f"완료! 성공: {success_count}, 실패: {failed_count}, 첨부파일 없음: {no_attach_count}")


class DownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("법원 공고 파일 다운로더")
        self.root.geometry("800x600")

        self.log_queue = queue.Queue()
        self.downloading = False

        self.setup_ui()
        self.check_queue()

    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 제목
        title_label = ttk.Label(main_frame, text="법원 회생·파산 자산매각 공고 다운로더",
                                font=('맑은 고딕', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # 옵션 프레임
        option_frame = ttk.LabelFrame(main_frame, text="다운로드 옵션", padding="10")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # 라디오 버튼
        self.download_option = tk.StringVar(value="specific")

        ttk.Radiobutton(option_frame, text="특정 페이지만 다운로드",
                        variable=self.download_option, value="specific",
                        command=self.toggle_input).grid(row=0, column=0, sticky=tk.W)

        ttk.Radiobutton(option_frame, text="전체 다운로드 (1-44페이지)",
                        variable=self.download_option, value="all",
                        command=self.toggle_input).grid(row=1, column=0, sticky=tk.W, pady=5)

        # 페이지 입력 프레임
        self.page_frame = ttk.Frame(option_frame)
        self.page_frame.grid(row=0, column=1, padx=20)

        ttk.Label(self.page_frame, text="페이지 번호:").grid(row=0, column=0)
        self.page_entry = ttk.Entry(self.page_frame, width=30)
        self.page_entry.grid(row=0, column=1, padx=5)
        self.page_entry.insert(0, "1,2,3")

        ttk.Label(self.page_frame, text="(쉼표로 구분)",
                  font=('맑은 고딕', 9)).grid(row=1, column=1)

        # 다운로드 버튼
        self.download_btn = ttk.Button(main_frame, text="다운로드 시작",
                                       command=self.start_download,
                                       style='Accent.TButton')
        self.download_btn.grid(row=2, column=0, columnspan=2, pady=20)

        # 로그 창
        log_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 프로그레스 바
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def toggle_input(self):
        if self.download_option.get() == "specific":
            self.page_entry.config(state='normal')
        else:
            self.page_entry.config(state='disabled')

    def start_download(self):
        if self.downloading:
            messagebox.showwarning("경고", "이미 다운로드가 진행 중입니다.")
            return

        self.downloading = True
        self.download_btn.config(state='disabled')
        self.progress.start()
        self.log_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self.download_thread)
        thread.daemon = True
        thread.start()

    def download_thread(self):
        try:
            downloader = CourtNoticeDownloader(log_queue=self.log_queue)

            if self.download_option.get() == "specific":
                pages_text = self.page_entry.get()
                pages = [int(p.strip()) for p in pages_text.split(',') if p.strip()]

                if not pages:
                    raise ValueError("페이지 번호를 입력해주세요.")

                downloader.download_specific_pages(pages)
            else:
                # 전체 다운로드
                downloader.download_specific_pages(list(range(1, 45)))

            self.log_queue.put("다운로드 완료!")

        except Exception as e:
            self.log_queue.put(f"오류 발생: {str(e)}")

        finally:
            self.downloading = False
            self.root.after(0, self.download_complete)

    def download_complete(self):
        self.download_btn.config(state='normal')
        self.progress.stop()
        messagebox.showinfo("완료", "다운로드가 완료되었습니다.")

    def check_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass

        self.root.after(100, self.check_queue)


if __name__ == "__main__":
    root = tk.Tk()
    app = DownloaderGUI(root)
    root.mainloop()