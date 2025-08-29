import requests
from bs4 import BeautifulSoup
import time
import os
import csv
from datetime import datetime
import logging

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class CourtNoticeDownloader:
    def __init__(self, base_dir="court_downloads"):
        self.base_url = "https://www.scourt.go.kr"
        self.list_url = f"{self.base_url}/portal/notice/realestate/RealNoticeList.work"
        self.view_url = f"{self.base_url}/portal/notice/realestate/RealNoticeView.work"
        self.download_url = "https://file.scourt.go.kr/AttachDownload"
        self.base_dir = base_dir
        self.session = requests.Session()

        # 디렉토리 생성
        os.makedirs(base_dir, exist_ok=True)

    def get_all_notices(self, start_page=1, end_page=44):
        """모든 공고 목록 수집"""
        all_notices = []

        for page in range(start_page, end_page + 1):
            logging.info(f"페이지 {page}/{end_page} 수집 중...")

            params = {
                'pageIndex': page,
                'bub_cd': '',
                'searchWord': '',
                'searchOption': ''
            }

            try:
                response = self.session.get(self.list_url, params=params)
                response.encoding = 'euc-kr'  # 한글 인코딩 설정
                soup = BeautifulSoup(response.text, 'html.parser')

                # 테이블에서 각 행 추출
                rows = soup.select('table.tableHor tbody tr')

                for row in rows:
                    cells = row.select('td')
                    if len(cells) >= 5:
                        # 링크에서 seq_id 추출
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

                time.sleep(0.5)  # 서버 부하 방지

            except Exception as e:
                logging.error(f"페이지 {page} 수집 실패: {str(e)}")
                continue

        logging.info(f"총 {len(all_notices)}개 공고 수집 완료")
        return all_notices

    def get_notice_detail(self, seq_id):
        """상세 페이지에서 첨부파일 정보 추출"""
        url = f"{self.view_url}?seq_id={seq_id}"

        try:
            response = self.session.get(url)
            response.encoding = 'euc-kr'
            soup = BeautifulSoup(response.text, 'html.parser')

            # 첨부파일 링크 찾기
            attachments = []
            for link in soup.select('a[href*="javascript:download"]'):
                onclick = link.get('href', '')
                if 'download(' in onclick:
                    # JavaScript 함수에서 파라미터 추출
                    import re
                    match = re.search(r"download\('([^']+)','([^']+)'\)", onclick)
                    if match:
                        file_id = match.group(1)
                        file_name = match.group(2)
                        attachments.append({
                            'file_id': file_id,
                            'file_name': file_name
                        })

            # 추가 정보 추출
            detail_info = {}
            table = soup.select_one('table.tableVer')
            if table:
                for row in table.select('tr'):
                    th = row.select_one('th')
                    td = row.select_one('td')
                    if th and td:
                        key = th.text.strip()
                        value = td.text.strip()
                        detail_info[key] = value

            return attachments, detail_info

        except Exception as e:
            logging.error(f"상세 페이지 {seq_id} 조회 실패: {str(e)}")
            return [], {}

    def download_file(self, file_id, file_name, court, seq_id):
        """파일 다운로드"""
        # 법원별 폴더 생성
        court_dir = os.path.join(self.base_dir, court.replace('/', '_'))
        os.makedirs(court_dir, exist_ok=True)

        # 파일명 안전하게 변경 (seq_id 포함)
        safe_filename = f"{seq_id}_{file_name}".replace('/', '_').replace('\\', '_')
        file_path = os.path.join(court_dir, safe_filename)

        # 이미 다운로드된 파일인지 확인
        if os.path.exists(file_path):
            logging.info(f"이미 존재: {safe_filename}")
            return True

        # POST 데이터 준비
        data = {
            'file': file_id,
            'path': '011',
            'downFile': file_name
        }

        try:
            response = self.session.post(self.download_url, data=data, stream=True)

            if response.status_code == 200:
                # 파일 저장
                with open(file_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                logging.info(f"다운로드 완료: {safe_filename}")
                return True
            else:
                logging.error(f"다운로드 실패 ({response.status_code}): {file_name}")
                return False

        except Exception as e:
            logging.error(f"다운로드 오류: {str(e)}")
            return False

    def save_metadata(self, notices_with_details):
        """메타데이터를 CSV로 저장"""
        csv_path = os.path.join(self.base_dir, f"metadata_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            fieldnames = ['seq_id', 'number', 'court', 'agency', 'title',
                          'views', 'page', 'file_name', 'download_status']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

            for notice in notices_with_details:
                writer.writerow(notice)

        logging.info(f"메타데이터 저장: {csv_path}")

    def run(self, start_page=1, end_page=44, delay=1):
        """전체 프로세스 실행"""
        # 1단계: 모든 공고 수집
        notices = self.get_all_notices(start_page, end_page)

        # 2단계: 각 공고의 상세 정보 및 파일 다운로드
        results = []
        total = len(notices)

        for idx, notice in enumerate(notices, 1):
            logging.info(f"처리 중 [{idx}/{total}]: {notice['title']}")

            # 상세 정보 가져오기
            attachments, detail_info = self.get_notice_detail(notice['seq_id'])

            if attachments:
                for attachment in attachments:
                    # 파일 다운로드
                    success = self.download_file(
                        attachment['file_id'],
                        attachment['file_name'],
                        notice['court'],
                        notice['seq_id']
                    )

                    # 결과 저장
                    result = notice.copy()
                    result['file_name'] = attachment['file_name']
                    result['download_status'] = 'success' if success else 'failed'
                    results.append(result)
            else:
                # 첨부파일 없음
                result = notice.copy()
                result['file_name'] = ''
                result['download_status'] = 'no_attachment'
                results.append(result)

            time.sleep(delay)  # 서버 부하 방지

        # 3단계: 메타데이터 저장
        self.save_metadata(results)

        # 통계 출력
        success_count = sum(1 for r in results if r['download_status'] == 'success')
        failed_count = sum(1 for r in results if r['download_status'] == 'failed')
        no_attach_count = sum(1 for r in results if r['download_status'] == 'no_attachment')

        logging.info("=" * 50)
        logging.info(f"완료! 성공: {success_count}, 실패: {failed_count}, 첨부파일 없음: {no_attach_count}")

    def download_specific_pages(self, pages, delay=1):
        """특정 페이지들의 파일만 다운로드"""
        if isinstance(pages, int):
            pages = [pages]

        all_notices = []

        # 지정된 페이지들만 수집
        for page in pages:
            logging.info(f"페이지 {page} 수집 중...")
            notices = self.get_all_notices(start_page=page, end_page=page)
            all_notices.extend(notices)

        # 파일 다운로드 처리
        results = []
        total = len(all_notices)

        for idx, notice in enumerate(all_notices, 1):
            logging.info(f"처리 중 [{idx}/{total}]: {notice['title']}")

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

        # 메타데이터 저장
        self.save_metadata(results)

        # 통계 출력
        success_count = sum(1 for r in results if r['download_status'] == 'success')
        failed_count = sum(1 for r in results if r['download_status'] == 'failed')
        no_attach_count = sum(1 for r in results if r['download_status'] == 'no_attachment')

        logging.info("=" * 50)
        logging.info(f"페이지 {pages} 완료!")
        logging.info(f"성공: {success_count}, 실패: {failed_count}, 첨부파일 없음: {no_attach_count}")


# 실행 예제
if __name__ == "__main__":
    downloader = CourtNoticeDownloader(base_dir="court_downloads")

    downloader.download_specific_pages([1])
