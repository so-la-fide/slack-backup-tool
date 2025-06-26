import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext, Canvas
import threading
import json
import os
from datetime import datetime, timedelta
import requests
from pathlib import Path
import webbrowser
import platform
import subprocess

class SlackBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Slack 백업 도구 v1.0")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        
        # 스크롤 가능한 캔버스 생성
        canvas = Canvas(root)
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding="20")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        main_frame = scrollable_frame
        
        # 타이틀
        title_label = ttk.Label(main_frame, text="🔄 Slack 백업 도구", font=('Arial', 20, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        subtitle_label = ttk.Label(main_frame, text="회사 Slack의 대화 내용을 안전하게 백업하세요", font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        # Step 1: 토큰 입력
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Slack Token 입력", padding="10")
        step1_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(step1_frame, text="OAuth Token (xoxp-로 시작):").grid(row=0, column=0, sticky=tk.W)
        self.token_entry = ttk.Entry(step1_frame, width=50, show="*")
        self.token_entry.grid(row=1, column=0, pady=(5, 0))
        
        help_button = ttk.Button(step1_frame, text="토큰 발급 방법", command=self.show_token_help)
        help_button.grid(row=1, column=1, padx=(10, 0))
        
        # Step 2: 백업 옵션
        step2_frame = ttk.LabelFrame(main_frame, text="Step 2: 백업 옵션 선택", padding="10")
        step2_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # 체크박스들
        self.public_channels_var = tk.BooleanVar(value=True)
        self.private_channels_var = tk.BooleanVar(value=True)
        self.direct_messages_var = tk.BooleanVar(value=True)
        self.threads_var = tk.BooleanVar(value=True)
        self.files_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(step2_frame, text="공개 채널", variable=self.public_channels_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="비공개 채널 (참여한 채널만)", variable=self.private_channels_var).grid(row=0, column=1, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="다이렉트 메시지", variable=self.direct_messages_var).grid(row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="스레드 댓글", variable=self.threads_var).grid(row=1, column=1, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="파일 다운로드", variable=self.files_var).grid(row=2, column=0, sticky=tk.W)
        
        # 기간 설정
        ttk.Label(step2_frame, text="백업 기간:").grid(row=3, column=0, sticky=tk.W, pady=(15, 5))
        self.period_var = tk.StringVar(value="all")
        period_combo = ttk.Combobox(step2_frame, textvariable=self.period_var, width=20, state="readonly")
        period_combo['values'] = ("전체 기간", "최근 1개월", "최근 3개월", "최근 6개월", "최근 1년")
        period_combo.current(0)
        period_combo.grid(row=4, column=0, sticky=tk.W)
        
        # 출력 형식
        ttk.Label(step2_frame, text="출력 형식:").grid(row=3, column=1, sticky=tk.W, pady=(15, 5))
        self.format_var = tk.StringVar(value="html")
        format_combo = ttk.Combobox(step2_frame, textvariable=self.format_var, width=20, state="readonly")
        format_combo['values'] = ("HTML", "PDF", "HTML + PDF")
        format_combo.current(0)
        format_combo.grid(row=4, column=1, sticky=tk.W)
        
        # Step 3: 저장 위치
        step3_frame = ttk.LabelFrame(main_frame, text="Step 3: 저장 위치 선택", padding="10")
        step3_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.save_path_var = tk.StringVar(value=str(Path.home() / "Downloads" / "slack_backup"))
        ttk.Label(step3_frame, text="저장 폴더:").grid(row=0, column=0, sticky=tk.W)
        path_entry = ttk.Entry(step3_frame, textvariable=self.save_path_var, width=40)
        path_entry.grid(row=1, column=0, pady=(5, 0))
        
        browse_button = ttk.Button(step3_frame, text="찾아보기", command=self.browse_folder)
        browse_button.grid(row=1, column=1, padx=(10, 0))
        
        # 경고 메시지
        warning_frame = ttk.Frame(main_frame)
        warning_frame.grid(row=5, column=0, columnspan=2, pady=(0, 15))
        warning_label = ttk.Label(warning_frame, text="⚠️ DM도 백업되므로 백업 파일 공유 시 주의하세요!", 
                                 foreground="orange", font=('Arial', 9, 'bold'))
        warning_label.pack()
        
        # 버튼들
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        self.backup_button = ttk.Button(button_frame, text="백업 시작", command=self.start_backup, 
                                       style='Accent.TButton')
        self.backup_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_button = ttk.Button(button_frame, text="취소", command=self.cancel_backup, 
                                       state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT)
        
        # 진행 상황
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 5))
        
        self.status_label = ttk.Label(main_frame, text="", font=('Arial', 9))
        self.status_label.grid(row=8, column=0, columnspan=2)
        
        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="실행 로그", padding="10")
        log_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, width=65, wrap=tk.WORD)
        self.log_text.pack()
        
        # 백업 중단 플래그
        self.stop_backup = False
        
    def show_token_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Slack Token 발급 방법")
        help_window.geometry("500x400")
        
        help_text = """Slack App 토큰 발급 방법:

1. https://api.slack.com/apps 접속
2. "Create New App" 클릭
3. "From scratch" 선택
4. App Name 입력, Workspace 선택
5. "OAuth & Permissions" 메뉴 클릭
6. User Token Scopes에 다음 권한 추가:
   • channels:read, channels:history
   • groups:read, groups:history  
   • mpim:read, mpim:history
   • im:read, im:history
   • users:read, files:read
7. "Install to Workspace" 클릭
8. xoxp로 시작하는 토큰 복사

※ 토큰은 비밀번호처럼 안전하게 보관하세요!"""
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=20, pady=20)
        text_widget.insert(1.0, help_text)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        open_button = ttk.Button(help_window, text="Slack API 페이지 열기", 
                                command=lambda: webbrowser.open("https://api.slack.com/apps"))
        open_button.pack(pady=10)
        
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.save_path_var.set(folder)
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_progress(self, value, status=""):
        self.progress_var.set(value)
        if status:
            self.status_label.config(text=status)
        self.root.update()
    
    def start_backup(self):
        # 입력 검증
        if not self.token_entry.get():
            messagebox.showerror("오류", "Slack Token을 입력해주세요.")
            return
        
        if not self.token_entry.get().startswith("xoxp-"):
            messagebox.showerror("오류", "올바른 User Token이 아닙니다.\nxoxp-로 시작하는 토큰을 입력해주세요.")
            return
        
        # 버튼 상태 변경
        self.backup_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.stop_backup = False
        
        # 백업 스레드 시작
        backup_thread = threading.Thread(target=self.run_backup)
        backup_thread.start()
    
    def cancel_backup(self):
        self.stop_backup = True
        self.log("백업을 중단하는 중...")
        self.cancel_button.config(state=tk.DISABLED)
    
    def run_backup(self):
        try:
            token = self.token_entry.get()
            save_path = Path(self.save_path_var.get())
            
            # 저장 폴더 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_folder = save_path / f"slack_backup_{timestamp}"
            backup_folder.mkdir(parents=True, exist_ok=True)
            
            self.log(f"백업 시작: {backup_folder}")
            self.update_progress(10, "Slack 연결 확인 중...")
            
            # Slack API 테스트
            headers = {"Authorization": f"Bearer {token}"}
            test_response = requests.get("https://slack.com/api/auth.test", headers=headers)
            
            if not test_response.json().get("ok"):
                raise Exception("토큰이 유효하지 않습니다.")
            
            team_name = test_response.json().get("team", "workspace")
            self.log(f"워크스페이스 확인: {team_name}")
            
            if self.stop_backup:
                self.cleanup_backup()
                return
            
            # 사용자 정보 가져오기
            self.update_progress(15, "사용자 정보를 가져오는 중...")
            users = {}
            users_response = requests.get("https://slack.com/api/users.list", headers=headers)
            if users_response.json().get("ok"):
                for user in users_response.json().get("members", []):
                    users[user['id']] = user.get('real_name', user.get('name', 'Unknown'))
            
            # 채널 목록 가져오기
            self.update_progress(20, "채널 목록을 가져오는 중...")
            channels = self.get_channels(headers)
            self.log(f"총 {len(channels)}개 채널 발견")
            
            if self.stop_backup:
                self.cleanup_backup()
                return
            
            # 메시지 백업
            self.update_progress(30, "메시지를 다운로드하는 중...")
            all_messages = {}
            
            for i, channel in enumerate(channels):
                if self.stop_backup:
                    self.cleanup_backup()
                    return
                
                progress = 30 + (40 * i / len(channels))
                
                # 채널 이름 안전하게 가져오기
                channel_name = channel.get('name')
                if not channel_name:
                    # DM의 경우 user ID를 사용
                    if channel.get('is_im'):
                        user_id = channel.get('user')
                        channel_name = f"DM-{users.get(user_id, user_id)}"
                    else:
                        channel_name = channel.get('id', 'Unknown')
                
                self.update_progress(progress, f"채널 백업 중: {channel_name}")
                self.log(f"채널 백업: {channel_name}")
                
                try:
                    messages = self.get_channel_messages(headers, channel['id'])
                    if messages:
                        all_messages[channel_name] = {
                            'channel_info': channel,
                            'messages': messages
                        }
                except Exception as e:
                    self.log(f"채널 {channel_name} 백업 중 오류: {str(e)}")
                    continue
            
            # 데이터 저장
            self.update_progress(80, "파일로 저장하는 중...")
            
            # JSON 저장
            json_file = backup_folder / "slack_backup.json"
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(all_messages, f, ensure_ascii=False, indent=2)
            self.log(f"JSON 파일 저장: {json_file}")
            
            # HTML 생성
            if self.format_var.get() in ["HTML", "HTML + PDF"]:
                self.update_progress(90, "HTML 파일 생성 중...")
                html_file = self.create_html_output(all_messages, backup_folder, users)
                self.log(f"HTML 파일 생성: {html_file}")
            
            # 완료
            self.update_progress(100, "백업 완료!")
            self.log("백업이 성공적으로 완료되었습니다.")
            
            # 완료 메시지
            if messagebox.askyesno("백업 완료", 
                                  f"백업이 완료되었습니다!\n\n저장 위치:\n{backup_folder}\n\n폴더를 열어보시겠습니까?"):
                if platform.system() == 'Windows':
                    os.startfile(backup_folder)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', backup_folder])
                else:  # Linux
                    subprocess.run(['xdg-open', backup_folder])
            
        except Exception as e:
            self.log(f"오류 발생: {str(e)}")
            messagebox.showerror("백업 실패", f"백업 중 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            self.cleanup_backup()
    
    def cleanup_backup(self):
        self.backup_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.update_progress(0, "")
    
    def get_channels(self, headers):
        channels = []
        
        # 공개 채널
        if self.public_channels_var.get():
            response = requests.get(
                "https://slack.com/api/conversations.list",
                headers=headers,
                params={"types": "public_channel", "limit": 1000}
            )
            if response.json().get("ok"):
                channels.extend(response.json().get("channels", []))
        
        # 비공개 채널
        if self.private_channels_var.get():
            response = requests.get(
                "https://slack.com/api/conversations.list",
                headers=headers,
                params={"types": "private_channel", "limit": 1000}
            )
            if response.json().get("ok"):
                channels.extend(response.json().get("channels", []))
        
        # DM
        if self.direct_messages_var.get():
            response = requests.get(
                "https://slack.com/api/conversations.list",
                headers=headers,
                params={"types": "im", "limit": 1000}
            )
            if response.json().get("ok"):
                channels.extend(response.json().get("channels", []))
        
        return channels
    
    def get_channel_messages(self, headers, channel_id):
        # 기간 계산
        oldest = 0
        if self.period_var.get() != "전체 기간":
            period_map = {
                "최근 1개월": 30,
                "최근 3개월": 90,
                "최근 6개월": 180,
                "최근 1년": 365
            }
            days = period_map.get(self.period_var.get(), 0)
            if days:
                oldest = int((datetime.now() - timedelta(days=days)).timestamp())
        
        # 메시지 가져오기
        messages = []
        cursor = None
        request_count = 0
        
        while True:
            params = {
                "channel": channel_id,
                "limit": 15,  # 새로운 제한: 최대 15개
                "oldest": oldest
            }
            if cursor:
                params["cursor"] = cursor
            
            try:
                response = requests.get(
                    "https://slack.com/api/conversations.history",
                    headers=headers,
                    params=params
                )
                
                if response.status_code == 429:  # Rate limit
                    retry_after = int(response.headers.get('Retry-After', 60))
                    self.log(f"API 제한 도달. {retry_after}초 대기...")
                    time.sleep(retry_after)
                    continue
                
                if not response.json().get("ok"):
                    break
                
                messages.extend(response.json().get("messages", []))
                
                # 페이징
                response_metadata = response.json().get("response_metadata", {})
                cursor = response_metadata.get("next_cursor")
                
                request_count += 1
                if cursor and request_count < 10:  # 채널당 최대 10회 요청
                    self.log(f"  - {len(messages)}개 메시지 수집 중...")
                    time.sleep(60)  # 분당 1회 제한
                else:
                    break
                    
            except Exception as e:
                self.log(f"  오류: {str(e)}")
                break
        
        return messages
    
    def create_html_output(self, all_messages, backup_folder, users):
        html_content = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slack 백업</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            background: #f5f5f5;
        }
        .header {
            background: #4a154b;
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
        }
        .channel {
            background: white;
            margin-bottom: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .channel-header {
            background: #f8f8f8;
            padding: 15px 20px;
            border-bottom: 1px solid #e0e0e0;
            font-weight: bold;
            font-size: 18px;
        }
        .messages {
            padding: 20px;
        }
        .message {
            margin-bottom: 20px;
            padding-bottom: 20px;
            border-bottom: 1px solid #f0f0f0;
        }
        .message:last-child {
            border-bottom: none;
        }
        .message-header {
            display: flex;
            align-items: center;
            margin-bottom: 8px;
        }
        .username {
            font-weight: bold;
            color: #1264a3;
            margin-right: 10px;
        }
        .timestamp {
            color: #666;
            font-size: 12px;
        }
        .message-text {
            color: #333;
            white-space: pre-wrap;
        }
        .thread-count {
            margin-top: 10px;
            color: #1264a3;
            font-size: 13px;
        }
        .nav {
            position: fixed;
            right: 20px;
            top: 20px;
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            max-height: 80vh;
            overflow-y: auto;
        }
        .nav h3 {
            margin-top: 0;
            color: #4a154b;
        }
        .nav ul {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .nav li {
            margin: 5px 0;
        }
        .nav a {
            color: #1264a3;
            text-decoration: none;
            font-size: 14px;
        }
        .nav a:hover {
            text-decoration: underline;
        }
        @media (max-width: 1200px) {
            .nav {
                position: static;
                margin-bottom: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="nav">
        <h3>채널 목록</h3>
        <ul>
"""
        
        # 네비게이션 추가
        for channel_name in sorted(all_messages.keys()):
            safe_id = channel_name.replace("#", "").replace(" ", "_").replace("-", "_")
            html_content += f'            <li><a href="#{safe_id}">{channel_name}</a></li>\n'
        
        html_content += """
        </ul>
    </div>
    
    <div class="header">
        <h1>Slack 백업</h1>
        <p>백업 일시: """ + datetime.now().strftime("%Y년 %m월 %d일 %H:%M") + """</p>
    </div>
"""
        
        # 채널별 메시지 추가
        for channel_name, data in sorted(all_messages.items()):
            safe_id = channel_name.replace("#", "").replace(" ", "_").replace("-", "_")
            html_content += f"""
    <div class="channel" id="{safe_id}">
        <div class="channel-header">#{channel_name}</div>
        <div class="messages">
"""
            
            # 메시지 정렬 (오래된 것부터)
            messages = sorted(data['messages'], key=lambda x: float(x.get('ts', 0)))
            
            for msg in messages:
                timestamp = datetime.fromtimestamp(float(msg.get('ts', 0))).strftime("%Y-%m-%d %H:%M:%S")
                user_id = msg.get('user', 'Unknown')
                username = users.get(user_id, user_id)
                text = msg.get('text', '')
                
                # HTML 이스케이프
                text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                
                html_content += f"""
            <div class="message">
                <div class="message-header">
                    <span class="username">{username}</span>
                    <span class="timestamp">{timestamp}</span>
                </div>
                <div class="message-text">{text}</div>
"""
                
                if msg.get('reply_count', 0) > 0:
                    html_content += f'                <div class="thread-count">💬 {msg["reply_count"]}개의 댓글</div>\n'
                
                html_content += "            </div>\n"
            
            html_content += """
        </div>
    </div>
"""
        
        html_content += """
</body>
</html>
"""
        
        # HTML 파일 저장
        html_file = backup_folder / "slack_backup.html"
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return html_file

def main():
    root = tk.Tk()
    app = SlackBackupTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
