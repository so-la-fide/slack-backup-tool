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
import html
import re

class SlackBackupTool:
    """
    Slack 대화 내용을 백업하는 GUI 애플리케이션입니다.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Slack 백업 도구 v1.5")
        self.root.geometry("650x750")
        self.root.resizable(False, False)
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Accent.TButton', foreground='white', background='#4a154b')
        
        # 스크롤 가능한 프레임 설정
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
        
        # --- UI 요소 배치 ---

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
        
        self.public_channels_var = tk.BooleanVar(value=True)
        self.private_channels_var = tk.BooleanVar(value=True)
        self.direct_messages_var = tk.BooleanVar(value=True)
        self.threads_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(step2_frame, text="공개 채널", variable=self.public_channels_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="비공개 채널 (그룹DM 포함)", variable=self.private_channels_var).grid(row=0, column=1, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="다이렉트 메시지", variable=self.direct_messages_var).grid(row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(step2_frame, text="스레드 댓글", variable=self.threads_var).grid(row=1, column=1, sticky=tk.W)
        
        ttk.Label(step2_frame, text="백업 기간:").grid(row=3, column=0, sticky=tk.W, pady=(15, 5))
        self.period_var = tk.StringVar(value="전체 기간")
        period_combo = ttk.Combobox(step2_frame, textvariable=self.period_var, width=20, state="readonly")
        period_combo['values'] = ("전체 기간", "최근 1개월", "최근 3개월", "최근 6개월", "최근 1년")
        period_combo.current(0)
        period_combo.grid(row=4, column=0, sticky=tk.W)
        
        ttk.Label(step2_frame, text="출력 형식:").grid(row=3, column=1, sticky=tk.W, pady=(15, 5))
        self.format_var = tk.StringVar(value="HTML")
        format_combo = ttk.Combobox(step2_frame, textvariable=self.format_var, width=20, state="readonly")
        format_combo['values'] = ("HTML",) # PDF 기능은 추후 확장 예정
        format_combo.current(0)
        format_combo.grid(row=4, column=1, sticky=tk.W)
        
        # Step 3: 채널 선택
        channel_frame = ttk.LabelFrame(main_frame, text="Step 3: 채널 선택 (선택사항)", padding="10")
        channel_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(channel_frame, text="특정 채널만 백업하려면 아래에서 선택하세요 (미선택 시 전체):").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        self.load_channels_button = ttk.Button(channel_frame, text="채널 목록 불러오기", command=self.load_channels)
        self.load_channels_button.grid(row=1, column=0, pady=(10, 0), sticky=tk.W)
        
        button_frame = ttk.Frame(channel_frame)
        button_frame.grid(row=1, column=1, pady=(10, 0), sticky=tk.W)
        
        self.select_all_button = ttk.Button(button_frame, text="전체 선택", command=self.select_all_channels, state=tk.DISABLED)
        self.select_all_button.pack(side=tk.LEFT, padx=(5, 0))
        
        self.deselect_all_button = ttk.Button(button_frame, text="전체 해제", command=self.deselect_all_channels, state=tk.DISABLED)
        self.deselect_all_button.pack(side=tk.LEFT, padx=(5, 0))
        
        list_frame = ttk.Frame(channel_frame)
        list_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        channel_scrollbar = ttk.Scrollbar(list_frame)
        channel_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.channel_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=8, yscrollcommand=channel_scrollbar.set)
        self.channel_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        channel_scrollbar.config(command=self.channel_listbox.yview)
        
        self.selected_count_label = ttk.Label(channel_frame, text="")
        self.selected_count_label.grid(row=3, column=0, columnspan=2, pady=(5, 0))
        
        self.available_channels = []
        self.channel_listbox.bind('<<ListboxSelect>>', self.update_selected_count)
        
        # Step 4: 저장 위치
        step4_frame = ttk.LabelFrame(main_frame, text="Step 4: 저장 위치 선택", padding="10")
        step4_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.save_path_var = tk.StringVar(value=str(Path.home() / "Downloads"))
        ttk.Label(step4_frame, text="저장 폴더:").grid(row=0, column=0, sticky=tk.W)
        path_entry = ttk.Entry(step4_frame, textvariable=self.save_path_var, width=40)
        path_entry.grid(row=1, column=0, pady=(5, 0))
        
        browse_button = ttk.Button(step4_frame, text="찾아보기", command=self.browse_folder)
        browse_button.grid(row=1, column=1, padx=(10, 0))
        
        # 경고 메시지
        warning_frame = ttk.Frame(main_frame)
        warning_frame.grid(row=6, column=0, columnspan=2, pady=(0, 15))
        warning_label = ttk.Label(warning_frame, text="⚠️ DM도 백업되므로 백업 파일 공유 시 주의하세요!", 
                                  foreground="orange", font=('Arial', 9, 'bold'))
        warning_label.pack()
        
        # 버튼들
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=(10, 0))
        
        self.backup_button = ttk.Button(button_frame, text="백업 시작", command=self.start_backup, 
                                          style='Accent.TButton')
        self.backup_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_button = ttk.Button(button_frame, text="취소", command=self.cancel_backup, 
                                          state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT)
        
        # 진행 상황
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 5))
        
        self.status_label = ttk.Label(main_frame, text="", font=('Arial', 9))
        self.status_label.grid(row=9, column=0, columnspan=2)
        
        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="실행 로그", padding="10")
        log_frame.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, width=65, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 백업 중단 플래그
        self.stop_backup = False
        self.save_on_cancel = False
    
    def show_token_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Slack Token 발급 방법")
        help_window.geometry("500x450")
        help_window.transient(self.root) # 부모 창 위에 표시

        help_text_content = """
    1. Slack API 사이트 접속
       https://api.slack.com/apps

    2. "Create New App" 클릭 → "From scratch" 선택

    3. 앱 이름(예: MyBackup) 입력 후 워크스페이스 선택

    4. 앱 생성 후 "OAuth & Permissions" 메뉴로 이동

    5. Scopes 섹션의 "User Token Scopes"에서 
       "Add an OAuth Scope" 버튼 클릭 후 아래 권한 추가:

       • channels:history (공개 채널 메시지 읽기)
       • channels:read (공개 채널 목록 읽기)
       • groups:history (비공개 채널 메시지 읽기)
       • groups:read (비공개 채널 목록 읽기)
       • im:history (DM 메시지 읽기)
       • im:read (DM 목록 읽기)
       • users:read (사용자 정보 읽기)

    6. 페이지 상단의 "Install to Workspace" 클릭 및 허용

    7. "User OAuth Token" 항목의 xoxp-로 시작하는 
       토큰을 복사하여 프로그램에 붙여넣기

    ※ 이 토큰은 개인 비밀번호와 같으니 절대 외부에 
       노출하지 마세요!
    """
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=20, pady=20, font=('Arial', 10))
        text_widget.insert(1.0, help_text_content)
        text_widget.config(state=tk.DISABLED, bg='#f0f0f0')
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        open_button = ttk.Button(help_window, text="Slack API 페이지 열기", 
                                 command=lambda: webbrowser.open("https://api.slack.com/apps"))
        open_button.pack(pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.save_path_var.set(folder)

    def load_channels(self):
        token = self.token_entry.get()
        if not token:
            messagebox.showerror("오류", "먼저 Slack Token을 입력해주세요.")
            return
        
        self.load_channels_button.config(state=tk.DISABLED, text="불러오는 중...")
        self.channel_listbox.delete(0, tk.END)
        self.update_selected_count()
        
        threading.Thread(target=self._fetch_channels_worker, args=(token,), daemon=True).start()

    def _fetch_channels_worker(self, token):
        try:
            headers = {"Authorization": f"Bearer {token}"}
            
            test_response = requests.post("https://slack.com/api/auth.test", headers=headers)
            test_response.raise_for_status()
            if not test_response.json().get("ok"):
                raise ValueError(f"토큰이 유효하지 않습니다: {test_response.json().get('error')}")

            self.log("사용자 정보 가져오는 중...")
            users_map = {user['id']: user.get('real_name', user.get('name', 'Unknown'))
                         for user in self._fetch_all_pages("users.list", headers, "members")}
            
            # Step 2의 옵션에 따라 가져올 채널 유형 결정
            types_to_fetch = []
            if self.public_channels_var.get():
                types_to_fetch.append("public_channel")
            if self.private_channels_var.get():
                types_to_fetch.append("private_channel")
                types_to_fetch.append("mpim")
            if self.direct_messages_var.get():
                types_to_fetch.append("im")
            
            if not types_to_fetch:
                self.log("선택된 채널 유형이 없습니다.")
                self.available_channels = []
                self.root.after(0, self._update_channel_listbox)
                return

            self.log(f"채널 목록 가져오는 중 ({', '.join(types_to_fetch)})...")
            types_string = ",".join(types_to_fetch)
            all_convos = self._fetch_all_pages(
                "conversations.list", headers, "channels", 
                params={"types": types_string, "limit": 200}
            )

            fetched_channels = []
            for c in all_convos:
                name = "unknown"
                if c.get('is_im'):
                    name = f"💬 {users_map.get(c.get('user'), 'Unknown User')}"
                elif c.get('is_mpim') or c.get('is_group'):
                     name = f"🔒{c.get('name', c.get('purpose', {}).get('value', '비공개 채널'))}"
                elif c.get('is_channel'):
                    name = f"#{c.get('name', '채널')}"
                
                fetched_channels.append({'id': c['id'], 'name': name, 'original': c})

            self.available_channels = sorted(fetched_channels, key=lambda x: x['name'].lower())

            self.root.after(0, self._update_channel_listbox)
            self.log(f"총 {len(self.available_channels)}개 채널/DM을 불러왔습니다.")

        except Exception as e:
            self.log(f"채널 목록 로드 실패: {e}")
            messagebox.showerror("오류", f"채널 목록을 불러오는 데 실패했습니다:\n{e}")
        finally:
            self.root.after(0, lambda: self.load_channels_button.config(state=tk.NORMAL, text="채널 목록 새로고침"))

    def _fetch_all_pages(self, api_method, headers, response_key, params=None):
        if params is None: params = {}
        items = []
        cursor = None
        while True:
            page_params = params.copy()
            if cursor:
                page_params["cursor"] = cursor
            
            response = requests.get(f"https://slack.com/api/{api_method}", headers=headers, params=page_params)
            
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 30))
                self.log(f"API 제한 도달. {retry_after}초 대기...")
                time.sleep(retry_after)
                continue

            response.raise_for_status()
            data = response.json()
            
            if not data.get("ok"):
                raise Exception(f"API 오류 ({api_method}): {data.get('error', 'Unknown error')}")
            
            items.extend(data.get(response_key, []))
            cursor = data.get("response_metadata", {}).get("next_cursor")
            if not cursor:
                break
            time.sleep(1.2)
        return items

    def _update_channel_listbox(self):
        self.channel_listbox.delete(0, tk.END)
        for ch in self.available_channels:
            self.channel_listbox.insert(tk.END, ch['name'])
        
        self.select_all_button.config(state=tk.NORMAL)
        self.deselect_all_button.config(state=tk.NORMAL)
        self.update_selected_count()

    def select_all_channels(self):
        self.channel_listbox.select_set(0, tk.END)
        self.update_selected_count()

    def deselect_all_channels(self):
        self.channel_listbox.select_clear(0, tk.END)
        self.update_selected_count()

    def update_selected_count(self, event=None):
        count = len(self.channel_listbox.curselection())
        total = self.channel_listbox.size()
        self.selected_count_label.config(text=f"선택됨: {count} / {total}개")

    def log(self, message):
        def _log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)
        self.root.after(0, _log)

    def update_progress(self, value, status=""):
        def _update():
            self.progress_var.set(value)
            if status:
                self.status_label.config(text=status)
        self.root.after(0, _update)

    def start_backup(self):
        if not self.token_entry.get().startswith("xoxp-"):
            messagebox.showerror("오류", "올바른 User Token(xoxp-...)을 입력해주세요.")
            return
        
        self.backup_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.stop_backup = False
        self.save_on_cancel = False
        
        threading.Thread(target=self.run_backup_worker, daemon=True).start()

    def cancel_backup(self):
        if messagebox.askyesno("백업 취소", "정말로 백업을 중단하시겠습니까?\n현재까지 진행된 내용은 저장됩니다."):
            self.stop_backup = True
            self.save_on_cancel = True
            self.log("백업을 중단하는 중... 잠시만 기다려주세요.")
            self.cancel_button.config(state=tk.DISABLED)

    def run_backup_worker(self):
        try:
            token = self.token_entry.get()
            save_path = Path(self.save_path_var.get())
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_folder = save_path / f"slack_backup_{timestamp}"
            backup_folder.mkdir(parents=True, exist_ok=True)
            
            self.log(f"백업 시작: {backup_folder}")
            self.update_progress(5, "Slack 연결 및 사용자 정보 확인 중...")
            
            headers = {"Authorization": f"Bearer {token}"}
            test_response = requests.post("https://slack.com/api/auth.test", headers=headers).json()
            if not test_response.get("ok"):
                raise Exception(f"토큰이 유효하지 않습니다: {test_response.get('error')}")
            
            team_name = test_response.get("team", "workspace")
            self.log(f"워크스페이스 확인: {team_name}")
            
            users_map = {user['id']: user.get('real_name', user.get('name', 'Unknown'))
                         for user in self._fetch_all_pages("users.list", headers, "members")}
            
            self.update_progress(15, "백업 대상 채널 목록 구성 중...")
            target_channels = self.get_target_channels()
            if not target_channels:
                raise ValueError("백업할 채널이 없습니다. 옵션을 확인해주세요.")
            self.log(f"총 {len(target_channels)}개 채널을 백업합니다.")
            
            if self.stop_backup: return

            self.update_progress(20, "메시지 다운로드 시작...")
            all_messages = {}
            total_channels = len(target_channels)
            
            for i, channel_info in enumerate(target_channels):
                if self.stop_backup:
                    self.log("사용자 요청으로 백업을 중단합니다.")
                    break
                
                progress = 20 + (70 * (i / total_channels))
                channel_name = channel_info['name']
                self.update_progress(progress, f"채널 백업 중 ({i+1}/{total_channels}): {channel_name}")
                self.log(f"채널 백업 시작: {channel_name}")
                
                try:
                    messages = self.get_channel_messages(headers, channel_info['id'])
                    if messages:
                        all_messages[channel_name] = {'messages': messages}
                except Exception as e:
                    self.log(f"채널 {channel_name} 백업 중 오류: {e}")
                
            if self.save_on_cancel and all_messages:
                self.log("중단됨 - 현재까지 수집한 데이터를 저장합니다...")
                self.save_backup_files(all_messages, backup_folder, users_map, is_partial=True)
                messagebox.showinfo("백업 중단", f"백업이 중단되었지만, 현재까지의 내용은 아래 폴더에 저장되었습니다:\n{backup_folder}")
                return

            self.update_progress(90, "백업 파일 생성 중...")
            self.save_backup_files(all_messages, backup_folder, users_map)
            
            self.update_progress(100, "백업 완료!")
            self.log("백업이 성공적으로 완료되었습니다.")
            
            if messagebox.askyesno("백업 완료", f"백업이 완료되었습니다!\n\n저장 위치:\n{backup_folder}\n\n폴더를 열어보시겠습니까?"):
                self.open_folder(backup_folder)

        except Exception as e:
            self.log(f"오류 발생: {e}")
            messagebox.showerror("백업 실패", f"백업 중 오류가 발생했습니다:\n{e}")
        finally:
            self.root.after(0, self.cleanup_ui)

    def cleanup_ui(self):
        self.backup_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.update_progress(0, "대기 중")

    def get_target_channels(self):
        selected_indices = self.channel_listbox.curselection()
        if selected_indices:
            self.log(f"선택된 {len(selected_indices)}개 채널을 백업합니다.")
            return [self.available_channels[i] for i in selected_indices]
        
        self.log("모든 채널을 대상으로 백업합니다 (선택된 채널 없음).")
        target_channels = []
        options = {
            'public': self.public_channels_var.get(),
            'private': self.private_channels_var.get(),
            'dm': self.direct_messages_var.get(),
            'mpim': self.private_channels_var.get(),
        }

        for ch_info in self.available_channels:
            ch = ch_info['original']
            ch_type = 'unknown'
            if ch.get('is_im'): ch_type = 'dm'
            elif ch.get('is_mpim'): ch_type = 'mpim'
            elif ch.get('is_group'): ch_type = 'private'
            elif ch.get('is_channel'): ch_type = 'public'

            if options.get(ch_type):
                target_channels.append(ch_info)
        return target_channels

    def get_channel_messages(self, headers, channel_id):
        params = {"channel": channel_id, "limit": 200}
        period_map = {"최근 1개월": 30, "최근 3개월": 90, "최근 6개월": 180, "최근 1년": 365}
        period_str = self.period_var.get()
        if period_str in period_map:
            days = period_map[period_str]
            oldest_ts = (datetime.now() - timedelta(days=days)).timestamp()
            params['oldest'] = str(oldest_ts)
        
        messages = self._fetch_all_pages("conversations.history", headers, "messages", params=params)
        self.log(f"  - 총 {len(messages)}개 메시지 수집 완료.")
        return messages

    def save_backup_files(self, messages_data, backup_folder, users_map, is_partial=False):
        status_prefix = "부분 " if is_partial else ""
        file_suffix = "_PARTIAL" if is_partial else ""

        json_file = backup_folder / f"slack_backup{file_suffix}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(messages_data, f, ensure_ascii=False, indent=2)
        self.log(f"{status_prefix}JSON 파일 저장 완료: {json_file}")
        
        if self.format_var.get() == "HTML":
            self.create_html_output(messages_data, backup_folder, users_map, is_partial)
    
    def open_folder(self, path):
        try:
            if platform.system() == 'Windows':
                os.startfile(path)
            elif platform.system() == 'Darwin':
                subprocess.run(['open', path], check=True)
            else:
                subprocess.run(['xdg-open', path], check=True)
        except Exception as e:
            self.log(f"폴더를 여는 데 실패했습니다: {e}")

    def create_html_output(self, all_messages, backup_folder, users, is_partial=False):
        """
        HTML 파일을 생성합니다. 가독성과 유지보수성을 위해 각 부분을 분리하여 생성 후 조립합니다.
        """
        file_suffix = "_PARTIAL" if is_partial else ""
        html_file_path = backup_folder / f"backup{file_suffix}.html"
        
        status_title = " (부분 백업)" if is_partial else ""
        status_warning = '<p style="color: orange; font-weight: bold;">⚠️ 백업이 중단되어 일부 채널만 포함되어 있을 수 있습니다.</p>' if is_partial else ''

        # --- HTML 각 부분을 명확하게 분리하여 생성 ---
        
        css_style = """
<style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; line-height: 1.6; color: #262626; margin: 0; background: #f8f8f8; }
    .container { display: flex; align-items: flex-start; }
    .nav { width: 260px; flex-shrink: 0; background: white; padding: 20px; border-right: 1px solid #e0e0e0; height: 100vh; overflow-y: auto; position: sticky; top: 0; }
    .content { flex-grow: 1; padding: 30px; max-width: 900px; }
    .header { background: #4a154b; color: white; padding: 25px; border-radius: 10px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 24px; }
    .channel { background: white; margin-bottom: 20px; border-radius: 8px; border: 1px solid #e8e8e8; overflow: hidden; display: none;}
    .channel-header { background: #f9f9f9; padding: 12px 20px; border-bottom: 1px solid #e8e8e8; font-weight: bold; font-size: 18px; color: #1d1c1d; }
    .messages { padding: 20px; }
    .message { display: flex; padding-bottom: 15px; margin-bottom: 15px; border-bottom: 1px solid #f0f0f0; }
    .message:last-child { border-bottom: none; }
    .message-avatar { width: 36px; height: 36px; border-radius: 4px; background: #ddd; margin-right: 12px; flex-shrink: 0; }
    .message-content { flex-grow: 1; }
    .message-header { display: flex; align-items: center; margin-bottom: 5px; }
    .username { font-weight: 900; color: #1d1c1d; margin-right: 8px; }
    .timestamp { color: #616061; font-size: 12px; }
    .message-text { color: #1d1c1d; white-space: pre-wrap; word-wrap: break-word; }
    .message-text a { color: #1264a3; text-decoration: underline; }
    .nav h3 { margin-top: 0; margin-bottom: 15px; color: #4a154b; font-size: 18px; }
    .nav ul { list-style: none; padding: 0; margin: 0; }
    .nav li a { color: #1264a3; text-decoration: none; font-size: 14px; display: block; padding: 8px 12px; border-radius: 5px; cursor: pointer; transition: background-color 0.2s; }
    .nav li a:hover, .nav li a.active { background-color: #e8e8e8; }
    @media (max-width: 900px) { 
        .container { flex-direction: column; } 
        .nav { position: static; width: 100%; height: auto; max-height: 30vh; margin-bottom: 20px; box-sizing: border-box; border-right: none; } 
    }
</style>
"""
        javascript_code = """
<script>
    function showChannel(targetId, element) {
        document.querySelectorAll('.channel').forEach(c => { c.style.display = 'none'; });
        document.getElementById(targetId).style.display = 'block';
        document.querySelectorAll('.nav a').forEach(a => { a.classList.remove('active'); });
        element.classList.add('active');
    }
    window.onload = () => {
        const firstChannelLink = document.querySelector('.nav li a');
        if (firstChannelLink) {
            firstChannelLink.click();
        }
    };
</script>
"""
        # --- HTML 본문 생성 시작 ---
        
        # 내비게이션 메뉴 생성
        nav_menu_items = []
        for channel_name in sorted(all_messages.keys()):
            safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', channel_name)
            display_name = html.escape(channel_name)
            nav_menu_items.append(f'<li><a onclick="showChannel(\'{safe_id}\', this)">{display_name}</a></li>')
        nav_menu_html = f'<div class="nav"><h3>채널 목록</h3><ul>{"".join(nav_menu_items)}</ul></div>'
        
        # 메인 콘텐츠 헤더 생성
        main_content_parts = [
            f'<div class="content">',
            f'<div class="header"><h1>Slack 대화 백업{status_title}</h1><p>백업 일시: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>{status_warning}</div>'
        ]

        # 각 채널의 메시지 콘텐츠 생성
        for channel_name, data in sorted(all_messages.items()):
            safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', channel_name)
            display_name = html.escape(channel_name)
            
            channel_html_parts = [f'<div class="channel" id="{safe_id}">', f'<div class="channel-header">{display_name}</div>', '<div class="messages">']
            
            for msg in sorted(data.get('messages', []), key=lambda x: float(x.get('ts', 0))):
                user_id = msg.get('user', 'Unknown')
                username = html.escape(users.get(user_id, user_id))
                timestamp = datetime.fromtimestamp(float(msg.get('ts', 0))).strftime("%Y-%m-%d %H:%M:%S")
                
                text_raw = msg.get('text', '')
                text_escaped = html.escape(text_raw)
                text_linked = re.sub(r'&lt;(https?://[^|]+?)\|(.+?)&gt;', r'<a href="\\1" target="_blank">\\2</a>', text_escaped)
                text_linked = re.sub(r'&lt;(https?://[^>]+?)&gt;', r'<a href="\\1" target="_blank">\\1</a>', text_linked)
                text_final = text_linked.replace('\n', '<br>')

                message_html = f"""
                <div class="message">
                    <div class="message-avatar"></div>
                    <div class="message-content">
                        <div class="message-header"><span class="username">{username}</span><span class="timestamp">{timestamp}</span></div>
                        <div class="message-text">{text_final}</div>
                    </div>
                </div>"""
                channel_html_parts.append(message_html)
                
            channel_html_parts.extend(['</div></div>'])
            main_content_parts.append("".join(channel_html_parts))

        main_content_parts.append('</div>')
        
        # --- 모든 HTML 조각들을 최종본으로 조립 ---
        final_html_parts = [
            '<!DOCTYPE html>',
            f'<html lang="ko">',
            '<head>',
            '    <meta charset="UTF-8">',
            '    <meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'    <title>Slack 백업{status_title}</title>',
            css_style,
            '</head>',
            '<body>',
            '    <div class="container">',
            nav_menu_html,
            "".join(main_content_parts),
            '    </div>',
            javascript_code,
            '</body>',
            '</html>'
        ]
        
        final_html = "\n".join(final_html_parts)

        # 파일에 쓰기
        try:
            with open(html_file_path, 'w', encoding='utf-8') as f:
                f.write(final_html)
            self.log(f"HTML 파일 생성 완료: {html_file_path}")
        except Exception as e:
            self.log(f"HTML 파일 저장 중 오류 발생: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SlackBackupTool(root)
    root.mainloop()
