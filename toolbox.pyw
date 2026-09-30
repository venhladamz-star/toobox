import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.scrolledtext as st
import threading
import time
import os
import sys
import re
import socket
import statistics
import subprocess
import urllib.request
import zipfile
import shutil
import urllib.parse
import ssl
import json
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass
import io
import tkinter.filedialog as fd
from collections import deque
class DummyStream(io.StringIO):
    def write(self, s):
        return None
    def flush(self):
        return None
    def fileno(self):
        return 1
if sys.stdout is None or not hasattr(sys.stdout, 'fileno'):
    sys.stdout = DummyStream()
if sys.stderr is None or not hasattr(sys.stderr, 'fileno'):
    sys.stderr = DummyStream()
try:
    import speedtest
    HAVE_SPEEDTEST = True
except ImportError:
    HAVE_SPEEDTEST = False
try:
    import win32api
    HAVE_WIN32 = True
except ImportError:
    HAVE_WIN32 = False
BG = '#111822'
PANEL = '#172233'
CARD = '#1e2c42'
INPUT = '#26344d'
BORDER = '#364a66'
ACCENT = '#a78bfa'
ACCENT2 = '#34d399'
ACCENT3 = '#22d3ee'
ACCENT4 = '#fbbf24'
DANGER = '#f87171'
WARN = '#fbbf24'
TXT = '#f0f4f8'
MUTED = '#a6b3cc'
class SafeStdoutRedirector:
    def __init__(self, app_instance, text_widget):
        self.app = app_instance
        self.widget = text_widget
    def write(self, string):
        if not string:
            return None
        else:
            self.app.root.after(0, self._safe_write, string)
    def _safe_write(self, string):
        try:
            cleaned = string.replace('\r', '')
            self._insert_ansi(cleaned)
        except Exception:
            return None
    def _insert_ansi(self, text):
        ansi_pattern = re.compile('\033\\[([0-9;]*)m')
        parts = ansi_pattern.split(text)
        current_tag = 'n'
        self.widget.configure(state='normal')
        for idx, part in enumerate(parts):
            if idx % 2 == 0:
                if part:
                    self.widget.insert('end', part, current_tag)
                continue
            else:
                code = part
                if code == '0' or not code:
                    current_tag = 'n'
                    continue
                else:
                    if code.startswith('38;2;'):
                        rgb_parts = code.split(';')
                        if len(rgb_parts) >= 5:
                            try:
                                r, g, b = (int(rgb_parts[2]), int(rgb_parts[3]), int(rgb_parts[4]))
                                hex_color = f'#{r:02x}{g:02x}{b:02x}'
                                tag_name = f'ansi_{r}_{g}_{b}'
                                self.widget.tag_config(tag_name, foreground=hex_color)
                                current_tag = tag_name
                            except Exception as e:
                                pass
                    else:
                        if '1' in code.split(';'):
                            continue
                        else:
                            codes = code.split(';')
                            for c in codes:
                                if c == '31' or c == '91':
                                    current_tag = 'danger'
                                else:
                                    if c == '32' or c == '92':
                                        current_tag = 'ok'
                                    else:
                                        if c == '33' or c == '93':
                                            current_tag = 'warn'
                                        else:
                                            if c == '36' or c == '96':
                                                current_tag = 'cyber'
                                            else:
                                                if c == '35' or c == '95':
                                                    current_tag = 'purple'
                                                else:
                                                    if c == '90':
                                                        current_tag = 'muted'
        self.widget.see('end')
        self.widget.configure(state='disabled')
    def flush(self):
        return None
class App:
    # Nguồn bộ cài Office chính thức (Microsoft CDN). Thêm key mới vào dict này
    # là tab Kích Hoạt tự có thêm nút tải — không phải đụng tới phần giao diện.
    OFFICE_SOURCES = {
        'proplus': {
            'label': '  TẢI OFFICE 2024 PROPLUS   📥',
            'desc': 'Bộ cài online Office 2024 ProPlus (.exe), chính thức từ Microsoft',
            'name': 'Office 2024 ProPlus',
            'file': 'OfficeSetup2024.exe',
            'url': ('https://c2rsetup.officeapps.live.com/c2r/download.aspx'
                    '?ProductreleaseID=ProPlus2024Retail'
                    '&platform=x64&language=en-us&version=O16GA'),
        },
        'hs2016': {
            'label': '  TẢI OFFICE 2016 HOME & STUDENT   📥',
            'desc': 'Office 2016 Home & Student — ảnh đĩa .img từ Microsoft (Win 10/11 mở bằng nhấp đúp)',
            'name': 'Office 2016 Home & Student',
            'file': 'Office2016HomeStudent.img',
            'url': ('https://officecdn.microsoft.com/db/492350f6-3a01-4f97-'
                    'b9c0-c7c6ddf67d60/media/en-us/HomeStudentRetail.img'),
        },
        'home2024': {
            'label': '  TẢI OFFICE 2024 HOME   📥',
            'desc': 'Office 2024 Home — ảnh đĩa .img từ Microsoft (Win 10/11 mở bằng nhấp đúp)',
            'name': 'Office 2024 Home',
            'file': 'Office2024Home.img',
            'url': ('https://officecdn.microsoft.com/db/492350f6-3a01-4f97-'
                    'b9c0-c7c6ddf67d60/media/en-us/Home2024Retail.img'),
        },
    }

    def __init__(self, root):
        self.root = root
        self.root.title('Toolbox')
        self.root.configure(bg=BG)
        self.root.geometry('380x440')
        self.root.minsize(360, 380)
        self.root.resizable(True, True)
        self.root.wm_attributes('-topmost', False)
        def _on_mousewheel(event):
            widget = event.widget.winfo_containing(event.x_root, event.y_root)
            while widget:
                try:
                    widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
                    break
                except (AttributeError, tk.TclError):
                    widget = getattr(widget, 'master', None)
                    if not isinstance(widget, tk.Widget):
                        break
        self.root.bind_all('<MouseWheel>', _on_mousewheel)
        self.v_mouse = tk.StringVar(value='X:- Y:-')
        self.v_dl_url = tk.StringVar()
        self.v_dl_mode = tk.StringVar(value='1')
        self.v_dl_fmt = tk.StringVar(value='MP3 - 320 kbps (Tốt nhất)')
        self.v_dl_quality = tk.StringVar(value='Tự động (Cao nhất)')
        self.is_downloading = False
        self.is_scanning = False
        self._build()
        self.root.protocol('WM_DELETE_WINDOW', self._safe_close)
        self._track_mouse()
        self._start_silent_downloads()
        self._cleanup_temp()
    def _ask_on_main(self, kind, title, msg, parent=None):
        """Gọi messagebox một cách an toàn từ luồng nền.
        kind: 'yesno' -> bool, 'info' | 'error' -> None"""
        if threading.current_thread() is threading.main_thread():
            if kind == 'yesno':
                return messagebox.askyesno(title, msg, parent=parent or self.root)
            if kind == 'error':
                messagebox.showerror(title, msg, parent=parent or self.root)
            else:
                messagebox.showinfo(title, msg, parent=parent or self.root)
            return None

        result = {'v': False}
        done = threading.Event()

        def _run():
            try:
                if kind == 'yesno':
                    result['v'] = bool(messagebox.askyesno(title, msg, parent=parent or self.root))
                elif kind == 'error':
                    messagebox.showerror(title, msg, parent=parent or self.root)
                else:
                    messagebox.showinfo(title, msg, parent=parent or self.root)
            except Exception:
                result['v'] = False
            finally:
                done.set()

        try:
            if not self._post(_run):
                return False
        except Exception:
            return False
        # Chờ trả lời trong khi cửa sổ app còn sống — không timeout cứng:
        # dialog mở lâu từng bị bỏ qua âm thầm (user bấm Yes sau đó bị ignore),
        # còn app chết thì phải thoát để không kẹt thread.
        while not done.wait(timeout=120):
            try:
                if not self.root.winfo_exists():
                    return False
            except Exception:
                return False
        return result['v']

    def _safe_close(self):
        """Chặn đóng app khi tab 'Sau cài Win' đang chạy dở — đóng giữa chừng
        có thể để services Windows kẹt STOP / registry ghi dở dang."""
        if getattr(self, '_post_running', False):
            messagebox.showwarning(
                'Đang chạy dở',
                'Có thao tác "Sau cài Win" đang chạy.\n'
                'Hãy chờ hoàn tất rồi mới đóng app — đóng giữa chừng có thể '
                'để services Windows ở trạng thái dở dang.',
                parent=self.root)
            return
        self.root.destroy()

    def _build(self):
        self._top_bar()
        self._divider()
        self._tab_bar()
        # Lưu divider này để dùng làm anchor khi switch tab
        self.div_main = tk.Frame(self.root, bg=BORDER, height=1)
        self.div_main.pack(fill='x')
        # ── Pack bottom section TRƯỚC (side='bottom' phải pack trước expand=True) ──
        self._footer()   # side='bottom' → luôn hiện ở dưới cùng
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill='x', side='bottom')

        # ── Tab content (expand để lấp đầy phần còn lại) ──
        self._build_dl_tab()
        self._build_app_tab()
        self._build_net_tab()
        self._build_act_tab()
        self._build_post_tab()
        # Tab mặc định: Tải Nhạc/Video
        self._switch_tab(1)
    def _divider(self):
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill='x')
    def _top_bar(self):
        bar = tk.Frame(self.root, bg=PANEL, height=30)
        bar.pack(fill='x')
        bar.pack_propagate(False)
        tk.Label(bar, text='Toolbox', bg=PANEL, fg=TXT, font=('Segoe UI', 10, 'bold')).pack(side='left', padx=8)
        self.lbl_status = tk.Label(bar, text='Chờ...', bg=PANEL, fg=MUTED, font=('Segoe UI', 8))
        self.lbl_status.pack(side='left', padx=2)
    def _tab_bar(self):
        bar = tk.Frame(self.root, bg=PANEL, height=28)
        bar.pack(fill='x')
        bar.pack_propagate(False)

        self.btn_tab_dl = tk.Button(bar, text='Tải Nhạc/Vid', command=lambda: self._switch_tab(1), bg=INPUT, fg=MUTED, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2')
        self.btn_tab_dl.pack(side='left', fill='both', expand=True)
        self.btn_tab_app = tk.Button(bar, text='🧩 Tải App', command=lambda: self._switch_tab(2), bg=INPUT, fg=MUTED, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2')
        self.btn_tab_app.pack(side='left', fill='both', expand=True)
        self.btn_tab_net = tk.Button(bar, text='📶 WiFi', command=lambda: self._switch_tab(3), bg=INPUT, fg=MUTED, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2')
        self.btn_tab_net.pack(side='left', fill='both', expand=True)
        self.btn_tab_act = tk.Button(bar, text='Local', command=lambda: self._switch_tab(4), bg=INPUT, fg=MUTED, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2')
        self.btn_tab_act.pack(side='left', fill='both', expand=True)
        self.btn_tab_post = tk.Button(bar, text='🛠 Sau cài Win', command=lambda: self._switch_tab(5), bg=INPUT, fg=MUTED, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2')
        self.btn_tab_post.pack(side='left', fill='both', expand=True)
    def _switch_tab(self, tab_idx):
        self.tab_dl_frame.pack_forget()
        self.tab_app_frame.pack_forget()
        self.tab_net_frame.pack_forget()
        self.tab_act_frame.pack_forget()
        self.tab_post_frame.pack_forget()
        self.btn_tab_dl.config(bg=INPUT, fg=MUTED)
        self.btn_tab_app.config(bg=INPUT, fg=MUTED)
        self.btn_tab_net.config(bg=INPUT, fg=MUTED)
        self.btn_tab_act.config(bg=INPUT, fg=MUTED)
        self.btn_tab_post.config(bg=INPUT, fg=MUTED)
        if tab_idx == 1:
            self.root.geometry('460x520')
            self.btn_tab_dl.config(bg=ACCENT3, fg=BG)
            self.tab_dl_frame.pack(fill='both', expand=True, after=self.div_main)
        elif tab_idx == 2:
            self.root.geometry('520x620')
            self.btn_tab_app.config(bg=ACCENT2, fg=BG)
            self.tab_app_frame.pack(fill='both', expand=True, after=self.div_main)
        elif tab_idx == 3:
            self.root.geometry('540x580')
            self.btn_tab_net.config(bg=ACCENT4, fg=BG)
            self.tab_net_frame.pack(fill='both', expand=True, after=self.div_main)
        elif tab_idx == 4:
            self.root.geometry('520x600')
            self.btn_tab_act.config(bg=ACCENT, fg='white')
            self.tab_act_frame.pack(fill='both', expand=True, after=self.div_main)
        elif tab_idx == 5:
            self.root.geometry('520x620')
            self.btn_tab_post.config(bg=DANGER, fg='white')
            self.tab_post_frame.pack(fill='both', expand=True, after=self.div_main)

    # ── Tab "Tải App": 16 app, tick rồi bấm CÀI ĐẶT ──────────────────────
    def _build_app_tab(self):
        self.tab_app_frame = tk.Frame(self.root, bg=BG)
        self._app_batch_running = False
        # Hàng thao tác trên cùng: nút Cài đặt + bộ đếm tick
        top = tk.Frame(self.tab_app_frame, bg=BG)
        top.pack(fill='x', padx=10, pady=(10, 2))
        self.btn_app_install = tk.Button(
            top, text='  ⚙ CẢI ĐẶT  ▶', command=self._start_app_batch,
            bg=ACCENT2, fg=BG, font=('Segoe UI', 9, 'bold'),
            relief='flat', cursor='hand2')
        self.btn_app_install.pack(side='left', fill='x', expand=True, ipady=4)
        self.lbl_app_picked = tk.Label(top, text='Đã tick: 0/16', bg=BG,
                                       fg=MUTED, font=('Segoe UI', 8))
        self.lbl_app_picked.pack(side='right', padx=(8, 0))
        # Tiến độ tổng: (số app đã xong + % file hiện tại) / tổng app
        self.app_canvas = tk.Canvas(self.tab_app_frame, bg='#1a2638', height=8,
                                    highlightthickness=0)
        self.app_canvas.pack(fill='x', padx=10, pady=(4, 0))
        self.lbl_app_pct = tk.Label(self.tab_app_frame, text='0%', bg=BG,
                                    fg=ACCENT3, font=('Segoe UI', 8, 'bold'))
        self.lbl_app_pct.pack(anchor='e', padx=10)
        # Danh sách 2 cột × 4 nhóm — cột dài nhất 9 dòng nên vừa khung, không cuộn
        cols = tk.Frame(self.tab_app_frame, bg=BG)
        cols.pack(fill='x', padx=8, pady=(2, 4))
        self.app_vars = {}
        for col_idx, groups in enumerate((('browser', 'chat'), ('input', 'tools'))):
            col = tk.Frame(cols, bg=BG)
            col.pack(side='left', fill='both', expand=True,
                     padx=(0, 6) if col_idx == 0 else (6, 0))
            for g in groups:
                tk.Label(col, text=GROUP_LABELS[g], bg=PANEL, fg=TXT,
                         font=('Segoe UI', 8, 'bold'), anchor='w'
                         ).pack(fill='x', pady=(4, 1))
                for ap in [x for x in APP_SOURCES if x['group'] == g]:
                    var = tk.BooleanVar(value=False)
                    tk.Checkbutton(
                        col, text=ap['name'], variable=var,
                        command=self._update_app_picked_label,
                        bg=BG, fg=TXT, selectcolor=INPUT,
                        activebackground=BG, activeforeground=TXT,
                        font=('Segoe UI', 8), anchor='w', justify='left',
                        cursor='hand2', highlightthickness=0
                    ).pack(fill='x', anchor='w', ipady=1)
                    self.app_vars[ap['key']] = var
        # Console log — luồng tải ghi từ thread nên mọi cập nhật đi qua _post()
        self.app_log_txt = st.ScrolledText(self.tab_app_frame, bg='#0f1522',
                                           fg=MUTED, font=('Courier New', 8),
                                           relief='flat', wrap='word', height=7)
        for t, c in [('n', MUTED), ('ok', ACCENT2), ('warn', WARN),
                     ('danger', DANGER), ('cyber', ACCENT3)]:
            self.app_log_txt.tag_config(t, foreground=c)
        self.app_log_txt.pack(fill='x', padx=10, pady=(2, 8))
        self._log_app_console(' Tick app cần cài rồi bấm "CẢI ĐẶT" — '
                              'tải hết trước, rồi cài lần lượt từng app.', 'cyber')

    def _update_app_picked_label(self):
        n = sum(1 for v in self.app_vars.values() if v.get())
        self.lbl_app_picked.config(text='Đã tick: %d/%d' % (n, len(APP_SOURCES)))

    def _log_app_console(self, text, tag='n'):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._log_app_console, text, tag)
            return
        self.app_log_txt.configure(state='normal')
        self.app_log_txt.insert('end', text + '\n', tag)
        self.app_log_txt.see('end')
        self.app_log_txt.configure(state='disabled')

    def _clear_app_console(self):
        self.app_log_txt.configure(state='normal')
        self.app_log_txt.delete('1.0', 'end')
        self.app_log_txt.configure(state='disabled')

    def _update_app_progress(self, pct):
        try:
            self.app_canvas.delete('bar')
            w = self.app_canvas.winfo_width()
            h = self.app_canvas.winfo_height()
            self.app_canvas.create_rectangle(0, 0, int(pct / 100.0 * w), h,
                                             fill=ACCENT3, outline='', tags='bar')
            self.lbl_app_pct.config(text='%.0f%%' % pct)
        except Exception:
            pass

    def _start_app_batch(self):
        """Bấm CẢI ĐẶT: tải hết các app đã tick trước, rồi cài lần lượt."""
        if self._app_batch_running:
            return
        picked = [a for a in APP_SOURCES if self.app_vars[a['key']].get()]
        if not picked:
            self._log_app_console('⚠️ Chưa tick app nào — hãy chọn ít nhất 1 app.', 'warn')
            return
        self._app_batch_running = True
        self.btn_app_install.config(state='disabled', bg=BORDER, fg=MUTED,
                                    text=' ĐANG TẢI… ')
        self._clear_app_console()
        self._update_app_progress(0)
        self._log_app_console('🚀 Đã chọn %d app — tải hết trước, '
                              'rồi cài lần lượt từng app.' % len(picked), 'cyber')
        threading.Thread(target=self._run_app_batch, args=(picked,),
                         daemon=True).start()

    def _run_app_batch(self, picked):
        # Phase A — tải tuần tự (lỗi 1 app được log + bỏ qua, không dừng loạt)
        dest = apps_download_dir()
        try:
            res = download_apps(
                picked, dest,
                progress_cb=lambda pct, msg='': self._post(self._update_app_progress, pct),
                log_cb=self._log_app_console)
        except Exception as e:
            self._log_app_console('❌ Lỗi hệ thống khi tải: %s' % e, 'danger')
            self._post(self._finish_app_batch)
            return
        # Phase B — cài lần lượt: app kế chỉ mở khi người dùng đóng wizard app trước
        for key in res['ok']:
            app = next(a for a in APP_SOURCES if a['key'] == key)
            path = res['paths'][key]
            if app.get('zip'):
                # Bản portable (EVKey): giải nén rồi mở trực tiếp — không có wizard
                try:
                    xdir = os.path.join(dest, app['key'])
                    with zipfile.ZipFile(path) as zf:
                        zf.extractall(xdir)
                    target = os.path.join(xdir, app['run'].replace('/', os.sep))
                    self._log_app_console('📦 %s là bản portable — giải nén và '
                                          'mở trực tiếp (không có wizard).'
                                          % app['name'], 'n')
                    os.startfile(target)
                    self._log_app_console('✓ Đã mở %s.' % app['name'], 'ok')
                except Exception as e:
                    self._log_app_console('❌ %s — không mở được bản portable: %s'
                                          % (app['name'], e), 'danger')
                continue
            self._log_app_console('▶ Mở cài đặt %s — bấm Next trong cửa sổ vừa hiện.'
                                  % app['name'], 'cyber')
            try:
                # List-form để path có khoảng trắng ("Toolbox Apps") không bị tách sai
                cmd = ['msiexec', '/i', path] if path.lower().endswith('.msi') else [path]
                proc = subprocess.Popen(cmd)
                proc.wait()
                self._log_app_console('✓ Đã đóng trình cài %s — sang app kế.'
                                      % app['name'], 'ok')
            except Exception as e:
                self._log_app_console('❌ %s — không mở được bản cài: %s'
                                      % (app['name'], e), 'danger')
        # Tổng kết
        names = {a['key']: a['name'] for a in APP_SOURCES}
        lines = ['🏁 Tải xong %d/%d app.' % (len(res['ok']), len(picked))]
        for key, reason in res['failed']:
            lines.append('   ❌ %s — %s' % (names.get(key, key), reason))
        lines.append('📁 File cài nằm ở: %s' % dest)
        lines.append('▶ Cài đặt: bấm Next từng wizard — app kế tự mở sau khi '
                     'bạn đóng app trước.')
        self._log_app_console('\n'.join(lines),
                              'ok' if not res['failed'] else 'warn')
        self._post(self._finish_app_batch)

    def _finish_app_batch(self):
        self._app_batch_running = False
        self.btn_app_install.config(state='normal', bg=ACCENT2, fg=BG,
                                    text='  ⚙ CẢI ĐẶT  ▶')

    # ── Tab "🛠 Sau cài Win": 7 tweak + quét ổ EFI + dọn Update Cache ───────
    def _build_post_tab(self):
        self.tab_post_frame = tk.Frame(self.root, bg=BG)
        self._post_running = False
        # Hàng nút: Áp dụng batch (chính) + 2 nút chạy riêng
        top = tk.Frame(self.tab_post_frame, bg=BG)
        top.pack(fill='x', padx=10, pady=(10, 2))
        self.btn_post_apply = tk.Button(
            top, text='  ⚙ ÁP DỤT ĐÃ TICK  ▶', command=self._start_post_batch,
            bg=ACCENT, fg='white', font=('Segoe UI', 9, 'bold'),
            relief='flat', cursor='hand2')
        self.btn_post_apply.pack(side='left', fill='x', expand=True, ipady=4)
        self.btn_post_efi = tk.Button(
            top, text='🔍 Quét ổ EFI', command=self._start_efi_scan,
            bg=INPUT, fg=ACCENT3, font=('Segoe UI', 8, 'bold'),
            relief='flat', cursor='hand2')
        self.btn_post_efi.pack(side='left', fill='y', padx=(6, 0), ipady=4)
        self.btn_post_wu = tk.Button(
            top, text='🧹 Dọn Update Cache', command=self._start_wu_clean,
            bg=INPUT, fg=WARN, font=('Segoe UI', 8, 'bold'),
            relief='flat', cursor='hand2')
        self.btn_post_wu.pack(side='left', fill='y', padx=(6, 0), ipady=4)
        # Tiến độ batch
        self.post_canvas = tk.Canvas(self.tab_post_frame, bg='#1a2638', height=8,
                                     highlightthickness=0)
        self.post_canvas.pack(fill='x', padx=10, pady=(6, 0))
        self.lbl_post_pct = tk.Label(self.tab_post_frame, text='0%', bg=BG,
                                     fg=ACCENT3, font=('Segoe UI', 8, 'bold'))
        self.lbl_post_pct.pack(anchor='e', padx=10)
        # 7 checkbox — default theo POST_TWEAKS (5 bật sẵn)
        cb = tk.Frame(self.tab_post_frame, bg=BG)
        cb.pack(fill='x', padx=10, pady=(2, 4))
        self.post_vars = {}
        for t in POST_TWEAKS:
            var = tk.BooleanVar(value=t['default'])
            tk.Checkbutton(
                cb, text=t['name'], variable=var, bg=BG, fg=TXT, selectcolor=INPUT,
                activebackground=BG, activeforeground=TXT, font=('Segoe UI', 8),
                anchor='w', justify='left', cursor='hand2', highlightthickness=0
            ).pack(fill='x', anchor='w', ipady=1)
            self.post_vars[t['id']] = var
        # Console log — luồng chạy nền ghi từ thread nên mọi cập nhật qua _post()
        self.post_log_txt = st.ScrolledText(self.tab_post_frame, bg='#0f1522',
                                            fg=MUTED, font=('Courier New', 8),
                                            relief='flat', wrap='word', height=8)
        for t, c in [('n', MUTED), ('ok', ACCENT2), ('warn', WARN),
                     ('danger', DANGER), ('cyber', ACCENT3)]:
            self.post_log_txt.tag_config(t, foreground=c)
        self.post_log_txt.pack(fill='x', padx=10, pady=(2, 6))
        tk.Label(self.tab_post_frame,
                 text='⚠ Không hoàn tác được — kiểm tra kỹ trước khi chạy.',
                 bg=BG, fg=WARN, font=('Segoe UI', 7)
                 ).pack(anchor='w', padx=10, pady=(0, 8))
        self._log_post_console(' Tick tweak rồi bấm "ÁP DỤT ĐÃ TICK" — hoặc '
                               'dùng 2 nút chạy riêng bên cạnh.', 'cyber')

    def _log_post_console(self, text, tag='n'):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._log_post_console, text, tag)
            return
        self.post_log_txt.configure(state='normal')
        self.post_log_txt.insert('end', text + '\n', tag)
        self.post_log_txt.see('end')
        self.post_log_txt.configure(state='disabled')

    def _clear_post_console(self):
        self.post_log_txt.configure(state='normal')
        self.post_log_txt.delete('1.0', 'end')
        self.post_log_txt.configure(state='disabled')

    def _update_post_progress(self, pct):
        try:
            self.post_canvas.delete('bar')
            w = self.post_canvas.winfo_width()
            h = self.post_canvas.winfo_height()
            self.post_canvas.create_rectangle(0, 0, int(pct / 100.0 * w), h,
                                              fill=ACCENT3, outline='', tags='bar')
            self.lbl_post_pct.config(text='%.0f%%' % pct)
        except Exception:
            pass

    # ——— Flow chạy (T5) ———
    def _post_buttons(self, state):
        if state == 'disabled':
            for b in (self.btn_post_apply, self.btn_post_efi, self.btn_post_wu):
                b.config(state='disabled', bg=BORDER, fg=MUTED)
        else:
            self.btn_post_apply.config(state='normal', bg=ACCENT, fg='white')
            self.btn_post_efi.config(state='normal', bg=INPUT, fg=ACCENT3)
            self.btn_post_wu.config(state='normal', bg=INPUT, fg=WARN)

    def _start_post_batch(self):
        """Bấm ÁP DỤT: xác nhận 'không hoàn tác được' rồi chạy batch trong thread."""
        if self._post_running:
            return
        picked = [t for t in POST_TWEAKS if self.post_vars[t['id']].get()]
        if not picked:
            self._log_post_console('⚠️ Chưa tick tweak nào — hãy chọn ít nhất 1.',
                                   'warn')
            return
        if not self._ask_on_main(
                'yesno', 'Xác nhận áp dụng',
                'Áp dụng %d tweak?\n\n⚠️ KHÔNG HOÀN TÁC ĐƯỢC — kiểm tra kỹ '
                'trước khi chạy.' % len(picked)):
            self._log_post_console('⛔ Đã hủy — chưa áp dụng gì.', 'n')
            return
        self._post_running = True
        self._post_buttons('disabled')
        self._clear_post_console()
        self._update_post_progress(0)
        self._log_post_console('🚀 Áp dụng %d tweak theo thứ tự…' % len(picked),
                               'cyber')
        threading.Thread(target=self._run_post_batch, args=(picked,),
                         daemon=True).start()

    def _run_post_batch(self, picked):
        # finally: DÙ có exception không lường trước cũng phải trả nút —
        # nếu để thread chết thì tab kẹt disabled cho đến khi mở lại app (B1)
        try:
            n = len(picked)
            failed = []
            for i, t in enumerate(picked, 1):
                self._log_post_console('[%d/%d] %s' % (i, n, t['name']), 'cyber')
                if t['id'] == 'efi':
                    ok = self._do_efi_hide()
                else:
                    ok = run_tweak(t['id'], log=self._log_post_console)
                if ok:
                    self._log_post_console('   → OK', 'ok')
                else:
                    failed.append(t['name'])
                    self._log_post_console('   → CÓ LỖI — xem log bên trên.',
                                           'warn')
                self._post(self._update_post_progress, i * 100.0 / n)
            self._post(self._update_post_progress, 100)
            if failed:
                lines = ['🏁 Xong %d/%d tweak — %d lỗi:'
                         % (n - len(failed), n, len(failed))]
                lines += ['   ❌ %s' % name for name in failed]
                self._log_post_console('\n'.join(lines), 'warn')
            else:
                self._log_post_console('🏁 Hoàn tất %d/%d tweak — không có lỗi.'
                                       % (n, n), 'ok')
        except Exception as e:
            self._log_post_console('❌ Lỗi không lường trước khi chạy batch: %s'
                                   % e, 'danger')
        finally:
            self._post(self._finish_post_batch)

    def _finish_post_batch(self):
        self._post_running = False
        self._post_buttons('normal')
        self._update_post_progress(100)   # efi/wu không set progress — reset ở đây

    def _do_efi_hide(self):
        """Quét → liệt kê console → xác nhận 1 lần → ẩn tất cả (mountvol /D)."""
        self._log_post_console('   🔍 Đang quét phân vùng EFI/Recovery…', 'n')
        scan_ok, items = list_efi_partitions()
        if not scan_ok:
            # Lỗi quét KHÁC rỗng thật (plan §6): ❌ + không hiện confirm
            self._log_post_console('   ❌ Không quét được phân vùng '
                                   '(PowerShell lỗi) — bỏ qua bước này.',
                                   'danger')
            return False
        if not items:
            self._log_post_console('   ✓ Không có phân vùng EFI/Recovery nào '
                                   'đang gắn ổ letter.', 'ok')
            return True
        for i, it in enumerate(items, 1):
            self._log_post_console('   [%d] Ổ %s: (%s — %s, %s)'
                                   % (i, it['letter'], it['type'],
                                      it['disk'], it['part']), 'n')
        msg = ('Tìm thấy %d ổ:\n%s\n\nẨn tất cả?\n'
               '(chỉ gỡ ổ letter trên Explorer — KHÔNG xóa dữ liệu)'
               % (len(items),
                  '\n'.join('  • Ổ %s: (%s)' % (i['letter'], i['type'])
                            for i in items)))
        if not self._ask_on_main('yesno', 'Ẩn phân vùng EFI/Recovery', msg):
            self._log_post_console('   ℹ Bạn chọn bỏ qua — không ẩn ổ nào.', 'n')
            return True
        n_ok, n_fail = hide_partitions(items, log=self._log_post_console)
        self._log_post_console('   → Ẩn %d/%d ổ.'
                               % (n_ok, len(items)), 'ok' if n_fail == 0 else 'warn')
        return n_fail == 0

    def _start_efi_scan(self):
        if self._post_running:
            return
        self._post_running = True
        self._post_buttons('disabled')
        self._clear_post_console()
        self._update_post_progress(0)
        self._log_post_console('🔍 Quét phân vùng EFI/Recovery…', 'cyber')
        threading.Thread(target=self._run_efi_scan, daemon=True).start()

    def _run_efi_scan(self):
        # finally như _run_post_batch: không được để thread chết làm kẹt tab (B1)
        try:
            ok = self._do_efi_hide()
            self._log_post_console('🏁 Quét/ẩn ổ EFI xong — %s.'
                                   % ('không có lỗi' if ok else 'có lỗi, xem log'),
                                   'ok' if ok else 'warn')
        except Exception as e:
            self._log_post_console('❌ Lỗi không lường trước khi quét ẩn ổ: %s'
                                   % e, 'danger')
        finally:
            self._post(self._finish_post_batch)

    def _start_wu_clean(self):
        if self._post_running:
            return
        if not self._ask_on_main(
                'yesno', 'Dọn Update Cache',
                'Dọn cache Windows Update?\n\n'
                '(dừng dịch vụ → xóa/đổi tên cache → chạy lại dịch vụ)\n'
                'Có thể mất vài chục giây.'):
            self._log_post_console('⛔ Đã hủy — chưa dọn gì.', 'n')
            return
        self._post_running = True
        self._post_buttons('disabled')
        self._clear_post_console()
        self._update_post_progress(0)
        self._log_post_console('🧹 Dọn cache Windows Update…', 'cyber')
        threading.Thread(target=self._run_wu_clean, daemon=True).start()

    def _run_wu_clean(self):
        # finally như trên: service đã dọc chừng cũng phải trả nút (B1)
        try:
            ok, msg = clean_update_cache(log=self._log_post_console)
            self._log_post_console('🏁 %s' % msg, 'ok' if ok else 'danger')
        except Exception as e:
            self._log_post_console('❌ Lỗi hệ thống khi dọn Update Cache: %s' % e,
                                   'danger')
        finally:
            self._post(self._finish_post_batch)

    def _build_dl_tab(self):
        self.tab_dl_frame = tk.Frame(self.root, bg=BG)
        er = tk.Frame(self.tab_dl_frame, bg=BG)
        er.pack(fill='x', padx=10, pady=(10, 4))
        tk.Label(er, text='Link (tự phân tích):', bg=BG, fg=TXT, font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.dl_entry = tk.Entry(er, textvariable=self.v_dl_url, bg=INPUT, fg=TXT, font=('Segoe UI', 9), relief='flat', bd=0, insertbackground=TXT)
        self.dl_entry.pack(side='left', fill='x', expand=True, padx=6, ipady=3)
        def paste_link():
            try:
                self.v_dl_url.set(self.root.clipboard_get().strip())
            except Exception as e:
                return None
        tk.Button(er, text='Dán', command=paste_link, bg=INPUT, fg=ACCENT3, font=('Segoe UI', 8, 'bold'), relief='flat', cursor='hand2', padx=6).pack(side='right')
        # Không còn nút 🔍 Phân tích: dán/nhập xong link là tool tự phân tích.
        # trace_add ở dưới debounce 700ms để không chạy theo từng phím, và URL đã
        # phân tích thành công thì không chạy lại.
        self._dl_format_map = {}
        self._dl_probe_job = None
        self._dl_probed_url = ''
        self.v_dl_url.trace_add('write', self._on_dl_url_changed)
        opt_f = tk.Frame(self.tab_dl_frame, bg=PANEL, padx=8, pady=6)
        opt_f.pack(fill='x', padx=10, pady=4)
        mr = tk.Frame(opt_f, bg=PANEL)
        mr.pack(fill='x', pady=2)
        tk.Label(mr, text='Định dạng tải:', bg=PANEL, fg=MUTED, font=('Segoe UI', 9)).pack(side='left')
        def on_mode_change():
            if self.v_dl_mode.get() == '1':
                self.cb_fmt.pack(side='left', padx=10)
                self.cb_quality.pack_forget()
            else:
                self.cb_fmt.pack_forget()
                self.cb_quality.pack(side='left', padx=10)
        tk.Radiobutton(mr, text='Tải Nhạc', variable=self.v_dl_mode, value='1', command=on_mode_change, bg=PANEL, fg=TXT, selectcolor=INPUT, activebackground=PANEL, activeforeground=TXT, font=('Segoe UI', 8, 'bold')).pack(side='left', padx=(10, 4))
        tk.Radiobutton(mr, text='Tải Video', variable=self.v_dl_mode, value='2', command=on_mode_change, bg=PANEL, fg=TXT, selectcolor=INPUT, activebackground=PANEL, activeforeground=TXT, font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4)
        cbr = tk.Frame(opt_f, bg=PANEL)
        cbr.pack(fill='x', pady=2)
        tk.Label(cbr, text='Cấu hình tùy chọn:', bg=PANEL, fg=MUTED, font=('Segoe UI', 9)).pack(side='left')
        self.cb_fmt = ttk.Combobox(cbr, textvariable=self.v_dl_fmt, state='readonly', width=22)
        self.cb_fmt['values'] = ('MP3 - 320 kbps (Tốt nhất)', 'MP3 - 192 kbps (Khuyên dùng)', 'FLAC (Chất lượng gốc)', 'WAV (Mặc định Microsoft)')
        self.cb_fmt.pack(side='left', padx=10)
        self.cb_quality = ttk.Combobox(cbr, textvariable=self.v_dl_quality, state='readonly', width=22)
        self.cb_quality['values'] = ('Tự động (Cao nhất)', '4K (Ultra HD)', '2K (Quad HD)', '1080p (Full HD)', '720p (HD)', '480p (SD)')
        ap = tk.Frame(self.tab_dl_frame, bg=BG)
        ap.pack(fill='x', padx=10, pady=4)
        self.btn_download = tk.Button(ap, text='Bắt đầu tải  📥', command=self._start_download_thread, bg=ACCENT3, fg=BG, font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2')
        self.btn_download.pack(side='left', fill='x', expand=True, ipady=3)
        self.lbl_dl_percent = tk.Label(ap, text='0.0%', bg=BG, fg=ACCENT3, font=('Segoe UI', 8, 'bold'))
        self.lbl_dl_percent.pack(side='right', padx=(8, 0))
        self.dl_canvas = tk.Canvas(self.tab_dl_frame, bg='#1a2638', height=8, highlightthickness=0)
        self.dl_canvas.pack(fill='x', padx=10, pady=4)
        self.dl_log_txt = st.ScrolledText(self.tab_dl_frame, bg='#0f1522', fg=MUTED, font=('Courier New', 8), relief='flat', wrap='word', height=6)
        for t, c in [('n', MUTED), ('ok', ACCENT2), ('warn', WARN), ('danger', DANGER), ('cyber', ACCENT3), ('purple', ACCENT)]:
            self.dl_log_txt.tag_config(t, foreground=c)
        self.dl_log_txt.tag_config('bold', font=('Courier New', 8, 'bold'))
        self._log_to_dl_console('⚡ HỆ THỐNG TẢI MULTIMEDIA ĐA NỀN TẢNG\nHỗ trợ: YouTube, TikTok, Facebook, Instagram, Twitter (X), SoundCloud, ZingMP3, NhacCuaTui, CapCut, Threads, Pinterest, Reddit, Twitch, Vimeo, Bilibili...\n', 'cyber')
    def _log_to_dl_console(self, text, tag='n'):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._log_to_dl_console, text, tag)
            return
        self.dl_log_txt.configure(state='normal')
        self.dl_log_txt.insert('end', text, tag)
        self.dl_log_txt.see('end')
        self.dl_log_txt.configure(state='disabled')
    def _clear_dl_console(self):
        self.dl_log_txt.configure(state='normal')
        self.dl_log_txt.delete('1.0', 'end')
        self.dl_log_txt.configure(state='disabled')
    def _update_dl_progress_ui(self, percent):
        try:
            self.dl_canvas.delete('progress')
            w = self.dl_canvas.winfo_width()
            h = self.dl_canvas.winfo_height()
            fill_w = int(percent / 100.0 * w)
            self.dl_canvas.create_rectangle(0, 0, fill_w, h, fill=ACCENT3, outline='', tags='progress')
            self.lbl_dl_percent.config(text=f'{percent:.1f}%')
        except Exception as e:
            return None
    def _start_download_thread(self):
        url = self.v_dl_url.get().strip()
        mode = self.v_dl_mode.get()
        fmt_str = self.v_dl_fmt.get()
        qual_str = self.v_dl_quality.get()
        if not url:
            messagebox.showwarning('Chú ý', 'Vui lòng nhập link cần tải!', parent=self.root)
            return None
        else:
            if self.is_downloading:
                messagebox.showinfo('Thông báo', 'Đang có tác vụ tải khác chạy dưới nền!', parent=self.root)
                return None
            else:
                self.is_downloading = True
                self.btn_download.config(state='disabled', bg=BORDER, fg=MUTED, text='Đang tải...')
                self._clear_dl_console()
                self._update_dl_progress_ui(0.0)
                threading.Thread(target=self._download_worker, args=(url, mode, fmt_str, qual_str), daemon=True).start()
    def _download_worker(self, url, mode, fmt_str, qual_str):
        """Bọc toàn bộ worker bằng try/except: mọi lỗi ở phần chuẩn bị
        (makedirs, yt-dlp -U, ffmpeg) cũng phải khôi phục nút Tải —
        trước đây phần này nằm NGOÀI try nên btn_download bị khóa vĩnh viễn."""
        try:
            self._download_worker_impl(url, mode, fmt_str, qual_str)
        except Exception as e:
            self._log_to_dl_console(f'\n❌ Có lỗi xảy ra: {str(e)}\n', 'danger')
            self._post(messagebox.showerror, 'Lỗi', f'Có lỗi xảy ra:\n{str(e)}')
            self._download_finished(False)

    def _download_worker_impl(self, url, mode, fmt_str, qual_str):
        paths = get_toolbox_paths()
        os.makedirs(paths['output'], exist_ok=True)
        self._log_to_dl_console('⏳ Đang kiểm tra hệ thống...\n', 'cyber')
        if not os.path.exists(paths['yt_dlp']):
            self._log_to_dl_console('⚡ Không tìm thấy yt-dlp.exe. Đang tiến hành tải tự động...\n', 'warn')
            try:
                urllib.request.urlretrieve('https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe', paths['yt_dlp'])
                self._log_to_dl_console('✓ Cài đặt yt-dlp.exe thành công!\n', 'ok')
            except Exception as e:
                self._log_to_dl_console(f'✗ Thất bại khi tải yt-dlp: {str(e)}\n Kiểm tra kết nối mạng!\n', 'danger')
                self._download_finished(False)
                return None
        else:
            self._log_to_dl_console('🌐 Đang kiểm tra phiên bản mới nhất của yt-dlp...\n', 'n')
            try:
                # Chạy lệnh cập nhật chính thức của yt-dlp (timeout 60s, xét returncode)
                proc_u = subprocess.run([paths['yt_dlp'], '-U'], capture_output=True, text=True,
                                        timeout=60, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                if proc_u.returncode != 0:
                    self._log_to_dl_console('⚠️ Cập nhật yt-dlp thất bại, dùng bản hiện tại.\n', 'warn')
                elif 'is up to date' in proc_u.stdout or 'đã ở phiên bản mới nhất' in proc_u.stdout:
                    self._log_to_dl_console('✓ yt-dlp đã ở phiên bản mới nhất!\n', 'ok')
                else:
                    self._log_to_dl_console('✓ Tự động cập nhật yt-dlp hoàn tất!\n', 'ok')
            except Exception as e:
                self._log_to_dl_console(f'⚠️ Không thể cập nhật tự động yt-dlp: {str(e)}\n', 'warn')
        
        if not os.path.exists(paths['aria2c']):
            self._log_to_dl_console('⚡ Không tìm thấy Aria2c tích hợp sẵn, sẽ dùng bộ tải thường.\n', 'warn')
        
        platform_name = get_platform_folder(url)
        platform_folder = os.path.join(paths['output'], platform_name)
        os.makedirs(platform_folder, exist_ok=True)
        output_template = os.path.join(platform_folder, '%(title)s.%(ext)s')
        self._log_to_dl_console('🎬 Kiểm tra bộ giải mã FFmpeg...\n', 'n')
        has_ffmpeg = download_ffmpeg_auto_redirect(self, paths)
        # Cookie cho TikTok phải được thêm TRƯỚC URL (danh sách extra_args)
        extra_args = []
        if 'tiktok.com' in url.lower():
            _apply_tiktok_cookie_redirect(self, extra_args)
        cmd = build_download_cmd(paths, url, output_template, mode, fmt_str, qual_str,
                                 has_ffmpeg, aria2c_exists=os.path.exists(paths['aria2c']),
                                 format_id=self._dl_format_map.get(qual_str) if mode == '2' else None,
                                 extra_args=extra_args)
        if mode == '1':
            self._log_to_dl_console(f'🎵 Đang tiến hành tải nhạc {fmt_str}...\n', 'cyber')
        else:
            self._log_to_dl_console(f'🎬 Đang tải video MP4 — chất lượng: {qual_str}...\n', 'warn')
        try:
            self._log_to_dl_console('▶ Đang thiết lập kết nối tới máy chủ...\n', 'n')
            proc = subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, encoding='utf-8', errors='replace', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
            line_queue = deque()
            reader_done = threading.Event()

            def output_reader():
                try:
                    if proc.stdout:
                        for raw in iter(proc.stdout.readline, ''):
                            line_queue.append(raw)
                finally:
                    reader_done.set()

            reader_th = threading.Thread(target=output_reader, daemon=True)
            reader_th.start()
            last_pct = 0.0
            last_log_step = -1
            err_line = ''
            while True:
                drained = False
                while line_queue:
                    drained = True
                    line = line_queue.popleft()

                    if '[download]' in line and '%' in line:
                        pct, detail = parse_ytdlp_progress_line(line)
                        if pct is not None:
                            last_pct = pct
                            self._post(self._update_dl_progress_ui, last_pct)
                            d_clean = " ".join(detail.split())
                            self._post(self.lbl_dl_percent.config, text=d_clean)
                            # Ghi log theo mốc 5% (không lặp dòng như kiểu % 5 == 0)
                            step = int(pct) // 5
                            if step != last_log_step or pct >= 100.0:
                                last_log_step = step
                                self._log_to_dl_console(line, 'n')
                        continue

                    if line.startswith('[#') and 'DL:' in line:
                        # Trích xuất phần trăm từ dòng log Aria2 để hiển thị đẹp mắt
                        m = re.search(r'\((\d+)%\)', line)
                        if m:
                            pct = float(m.group(1))
                            self._post(self._update_dl_progress_ui, pct)
                            step = int(pct) // 5
                            if step != last_log_step or pct >= 100.0:
                                last_log_step = step
                                self._log_to_dl_console(line, 'n')
                        continue

                    self._log_to_dl_console(line, 'n')
                    if 'ERROR:' in line:
                        err_line = line.strip()
                    pct, detail = parse_ytdlp_progress_line(line)
                    if pct is not None:
                        last_pct = pct
                        self._post(self._update_dl_progress_ui, last_pct)
                    if 'Merger' in line or 'ExtractAudio' in line or 'Post-process' in line:
                        self._post(self._update_dl_progress_ui, 99.0)
                # Thoát khi process kết thúc VÀ thread đọc đã báo xong VÀ hết hàng đợi
                # (bỏ điều kiện cũ gây race → mất log cuối)
                if proc.poll() is not None and reader_done.is_set() and not line_queue:
                    break
                if not drained:
                    time.sleep(0.05)
            reader_th.join(timeout=3)
            proc.wait(timeout=10)
            if proc.returncode != 0:
                hint = f'\n▶ Nguyên nhân: {err_line}' if err_line else ''
                self._log_to_dl_console(f'\n❌ Tải thất bại với lỗi hệ thống.{hint}\n', 'danger')
                self._post(self._update_dl_progress_ui, 0.0)
                self._post(messagebox.showerror, "Thất bại",
                           f"Không thể tải file!\nXem nhật ký để biết chi tiết.{hint}\n\n{platform_folder}")
                self._download_finished(False)
            else:
                self._log_to_dl_console(f'\n✅ Tải thành công!\n📁 Lưu tại: {platform_folder}\n', 'ok')
                self._post(self._update_dl_progress_ui, 100.0)
                self._post(messagebox.showinfo, "Thành công", f"Tải thành công!\nĐã lưu tại:\n{platform_folder}")
                self._download_finished(True)
        except Exception as e:
            self._log_to_dl_console(f'\n❌ Có lỗi xảy ra: {str(e)}\n', 'danger')
            self._post(messagebox.showerror, "Lỗi", f"Có lỗi xảy ra:\n{str(e)}")
            self._download_finished(False)

    def _post(self, fn, *args, **kwargs):
        """Đưa callback về main loop một cách an toàn từ luồng nền.

        `root.after` là lời gọi Tcl: khi main thread chưa vào mainloop (đang khởi
        động) hoặc cửa sổ đã đóng, nó raise RuntimeError/TclError. Retry NGAY TRONG
        LUỒNG GỌI có giới hạn (10 x 50ms) — không dùng threading.Timer vì mỗi lần
        fail mà tạo timer sẽ sinh bão timer/threads khi mainloop chưa chạy.
        Trả về True nếu đã xếp được việc.

        Nhận cả **kwargs: nhiều nơi gọi `_post(lbl.config, state=..., text=...)`;
        nếu chỉ nhận positional thì TypeError ném ra NGAY ở luồng gọi và làm hỏng
        luồng tải (đã từng hiện dialog 'unexpected keyword argument state').
        """
        for _ in range(10):
            try:
                if not self.root.winfo_exists():
                    return False
                # Tk after(delay, func, *args) KHÔNG nhận **kwargs: truyền thẳng
                # sẽ raise "Misc.after() got an unexpected keyword argument".
                # Bọc lambda để kwargs đi tới fn chứ không tới after().
                self.root.after(0, lambda: fn(*args, **kwargs))
                return True
            except (RuntimeError, tk.TclError):
                time.sleep(0.05)
        return False

    def _on_dl_url_changed(self, *_):
        """Tự phân tích khi link được dán/nhập (nút 🔍 đã bị ẩn).

        Debounce 700ms: gõ từng phím không kích hoạt, chỉ chạy khi người dùng
        ngừng nhập. Không bắt URL kiểu 'xxx' — chờ tới khi có http(s).
        """
        if self._dl_probe_job:
            try:
                self.root.after_cancel(self._dl_probe_job)
            except tk.TclError:
                pass
            self._dl_probe_job = None
        url = self.v_dl_url.get().strip()
        if 'http://' not in url and 'https://' not in url:
            return
        self._dl_probe_job = self.root.after(700, self._auto_probe_now, url)

    def _auto_probe_now(self, url):
        self._dl_probe_job = None
        if url != self.v_dl_url.get().strip():
            return  # link đổi trong lúc chờ debounce
        self._probe_link_thread(auto=True)

    def _probe_link_thread(self, auto=False):
        """Phân tích link (logic /api/info của reclip): lấy title + danh sách quality thật.

        auto=True: chạy từ trace khi dán link — không hiện modal, không chạy
        lại URL đã phân tích thành công.
        """
        url = self.v_dl_url.get().strip()
        if not url:
            if not auto:
                messagebox.showwarning('Chú ý', 'Vui lòng nhập link cần phân tích!', parent=self.root)
            return None
        if auto and url == self._dl_probed_url:
            return None
        if self.is_downloading:
            if not auto:
                messagebox.showinfo('Thông báo', 'Đang có tác vụ tải khác chạy dưới nền!', parent=self.root)
            return None
        paths = get_toolbox_paths()
        if not os.path.exists(paths['yt_dlp']):
            self._log_to_dl_console('⚡ Chưa có yt-dlp.exe — hãy bấm "Bắt đầu tải" một lần để tool tự cài.\n', 'warn')
            return None
        self._log_to_dl_console('🔍 Đang phân tích link...\n', 'cyber')

        def worker():
            try:
                info = probe_video_info(url, paths['yt_dlp'])
            except Exception as e:
                self._log_to_dl_console(f'✗ Không phân tích được link: {e}\n', 'danger')
                return None
            options = build_quality_options(info)
            title = info.get('title') or '(không có tiêu đề)'
            uploader = info.get('uploader') or info.get('channel') or '?'
            dur = info.get('duration')
            dur_s = f'{int(dur) // 60}:{int(dur) % 60:02d}' if dur else '?'

            def apply_ui():
                # Link đổi giữa chừng → bỏ kết quả cũ, đừng ghi đè danh sách quality
                if url != self.v_dl_url.get().strip():
                    return
                self._dl_probed_url = url
                self._log_to_dl_console(f'✓ {sanitize_title(title)}\n  Kênh: {uploader} | Thời lượng: {dur_s} | {len(options)} chất lượng video\n', 'ok')
                if options:
                    labels = ['Tự động (Cao nhất)'] + [o['label'] for o in options]
                    self._dl_format_map = {o['label']: o['format_id'] for o in options}
                    self.cb_quality['values'] = labels
                    if self.v_dl_quality.get() not in labels:
                        self.v_dl_quality.set('Tự động (Cao nhất)')
                    self._log_to_dl_console('  → Đã nạp danh sách chất lượng thật vào ô "Cấu hình tùy chọn".\n', 'cyber')

            self._post(apply_ui)
            return None

        threading.Thread(target=worker, daemon=True).start()

    def _download_finished(self, success):
        self.is_downloading = False
        self._post(self.btn_download.config, state='normal', bg=ACCENT3, fg=BG, text='Bắt đầu tải  📥')

    def _opt_log(self, text, tag='info'):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._opt_log, text, tag)
            return
        self.opt_log_txt.configure(state='normal')
        self.opt_log_txt.insert('end', text, tag)
        self.opt_log_txt.see('end')
        self.opt_log_txt.configure(state='disabled')

    def _update_opt_progress(self, percent, stage_text=""):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._update_opt_progress, percent, stage_text)
            return
        try:
            self.opt_progress_canvas.delete('bar')
            w = self.opt_progress_canvas.winfo_width()
            h = self.opt_progress_canvas.winfo_height()
            if w <= 1:
                w = 510
            fill_w = int(percent / 100.0 * w)
            self.opt_progress_canvas.create_rectangle(0, 0, fill_w, h, fill=ACCENT4, outline='', tags='bar')
            if stage_text:
                self.lbl_opt_stage.config(text=f"Tiến độ: {percent:.1f}% | {stage_text}")
        except Exception:
            pass

    def _build_net_tab(self):
        def install_psutil():
            try:
                import psutil
            except ImportError:
                try:
                    subprocess.run([sys.executable, "-m", "pip", "install", "psutil", "--quiet"], timeout=20, capture_output=True, creationflags=0x08000000 if os.name == 'nt' else 0)
                except Exception:
                    pass
        threading.Thread(target=install_psutil, daemon=True).start()

        self.tab_net_frame = tk.Frame(self.root, bg=BG)
        
        # Sub-tab bar
        sub_bar = tk.Frame(self.tab_net_frame, bg=PANEL, height=34)
        sub_bar.pack(fill='x')
        sub_bar.pack_propagate(False)

        self.btn_sub_net_optimize = tk.Button(sub_bar, text='🚀 Tối Ưu Mạng', bg=INPUT, fg=MUTED, font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2')
        self.btn_sub_sub_optimize_pack = self.btn_sub_net_optimize.pack(side='left', fill='both', expand=True)

        self.btn_sub_wifi_pw = tk.Button(sub_bar, text='🔑 Xem Mật Khẩu WiFi', bg=INPUT, fg=MUTED, font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2')
        self.btn_sub_wifi_pw.pack(side='left', fill='both', expand=True)

        # Divider
        tk.Frame(self.tab_net_frame, bg=BORDER, height=1).pack(fill='x')

        # 2 content frames
        self.sub_net_optimize_frame = tk.Frame(self.tab_net_frame, bg=BG)
        self.sub_wifi_pw_frame = tk.Frame(self.tab_net_frame, bg=BG)

        def _switch_sub_net(sub_idx):
            self.sub_net_optimize_frame.pack_forget()
            self.sub_wifi_pw_frame.pack_forget()
            self.btn_sub_net_optimize.config(bg=INPUT, fg=MUTED)
            self.btn_sub_wifi_pw.config(bg=INPUT, fg=MUTED)
            if sub_idx == 1:
                self.btn_sub_net_optimize.config(bg=ACCENT4, fg=BG)
                self.sub_net_optimize_frame.pack(fill='both', expand=True)
            elif sub_idx == 2:
                self.btn_sub_wifi_pw.config(bg=ACCENT4, fg=BG)
                self.sub_wifi_pw_frame.pack(fill='both', expand=True)

        self.btn_sub_net_optimize.config(command=lambda: _switch_sub_net(1))
        self.btn_sub_wifi_pw.config(command=lambda: _switch_sub_net(2))

        # Build Sub-tab 1: Tối Ưu Mạng (re-housed to self.sub_net_optimize_frame)
        hdr = tk.Frame(self.sub_net_optimize_frame, bg=PANEL, height=40)
        hdr.pack(fill='x', pady=(0, 10))
        hdr.pack_propagate(False)
        tk.Label(hdr, text='🚀 TỐI ƯU MẠNG TỰ ĐỘNG', bg=PANEL, fg=ACCENT4, font=('Segoe UI', 12, 'bold')).pack(side='left', padx=15)
        
        main_f = tk.Frame(self.sub_net_optimize_frame, bg=BG, padx=15, pady=5)
        main_f.pack(fill='both', expand=True)
        
        self.btn_opt_start = tk.Button(main_f, text='🚀 Bắt đầu tối ưu', command=self._start_net_optimization, bg=ACCENT4, fg=BG, font=('Segoe UI', 10, 'bold'), relief='flat', cursor='hand2')
        self.btn_opt_start.pack(fill='x', pady=(0, 10), ipady=8)
        
        prog_f = tk.Frame(main_f, bg=BG)
        prog_f.pack(fill='x', pady=(0, 10))
        
        self.lbl_opt_stage = tk.Label(prog_f, text='Trạng thái: Sẵn sàng tối ưu', bg=BG, fg=MUTED, font=('Segoe UI', 9, 'bold'))
        self.lbl_opt_stage.pack(anchor='w', pady=(0, 4))
        
        self.opt_progress_canvas = tk.Canvas(prog_f, bg='#1a2638', height=14, highlightthickness=0)
        self.opt_progress_canvas.pack(fill='x')
        
        log_lbl = tk.Label(main_f, text='Nhật ký tối ưu hóa mạng:', bg=BG, fg=MUTED, font=('Segoe UI', 9, 'bold'))
        log_lbl.pack(anchor='w', pady=(5, 2))
        
        self.opt_log_txt = st.ScrolledText(main_f, bg='#0c1017', fg='#d1d5db', insertbackground='white', font=('Consolas', 9), relief='flat', wrap='word')
        self.opt_log_txt.pack(fill='both', expand=True, pady=(0, 10))
        
        for tag, color in [
            ('info', MUTED),
            ('ok', ACCENT2),
            ('warn', WARN),
            ('danger', DANGER),
            ('accent', ACCENT),
            ('accent3', ACCENT3),
            ('header', '#ffffff'),
        ]:
            self.opt_log_txt.tag_config(tag, foreground=color)
        
        self._opt_log('⚡ HỆ THỐNG TỐI ƯU MẠNG TỰ ĐỘNG\n', 'accent3')
        self._opt_log('Hệ thống sẽ chạy chẩn đoán toàn diện, sửa lỗi và kiểm tra lại tốc độ mạng.\nBấm nút "Bắt đầu tối ưu" ở trên để bắt đầu.\n', 'info')

        # Build Sub-tab 2: Xem Mật Khẩu WiFi
        self._build_wifi_pw_sub()

        # Default sub-tab
        _switch_sub_net(1)

    def _build_wifi_pw_sub(self):
        # Thiết lập style cho Treeview
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Wifi.Treeview',
            background=INPUT, foreground=TXT, fieldbackground=INPUT,
            rowheight=26, font=('Segoe UI', 9))
        style.configure('Wifi.Treeview.Heading',
            background=PANEL, foreground=ACCENT4, font=('Segoe UI', 9, 'bold'))
        style.map('Wifi.Treeview', background=[('selected', ACCENT4)], foreground=[('selected', BG)])
        
        main_f = tk.Frame(self.sub_wifi_pw_frame, bg=BG, padx=15, pady=10)
        main_f.pack(fill='both', expand=True)
        
        # Nút Quét WiFi
        self.btn_wifi_scan = tk.Button(main_f, text='🔍 Quét WiFi', command=self._scan_wifi_thread, bg=ACCENT4, fg=BG, font=('Segoe UI', 10, 'bold'), relief='flat', cursor='hand2')
        self.btn_wifi_scan.pack(fill='x', pady=(0, 10), ipady=8)
        
        # Divider
        tk.Frame(main_f, bg=BORDER, height=1).pack(fill='x', pady=(0, 10))
        
        # Container cho Treeview và Scrollbar
        tree_f = tk.Frame(main_f, bg=BG)
        tree_f.pack(fill='both', expand=True, pady=(0, 10))
        
        self.wifi_tree = ttk.Treeview(tree_f, columns=('#', 'SSID', 'Password', 'Security'), show='headings', style='Wifi.Treeview')
        self.wifi_tree.heading('#', text='STT')
        self.wifi_tree.heading('SSID', text='Tên WiFi (SSID)')
        self.wifi_tree.heading('Password', text='Mật khẩu')
        self.wifi_tree.heading('Security', text='Bảo mật')
        
        self.wifi_tree.column('#', width=40, anchor='center', stretch=False)
        self.wifi_tree.column('SSID', width=150, anchor='w')
        self.wifi_tree.column('Password', width=180, anchor='w')
        self.wifi_tree.column('Security', width=100, anchor='w')
        
        vsb = tk.Scrollbar(tree_f, orient='vertical', command=self.wifi_tree.yview)
        self.wifi_tree.configure(yscrollcommand=vsb.set)
        
        self.wifi_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        
        # Double-click row → copy mật khẩu
        self.wifi_tree.bind('<Double-Button-1>', self._copy_selected_wifi_pw)
        
        # Xác nhận label
        self.lbl_wifi_confirm = tk.Label(main_f, text='', bg=BG, fg=ACCENT2, font=('Segoe UI', 9, 'bold'))
        self.lbl_wifi_confirm.pack(pady=(0, 5))
        
        # Footer buttons
        btn_f = tk.Frame(main_f, bg=BG)
        btn_f.pack(fill='x')
        
        btn_copy = tk.Button(btn_f, text='📋 Sao chép mật khẩu đã chọn', command=self._copy_selected_wifi_pw, bg=INPUT, fg=TXT, font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2')
        btn_copy.pack(side='left', fill='x', expand=True, padx=(0, 5), ipady=6)
        
        btn_export = tk.Button(btn_f, text='💾 Xuất ra file .txt', command=self._export_wifi_txt, bg=INPUT, fg=ACCENT4, font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2')
        btn_export.pack(side='right', fill='x', expand=True, padx=(5, 0), ipady=6)
        
        self.wifi_results = []

    def _copy_selected_wifi_pw(self, event=None):
        selected = self.wifi_tree.selection()
        if not selected:
            return None
        item = self.wifi_tree.item(selected[0])
        values = item.get('values', [])
        if len(values) >= 3:
            password = str(values[2])
            self.root.clipboard_clear()
            self.root.clipboard_append(password)
            # Không gọi root.update() ở đây: re-entrancy giữa callback chuột có thể
            # dispatch tiếp event đang xếp hàng (double-click chạy 2 lần)

            self.lbl_wifi_confirm.config(text="✅ Đã sao chép mật khẩu!", fg=ACCENT2)
            
            def clear_confirm():
                try:
                    self.lbl_wifi_confirm.config(text="")
                except Exception:
                    pass
            self.root.after(2000, clear_confirm)

    def _scan_wifi_passwords(self):
        results = []
        import subprocess, platform, re, os
        
        if platform.system() == 'Windows':
            try:
                out = subprocess.run(
                    ['netsh', 'wlan', 'show', 'profiles'],
                    capture_output=True, text=True, timeout=10,
                    creationflags=0x08000000 if os.name == 'nt' else 0
                ).stdout
            except Exception as e:
                return [{'ssid': 'Lỗi chạy netsh', 'password': str(e), 'auth': 'Lỗi'}]
            
            profiles = re.findall(r'(?:All User Profile|Hồ sơ tất cả người dùng)\s*:\s*(.+)', out)
            
            for ssid in profiles:
                ssid = ssid.strip()
                try:
                    detail = subprocess.run(
                        ['netsh', 'wlan', 'show', 'profile', ssid, 'key=clear'],
                        capture_output=True, text=True, timeout=10,
                        creationflags=0x08000000 if os.name == 'nt' else 0
                    ).stdout
                    
                    pw_match = re.search(r'(?:Key Content|Nội dung khóa)\s*:\s*(.+)', detail)
                    password = pw_match.group(1).strip() if pw_match else '(Không có / Không có quyền)'
                    
                    auth_match = re.search(r'(?:Authentication|Xác thực)\s*:\s*(.+)', detail)
                    auth = auth_match.group(1).strip() if auth_match else 'Unknown'
                    
                    results.append({'ssid': ssid, 'password': password, 'auth': auth})
                except Exception:
                    results.append({'ssid': ssid, 'password': '(Lỗi đọc / Cần quyền Admin)', 'auth': 'Unknown'})
        
        elif platform.system() == 'Darwin':
            try:
                out = subprocess.run(
                    ['/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport', '-s'],
                    capture_output=True, text=True, timeout=10
                ).stdout
                ssids = [line.split()[0] for line in out.strip().splitlines()[1:] if line.strip()]
                for ssid in ssids:
                    try:
                        pw = subprocess.run(
                            ['security', 'find-generic-password', '-wa', ssid],
                            capture_output=True, text=True, timeout=5
                        ).stdout.strip()
                        results.append({'ssid': ssid, 'password': pw or '(Không có)', 'auth': 'WPA'})
                    except Exception:
                        results.append({'ssid': ssid, 'password': '(Lỗi)', 'auth': 'Unknown'})
            except Exception:
                pass
        
        elif platform.system() == 'Linux':
            import os, glob
            config_paths = glob.glob('/etc/NetworkManager/system-connections/*.nmconnection') + \
                           glob.glob('/etc/NetworkManager/system-connections/*')
            for path in config_paths:
                try:
                    content = open(path, 'r', errors='ignore').read()
                    ssid_m = re.search(r'ssid=(.+)', content)
                    pw_m = re.search(r'psk=(.+)', content)
                    auth_m = re.search(r'key-mgmt=(.+)', content)
                    if ssid_m:
                        results.append({
                            'ssid': ssid_m.group(1).strip(),
                            'password': pw_m.group(1).strip() if pw_m else '(Không có)',
                            'auth': auth_m.group(1).strip() if auth_m else 'Unknown'
                        })
                except Exception:
                    pass
        
        return results

    def _scan_wifi_thread(self):
        self.btn_wifi_scan.config(state='disabled', text='⌛ Đang quét WiFi...')
        for item in self.wifi_tree.get_children():
            self.wifi_tree.delete(item)
        self.lbl_wifi_confirm.config(text="⌛ Đang đọc danh sách mạng lưu trên máy...", fg=MUTED)
        
        def worker():
            results = self._scan_wifi_passwords()
            self.wifi_results = results
            
            def update_ui():
                self.btn_wifi_scan.config(state='normal', text='🔍 Quét WiFi')
                for item in self.wifi_tree.get_children():
                    self.wifi_tree.delete(item)
                
                if not results:
                    self.lbl_wifi_confirm.config(text="⚠️ Không tìm thấy WiFi nào hoặc cần chạy với quyền Admin.", fg=DANGER)
                    return
                
                for idx, res in enumerate(results, 1):
                    self.wifi_tree.insert('', 'end', values=(idx, res['ssid'], res['password'], res['auth']))
                
                self.lbl_wifi_confirm.config(text=f"✅ Đã quét xong! Tìm thấy {len(results)} mạng WiFi.", fg=ACCENT2)
                
            self._post(update_ui)
            
        threading.Thread(target=worker, daemon=True).start()

    def _export_wifi_txt(self):
        if not hasattr(self, 'wifi_results') or not self.wifi_results:
            messagebox.showwarning('Chú ý', 'Không có dữ liệu WiFi để xuất. Hãy thực hiện quét trước!', parent=self.root)
            return None
        
        file_path = fd.asksaveasfilename(
            defaultextension='.txt',
            filetypes=[('Text files', '*.txt')],
            title="Xuất danh sách mật khẩu WiFi",
            parent=self.root
        )
        if not file_path:
            return
            
        try:
            import datetime
            now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            
            with open(file_path, 'w', encoding='utf-8-sig') as f:
                f.write("DANH SÁCH MẬT KHẨU WIFI\n")
                f.write(f"Xuất lúc: {now_str}\n")
                f.write("================================================================================\n")
                f.write(f"{'STT':<5} {'SSID':<30} {'Mật khẩu':<30} {'Bảo mật':<15}\n")
                f.write("--------------------------------------------------------------------------------\n")
                for idx, res in enumerate(self.wifi_results, 1):
                    f.write(f"{idx:<5} {res['ssid']:<30} {res['password']:<30} {res['auth']:<15}\n")
                f.write("================================================================================\n")
            
            self.lbl_wifi_confirm.config(text=f"✅ Đã xuất file thành công!", fg=ACCENT2)
            
            def clear_confirm():
                try:
                    self.lbl_wifi_confirm.config(text="")
                except Exception:
                    pass
            self.root.after(3000, clear_confirm)
            
            messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra file:\n{file_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {str(e)}", parent=self.root)

    def _start_net_optimization(self):
        warning_msg = """⚠️ Lưu ý trước khi tối ưu:

- Giai đoạn đo tốc độ (Speedtest) sẽ chiếm toàn bộ băng thông ~30 giây.
  Trình duyệt vẫn hoạt động nhưng trang web có thể load chậm hơn.

- Tất cả ứng dụng đang mở (trình duyệt, Zalo, v.v.) sẽ KHÔNG bị ảnh hưởng.
- Tool chỉ tác động đến tiến trình ẩn chạy nền không rõ nguồn gốc.
- Mọi thay đổi cài đặt mạng đều hỏi xác nhận trước khi thực hiện.

Bấm OK để tiếp tục."""
        messagebox.showinfo("Lưu ý trước khi tối ưu", warning_msg, parent=self.root)
        
        self.btn_opt_start.config(state='disabled', text='⌛ Đang tối ưu...')
        
        self.opt_log_txt.configure(state='normal')
        self.opt_log_txt.delete('1.0', 'end')
        self.opt_log_txt.configure(state='disabled')
        
        self._opt_log('⚡ BẮT ĐẦU QUY TRÌNH TỐI ƯU MẠNG TỰ ĐỘNG\n', 'accent3')
        
        threading.Thread(target=self._net_optimization_thread, daemon=True).start()

    def _net_optimization_thread(self):
        """Bọc ngoài để LUÔN luôn khôi phục nút Bắt đầu, kể cả khi worker lỗi ở giữa chừng."""
        try:
            self._net_optimization_worker()
        except Exception as e:
            self._opt_log(f'⚠️ Lỗi không mong đợi khi tối ưu: {e}\n', 'danger')
        finally:
            # Chống "process bị đóng băng vĩnh viễn": nếu worker rớt sau khi đã
            # suspend mà trước khi resume, ở đây sẽ rescue nốt.
            for pid in list(getattr(self, '_susp_rescue', ())):
                try:
                    import psutil as _psutil
                    _psutil.Process(pid).resume()
                    self._opt_log(f"▶️ Đã phục hồi (rescue) PID {pid}\n", "ok")
                except Exception:
                    pass
                finally:
                    self._susp_rescue.discard(pid)
            self._post(lambda: self.btn_opt_start.config(state='normal', text='🚀 Bắt đầu tối ưu'))
            self._post(lambda: self.lbl_opt_stage.config(text='Trạng thái: Hoàn tất tối ưu!'))

    def _net_optimization_worker(self):
        import platform
        import subprocess
        import sys
        import time
        import socket
        import re
        import importlib.util
        
        try:
            import psutil
        except ImportError:
            self._opt_log("⚡ Đang cài đặt thư viện psutil...\n", "warn")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "psutil", "--quiet"], timeout=20, capture_output=True, creationflags=0x08000000 if os.name == 'nt' else 0)
                import psutil
            except Exception as e:
                self._opt_log(f"⚠️ Cài đặt psutil lỗi: {e}\n", "danger")
                self._post(lambda: self.btn_opt_start.config(state='normal', text='🚀 Bắt đầu tối ưu'))
                return
        
        system_os = platform.system()
        
        self._update_opt_progress(0, "Đang chụp snapshot các ứng dụng hoạt động...")
        self._opt_log("⏳ Đang snapshot các ứng dụng đang chạy...\n", "info")
        
        def is_system_process(name):
            SYSTEM_WHITELIST = {
                'chrome.exe', 'brave.exe', 'msedge.exe', 'firefox.exe', 'opera.exe',
                'svchost.exe', 'lsass.exe', 'system', 'smss.exe', 'csrss.exe',
                'winlogon.exe', 'services.exe', 'explorer.exe', 'dwm.exe',
                'taskhostw.exe', 'runtimebroker.exe', 'sihost.exe',
                'msmpeng.exe', 'msseces.exe', 'defender.exe',
                'python.exe', 'pythonw.exe', 'python3.exe',
                'zalo.exe', 'telegram.exe', 'discord.exe', 'zoom.exe', 'teams.exe',
                'code.exe', 'cursor.exe', 'notepad.exe', 'notepad++.exe',
                'winrar.exe', '7z.exe', 'ultraviewer.exe',
            }
            return (name or '').lower() in SYSTEM_WHITELIST

        def get_running_apps_snapshot():
            # CHỈ chụp tiến trình không phải hệ thống. Trước đây là set(psutil.pids())
            # = mọi process đang sống → bước 4 "bảo vệ app người dùng" che hết,
            # whitelist becomes dead logic và luôn kết luận "không có gì đáng ngờ".
            snapshot = set()
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if not is_system_process(proc.info['name']):
                        snapshot.add(proc.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            return snapshot

        baseline_snapshot = get_running_apps_snapshot()
        self._opt_log(f"✅ Đã bảo vệ {len(baseline_snapshot)} tiến trình người dùng đang mở.\n", "ok")

        def is_user_app(pid, baseline_snapshot):
            return pid in baseline_snapshot

        self._update_opt_progress(16.6, "Giai đoạn 1/6: Đo tốc độ TRƯỚC (Baseline)")
        self._opt_log("⏳ Khởi động đo tốc độ mạng trước tối ưu (Baseline)...\n", "info")
        
        if importlib.util.find_spec('speedtest') is None:
            self._opt_log("⚡ Đang cài đặt speedtest-cli...\n", "warn")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "speedtest-cli", "--quiet"], timeout=20, capture_output=True, creationflags=0x08000000 if os.name == 'nt' else 0)
            except Exception as e:
                self._opt_log(f"⚠️ Cài đặt speedtest-cli thất bại: {e}\n", "danger")
        
        before = {'download': 0.0, 'upload': 0.0, 'ping': 999.0}
        try:
            import speedtest
            self._opt_log("📊 Đang đo Speedtest (Download/Upload/Ping)... Vui lòng đợi ~30 giây.\n", "info")
            st = speedtest.Speedtest()
            st.get_best_server()
            download = st.download() / 1_000_000
            upload = st.upload() / 1_000_000
            ping = st.results.ping
            before = {'download': download, 'upload': upload, 'ping': ping}
            
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
            self._opt_log(f"📊 [TRƯỚC] Download: {download:.2f} Mbps\n", "header")
            self._opt_log(f"           Upload:   {upload:.2f} Mbps\n", "header")
            self._opt_log(f"           Ping:     {ping:.1f} ms\n", "header")
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi chạy Speedtest: {e}\n", "danger")
        
        self._update_opt_progress(33.3, "Giai đoạn 2/6: Chẩn đoán mạng toàn diện")
        self._opt_log("⏳ Bắt đầu chẩn đoán kết nối toàn diện...\n", "info")
        
        issues = []
        
        dns_targets = ['google.com', 'facebook.com', 'cloudflare.com', 'youtube.com']
        dns_times = []
        try:
            for host in dns_targets:
                try:
                    start_dns = time.perf_counter()
                    socket.getaddrinfo(host, None)
                    elapsed_dns = (time.perf_counter() - start_dns) * 1000
                    dns_times.append(elapsed_dns)
                    self._opt_log(f"🔍 DNS: {host} → {elapsed_dns:.1f}ms\n", "info")
                except Exception:
                    dns_times.append(999.0)
                    self._opt_log(f"🔍 DNS: {host} → Lỗi kết nối!\n", "danger")
            avg_dns = sum(dns_times) / len(dns_times)
            if avg_dns > 150:
                issues.append('dns_slow')
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi chẩn đoán DNS: {e}\n", "danger")
            avg_dns = 999.0
            
        optimum_size = 1500 - 28
        current_mtu = 1500
        try:
            test_sizes = [1400, 1420, 1440, 1450, 1460, 1472, 1480, 1492, 1500]
            working_sizes = []
            for size in test_sizes:
                try:
                    if system_os == 'Windows':
                        cmd = ['ping', '-f', '-l', str(size), '-n', '2', '8.8.8.8']
                    else:
                        cmd = ['ping', '-M', 'do', '-s', str(size), '-c', '2', '8.8.8.8']
                    
                    r = subprocess.run(cmd, capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                    out = r.stdout or ""
                    err = r.stderr or ""
                    
                    fragmented = ("fragment" in out.lower() or "fragment" in err.lower() or "phân mảnh" in out.lower() or "phân mảnh" in err.lower())
                    if r.returncode == 0 and not fragmented:
                        working_sizes.append(size)
                except Exception:
                    pass
            if working_sizes:
                optimum_size = max(working_sizes)
            optimum_mtu = optimum_size + 28
            
            if system_os == 'Windows':
                # `show subinterface` KHÔNG có cột tên adapter → nhánh match
                # 'wi-fi/ethernet' không bao giờ trúng, luôn rơi vào else lấy MTU
                # của dòng ĐẦU TIÊN (thường là VPN/vEthernet/tunnel → chẩn đoán sai)
                r_mtu = subprocess.run(['netsh', 'interface', 'ipv4', 'show', 'interfaces'], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                mtu_rows = []
                for line in r_mtu.stdout.splitlines():
                    parts = line.split()
                    # Cột: Idx | MTU | State | Name...
                    if len(parts) >= 4 and parts[0].isdigit() and re.fullmatch(r'\d{4,5}', parts[1]):
                        mtu_rows.append((int(parts[1]), ' '.join(parts[3:]).lower()))
                picked = None
                for kw in ('wi-fi', 'wlan'):
                    picked = next((m for m, n in mtu_rows if kw in n), None)
                    if picked is not None:
                        break
                if picked is None:
                    picked = next((m for m, n in mtu_rows if 'ethernet' in n), None)
                if picked is not None:
                    current_mtu = picked
            self._opt_log(f"🔍 MTU tối ưu phát hiện: {optimum_mtu} bytes (MTU thực tế: {current_mtu})\n", "info")
            if abs(current_mtu - optimum_mtu) > 28:
                issues.append('mtu_suboptimal')
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi chẩn đoán MTU: {e}\n", "danger")
            
        loss_pct = 0.0
        try:
            if system_os == 'Windows':
                cmd = ['ping', '-n', '10', '8.8.8.8']
            else:
                cmd = ['ping', '-c', '10', '8.8.8.8']
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
            out = r.stdout or ""
            match = re.search(r'(\d+)%\s*(?:loss|mất|packet loss)', out.lower())
            if match:
                loss_pct = float(match.group(1))
            else:
                match_paren = re.search(r'\((\d+)%\s*(?:loss|lost|mất)?\)', out.lower())
                if match_paren:
                    loss_pct = float(match_paren.group(1))
            
            if loss_pct > 2.0:
                issues.append('packet_loss')
                self._opt_log(f"⚠️ Packet loss: {loss_pct:.0f}% — có thể ảnh hưởng chất lượng kết nối\n", "danger")
            else:
                self._opt_log(f"🔍 Packet loss: {loss_pct:.0f}% (Packets OK)\n", "ok")
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi chẩn đoán Packet Loss: {e}\n", "danger")
            
        wifi_signal = 100
        is_wifi = False
        try:
            if system_os == 'Windows':
                r = subprocess.run(['netsh', 'wlan', 'show', 'interfaces'], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                out = r.stdout or ""
                for line in out.splitlines():
                    if 'signal' in line.lower() or 'tín hiệu' in line.lower():
                        match = re.search(r'(\d+)%', line)
                        if match:
                            wifi_signal = int(match.group(1))
                            is_wifi = True
                            break
            elif system_os == 'Darwin':
                r = subprocess.run(['/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport', '-I'], capture_output=True, text=True, timeout=15)
                out = r.stdout or ""
                for line in out.splitlines():
                    if 'agrCtlRSSI' in line:
                        match = re.search(r'-\d+', line)
                        if match:
                            rssi = int(match.group(0))
                            wifi_signal = min(max(2 * (rssi + 100), 0), 100)
                            is_wifi = True
                            break
            else:
                r = subprocess.run(['iwconfig'], capture_output=True, text=True, timeout=15)
                out = r.stdout or ""
                for line in out.splitlines():
                    if 'signal level' in line.lower() or 'quality' in line.lower():
                        match = re.search(r'Signal level=(-\d+|\d+)', line.lower())
                        if match:
                            val = int(match.group(1))
                            if val < 0:
                                wifi_signal = min(max(2 * (val + 100), 0), 100)
                            else:
                                wifi_signal = val
                            is_wifi = True
                            break
            if is_wifi:
                if wifi_signal < 60:
                    issues.append('wifi_weak')
                    self._opt_log(f"⚠️ Wi-Fi signal: {wifi_signal}% — yếu, cân nhắc di chuyển gần router\n", "danger")
                else:
                    self._opt_log(f"🔍 Wi-Fi signal: {wifi_signal}% (Tốt)\n", "ok")
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi chẩn đoán tín hiệu Wi-Fi: {e}\n", "danger")
            
        self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
        self._opt_log(f"🩺 Phát hiện {len(issues)} vấn đề:\n", "danger" if issues else "ok")
        if not issues:
            self._opt_log("  • Không phát hiện bất kỳ vấn đề mạng nào!\n", "ok")
        else:
            for issue in issues:
                if issue == 'dns_slow':
                    self._opt_log(f"  • DNS phân giải chậm (avg {avg_dns:.1f}ms)\n", "warn")
                elif issue == 'mtu_suboptimal':
                    self._opt_log(f"  • MTU chưa tối ưu (MTU hiện tại: {current_mtu} vs Tối ưu: {optimum_mtu})\n", "warn")
                elif issue == 'packet_loss':
                    self._opt_log(f"  • Có hiện tượng rớt gói tin (loss {loss_pct:.1f}%)\n", "danger")
                elif issue == 'wifi_weak':
                    self._opt_log(f"  • Cường độ sóng Wi-Fi yếu ({wifi_signal}%)\n", "warn")
        self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
        
        self._update_opt_progress(50.0, "Giai đoạn 3/6: Xóa Wi-Fi profiles cũ")
        self._opt_log("⏳ Đang quét danh sách Wi-Fi profiles...\n", "info")
        
        deleted_wifi = []
        if system_os == 'Windows':
            try:
                r_prof = subprocess.run(['netsh', 'wlan', 'show', 'profiles'], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                # Dùng chung regex với scanner (có biến thể tiếng Việt) — parser cũ
                # chỉ khớp nhãn tiếng Anh → netsh locale VN tìm thấy 0 profile
                profiles = [s.strip().strip('"').strip() for s in re.findall(
                    r'(?:All User Profile|Hồ sơ tất cả người dùng)\s*:\s*(.+)', r_prof.stdout)]
                profiles = [p for p in profiles if p]

                current_ssid = ""
                r_int = subprocess.run(['netsh', 'wlan', 'show', 'interfaces'], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                for line in r_int.stdout.splitlines():
                    if 'ssid' in line.lower() and 'bssid' not in line.lower():
                        k, _, v = line.partition(':')
                        current_ssid = v.strip().strip('"').strip()
                        break
                
                import winreg
                import datetime
                import struct
                
                reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles"
                profile_dates = {}
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                    guid_count = winreg.QueryInfoKey(key)[0]
                    for i in range(guid_count):
                        guid = winreg.EnumKey(key, i)
                        # key đã mở tới ...\NetworkList\Profiles → mở guid-relative
                        # (đường dẫn đầy đủ làm FileNotFoundError bị nuốt im lặng)
                        subkey = winreg.OpenKey(key, guid)
                        try:
                            p_name, _ = winreg.QueryValueEx(subkey, "ProfileName")
                            try:
                                last_conn_data, _ = winreg.QueryValueEx(subkey, "DateLastConnected")
                                # DateLastConnected là FILETIME 8 byte (100ns kể từ 1601),
                                # KHÔNG phải SYSTEMTIME 16 byte như code cũ giả định
                                if isinstance(last_conn_data, bytes) and len(last_conn_data) >= 8:
                                    ft = struct.unpack('<Q', last_conn_data[:8])[0]
                                    if ft:
                                        dt = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=ft // 10)
                                        profile_dates[p_name] = (dt.year, dt.month, dt.day)
                                    else:
                                        profile_dates[p_name] = None
                                else:
                                    profile_dates[p_name] = None
                            except Exception:
                                profile_dates[p_name] = None
                        except Exception:
                            pass
                        winreg.CloseKey(subkey)
                    winreg.CloseKey(key)
                except Exception:
                    pass
                
                to_delete = []
                now_dt = datetime.datetime.now()
                for p in profiles:
                    if p == current_ssid:
                        continue
                    date_info = profile_dates.get(p)
                    if date_info is None:
                        to_delete.append((p, "không có lịch sử kết nối"))
                    else:
                        try:
                            p_year, p_month, p_day = date_info
                            if p_year < 1900 or p_month < 1 or p_month > 12 or p_day < 1 or p_day > 31:
                                to_delete.append((p, "lịch sử kết nối không hợp lệ"))
                            else:
                                p_dt = datetime.datetime(p_year, p_month, p_day)
                                days_diff = (now_dt - p_dt).days
                                if days_diff > 30:
                                    to_delete.append((p, f"lần cuối kết nối: {days_diff} ngày trước"))
                        except Exception:
                            to_delete.append((p, "lịch sử kết nối lỗi"))
                
                # Nếu không đọc được lịch sử kết nối từ registry (permission/đổi key),
                # MỌI profile đều rơi vào "không có lịch sử" → dialog đề nghị xoá HẾT
                # mạng Wi-Fi đã lưu. Không có dữ liệu thời gian thì không được xoá.
                if to_delete and not profile_dates:
                    self._opt_log("ℹ️ Không đọc được lịch sử kết nối Wi-Fi — bỏ qua bước dọn dẹp profile.\n", "info")
                    to_delete = []

                if to_delete:
                    msg = f"Tìm thấy {len(to_delete)} Wi-Fi profiles cũ không còn dùng:\n"
                    for p, desc in to_delete[:8]:
                        msg += f"- {p} ({desc})\n"
                    if len(to_delete) > 8:
                        msg += f"... và {len(to_delete) - 8} mạng khác.\n"
                    msg += "\nXoá để giảm nhiễu khi scan mạng?"
                    
                    ans = self._ask_on_main('yesno', "Xác nhận xóa Wi-Fi profiles cũ", msg, parent=self.root)
                    if ans:
                        for p, _ in to_delete:
                            del_rc, _, _ = _run_cmd(['netsh', 'wlan', 'delete', 'profile', f'name={p}'])
                            if del_rc == 0:
                                deleted_wifi.append(p)
                                self._opt_log(f"🗑️ Đã xoá: {p}\n", "warn")
                        self._opt_log(f"✅ Đã dọn dẹp {len(deleted_wifi)} profile Wi-Fi thừa.\n", "ok")
                    else:
                        self._opt_log("ℹ️ Bỏ qua xóa Wi-Fi profiles cũ.\n", "info")
                else:
                    self._opt_log("✅ Không có Wi-Fi profile thừa.\n", "ok")
            except Exception as e:
                self._opt_log(f"⚠️ Lỗi dọn dẹp Wi-Fi profiles: {e}\n", "danger")
        else:
            self._opt_log("ℹ️ Bỏ qua dọn dẹp Wi-Fi profiles (chỉ hỗ trợ Windows).\n", "info")
            
        self._update_opt_progress(66.6, "Giai đoạn 4/6: Phát hiện tiến trình nền ngốn mạng")
        self._opt_log("⏳ Đang quét tiến trình nền tiêu thụ băng thông...\n", "info")
        
        suspects = []
        suspended_pids = []
        self._susp_rescue = set()
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    pid = proc.pid
                    name = proc.info['name']
                    raw_conns = []
                    try:
                        raw_conns = proc.connections(kind='inet')
                    except AttributeError:
                        try:
                            raw_conns = proc.net_connections(kind='inet')
                        except Exception:
                            pass
                    except Exception:
                        pass
                    conns = [c for c in (raw_conns or []) if c.status == 'ESTABLISHED']
                    
                    if is_user_app(pid, baseline_snapshot):
                        continue
                    if is_system_process(name):
                        continue
                    if len(conns) > 5:
                        suspects.append({'pid': pid, 'name': name, 'connections': len(conns), 'proc': proc})
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            if suspects:
                msg = f"Phát hiện {len(suspects)} tiến trình nền đang dùng mạng nhiều:\n\n"
                msg += f"{'PID':<8} {'Tên':<25} {'Connections':<12}\n"
                msg += "-" * 45 + "\n"
                for s in suspects[:8]:
                    msg += f"{s['pid']:<8} {s['name']:<25} {s['connections']:<12}\n"
                if len(suspects) > 8:
                    msg += f"... và {len(suspects) - 8} tiến trình khác.\n"
                msg += "\nCác app đang mở của bạn (Brave, Zalo...) KHÔNG nằm trong danh sách này.\n"
                msg += "Muốn tạm dừng các tiến trình trên không?"
                
                ans = self._ask_on_main('yesno', "Tạm dừng các tiến trình nền", msg, parent=self.root)
                if ans:
                    for s in suspects:
                        try:
                            s['proc'].suspend()
                            suspended_pids.append(s['pid'])
                            self._susp_rescue.add(s['pid'])
                            self._opt_log(f"⏸️ Đã tạm dừng: {s['name']} (PID {s['pid']})\n", "warn")
                        except Exception as e:
                            self._opt_log(f"⚠️ Lỗi tạm dừng PID {s['pid']}: {e}\n", "danger")
                else:
                    self._opt_log("ℹ️ Bỏ qua tạm dừng tiến trình nền.\n", "info")
            else:
                self._opt_log("✅ Không phát hiện tiến trình nền đáng ngờ.\n", "ok")
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi kiểm tra tiến trình nền: {e}\n", "danger")
            
        self._update_opt_progress(83.3, "Giai đoạn 5/6: Tối ưu DNS")
        
        changed_dns = False
        best_dns_ip = ""
        best_dns_name = ""
        if 'dns_slow' in issues:
            self._opt_log("⏳ Đang chạy DNS Benchmark để tìm máy chủ phân giải nhanh nhất...\n", "info")
            servers_to_test = [
                ('8.8.8.8', 'Google DNS', '8.8.4.4'),
                ('1.1.1.1', 'Cloudflare DNS', '1.0.0.1'),
                ('9.9.9.9', 'Quad9 DNS', '149.112.112.112'),
            ]
            
            current_dns_list = []
            try:
                if system_os == 'Windows':
                    r_dns = subprocess.run(['netsh', 'interface', 'ip', 'show', 'dns'], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                    ips = re.findall(r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', r_dns.stdout)
                    if ips:
                        current_dns_list = list(dict.fromkeys(ips))
                elif os.path.exists('/etc/resolv.conf'):
                    with open('/etc/resolv.conf', 'r') as f:
                        for line in f:
                            if line.strip().startswith('nameserver'):
                                parts = line.split()
                                if len(parts) >= 2:
                                    current_dns_list.append(parts[1])
            except Exception:
                pass
            
            if current_dns_list:
                servers_to_test.append((current_dns_list[0], 'DNS Hiện tại', current_dns_list[1] if len(current_dns_list) > 1 else ''))
            
            dns_results = []
            for ip, name, backup in servers_to_test:
                self._opt_log(f"  Đo tốc độ phản hồi: {ip} ({name})...\n", "info")
                times = []
                for _ in range(3):
                    try:
                        t0 = time.perf_counter()
                        subprocess.run(['nslookup', 'google.com', ip], capture_output=True, text=True, timeout=15, creationflags=0x08000000 if os.name == 'nt' else 0)
                        elapsed = (time.perf_counter() - t0) * 1000
                        times.append(elapsed)
                    except Exception:
                        times.append(999.0)
                median_time = sorted(times)[1]
                dns_results.append({'ip': ip, 'name': name, 'time': median_time, 'backup': backup})
            
            dns_results.sort(key=lambda x: x['time'])
            
            self._opt_log("🏁 Kết quả đo DNS:\n", "accent")
            for idx, r in enumerate(dns_results):
                suffix = "  ← NHANH NHẤT" if idx == 0 else ""
                self._opt_log(f"  {r['ip']:<15} ({r['name']:<15})  →  {r['time']:.1f}ms{suffix}\n", "ok" if idx == 0 else "info")
            best = dns_results[0]
            best_dns_ip = best['ip']
            best_dns_name = best['name']
            
            current_dns_info = next((r for r in dns_results if r['name'] == 'DNS Hiện tại'), None)
            is_current_dns_fastest = (best['name'] == 'DNS Hiện tại')
            
            if not is_current_dns_fastest and current_dns_info:
                diff_ms = current_dns_info['time'] - best['time']
                msg = f"DNS hiện tại ({current_dns_info['ip']}) chậm hơn {best['name']} ({best['ip']}) {diff_ms:.1f}ms.\n"
                msg += f"Đổi sang {best['ip']} để cải thiện tốc độ duyệt web?\n\n"
                msg += "Lưu ý: Thay đổi này có thể hoàn tác trong Network Settings."
                
                ans = self._ask_on_main('yesno', "Xác nhận đổi DNS", msg, parent=self.root)
                if ans:
                    try:
                        if system_os == 'Windows':
                            adapters = []
                            rc_ad, ad_out, _ = _run_cmd(['netsh', 'interface', 'show', 'interface'], timeout=10)
                            for line in ad_out.splitlines():
                                parts = line.split()
                                if len(parts) < 4:
                                    continue
                                # So khớp CỘT STATE (token chính xác) — check 'connected'
                                # trong nguyên dòng trước đây khớp cả "Disconnected"
                                cols = [w.lower() for w in parts[:2]]
                                if 'connected' in cols or 'đã kết nối' in cols:
                                    adapters.append(' '.join(parts[3:]))
                            if not adapters:
                                adapters = ['Wi-Fi', 'Ethernet']
                            adapter = adapters[0]
                            rc1, _, dns_err = _run_cmd(['netsh', 'interface', 'ip', 'set', 'dns', f'name="{adapter}"', 'static', best['ip']], timeout=15)
                            if rc1 == 0:
                                if best['backup']:
                                    rc2, _, dns_err2 = _run_cmd(['netsh', 'interface', 'ip', 'add', 'dns', f'name="{adapter}"', best['backup'], 'index=2'], timeout=15)
                                    if rc2 != 0:
                                        self._opt_log(f"⚠️ Đã đổi DNS chính, nhưng thêm DNS phụ thất bại: {(dns_err2 or '').strip()}\n", "warn")
                                changed_dns = True
                                self._opt_log(f"✅ Đã đổi DNS → {best['ip']} ({best['name']}) + {best['backup'] or 'backup'} (backup)\n", "ok")
                            else:
                                # returncode bị bỏ qua trước đây → báo thành công giả
                                self._opt_log(f"⚠️ netsh đổi DNS thất bại trên adapter '{adapter}': {(dns_err or '').strip()}\n", "danger")
                        elif system_os == 'Darwin':
                            subprocess.run(['networksetup', '-setdnsservers', 'Wi-Fi', best['ip'], best['backup'] or '8.8.8.8'], timeout=15, capture_output=True)
                            changed_dns = True
                            self._opt_log(f"✅ Đã đổi DNS → {best['ip']} ({best['name']}) + {best['backup'] or 'backup'} (backup)\n", "ok")
                        else:
                            with open('/etc/resolv.conf', 'w') as f:
                                f.write(f"nameserver {best['ip']}\n")
                                if best['backup']:
                                    f.write(f"nameserver {best['backup']}\n")
                            changed_dns = True
                            self._opt_log(f"✅ Đã đổi DNS → {best['ip']} ({best['name']}) + {best['backup'] or 'backup'} (backup)\n", "ok")
                    except Exception as e:
                        self._opt_log(f"⚠️ Lỗi cấu hình DNS: {e}\n", "danger")
                else:
                    self._opt_log("ℹ️ Bỏ qua đổi DNS.\n", "info")
            else:
                self._opt_log("✅ DNS hiện tại của hệ thống đã là nhanh nhất.\n", "ok")
        else:
            self._opt_log("ℹ️ Tốc độ DNS ổn định, không cần tối ưu.\n", "info")
            
        self._update_opt_progress(100.0, "Giai đoạn 6/6: Phục hồi ứng dụng & Đo tốc độ SAU")
        
        self._opt_log("🔄 Đang phục hồi các tiến trình nền đã tạm dừng...\n", "info")
        for pid in suspended_pids:
            try:
                psutil.Process(pid).resume()
                self._opt_log(f"▶️ Đã resume: PID {pid}\n", "ok")
            except Exception as e:
                self._opt_log(f"⚠️ Lỗi resume PID {pid}: {e}\n", "danger")
            finally:
                self._susp_rescue.discard(pid)
                
        self._opt_log("⏳ Bắt đầu đo tốc độ mạng sau tối ưu (Verify)...\n", "info")
        download_after, upload_after, ping_after = before['download'], before['upload'], before['ping']
        dl_pct = ul_pct = ping_pct = 0.0
        try:
            # Đo tốc độ là network call dễ hỏng (và `speedtest` có thể không import được
            # nếu cài đặt lúc nãy thất bại) — tách riêng để bảng KẾT QUẢ + danh sách
            # thay đổi bên dưới LUÔN luôn được in ra
            try:
                import speedtest
                st = speedtest.Speedtest()
                st.get_best_server()
                download_after = st.download() / 1_000_000
                upload_after = st.upload() / 1_000_000
                ping_after = st.results.ping

                dl_pct = ((download_after - before['download']) / before['download'] * 100) if before['download'] > 0 else 0.0
                ul_pct = ((upload_after - before['upload']) / before['upload'] * 100) if before['upload'] > 0 else 0.0
                ping_pct = ((ping_after - before['ping']) / before['ping'] * 100) if before['ping'] > 0 else 0.0
            except Exception as e:
                self._opt_log(f"⚠️ Không đo được tốc độ SAU tối ưu: {e}\n", "danger")

            dl_sign = "+" if dl_pct >= 0 else ""
            ul_sign = "+" if ul_pct >= 0 else ""
            ping_sign = "+" if ping_pct >= 0 else ""
            
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
            self._opt_log("📊 KẾT QUẢ TỐI ƯU\n", "header")
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
            self._opt_log(f"             {'TRƯỚC':<10} {'SAU':<10} {'THAY ĐỔI':<12}\n", "accent")
            self._opt_log(f"Download:   {before['download']:.1f} Mbps  {download_after:.1f} Mbps  {dl_sign}{dl_pct:.1f}% {'✅' if dl_pct > 0 else '❌'}\n", "ok" if dl_pct > 0 else "info")
            self._opt_log(f"Upload:     {before['upload']:.1f} Mbps  {upload_after:.1f} Mbps  {ul_sign}{ul_pct:.1f}% {'✅' if ul_pct > 0 else '❌'}\n", "ok" if ul_pct > 0 else "info")
            self._opt_log(f"Ping:       {before['ping']:.0f} ms      {ping_after:.0f} ms      {ping_sign}{ping_pct:.1f}% {'✅' if ping_pct < 0 else '❌'}\n", "ok" if ping_pct < 0 else "info")
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
            
            self._opt_log("🎯 Các thay đổi đã thực hiện:\n", "accent")
            has_changes = False
            if len(deleted_wifi) > 0:
                self._opt_log(f"  • Xoá {len(deleted_wifi)} Wi-Fi profiles cũ\n", "ok")
                has_changes = True
            if len(suspended_pids) > 0:
                self._opt_log(f"  • Tạm dừng {len(suspended_pids)} tiến trình nền\n", "ok")
                has_changes = True
            if changed_dns:
                self._opt_log(f"  • Đổi DNS → {best_dns_name} ({best_dns_ip})\n", "ok")
                has_changes = True
            if not has_changes:
                self._opt_log("  • Không cần thực hiện thay đổi nào (hệ thống đã ở trạng thái tối ưu sẵn)\n", "info")
            self._opt_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n", "accent3")
            
            improved_dl = dl_pct >= 5.0
            improved_ul = ul_pct >= 5.0
            improved_ping = ping_pct <= -5.0
            
            if not (improved_dl or improved_ul or improved_ping):
                self._opt_log("ℹ️ Tốc độ thay đổi không đáng kể.\n", "info")
                self._opt_log("   Vấn đề có thể nằm ở: phần cứng router, chất lượng đường truyền ISP,\n", "info")
                self._opt_log("   hoặc vị trí Wi-Fi. Cân nhắc liên hệ nhà mạng hoặc đổi vị trí router.\n", "info")
                
        except Exception as e:
            self._opt_log(f"⚠️ Lỗi phân tích kết quả verify: {e}\n", "danger")

    # ══════════════════════════════════════════════════════
    # TAB 5 — CÀI APP
    # ══════════════════════════════════════════════════════
    # ══════════════════════════════════════════════════════
    # TAB 6 — LOCAL (4 sub-tabs)
    # ══════════════════════════════════════════════════════
    # TAB 6 — LOCAL (4 sub-tabs)
    # ══════════════════════════════════════════════════════
    VENDOR_TOOLS = {
        'lenovo': [
            {'name': 'Lenovo Vantage',   'uri': 'lenovo-vantage:', 'exe': None},
            {'name': 'Lenovo System Update', 'uri': None,
             'exe': r'C:\Program Files (x86)\Lenovo\System Update\tvsuKernel.exe'},
        ],
        'dell': [
            {'name': 'Dell Update',       'uri': None,
             'exe': r'C:\Program Files\Dell\UpdateService\ServiceShell.exe'},
            {'name': 'Dell SupportAssist','uri': None,
             'exe': r'C:\Program Files\Dell\SARemediation\agent\SupportAssistAgent.exe'},
        ],
        'hp': [
            {'name': 'HP Support Assistant', 'uri': None,
             'exe': r'C:\Program Files (x86)\HP\HP Support Framework\HPSF.exe'},
        ],
        'asus': [
            {'name': 'MyASUS', 'uri': None,
             'exe': r'C:\Program Files\ASUS\MyASUS\MyASUS.exe'},
        ],
        'acer': [
            {'name': 'Acer Care Center', 'uri': None,
             'exe': r'C:\Program Files\Acer\Care Center\CareCenterService.exe'},
        ],
        'msi': [
            {'name': 'MSI Center', 'uri': None,
             'exe': r'C:\Program Files\MSI\MSI Center\MSICenter.exe'},
        ],
        'gigabyte': [
            {'name': 'Gigabyte App Center', 'uri': None,
             'exe': r'C:\Program Files (x86)\GIGABYTE\AppCenter\RunApp.exe'},
        ],
        'microsoft': [
            {'name': 'Windows Update', 'uri': 'ms-settings:windowsupdate', 'exe': None},
        ],
    }
    SUPPORT_URLS = {
        'lenovo':    'https://pcsupport.lenovo.com/products/search?q=',
        'dell':      'https://www.dell.com/support/home',
        'hp':        'https://support.hp.com/vn-en/drivers',
        'asus':      'https://www.asus.com/support/download-center/',
        'acer':      'https://www.acer.com/ac/en/US/content/support',
        'msi':       'https://www.msi.com/support',
        'gigabyte':  'https://www.gigabyte.com/Support',
        'microsoft': 'https://support.microsoft.com',
    }
    ERROR_CODES = {
        1:  'Lỗi chung - thiếu tài nguyên',
        3:  'Driver bị trả về - không tương thích',
        10: 'Không khởi động được',
        12: 'Xung đột tài nguyên',
        14: 'Cần khởi động lại',
        16: 'Cài đặt không đầy đủ',
        18: 'Cần cài lại driver',
        19: 'Registry hỏng',
        21: 'Đang chờ xóa',
        22: 'Đã bị tắt',
        24: 'Thiết bị không tồn tại',
        28: 'Cài đặt thất bại',
        43: 'Lỗi sau khi khởi động',
    }
    def _build_act_tab(self):
        self.tab_act_frame = tk.Frame(self.root, bg=BG)
        # Chỉ còn 1 nội dung là Kích Hoạt → bỏ hẳn thanh sub-tab
        # (đã xóa Auto Click / Ẩn App / Gỡ App)
        self.sub_frame_2 = tk.Frame(self.tab_act_frame, bg=BG)
        self._build_sub_activate()
        self.sub_frame_2.pack(fill='both', expand=True)

    def _build_sub_activate(self):
        p = self.sub_frame_2
        # Status cards
        sc = tk.Frame(p, bg=PANEL, padx=10, pady=6)
        sc.pack(fill='x', padx=8, pady=(8, 4))
        wr = tk.Frame(sc, bg=PANEL)
        wr.pack(fill='x', pady=2)
        tk.Label(wr, text='🪟 Windows:', bg=PANEL, fg=TXT,
            font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_win_status = tk.Label(wr, text='Đang kiểm tra...',
            bg=PANEL, fg=MUTED, font=('Segoe UI', 8))
        self.lbl_win_status.pack(side='left', padx=8)
        or_ = tk.Frame(sc, bg=PANEL)
        or_.pack(fill='x', pady=2)
        tk.Label(or_, text='📦 Office:', bg=PANEL, fg=TXT,
            font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_off_status = tk.Label(or_, text='Đang kiểm tra...',
            bg=PANEL, fg=MUTED, font=('Segoe UI', 8))
        self.lbl_off_status.pack(side='left', padx=8)
        # Action cards
        def make_card(parent, label, btn_text, cmd, btn_color=ACCENT3, fg_color=BG):
            card = tk.Frame(parent, bg=CARD, padx=8, pady=4)
            card.pack(fill='x', padx=8, pady=2)
            tk.Label(card, text=label, bg=CARD, fg=MUTED,
                font=('Segoe UI', 8)).pack(anchor='w')
            tk.Button(card, text=btn_text, command=cmd,
                bg=btn_color, fg=fg_color,
                font=('Segoe UI', 9, 'bold'),
                relief='flat', cursor='hand2').pack(
                fill='x', ipady=5, pady=2)
        make_card(p, 'Kích hoạt bản quyền vĩnh viễn qua HWID',
            '  KÍCH HOẠT WINDOWS  🚀',
            self._activate_windows, ACCENT, 'white')
        make_card(p, 'Kích hoạt Office 365/2019/2021/2024 qua Ohook',
            '  KÍCH HOẠT OFFICE   🚀',
            self._activate_office, ACCENT4, BG)
        make_card(p, 'Kích hoạt đồng thời Windows + Office',
            '  KÍCH HOẠT CẢ HAI  🚀',
            self._activate_both, ACCENT2, BG)
        # Office download — 1 nút mỗi bản, sinh từ OFFICE_SOURCES
        dl_card = tk.Frame(p, bg=CARD, padx=8, pady=4)
        dl_card.pack(fill='x', padx=8, pady=2)
        for _key, _src in self.OFFICE_SOURCES.items():
            tk.Label(dl_card, text=_src['desc'], bg=CARD, fg=MUTED,
                font=('Segoe UI', 8)).pack(anchor='w')
            _btn = tk.Button(dl_card, text=_src['label'],
                command=lambda k=_key: self._download_office(k),
                bg=INPUT, fg=MUTED,
                font=('Segoe UI', 9, 'bold'),
                relief='flat', cursor='hand2')
            _btn.pack(fill='x', ipady=4, pady=(0, 4))
            # Nút ProPlus giữ vai trò cũ: _activate_office chỉnh màu nó khi chưa cài
            if _key == 'proplus':
                self.btn_dl_office = _btn
        # Progress
        self.act_canvas = tk.Canvas(p, bg='#1a2638', height=6,
            highlightthickness=0)
        self.act_canvas.pack(fill='x', padx=8, pady=2)
        self.lbl_act_pct = tk.Label(p, text='',
            bg=BG, fg=MUTED, font=('Segoe UI', 8))
        self.lbl_act_pct.pack(anchor='e', padx=10)
        # Log
        self.act_log = st.ScrolledText(p, bg='#0f1522', fg=MUTED,
            font=('Courier New', 8), relief='flat',
            wrap='word', height=3)
        self.act_log.pack(fill='x', padx=8, pady=(0, 6))
        for t, c in [('n', MUTED), ('ok', ACCENT2),
                     ('warn', WARN), ('danger', DANGER),
                     ('cyber', ACCENT3)]:
            self.act_log.tag_config(t, foreground=c)
    def _log_activ(self, text, tag='n'):
        if threading.current_thread() is not threading.main_thread():
            self._post(self._log_activ, text, tag)
            return
        self.act_log.configure(state='normal')
        self.act_log.insert('end', text + '\n', tag)
        self.act_log.see('end')
        self.act_log.configure(state='disabled')
    def _update_act_progress(self, pct):
        try:
            self.act_canvas.delete('bar')
            w = self.act_canvas.winfo_width()
            h = self.act_canvas.winfo_height()
            fw = int(pct / 100.0 * w)
            self.act_canvas.create_rectangle(
                0, 0, fw, h, fill=ACCENT3, outline='', tags='bar')
            self.lbl_act_pct.config(text=f'{pct:.0f}%')
        except Exception:
            pass
    def _detect_activation_status(self):
        # Windows
        try:
            rc, out, _ = _run_cmd(['powershell', '-NoProfile', '-Command',
                '(Get-WmiObject SoftwareLicensingProduct '
                '-Filter "LicenseStatus=1 AND '
                'PartialProductKey IS NOT NULL").Name'],
                timeout=15)
            if rc == 0 and out.strip():
                self._post(self.lbl_win_status.config,
                    {'text': '✅ Đã kích hoạt', 'fg': ACCENT2})
            else:
                self._post(self.lbl_win_status.config,
                    {'text': '❌ Chưa kích hoạt', 'fg': DANGER})
        except Exception:
            self._post(self.lbl_win_status.config,
                {'text': '⚠ Không kiểm tra được', 'fg': WARN})
        # Office
        office_paths = [
            r'C:\Program Files\Microsoft Office',
            r'C:\Program Files (x86)\Microsoft Office',
            r'C:\Program Files\Microsoft Office 15',
        ]
        off_found = any(os.path.exists(p) for p in office_paths)
        if off_found:
            self._post(self.lbl_off_status.config,
                {'text': '✅ Đã cài Office', 'fg': ACCENT2})
        else:
            self._post(self.lbl_off_status.config,
                {'text': '❌ Chưa cài Office', 'fg': MUTED})
    def _run_mas(self, switch_primary, switch_fallback=None):
        fallback_sw = switch_fallback or switch_primary
        ps_cmds = [
            f'& ([ScriptBlock]::Create((irm https://get.activated.win))) {switch_primary}',
            f'& ([ScriptBlock]::Create((curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String))) {fallback_sw}',
            # Phải kèm switch — versions cũ bỏ switch → chạy menu MAS ẩn, vô vọng
            f'& ([ScriptBlock]::Create((irm https://get.activated.win))) {switch_primary}',
        ]
        for ps_cmd in ps_cmds:
            self._log_activ(f'⚡ Chạy: powershell {ps_cmd[:45]}...', 'cyber')
            try:
                proc = subprocess.Popen(
                    ['powershell', '-NoProfile',
                     '-ExecutionPolicy', 'Bypass',
                     '-Command', ps_cmd],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace',
                    creationflags=subprocess.CREATE_NO_WINDOW)
                # communicate(timeout) thay vì `for line in proc.stdout`: đọc stream
                # trực tiếp KHÔNG BAO GIỜ timeout (kẹt tiến trình con kế thừa pipe),
                # và nếu timeout thì tiến trình cũng không bị kill → chạy 2 MAS song song
                try:
                    out, _ = proc.communicate(timeout=120)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    try:
                        proc.communicate(timeout=10)
                    except Exception:
                        pass
                    self._log_activ('⚠ Hết 120s, đã hủy tiến trình MAS.', 'warn')
                    continue
                for ln in (out or '').splitlines():
                    if ln.strip():
                        self._log_activ(ln.strip(), 'n')
                if proc.returncode == 0:
                    self._log_activ('✅ Hoàn tất!', 'ok')
                    threading.Thread(
                        target=self._detect_activation_status,
                        daemon=True).start()
                    return
            except Exception as e:
                self._log_activ(f'⚠ Lỗi: {e}', 'warn')
        self._log_activ('❌ Thất bại sau 3 lần thử.', 'danger')
    def _activate_windows(self):
        if ctypes.windll.shell32.IsUserAnAdmin() == 0:
            messagebox.showwarning('Quyền Admin',
                'Cần chạy Toolbox với quyền Administrator!',
                parent=self.root)
            return
        self._log_activ('🚀 Kích hoạt Windows (HWID)...', 'cyber')
        threading.Thread(
            target=self._run_mas, args=('/HWID',), daemon=True).start()
    def _activate_office(self):
        office_paths = [
            r'C:\Program Files\Microsoft Office',
            r'C:\Program Files (x86)\Microsoft Office',
        ]
        if not any(os.path.exists(p) for p in office_paths):
            messagebox.showwarning('Chưa cài Office',
                'Vui lòng cài Office trước!\nNhấn một trong các nút "TẢI OFFICE" bên dưới.',
                parent=self.root)
            self.btn_dl_office.config(bg=ACCENT4, fg=BG)
            return
        self._log_activ('🚀 Kích hoạt Office (Ohook)...', 'cyber')
        threading.Thread(
            target=self._run_mas, args=('/Ohook',), daemon=True).start()
    def _activate_both(self):
        if ctypes.windll.shell32.IsUserAnAdmin() == 0:
            messagebox.showwarning('Quyền Admin',
                'Cần chạy Toolbox với quyền Administrator!',
                parent=self.root)
            return
        self._log_activ('🚀 Kích hoạt Windows + Office...', 'cyber')
        threading.Thread(
            target=self._run_mas,
            args=('/HWID /Ohook',), daemon=True).start()
    def _download_office(self, kind='proplus'):
        """Tải bộ cài Office về Desktop — nguồn lấy từ OFFICE_SOURCES."""
        src = self.OFFICE_SOURCES.get(kind) or self.OFFICE_SOURCES['proplus']
        url = src['url']
        dest = os.path.join(os.path.expanduser('~'), 'Desktop', src['file'])
        # File đã tải rồi: hỏi mở hay tải lại, tránh tải nhầm lại vài GB
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            if self._ask_on_main('yesno', 'Đã có file',
                    f'{src["name"]} đã có tại Desktop:\n{dest}\n\n'
                    'Yes = mở ngay, No = tải lại từ đầu', parent=self.root):
                try:
                    os.startfile(dest)
                except Exception as e:
                    self._log_activ(f'❌ Không mở được file: {e}', 'danger')
                return
            try:
                os.remove(dest)
            except OSError:
                pass
        self._log_activ(f'📥 Đang tải {src["name"]} về Desktop...', 'warn')
        self._update_act_progress(0)
        def worker():
            try:
                def hook(count, block, total):
                    if total > 0:
                        pct = min(count * block * 100 // total, 99)
                        self._post(self._update_act_progress, pct)
                urllib.request.urlretrieve(url, dest, hook)
                self._post(self._update_act_progress, 100)
                self._log_activ(
                    f'✅ Tải xong: {dest}', 'ok')
                if dest.endswith('.img'):
                    self._log_activ(
                        '  → Nhấp đúp file .img để Windows gắn ổ ảo, rồi chạy Setup.exe.', 'cyber')
                # Hỏi trên MAIN thread (worker thread không được tạo modal dialog Tk)
                if self._ask_on_main('yesno', 'Mở file?',
                                     'Tải xong! Mở bộ cài Office ngay?', parent=self.root):
                    os.startfile(dest)
            except Exception as e:
                self._log_activ(f'❌ Lỗi tải: {e}', 'danger')
        threading.Thread(target=worker, daemon=True).start()
    def _track_mouse(self):
        def _loop():
            while True:
                try:
                    if HAVE_WIN32:
                        x, y = win32api.GetCursorPos()
                        self._post(self.v_mouse.set, f'X:{x} Y:{y}')
                except Exception as e:
                    pass
                time.sleep(0.05)
        # Khởi động sau khi mainloop đã chạy (sau 300ms) — nếu start ngay trong
        # __init__, luồng này gọi _post khi mainloop chưa lên → retry vô ích
        self.root.after(300, lambda: threading.Thread(target=_loop, daemon=True).start())
    def _start_silent_downloads(self):
        def _worker():
            paths = get_toolbox_paths()
            if not os.path.exists(paths['yt_dlp']):
                try:
                    os.makedirs(os.path.dirname(paths['yt_dlp']), exist_ok=True)
                    urllib.request.urlretrieve(
                        'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe', 
                        paths['yt_dlp']
                    )
                except Exception as e:
                    pass
        threading.Thread(target=_worker, daemon=True).start()
    def _do_exit(self):
        if hasattr(self, '_cleanup_temp'):
            self._cleanup_temp()
        self.root.destroy()
        import os
        os._exit(0)

    def _cleanup_temp(self):
        tmp_root = os.path.join(os.environ.get('TEMP', ''), 'ToolboxInst')
        if os.path.exists(tmp_root):
            try:
                shutil.rmtree(tmp_root, ignore_errors=True)
            except Exception as e:
                pass
    def _footer(self):
        foot = tk.Frame(self.root, bg=PANEL, height=22)
        foot.pack(fill='x', side='bottom')
        foot.pack_propagate(False)
        lbl_m = tk.Label(foot, textvariable=self.v_mouse, bg=PANEL, fg=MUTED, font=('Courier New', 8))
        lbl_m.pack(side='left', padx=8)
        f_inner = tk.Frame(foot, bg=PANEL)
        f_inner.pack(side='right', padx=8)
        lbl_by = tk.Label(f_inner, text='Build by ', bg=PANEL, fg=MUTED, font=('Segoe UI', 8, 'italic'))
        lbl_by.pack(side='left')
        self.lbl_vk = tk.Label(f_inner, text='Vũ Khuê', bg=PANEL, fg=ACCENT2, font=('Segoe UI', 10, 'bold'))
        self.lbl_vk.pack(side='left')
        self.vk_colors = ['#f87171', '#34d399', '#fbbf24', '#22d3ee', '#f472b6', '#c084fc', '#a78bfa']
        self.vk_color_idx = 0
        self._blink_vk()
    def _blink_vk(self):
        try:
            self.vk_color_idx = (self.vk_color_idx + 1) % len(self.vk_colors)
            color = self.vk_colors[self.vk_color_idx]
            self.lbl_vk.config(fg=color)
        except Exception:
            pass
        self.root.after(180, self._blink_vk)
# ── Tab Tải App: nguồn cài chính thức của 16 phần mềm (đã probe HEAD 200) ──
# kind='direct' → url cố định của hãng
# kind='github' → lấy file mới nhất từ GitHub API (repo hãng tự phát hành)
# kind='page'   → tải trang chính thức rồi regex ra link file
# Thêm app = thêm 1 dict vào danh sách này, giao diện tự hiện thêm dòng tick.
BROWSER_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
              '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36')
APP_SOURCES = [
    # 🌐 Trình duyệt
    {'key': 'chrome', 'name': 'Google Chrome', 'group': 'browser', 'kind': 'direct',
     'file': 'ChromeSetup.msi',
     'url': 'https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi'},
    {'key': 'firefox', 'name': 'Firefox', 'group': 'browser', 'kind': 'direct',
     'file': 'FirefoxSetup.exe',
     'url': 'https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=vi'},
    {'key': 'brave', 'name': 'Brave', 'group': 'browser', 'kind': 'direct',
     'file': 'BraveBrowserSetup.exe',
     'url': 'https://laptop-updates.brave.com/latest/win/x64'},
    {'key': 'coccoc', 'name': 'Cốc Cốc', 'group': 'browser', 'kind': 'direct',
     'file': 'CocCocSetup.exe',
     'url': 'https://files.coccoc.com/browser/x64/coccoc_standalone_vi.exe'},
    # 💬 Nhắn tin
    {'key': 'zalo', 'name': 'Zalo', 'group': 'chat', 'kind': 'direct',
     'file': 'ZaloSetup.exe', 'headers': {'User-Agent': BROWSER_UA},
     'url': 'https://zalo.me/download/zalo-pc'},
    {'key': 'telegram', 'name': 'Telegram', 'group': 'chat', 'kind': 'github',
     'repo': 'telegramdesktop/tdesktop', 'asset': r'^td-setup-win-x64-.*\.exe$',
     'file': 'TelegramSetup.exe'},
    {'key': 'discord', 'name': 'Discord', 'group': 'chat', 'kind': 'direct',
     'file': 'DiscordSetup.exe',
     'url': 'https://discord.com/api/download?platform=win'},
    # ⌨ Bộ gõ tiếng Việt
    {'key': 'evkey', 'name': 'EVKey (bản portable)', 'group': 'input', 'kind': 'github',
     'repo': 'lamquangminh/EVKey', 'asset': r'^EVKey\.zip$', 'file': 'EVKey.zip',
     'zip': True, 'run': 'x64/EVKey64.exe'},
    {'key': 'unikey', 'name': 'UniKey', 'group': 'input', 'kind': 'direct',
     'file': 'UniKeySetup.exe',
     'url': 'https://sourceforge.net/projects/unikey/files/latest/download'},
    # 🛠 Tiện ích
    {'key': 'ultraviewer', 'name': 'UltraViewer', 'group': 'tools', 'kind': 'page',
     'page': 'https://www.ultraviewer.net/en/download.html',
     'pattern': r'href="(UltraViewer[^"]+\.exe)"', 'file': 'UltraViewerSetup.exe'},
    {'key': 'anydesk', 'name': 'AnyDesk', 'group': 'tools', 'kind': 'direct',
     'file': 'AnyDesk.exe', 'url': 'https://download.anydesk.com/AnyDesk.exe'},
    {'key': '7zip', 'name': '7-Zip', 'group': 'tools', 'kind': 'github',
     'repo': 'ip7z/7zip', 'asset': r'^7z\d+-x64\.exe$', 'file': '7-Zip-x64.exe'},
    {'key': 'vlc', 'name': 'VLC media player', 'group': 'tools', 'kind': 'page',
     'page': 'https://www.videolan.org/vlc/download-windows.html',
     'pattern': r'href="(//get\.videolan\.org/vlc/[\d.]+/win64/vlc-[\d.]+-win64\.exe)"',
     'file': 'VLC-x64.exe'},
    {'key': 'notepadpp', 'name': 'Notepad++', 'group': 'tools', 'kind': 'github',
     'repo': 'notepad-plus-plus/notepad-plus-plus', 'asset': r'\.Installer\.x64\.exe$',
     'file': 'NotepadPlusPlus-x64.exe'},
    {'key': 'vscode', 'name': 'Visual Studio Code', 'group': 'tools', 'kind': 'direct',
     'file': 'VSCodeUserSetup.exe',
     'url': 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user'},
    {'key': 'zoom', 'name': 'Zoom Workplace', 'group': 'tools', 'kind': 'direct',
     'file': 'ZoomInstallerFull.exe',
     'url': 'https://zoom.us/client/latest/ZoomInstallerFull.exe'},
]
GROUP_LABELS = {'browser': '🌐 Trình duyệt', 'chat': '💬 Nhắn tin',
                'input': '⌨ Bộ gõ tiếng Việt', 'tools': '🛠 Tiện ích'}


class AppSourceError(Exception):
    """Không tìm được link tải chính thức — log lý do rồi bỏ qua app này."""


def _http_get_text(url, headers=None):
    req = urllib.request.Request(url, headers=headers or {'User-Agent': BROWSER_UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8', 'replace')


def resolve_app_url(app, http_get=None):
    """Trả URL tải file cài của app, hoặc raise AppSourceError để bỏ qua app."""
    if app.get('kind') == 'direct':
        return app['url']
    http_get = http_get or _http_get_text
    kind = app.get('kind')
    if kind == 'github':
        try:
            data = json.loads(http_get(
                'https://api.github.com/repos/%s/releases/latest' % app['repo']))
        except Exception as e:
            raise AppSourceError('không gọi được GitHub API: %s' % e)
        for asset in data.get('assets') or []:
            if re.search(app['asset'], asset.get('name', '')):
                return asset['browser_download_url']
        raise AppSourceError('bản phát hành mới nhất không có file phù hợp')
    if kind == 'page':
        try:
            html = http_get(app['page'])
        except Exception as e:
            raise AppSourceError('không tải được trang %s: %s' % (app['page'], e))
        m = re.search(app['pattern'], html)
        if not m:
            raise AppSourceError('trang %s không còn chứa link tải' % app['page'])
        return urllib.parse.urljoin(app['page'], m.group(1))
    raise AppSourceError('kind không hỗ trợ: %s' % kind)


def apps_download_dir():
    """Thư mục riêng chứa file cài: Desktop\\Toolbox Apps."""
    return os.path.join(os.path.expanduser('~'), 'Desktop', 'Toolbox Apps')


def _fetch_to_file(url, headers, dest, on_chunk):
    """Tải streaming về dest. on_chunk(số_byte_vừa_đọc, tổng_byte_hoặc_None)."""
    req = urllib.request.Request(url, headers=headers or {'User-Agent': BROWSER_UA})
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.headers.get('Content-Length')
        total = int(raw) if raw and raw.isdigit() else None
        n = 0
        with open(dest, 'wb') as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                n += len(chunk)
                on_chunk(len(chunk), total)
    return n


def download_apps(apps, dest_dir, progress_cb=None, log_cb=None, fetch=None, resolve=None):
    """Tải tuần tự các app đã tick về dest_dir.

    Lỗi 1 app (không tìm ra link / mạng đứt) → ghi log, bỏ qua, vẫn tải app kế.
    Trả {'ok': [key], 'failed': [(key, lý_do)], 'paths': {key: đường_dẫn}}.
    """
    fetch = fetch or _fetch_to_file
    resolve = resolve or resolve_app_url
    log = log_cb or (lambda text, tag='n': None)
    prog = progress_cb or (lambda pct, msg='': None)
    res = {'ok': [], 'failed': [], 'paths': {}}
    prog(0.0, '')
    if not apps:
        prog(100.0, '')
        return res
    os.makedirs(dest_dir, exist_ok=True)
    done = 0
    total_apps = float(len(apps))
    for app in apps:
        name = app['name']
        dest = os.path.join(dest_dir, app['file'])
        # File đã tải đủ từ trước → dùng lại, không tốn API/kiểm tra link
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            log('✓ %s — đã có %s, dùng lại' % (name, app['file']), 'ok')
            res['ok'].append(app['key'])
            res['paths'][app['key']] = dest
            done += 1
            prog(done / total_apps * 100.0, name)
            continue
        try:
            url = resolve(app)
        except AppSourceError as e:
            log('❌ %s — %s' % (name, e), 'danger')
            res['failed'].append((app['key'], str(e)))
            continue
        state = {'n': 0}

        def on_chunk(nbytes, total, _name=name, _state=state):
            _state['n'] += nbytes
            frac = min(1.0, _state['n'] / float(total)) if total else 0.0
            prog((done + frac) / total_apps * 100.0, _name)

        log('📥 Đang tải %s…' % name, 'warn')
        try:
            fetch(url, app.get('headers'), dest, on_chunk)
        except Exception as e:
            # File dở phải xoá: nếu không lượt sau sẽ thấy "đã có file" và dùng lại bản hỏng
            try:
                if os.path.exists(dest):
                    os.remove(dest)
            except OSError:
                pass
            log('❌ %s — lỗi tải: %s' % (name, e), 'danger')
            res['failed'].append((app['key'], str(e)))
            continue
        log('✓ %s — tải xong (%s)' % (name, app['file']), 'ok')
        res['ok'].append(app['key'])
        res['paths'][app['key']] = dest
        done += 1
        prog(done / total_apps * 100.0, name)
    prog(100.0, '')
    return res


# ── Tab "🛠 Sau cài Win": port từ WinCleanTool (F:\one.bat) ─────────────────
# Mọi lệnh đi qua runner injectable được (rc, out) → test không đụng registry thật.
POST_TWEAKS = [
    {'id': 'hiber', 'name': 'Tắt Hibernate & Fast Startup', 'default': True},
    {'id': 'efi', 'name': 'Ẩn phân vùng EFI/Recovery', 'default': True},
    {'id': 'visfx', 'name': 'Mở Visual Effects', 'default': False},
    {'id': 'cortana', 'name': 'Tắt Cortana & gợi ý/quảng cáo', 'default': True},
    {'id': 'menu', 'name': 'Tối ưu menu & tắt máy', 'default': True},
    {'id': 'bg', 'name': 'Tắt app chạy nền', 'default': True},
    {'id': 'msconfig', 'name': 'Mở msconfig', 'default': False},
]

# 8 khóa ContentDeliveryManager trong one.bat (gợi ý/quảng cáo/tips)
_CDM_KEYS = (
    'SubscribedContent-338389Enabled', 'SystemPaneSuggestionsEnabled',
    'SoftLandingEnabled', 'RotatingLockScreenOverlayEnabled',
    'SubscribedContent-310093Enabled', 'SubscribedContent-338388Enabled',
    'SubscribedContent-353694Enabled', 'SubscribedContent-353696Enabled',
)


def _reg_add(path, name, typ, value):
    return ['reg', 'add', path, '/v', name, '/t', typ, '/d', str(value), '/f']


def run_cmd(argv, timeout=60):
    """Chạy lệnh ẩn cửa sổ, trả (returncode, output gộp stdout+stderr)."""
    try:
        r = subprocess.run(argv, capture_output=True, text=True, timeout=timeout,
                           encoding='utf-8', errors='replace',
                           creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        return r.returncode, ((r.stdout or '') + (r.stderr or '')).strip()
    except subprocess.TimeoutExpired:
        return 1, 'Timeout %ss' % timeout
    except Exception as e:
        return 1, str(e)


def tweak_ops(tid):
    """Danh sách op của 1 tweak: ('run', argv) | ('open', exe) | ('efi',).
    KeyError nếu id lạ. Khớp từng lệnh với one.bat — chỉ bỏ ghi đôi HKU\\<SID>
    (plan A4: elevate cùng user thì HKCU == HKU\\<SID>)."""
    if tid == 'hiber':
        return [
            ('run', ['powercfg', '/hibernate', 'off']),
            ('run', _reg_add(r'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power',
                             'HiberbootEnabled', 'REG_DWORD', 0)),
        ]
    if tid == 'efi':
        return [('efi',)]
    if tid == 'visfx':
        return [('open', 'SystemPropertiesPerformance.exe')]
    if tid == 'cortana':
        ops = [('run', _reg_add(r'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search',
                                'AllowCortana', 'REG_DWORD', 0))]
        cdm = r'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
        ops += [('run', _reg_add(cdm, k, 'REG_DWORD', 0)) for k in _CDM_KEYS]
        return ops
    if tid == 'menu':
        desk = r'HKCU\Control Panel\Desktop'
        return ([('run', _reg_add(desk, n, 'REG_SZ', v))
                 for n, v in (('MenuShowDelay', '0'), ('WaitToKillAppTimeout', '2000'),
                              ('HungAppTimeout', '1000'), ('AutoEndTasks', '1'),
                              ('LowLevelHooksTimeout', '1000'))]
                + [('run', _reg_add(r'HKLM\SYSTEM\CurrentControlSet\Control',
                                    'WaitToKillServiceTimeout', 'REG_SZ', 5000))])
    if tid == 'bg':
        baa = r'HKCU\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications'
        pol = r'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy'
        return [('run', _reg_add(baa, 'GlobalUserDisabled', 'REG_DWORD', 1)),
                ('run', _reg_add(baa, 'LetAppsRunInBackground', 'REG_DWORD', 2)),
                # GPO chính thức "Force deny" (spec T6 yêu cầu HKLM Policies);
                # 2 key HKCU phía trên là kiểu one.bat/Settings
                ('run', _reg_add(pol, 'LetAppsRunInBackground', 'REG_DWORD', 2))]
    if tid == 'msconfig':
        return [('open', 'msconfig.exe')]
    raise KeyError(tid)


def run_tweak(tid, runner=None, log=None):
    """Chạy 1 tweak. Trả False nếu có lệnh rc != 0 — nhưng vẫn chạy hết op còn lại."""
    runner = runner or run_cmd
    log = log or (lambda text, tag='n': None)
    ok = True
    for kind, *payload in tweak_ops(tid):
        if kind == 'open':
            exe = payload[0]
            try:
                subprocess.Popen([exe])
                log('   ▶ Đã mở %s — tự cấu hình rồi đóng cửa sổ.' % exe, 'cyber')
            except Exception as e:
                ok = False
                log('   ❌ Không mở được %s: %s' % (exe, e), 'danger')
            continue
        if kind == 'efi':
            log('   ℹ Ổ EFI cần xác nhận — bấm "🔍 Quét ổ EFI".', 'n')
            continue
        argv = payload[0]
        rc, out = runner(argv, timeout=60)
        if rc == 0:
            log('   ✓ %s' % ' '.join(argv), 'ok')
        else:
            ok = False
            log('   ❌ rc=%s %s%s' % (rc, ' '.join(argv),
                                      (' — ' + out) if out else ''), 'danger')
    return ok


def list_efi_partitions(runner=None, sysdrive=None):
    """Quét phân vùng EFI/Recovery đang bị gán ổ letter (trừ ổ hệ thống).
    Trả (ok, items): ok=False nếu PowerShell lỗi/timeout — KHÁC với rỗng thật;
    items = list dict {'letter','type','disk','part'}."""
    runner = runner or run_cmd
    sysdrive = (sysdrive or os.environ.get('SystemDrive', 'C:')).rstrip(':').upper()
    script = (
        "Get-Partition -ErrorAction SilentlyContinue | Where-Object { "
        "$_.DriveLetter -and $_.DriveLetter -ne '%s' -and ( "
        "$_.GptType -eq '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' -or "
        "$_.GptType -eq '{de94bba4-06d1-4d40-a16a-bfd50179d6ac}' -or "
        "$_.MbrType -eq 39 ) } | ForEach-Object { "
        "'{0}|{1}|Disk {2}|Partition {3}' -f "
        "$_.DriveLetter,$_.Type,$_.DiskNumber,$_.PartitionNumber }"
    ) % sysdrive
    rc, out = runner(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                      '-Command', script], timeout=60)
    if rc != 0:
        return False, []
    items = []
    for line in out.splitlines():
        p = line.strip().split('|')
        if len(p) == 4 and len(p[0]) == 1 and p[0].isalpha():
            items.append({'letter': p[0].upper(), 'type': p[1],
                          'disk': p[2], 'part': p[3]})
    return True, items


def hide_partitions(items, runner=None, log=None):
    """Gỡ ổ letter của phân vùng EFI/Recovery (mountvol /D) — chỉ ẩn trên
    Explorer, KHÔNG đụng dữ liệu. Trả (n_ok, n_fail)."""
    runner = runner or run_cmd
    log = log or (lambda text, tag='n': None)
    n_ok = n_fail = 0
    for it in items:
        rc, out = runner(['mountvol', '%s:' % it['letter'], '/D'], timeout=30)
        if rc == 0:
            n_ok += 1
            log('   ✓ Đã ẩn ổ %s: (%s)' % (it['letter'], it['type']), 'ok')
        else:
            n_fail += 1
            log('   ❌ Không ẩn được ổ %s: %s'
                % (it['letter'], out or ('rc=%d' % rc)), 'danger')
    return n_ok, n_fail


def clean_update_cache(runner=None, log=None, isdir=None, windir=None):
    """Dọn cache Windows Update: stop dịch vụ → rd .old → ren → KHỞI ĐỘNG LẠI.

    Khác bản .bat: ren lỗi được báo FAIL thật (UPDATE_OK của .bat luôn =1 nên
    nhánh FAIL là dead code), và net start LUÔN chạy kể cả khi rename lỗi.
    Trả (ok, msg)."""
    runner = runner or run_cmd
    log = log or (lambda text, tag='n': None)
    isdir = isdir or os.path.isdir
    windir = windir or os.environ.get('windir', r'C:\Windows')
    ok = True
    failed = []
    try:
        log('■ Dừng dịch vụ Windows Update…', 'cyber')
        for svc in ('wuauserv', 'bits', 'cryptsvc'):
            rc, out = runner(['net', 'stop', svc, '/y'], timeout=60)
            if rc == 0:
                log('   ✓ Đã dừng %s' % svc, 'ok')
            else:
                log('   ⚠ net stop %s (có thể đã dừng): %s'
                    % (svc, out or ('rc=%d' % rc)), 'warn')
        log('■ Xóa/đổi tên thư mục cache…', 'cyber')
        folders = ((os.path.join(windir, 'SoftwareDistribution'), 'SoftwareDistribution.old'),
                   (os.path.join(windir, 'System32', 'catroot2'), 'catroot2.old'))
        for folder, oldname in folders:
            base = os.path.basename(folder)
            if not isdir(folder):
                log('   ℹ Không có %s — bỏ qua.' % base, 'n')
                continue
            if isdir(folder + '.old'):
                # Token RIÊNG cho cmd: quote nhúng trong 1 arg ("ren "..." ...") bị
                # cmd.exe parse sai → 'filename... syntax is incorrect' (bug thật)
                rc, out = runner(['cmd', '/c', 'rd', '/s', '/q', folder + '.old'],
                                 timeout=300)
                if rc != 0:
                    log('   ⚠ Không xóa được %s cũ: %s'
                        % (oldname, out or ('rc=%d' % rc)), 'warn')
            rc, out = runner(['cmd', '/c', 'ren', folder, oldname], timeout=120)
            if rc == 0:
                log('   ✓ Đã đổi tên %s → %s' % (base, oldname), 'ok')
            else:
                ok = False
                failed.append(base)
                log('   ❌ Đổi tên %s THẤT BẠI: %s'
                    % (base, out or ('rc=%d' % rc)), 'danger')
    finally:
        # Plan C3.4: KHỞI ĐỘNG LẠI luôn chạy — dù rename lỗi hay runner nổ
        # giữa chừng, services không được phép kẹt STOP.
        log('■ Khởi động lại dịch vụ…', 'cyber')
        for svc in ('cryptsvc', 'bits', 'wuauserv'):
            rc, out = runner(['net', 'start', svc], timeout=60)
            if rc == 0:
                log('   ✓ Đã khởi động %s' % svc, 'ok')
            else:
                log('   ⚠ net start %s: %s' % (svc, out or ('rc=%d' % rc)), 'warn')
    if ok:
        return True, 'Dọn Update Cache thành công.'
    return False, ('FAIL: không đổi tên được %s (xem log).'
                   % ', '.join(failed))


def get_toolbox_paths():
    exe_dir = os.path.dirname(sys.executable) if hasattr(sys, '_MEIPASS') else os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(exe_dir, 'Toolbox Downloads')
    lib_dir = os.path.join(exe_dir, 'important_library')
    os.makedirs(lib_dir, exist_ok=True)
    yt_dlp_path = os.path.join(lib_dir, 'yt-dlp.exe')
    
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    aria2c_bundled = os.path.join(base_path, 'resources', 'aria2c.exe')
    aria2c_path = aria2c_bundled if os.path.exists(aria2c_bundled) else os.path.join(lib_dir, 'aria2c.exe')
    
    ffmpeg_dir = os.path.join(lib_dir, 'ffmpeg')
    ffmpeg_exe = os.path.join(ffmpeg_dir, 'bin', 'ffmpeg.exe')
    return {'base_dir': exe_dir, 'yt_dlp': yt_dlp_path, 'aria2c': aria2c_path, 'ffmpeg_dir': ffmpeg_dir, 'ffmpeg_exe': ffmpeg_exe, 'output': output_dir}
def get_platform_folder(url):
    u = url.lower()
    if 'youtube.com' in u or 'youtu.be' in u:
        return 'YouTube'
    elif 'facebook.com' in u or 'fb.watch' in u:
        return 'Facebook'
    elif 'tiktok.com' in u:
        return 'TikTok'
    elif 'instagram.com' in u:
        return 'Instagram'
    elif 'twitter.com' in u or 'x.com' in u:
        return 'Twitter'
    elif 'soundcloud.com' in u:
        return 'SoundCloud'
    elif 'zingmp3.vn' in u:
        return 'ZingMP3'
    elif 'nhaccuatui.com' in u:
        return 'NhacCuaTui'
    elif 'capcut.com' in u:
        return 'CapCut'
    elif 'threads.net' in u:
        return 'Threads'
    elif 'pinterest.com' in u or 'pin.it' in u:
        return 'Pinterest'
    elif 'reddit.com' in u:
        return 'Reddit'
    elif 'twitch.tv' in u:
        return 'Twitch'
    elif 'vimeo.com' in u:
        return 'Vimeo'
    elif 'bilibili.com' in u or 'bili' in u:
        return 'Bilibili'
    else:
        return 'Other'
def parse_ytdlp_json(stdout):
    """Parse output JSON của `yt-dlp -j`.

    Port từ reclip-main/app.py:16-29 — với `--no-playlist` một số extractor
    vẫn in nhiều object, nên json.loads trực tiếp sẽ lỗi "Extra data".
    Trả về object hợp lệ đầu tiên.
    """
    for line in stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        return json.loads(line)
    raise ValueError("yt-dlp returned no data")


def sanitize_title(title):
    """Làm sạch tên file: bỏ ký tự cấm Windows, giới hạn 100 ký tự.

    Port từ reclip-main/app.py:78-81.
    """
    safe = "".join(c for c in (title or "") if c not in r'\/:*?"<>|').strip()[:100].strip()
    return safe


def build_quality_options(info):
    """Chọn format tốt nhất cho từng chiều cao, sắp giảm dần.

    Port từ reclip-main/app.py (best_by_height trong route /api/info).
    """
    best_by_height = {}
    for f in info.get('formats') or []:
        height = f.get('height')
        vcodec = (f.get('vcodec') or 'none').lower()
        if not height or vcodec == 'none':
            continue
        acodec = (f.get('acodec') or 'none').lower()
        # Điểm: 2 = H.264 (chạy được ở WMP/VLC/điện thoại), 1 = video-only
        # (khi ghép `id+bestaudio` với progressive sẽ ra 2 track audio)
        score = (2 if vcodec.startswith('avc') else 0) + (1 if acodec == 'none' else 0)
        cur = best_by_height.get(height)
        if cur is None:
            best_by_height[height] = (f, score)
            continue
        cur_f, cur_score = cur
        if score > cur_score or (score == cur_score
                                 and (f.get('tbr') or 0) > (cur_f.get('tbr') or 0)):
            best_by_height[height] = (f, score)
    options = [{'format_id': f['format_id'], 'label': f'{h}p', 'height': h}
               for h, (f, _) in best_by_height.items()]
    options.sort(key=lambda x: x['height'], reverse=True)
    return options


def probe_video_info(url, yt_dlp_path, timeout=60):
    """Truy vấn metadata + danh sách quality của một link.

    Port từ reclip-main/app.py route /api/info (`yt-dlp --no-playlist -j`).
    Nâng RuntimeError với dòng lỗi CUỐI của stderr (cách báo lỗi của reclip).
    """
    cmd = [yt_dlp_path, '--ignore-config', '--no-playlist', '-j', url]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8',
                           errors='replace', timeout=timeout,
                           creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
    except subprocess.TimeoutExpired:
        raise RuntimeError(f'Phân tích link quá {timeout} giây')
    if r.returncode != 0:
        err = (r.stderr or '').strip()
        raise RuntimeError(err.split('\n')[-1] if err else 'yt-dlp trả về lỗi không rõ')
    return parse_ytdlp_json(r.stdout)


def build_download_cmd(paths, url, output_template, mode, fmt_str, qual_str, has_ffmpeg,
                       aria2c_exists=False, format_id=None, extra_args=None):
    """Dựng lệnh yt-dlp cho tab Tải Nhạc/Video.

    Phần CHỌN ĐỊNH DẠNG lấy theo reclip-main/app.py (run_download):
      - audio  -> `-x --audio-format <fmt>`
      - video  -> `-f <id>+bestaudio/best --merge-output-format mp4`
    Phần TIẾN ĐỘ/thư mục/aria2c/ffmpeg giữ nguyên của toolbox.
    """
    cmd = [paths['yt_dlp'], '--ignore-config', '--no-playlist', '--newline', '--no-warnings',
           '--windows-filenames', '--no-mtime', '--retries', '10', '--fragment-retries', '10',
           '--file-access-retries', '30', '-o', output_template]
    if aria2c_exists:
        cmd += ['--downloader', paths['aria2c'], '--downloader-args', 'aria2c:-x 16 -s 16 -j 16 -k 1M']
    else:
        cmd += ['--concurrent-fragments', '8']
    if has_ffmpeg:
        cmd += ['--ffmpeg-location', paths['ffmpeg_exe']]

    if mode == '1':
        if 'FLAC' in fmt_str:
            audio_fmt, audio_q = 'flac', '0'
        elif 'WAV' in fmt_str:
            audio_fmt, audio_q = 'wav', '0'
        elif '192' in fmt_str:
            audio_fmt, audio_q = 'mp3', '5'
        else:
            audio_fmt, audio_q = 'mp3', '0'
        cmd += ['-x', '--audio-format', audio_fmt, '--audio-quality', audio_q]
        if audio_fmt in ('mp3', 'flac'):
            cmd += ['--embed-thumbnail', '--add-metadata']
    else:
        # File phải phát được ở MỌI trình phát. YouTube ưu tiên VP9/Opus, còn
        # --merge-output-format mp4 chỉ đổi CONTAINER (không transcode) → sinh ra
        # .mp4 chứa VP9/Opus bị Windows Media Player từ chối ("không hỗ trợ định
        # dạng"). Ép H.264 (avc1) + AAC (mp4a) trước, hạ dần nếu site không có.
        if format_id:
            video_f = (f'{format_id}+bestaudio[acodec^=mp4a]/'
                       f'{format_id}+bestaudio/{format_id}/best')
        else:
            video_f = ('bestvideo[vcodec^=avc1]+bestaudio[acodec^=mp4a]/'
                       'bestvideo[vcodec^=avc1]+bestaudio/'
                       'bestvideo+bestaudio/best[ext=mp4]/best')
            for token, height in (('4K', 2160), ('2K', 1440), ('1080p', 1080),
                                  ('720p', 720), ('480p', 480)):
                if token in qual_str:
                    # Không lọc height trên bestaudio: audio-only không có key
                    # height → bị loại hết, chain rơi xuống nhánh hỏng
                    video_f = (f'bestvideo[vcodec^=avc1][height<={height}]'
                               f'+bestaudio[acodec^=mp4a]/'
                               f'bestvideo[vcodec^=avc1][height<={height}]+bestaudio/'
                               f'bestvideo[height<={height}]+bestaudio/'
                               f'best[ext=mp4][height<={height}]/best[height<={height}]')
                    break
        cmd += ['-f', video_f, '--merge-output-format', 'mp4']

    cmd += list(extra_args or [])
    cmd.append(url)
    return cmd


def parse_ytdlp_progress_line(line):
    line_s = line.strip()
    if '[download]' in line_s:
        m = re.search('(\\d+(?:\\.\\d+)?)\\s*%', line_s)
        pct = float(m.group(1)) if m else None
        detail = line_s.replace('[download]', '').strip()
        if len(detail) > 50:
            detail = detail[:47] + '...'
        return (pct, detail)
    else:
        if 'Merger' in line_s or 'ExtractAudio' in line_s or 'Post-process' in line_s:
            return (None, 'Đang hậu kỳ nhạc/video...')
        else:
            if 'Writing video' in line_s or 'Writing audio' in line_s:
                return (None, line_s[:50])
            else:
                return (None, '')
def download_ffmpeg_auto_redirect(app, paths):
    version_file = os.path.join(paths['ffmpeg_dir'], 'ffmpeg_version.txt')
    local_etag = ""
    # Kiểm tra xem có phiên bản local đã được ghi lại ETag chưa
    if os.path.exists(version_file) and os.path.exists(paths['ffmpeg_exe']):
        try:
            with open(version_file, 'r', encoding='utf-8') as f:
                local_etag = f.read().strip()
        except Exception:
            pass

    app._log_to_dl_console('🌐 Đang kiểm tra phiên bản mới nhất của FFmpeg...\n', 'n')
    url = 'https://github.com/yt-dlp/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip'
    
    # Gửi yêu cầu HEAD siêu nhẹ tới GitHub để kiểm tra ETag / ngày chỉnh sửa cuối cùng
    remote_etag = ""
    try:
        import ssl
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
            remote_etag = resp.headers.get('ETag', '').strip()
    except Exception as e:
        app._log_to_dl_console(f'⚠️ Không thể kết nối kiểm tra bản cập nhật FFmpeg: {str(e)}\n', 'warn')
        # Nếu kiểm tra mạng lỗi nhưng máy đã có sẵn ffmpeg offline thì tiếp tục dùng phiên bản hiện có
        if os.path.exists(paths['ffmpeg_exe']):
            app._log_to_dl_console('✓ Sử dụng phiên bản FFmpeg hiện tại offline.\n', 'ok')
            return True
        else:
            return False

    # Nếu ETag khớp và file exe có thật thì bỏ qua không cần tải lại
    if os.path.exists(paths['ffmpeg_exe']) and remote_etag and local_etag == remote_etag:
        app._log_to_dl_console('✓ FFmpeg đã ở phiên bản mới nhất!\n', 'ok')
        return True

    # Tiến hành tải mới hoặc cập nhật lên bản mới hơn
    if os.path.exists(paths['ffmpeg_exe']):
        app._log_to_dl_console('⚡ Phát hiện phiên bản FFmpeg mới hơn! Đang tự động cập nhật...\n', 'warn')
    else:
        app._log_to_dl_console('⚠️ Thiếu bộ giải mã FFmpeg! Đang tự động tải...\n', 'warn')

    zip_path = os.path.join(os.path.dirname(paths['ffmpeg_dir']), 'ffmpeg_temp.zip')
    try:
        import ssl
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        app._log_to_dl_console('  🌐 Đang kết nối tới server (timeout 120s)...\n', 'n')
        with urllib.request.urlopen(req, timeout=120, context=ctx) as response:
            total_size = int(response.headers.get('Content-Length', 0))
            total_mb = total_size / 1024 / 1024 if total_size > 0 else 0
            if total_mb > 0:
                app._log_to_dl_console(f'  📦 Kích thước: {total_mb:.1f} MB — Bắt đầu tải...\n', 'n')
            downloaded = 0
            last_pct = -1
            chunk_size = 131072
            with open(zip_path, 'wb') as f:
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        pct = int(downloaded / total_size * 100)
                        if pct >= last_pct + 10:
                            last_pct = pct
                            mb_done = downloaded / 1024 / 1024
                            app.root.after(0, lambda p=pct, m=mb_done, t=total_mb:
                                app._log_to_dl_console(f'  ↳ {m:.1f}/{t:.1f} MB ({p}%)\n', 'n'))
        app._log_to_dl_console('  ✓ Tải xong! Đang giải nén...\n', 'ok')
        os.makedirs(os.path.join(paths['ffmpeg_dir'], 'bin'), exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            extract_root = os.path.dirname(paths['ffmpeg_dir'])
            top_folders = {name.split('/')[0] for name in zip_ref.namelist() if '/' in name}
            zip_root = list(top_folders)[0] if top_folders else None
            for entry in zip_ref.namelist():
                if entry.endswith('ffmpeg.exe'):
                    zip_ref.extract(entry, extract_root)
                    src = os.path.join(extract_root, entry)
                    dst = paths['ffmpeg_exe']
                    if os.path.exists(src):
                        if os.path.exists(dst):
                            os.remove(dst)
                        shutil.move(src, dst)
                    break
        if zip_root:
            leftover = os.path.join(extract_root, zip_root)
            if os.path.exists(leftover):
                shutil.rmtree(leftover, ignore_errors=True)
        if os.path.exists(zip_path):
            os.remove(zip_path)
            
        # Ghi lại ETag mới để dùng cho những lần kiểm tra sau
        if remote_etag:
            try:
                with open(version_file, 'w', encoding='utf-8') as f:
                    f.write(remote_etag)
            except Exception:
                pass
                
        app._log_to_dl_console('✓ Cài đặt FFmpeg thành công!\n', 'ok')
    except Exception as e:
        app._log_to_dl_console(f'⚠️ Lỗi khi tải cập nhật FFmpeg: {str(e)}\n', 'warn')
        if os.path.exists(zip_path):
            try:
                os.remove(zip_path)
            except Exception:
                pass
        # Nếu tải lỗi nhưng máy đã có sẵn file exe cũ thì vẫn tiếp tục dùng bản cũ, không bị lỗi
        if os.path.exists(paths['ffmpeg_exe']):
            app._log_to_dl_console('✓ Vẫn sử dụng bộ giải mã FFmpeg hiện có trên máy.\n', 'ok')
            return True
        return False
    else:
        return True
def _apply_tiktok_cookie_redirect(app, cmd):
    app._log_to_dl_console('🔑 Đang lấy cookie trình duyệt để hỗ trợ kết nối TikTok...\n', 'n')
    localappdata = os.environ.get('LOCALAPPDATA', '')
    appdata = os.environ.get('APPDATA', '')
    browsers = {'chrome': {'path': os.path.join(localappdata, 'Google', 'Chrome', 'User Data'), 'exe': 'chrome.exe'}, 'edge': {'path': os.path.join(localappdata, 'Microsoft', 'Edge', 'User Data'), 'exe': 'msedge.exe'}, 'brave': {'path': os.path.join(localappdata, 'BraveSoftware', 'Brave-Browser', 'User Data'), 'exe': 'brave.exe'}, 'firefox': {'path': os.path.join(appdata, 'Mozilla', 'Firefox', 'Profiles'), 'exe': 'firefox.exe'}, 'opera': {'path': os.path.join(appdata, 'Opera Software', 'Opera Stable'), 'exe': 'opera.exe'}}
    installed_browsers = [b for b, info in browsers.items() if os.path.exists(info['path'])]
    if not installed_browsers:
        return None
    else:
        rc, tasklist_raw, _ = _run_cmd(['tasklist', '/NH'])
        if rc != 0:
            tasklist_raw = ''
        tasklist = tasklist_raw.lower()
        selected_browser = None
        for b in installed_browsers:
            if browsers[b]['exe'] in tasklist:
                selected_browser = b
                break
        if not selected_browser:
            selected_browser = installed_browsers[0]
        cmd.extend(['--cookies-from-browser', selected_browser])
        app._log_to_dl_console(f'✓ Sử dụng cấu hình kết nối từ trình duyệt: {selected_browser.capitalize()}\n', 'ok')
DNS_SERVERS = {'Google': ('8.8.8.8', '8.8.4.4'), 'Cloudflare': ('1.1.1.1', '1.0.0.1'), 'OpenDNS': ('208.67.222.222', '208.67.220.220'), 'Quad9': ('9.9.9.9', '149.112.112.112')}
PING_TARGETS = [('Google DNS', '8.8.8.8'), ('Cloudflare DNS', '1.1.1.1'), ('Google Home', 'google.com'), ('Facebook Server', 'facebook.com')]
def _run_cmd(args, timeout=25):
    try:
        r = subprocess.run(args, capture_output=True, text=True, timeout=timeout, encoding='utf-8', errors='replace', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        return (r.returncode, r.stdout, r.stderr)
    except subprocess.TimeoutExpired:
        return ((-1), '', 'Timeout')
    except Exception as e:
        return ((-1), '', str(e))
def _parse_ping_output(stdout):
    """Parse Windows ping output, return (times_list, loss_pct)"""
    times = []
    loss = 0.0
    for line in stdout.splitlines():
        # Parse reply time: 'time=14ms' or 'time<1ms'
        for prefix in ['time=', 'time<']:
            if prefix in line.lower():
                for token in line.split():
                    tl = token.lower()
                    if prefix in tl:
                        try:
                            val = tl.split(prefix)[1].rstrip('ms').rstrip('ms ').strip()
                            times.append(float(val))
                        except Exception:
                            pass
                        break
        # Parse loss: 'Packets: Sent = 4, Received = 4, Lost = 0 (0% loss)'
        # PHẢI lấy % trong ngoặc — vòng cũ nhận token số ĐẦU TIÊN của dòng
        # (chính là "Sent = 4") nên luôn báo loss = 4% dù mất 0%.
        if 'loss' in line.lower() or 'lost' in line.lower() or 'mất' in line.lower():
            m = re.search(r'\(\s*(\d+(?:\.\d+)?)%\s*(?:loss|mất)\s*\)', line.lower())
            if m:
                loss = float(m.group(1))
    return times, loss
import struct
def _query_dns(server_ip, host='www.google.com', timeout=2.0):
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.settimeout(timeout)
    # Build minimal DNS query
    header = struct.pack('>HHHHHH', 0x1234, 0x0100, 1, 0, 0, 0)
    question = b''.join(len(p).to_bytes(1,'big') + p.encode() for p in host.split('.')) + b'\x00'
    question += struct.pack('>HH', 1, 1)
    try:
        t0 = time.perf_counter()
        sock.sendto(header + question, (server_ip, 53))
        sock.recv(512)
        return (time.perf_counter() - t0) * 1000
    except Exception:
        return 999.0
    finally:
        sock.close()

def smart_scan_engine(app):
    print('\033[38;2;100;200;255m🤖 Bước 1: Đang chẩn đoán cấu hình TCP...\033[0m')
    rc, tcp_out, _ = _run_cmd(['netsh', 'int', 'tcp', 'show', 'global'], timeout=10)
    tcp_settings = {}
    for line in tcp_out.splitlines():
        if ':' in line:
            k, _, v = line.partition(':')
            tcp_settings[k.strip().lower()] = v.strip().lower()
    adapters = []
    rc_ad, ad_out, _ = _run_cmd(['netsh', 'interface', 'show', 'interface'], timeout=10)
    for line in ad_out.splitlines():
        parts = line.split()
        if len(parts) >= 4 and ('Connected' in line or 'Đã kết nối' in line or 'connected' in line.lower()):
            adapters.append(' '.join(parts[3:]))
    if not adapters:
        adapters = ['Ethernet', 'Wi-Fi']
    print('\033[38;2;0;255;204m✓ Hoàn thành chẩn đoán TCP.\033[0m')
    print('\n\033[38;2;100;200;255m🤖 Bước 2: Đo lường độ trễ Ping tới các server ngoại...\033[0m')
    ping_results = []
    for label, host in PING_TARGETS:
        print(f'  Đang ping {label} ({host})...')
        p_rc, p_stdout, _ = _run_cmd(['ping', '-n', '3', host])
        times, loss = _parse_ping_output(p_stdout)
        avg_t = round(statistics.mean(times), 2) if times else 0.0
        ping_results.append((label, host, avg_t, loss))
        if times:
            print(f'    → Trung bình: {avg_t}ms | Mất gói: {loss}%')
        else:
            print(f'    → Không phản hồi (timeout/firewall)')
    print('\033[38;2;0;255;204m✓ Đã hoàn thành đo Ping.\033[0m')
    print('\n\033[38;2;100;200;255m🤖 Bước 3: Benchmark hiệu năng DNS Servers...\033[0m')
    dns_results = []
    for name, (primary, secondary) in DNS_SERVERS.items():
        print(f'  Kiểm tra tốc độ phản hồi: {name} ({primary})...')
        try:
            dns_ms = round(_query_dns(primary), 2)
            ok = (dns_ms < 999.0)
        except Exception as e:
            dns_ms = 999.0
            ok = False
        dns_results.append((name, primary, dns_ms, ok))
    dns_results.sort(key=lambda x: (not x[3], x[2]))
    best_dns_name, best_dns_ip = (dns_results[0][0], dns_results[0][1])
    print(f'\033[38;2;50;255;120m✓ DNS phản hồi nhanh nhất: {best_dns_name} ({best_dns_ip}) - {dns_results[0][2]}ms\033[0m')
    print('\n\033[38;2;255;215;0m📊 BẢN KHUYẾN NGHỊ SMART SCAN TỪ HỆ THỐNG:\033[0m')
    recs = []
    autotuning = tcp_settings.get('receive window auto-tuning level', '')
    if 'normal' not in autotuning:
        recs.append(('Tối ưu TCP buffer (Auto-Tuning)', 'autotuninglevel=normal'))
    ecn = tcp_settings.get('ecn capability', '')
    if 'enabled' not in ecn:
        recs.append(('Kích hoạt ECN (Explicit Congestion)', 'ecncapability=enabled'))
    rss = tcp_settings.get('receive-side scaling state', '')
    if 'enabled' not in rss:
        recs.append(('Bật RSS (Phân phối tải nhiều lõi CPU)', 'rss=enabled'))
    recs.append(('Xóa bộ nhớ đệm cache hệ thống (Flush DNS)', 'flushdns'))
    recs.append((f'Cấu hình DNS nhanh nhất ({best_dns_name})', 'set_dns', best_dns_ip, DNS_SERVERS[best_dns_name][1]))
    print('┌───────────────────────────────────┬──────────────┬───────────────┐')
    print('│ Đề xuất cấu hình                  │ Trạng thái   │ Mức ảnh hưởng │')
    print('├───────────────────────────────────┼──────────────┼───────────────┤')
    for name, *details in recs:
        print(f'│ {name:<33} │  \033[33mCần Tối Ưu\033[0m   │  \033[36mTrung-Cao\033[0m     │')
    print('└───────────────────────────────────┴──────────────┴───────────────┘')
    print('\n\033[38;2;50;255;120m⚡ Bắt đầu tiến hành tối ưu hóa hệ thống mạng tự động...\033[0m')
    for r in recs:
        action_name = r[0]
        action_key = r[1]
        print(f'  ➜ Đang xử lý: {action_name}...')
        if action_key == 'flushdns':
            _run_cmd(['ipconfig', '/flushdns'])
            print('    [✓] Đã xóa cache DNS.')
        elif action_key.startswith('autotuninglevel') or action_key.startswith('ecncapability') or action_key.startswith('rss'):
            _run_cmd(['netsh', 'int', 'tcp', 'set', 'global', action_key])
            print('    [✓] Đã cấu hình tham số hệ thống.')
        elif action_key == 'set_dns':
            target_ip = r[2]
            sec_ip = r[3]
            chosen_adapter = adapters[0]
            _run_cmd(['netsh', 'interface', 'ip', 'set', 'dns', f'name="{chosen_adapter}"', 'static', target_ip])
            print('    [✓] Đã cấu hình DNS.')
    print('\n \033[38;2;50;255;120m🎉 CHÚC MỪNG! HỆ THỐNG MẠNG ĐÃ ĐƯỢC TỐI ƯU HÓA HOÀN TOÀN! \033[0m')
    print(' \033[38;2;90;100;120mTất cả cấu hình đã có hiệu lực ngay lập tức. Hãy tận hưởng độ trễ cực thấp! 🚀 \033[0m')
def speed_test_engine(app):
    if not HAVE_SPEEDTEST:
        print(' \033[91m❌ Thư viện speedtest-cli chưa được cài đặt! \033[0m')
        if getattr(sys, 'frozen', False):
            print('  \033[33mKhông thể cài đặt tự động trong bản portable.\033[0m')
        else:
            print('  \033[33mĐang cài đặt speedtest-cli...\033[0m')
            rc, _, _ = _run_cmd(['pip', 'install', 'speedtest-cli'])
            if rc != 0:
                print(' \033[91m✗ Cài đặt thư viện thất bại! Vui lòng kiểm tra lại mạng. \033[0m')
                return None
    print('🔌 Kết nối tới server đo lường chất lượng cao nhất...')
    try:
        st = speedtest.Speedtest()
        st.get_best_server()
        print('📥 Đang đo tốc độ tải xuống (Download) - vui lòng đợi...')
        dl = round(st.download() / 1000000, 2)
        print('📤 Đang đo tốc độ tải lên (Upload) - vui lòng đợi...')
        ul = round(st.upload() / 1000000, 2)
        ping = round(st.results.dict().get('ping', 0), 2)
        server_name = st.results.dict().get('server', {}).get('name', 'Unknown')
        print('\n \033[38;2;50;255;120m📊 KẾT QUẢ ĐO TỐC ĐỘ MẠNG THỜI GIAN THỰC: \033[0m')
        print('┌───────────────────────────┬───────────────────────────┐')
        print('│ Chỉ Số Đo Lường           │ Giá Trị Đạt Được          │')
        print('├───────────────────────────┼───────────────────────────┤')
        print(f'│ 📥 Download Speed         │  \033[36m{dl:>18} Mbps \033[0m │')
        print(f'│ 📤 Upload Speed           │  \033[36m{ul:>18} Mbps \033[0m │')
        print(f'│ 📡 Độ trễ (Latency/Ping)  │  \033[32m{ping:>18} ms \033[0m   │')
        print(f'│ 🖥  Trạm kiểm định         │ {server_name:>23}   │')
        print('└───────────────────────────┴───────────────────────────┘')
    except Exception as e:
        print(f'[91m❌ Không thể thực hiện đo tốc độ. Lỗi: {str(e)}[0m')
def ping_latency_engine(app):
    print('📶 Đang ping kiểm tra đến 4 server máy chủ lớn...')
    for label, host in PING_TARGETS:
        print(f'  Đo độ trễ tới {label:<16} ({host})...')
        rc, out, _ = _run_cmd(['ping', '-n', '4', host])
        times, loss = _parse_ping_output(out)
        if times:
            avg_ms = round(statistics.mean(times), 2)
            min_ms = round(min(times), 2)
            max_ms = round(max(times), 2)
            print(f'    ✓ Trung bình: {avg_ms}ms | Cực tiểu/Cực đại: {min_ms}ms / {max_ms}ms | Mất gói: {loss}%')
        else:
            print('    ✗ Mất kết nối tới server (Request Timed Out hoặc chặn firewall).')
def tcp_optimize_engine(app):
    tweaks = [('Tự động co giãn TCP Buffer (Normal)', 'autotuninglevel=normal', 'Tối ưu dung lượng gói truyền tải'), ('Kích hoạt ECN (Explicit Congestion Notification)', 'ecncapability=enabled', 'Báo hiệu nghẽn trước, giảm rớt mạng'), ('Receive-Side Scaling (RSS) đa lõi', 'rss=enabled', 'Chia sẻ gánh nặng xử lý card mạng đều lên CPU'), ('Direct Cache Access (DCA) trực tiếp', 'dca=enabled', 'Tăng tốc độ truy vấn bộ nhớ đệm mạng')]
    print('⚙ Bắt đầu tối ưu hóa tham số mạng sâu trong hệ thống registry Windows...')
    success_cnt = 0
    for title, action, desc in tweaks:
        print(f'  ➜ Áp dụng: {title}...')
        rc, _, err = _run_cmd(['netsh', 'int', 'tcp', 'set', 'global', action])
        if rc == 0:
            print(f'    [32m[✓] Thành công: {desc}[0m')
            success_cnt += 1
        else:
            print('    [91m[✗] Thất bại (Cần khởi chạy ứng dụng với quyền Administrator)[0m')
    print(f'\n[32m✓ Hoàn tất cấu hình sâu. Đã áp dụng thành công {success_cnt}/{len(tweaks)} tùy chọn mạng![0m')
def dns_benchmark_engine(app):
    print('⚡ Bắt đầu đo tốc độ DNS Benchmark của các hãng lớn...')
    results = []
    for name, (primary, secondary) in DNS_SERVERS.items():
        print(f'  Gửi truy vấn test tới: {name:<12} ({primary})...')
        t_sum = 0
        success_runs = 0
        for _ in range(3):
            try:
                dns_ms = _query_dns(primary)
                if dns_ms < 999.0:
                    t_sum += dns_ms
                    success_runs += 1
            except Exception as e:
                pass
        if success_runs > 0:
            avg_dns = round(t_sum / success_runs, 2)
            results.append((name, primary, secondary, avg_dns, True))
            print(f'    ↳ Phản hồi: {avg_dns}ms')
        else:
            results.append((name, primary, secondary, 999.0, False))
            print('    ↳ [91m✗ Lỗi phản hồi DNS.[0m')
    results.sort(key=lambda x: (not x[4], x[3]))
    best = results[0]
    print(f'\n[32m★ DNS NHANH NHẤT HIỆN TẠI: {best[0]} ({best[1]}) - {best[3]}ms[0m')
    adapters = []
    rc_ad, ad_out, _ = _run_cmd(['netsh', 'interface', 'show', 'interface'], timeout=10)
    for line in ad_out.splitlines():
        parts = line.split()
        if len(parts) >= 4 and ('Connected' in line or 'Đã kết nối' in line or 'connected' in line.lower()):
                adapters.append(' '.join(parts[3:]))
    if not adapters:
        adapters = ['Ethernet', 'Wi-Fi']
    chosen = adapters[0]
    print(f'  ➜ Đang áp dụng thiết lập DNS của {best[0]} cho adapter \'{chosen}\'...')
    rc1, _, _ = _run_cmd(['netsh', 'interface', 'ip', 'set', 'dns', f'name="{chosen}"', 'static', best[1]])
    rc2, _, _ = _run_cmd(['netsh', 'interface', 'ip', 'add', 'dns', f'name="{chosen}"', best[2], 'index=2'])
    if rc1 == 0 and rc2 == 0:
        print(f'[32m✓ Hoàn thành set DNS nhanh nhất: {best[1]} và {best[2]}![0m')
    else:
        print('[91m✗ Thất bại khi set DNS. Hãy đảm bảo chạy Toolbox với quyền Administrator![0m')
def reset_network_engine(app):
    actions = [('Xóa DNS Cache (Flush DNS)', ['ipconfig', '/flushdns']), ('Reset bộ nhớ đệm NetBIOS', ['nbtstat', '-R']), ('Khởi tạo lại Winsock kết nối', ['netsh', 'winsock', 'reset']), ('Khởi tạo lại cấu trúc TCP/IP Stack', ['netsh', 'int', 'ip', 'reset'])]
    print('🔄 Bắt đầu dọn dẹp và reset toàn bộ card mạng hệ thống...')
    success_cnt = 0
    for label, cmd in actions:
        print(f'  ➜ Thực hiện: {label}...')
        rc, _, _ = _run_cmd(cmd)
        if rc == 0:
            print('    [32m✓ Thành công.[0m')
            success_cnt += 1
        else:
            print('    [33m⚠ Không phản hồi hoặc quyền bị giới hạn.[0m')
    print(f'\n[32m✓ Hoàn thành Reset mạng ({success_cnt}/{len(actions)} tác vụ thành công).[0m')
    print('[33m⚠ Lưu ý: Hãy khởi động lại máy tính (Restart) nếu bạn thấy mạng chập chờn để các thay đổi Winsock áp dụng hoàn toàn![0m')
def _is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False
if __name__ == '__main__':
    if os.name == 'nt':
        try:
            ctypes.windll.kernel32.SetConsoleMode(ctypes.windll.kernel32.GetStdHandle((-11)), 7)
        except Exception:
            pass
        # Yêu cầu quyền Administrator nếu chưa có
        if not _is_admin():
            import tkinter as _tk
            _r = _tk.Tk()
            _r.withdraw()
            answer = messagebox.askyesno(
                'Yêu cầu quyền Administrator',
                'Toolbox cần quyền Administrator để sử dụng đầy đủ tính năng\n'
                '(Tối ưu mạng, thay đổi TCP/IP, set DNS...)\n\n'
                'Bạn có muốn khởi động lại với quyền Admin không?',
                icon='warning'
            )
            _r.destroy()
            if answer:
                # Khởi động lại với quyền runas (UAC popup)
                ctypes.windll.shell32.ShellExecuteW(
                    None, 'runas',
                    sys.executable,
                    ' '.join(f'"{a}"' for a in sys.argv),
                    None, 1
                )
                sys.exit(0)
    root = tk.Tk()
    App(root)
    root.mainloop()
