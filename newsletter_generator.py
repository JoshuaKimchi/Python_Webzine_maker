import tkinter as tk
from tkinter import scrolledtext, messagebox, Canvas, Frame, Scrollbar, colorchooser, filedialog
import json
import webbrowser
import uuid
import re
from pathlib import Path

# =================================================================================
# 상수 정의
# =================================================================================
CWD = Path.cwd()
BACKUP_DIR = CWD / "backups"
DATA_FILE = CWD / "newsletter_data.json"
OUTPUT_FILE = CWD / "newsletter.html"

APP_FONT = "굴림 10"
RECOMMENDED_COLORS = ["#74438d", "#f1b34a", "#4a6da7", "#509598", "#616161"]
COLOR_SUCCESS = "#28a745"
COLOR_PRIMARY = "#007bff"
COLOR_DANGER = "#dc3545"
COLOR_INFO = "#6c757d"

# =================================================================================
# HTML 템플릿
# =================================================================================
HTML_TEMPLATE = """
<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{html_title}</title><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet"><style>body{{font-family:'Noto Sans KR',sans-serif;margin:0;background-color:#f0f2f5;color:#333}}.header{{display:flex;justify-content:space-between;align-items:center;background-image:url('your_header_img_here.jpg');background-size:cover;background-position:center;padding:20px 40px;color:white;min-height:150px;position:relative;z-index:1}}.header::before{{content:'';position:absolute;top:0;left:0;right:0;bottom:0;background-color:rgba(0,0,0,0.4);z-index:2}}.header .logo,.header .issue-info{{z-index:3;position:relative}}.header .logo-text{{font-size:2.5em;font-weight:bold;text-shadow:2px 2px 4px rgba(0,0,0,0.7)}}.header .issue-info{{font-size:1.2em;background-color:#f1b34a;padding:8px 15px;border-radius:5px;position:absolute;right:40px;bottom:20px}}.container{{padding:20px}}.info-card{{display:flex;width:100%;max-width:800px;margin:25px auto;border-radius:10px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.1);background-color:#fff}}.card-sidebar{{display:flex;justify-content:center;align-items:center;flex-shrink:0;width:70px;padding:20px 0;writing-mode:vertical-rl;text-orientation:mixed;color:white;font-size:1.5em;font-weight:700;letter-spacing:2px;text-align:center;transition:all 0.3s ease}}.card-main{{flex-grow:1;display:flex;flex-direction:column}}.main-header{{padding:12px 20px;color:white;font-size:1.2em;font-weight:700}}.main-content{{padding:20px;line-height:1.8}}.content-item{{margin-bottom:20px;border-bottom:1px solid #eee;padding-bottom:15px}}.content-item:last-child{{border-bottom:none;margin-bottom:0;padding-bottom:0}}.content-item-title{{font-size:1.1em;font-weight:500;margin-bottom:8px}}.content-item-body{{color:#555}}a{{text-decoration:none;color:inherit}}a:hover{{text-decoration:underline}}@media (max-width:768px){{.header{{padding:20px 15px}}.header .logo-text{{font-size:2em}}.header .issue-info{{font-size:1em;padding:6px 10px;right:15px;bottom:15px}}.container{{padding:15px 10px}}.info-card{{flex-direction:column;margin:15px auto;border-left:none}}.card-sidebar{{writing-mode:horizontal-tb;text-orientation:initial;width:auto;padding:10px 15px;justify-content:flex-start;font-size:1.3em}}.main-content{{padding:15px}}}}</style></head><body><header class="header"><div class="logo"><span class="logo-text">{main_title}</span></div><div class="issue-info">제 {issue_no}호 / {issue_date}</div></header><main class="container">{sections_html}</main></body></html>
"""

SECTION_TEMPLATE = """
<div class="info-card" style="border-left:5px solid {color};"><div class="card-sidebar" style="background-color:{color};">{sidebar_title}</div><div class="card-main"><div class="main-header" style="background-color:{color};">{section_title}</div><div class="main-content">{contents_html}</div></div></div>
"""

CONTENT_TEMPLATE = """
<div class="content-item"><div class="content-item-title" style="color:{color};font-weight:{font_weight};">{content_title}</div><div class="content-item-body">{content_body}</div></div>
"""

# =================================================================================
# 메인 애플리케이션 클래스
# =================================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("HTML 뉴스레터 생성기")
        self.root.geometry("900x800")
        self.root.option_add("*Font", APP_FONT)

        self.header_widgets = {}
        self.sections = []

        BACKUP_DIR.mkdir(parents=True, exist_ok=True)

        self.setup_ui()
        self.load_data()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        self.canvas = Canvas(self.root, bd=0, highlightthickness=0)
        self.scrollbar = Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_mousewheel, add="+")
        self.root.bind_all("<Button-5>", self._on_mousewheel, add="+")

        header_frame = tk.LabelFrame(self.scrollable_frame, text="헤더 정보", padx=15, pady=15)
        header_frame.pack(fill="x", padx=10, pady=10)
        self.create_header_widgets(header_frame)

        self.sections_frame = Frame(self.scrollable_frame)
        self.sections_frame.pack(fill="both", expand=True, padx=10, pady=0)
        
        control_frame = Frame(self.scrollable_frame)
        control_frame.pack(fill="x", padx=10, pady=20)
        tk.Button(control_frame, text="섹션 추가 (+)", command=self.add_section, bg=COLOR_PRIMARY, fg="white").pack(side="left", padx=5)
        tk.Button(control_frame, text="HTML 생성 및 열기", command=self.generate_html, font=(APP_FONT.split()[0], 12, "bold"), bg=COLOR_SUCCESS, fg="white").pack(side="right", padx=5)

    def create_header_widgets(self, parent):
        tk.Label(parent, text="메인 제목:").grid(row=0, column=0, sticky="w", padx=5)
        self.header_widgets['main_title'] = tk.Entry(parent)
        self.header_widgets['main_title'].grid(row=0, column=1, columnspan=3, sticky="ew")
        
        tk.Label(parent, text="발행 호수:").grid(row=1, column=0, sticky="w", padx=5, pady=(5,0))
        self.header_widgets['issue_no'] = tk.Entry(parent, width=10)
        self.header_widgets['issue_no'].grid(row=1, column=1, sticky="w", pady=(5,0))
        
        button_frame = Frame(parent)
        button_frame.grid(row=1, column=2, sticky="w", pady=(5,0))
        tk.Button(button_frame, text="백업 저장", command=self.manual_save).pack(side="left", padx=(10, 2))
        tk.Button(button_frame, text="백업 불러오기", command=self.manual_load).pack(side="left")

        tk.Label(parent, text="발행 날짜:").grid(row=2, column=0, sticky="w", padx=5, pady=(5,0))
        self.header_widgets['issue_date'] = tk.Entry(parent, width=20)
        self.header_widgets['issue_date'].grid(row=2, column=1, sticky="w", pady=(5,0))
        parent.columnconfigure(1, weight=1)

    def add_section(self, data=None):
        color_index = len(self.sections) % len(RECOMMENDED_COLORS)
        initial_color = data.get("color") if data else RECOMMENDED_COLORS[color_index]
        section = SectionFrame(self.sections_frame, self.remove_section, data, initial_color)
        section.pack(fill="x", pady=(0, 15), expand=True, padx=5)
        self.sections.append(section)

    def remove_section(self, section_to_remove):
        section_to_remove.destroy()
        self.sections.remove(section_to_remove)

    def get_data(self):
        return {
            "header": {key: widget.get() for key, widget in self.header_widgets.items()},
            "sections": [section.get_data() for section in self.sections]
        }

    def save_data(self):
        try:
            with DATA_FILE.open('w', encoding='utf-8') as f:
                json.dump(self.get_data(), f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("자동 저장 오류", f"데이터 자동 저장 중 오류가 발생했습니다:\n{e}")

    def load_data(self):
        if not DATA_FILE.exists(): return
        try:
            with DATA_FILE.open('r', encoding='utf-8') as f:
                data = json.load(f)
            self._populate_ui_from_data(data)
        except (json.JSONDecodeError, IOError):
            if messagebox.askyesno("데이터 파일 오류", f"데이터 파일({DATA_FILE.name})을 불러오는 데 실패했습니다.\n백업 파일을 만드시겠습니까?"):
                DATA_FILE.rename(f"{DATA_FILE}.bak_{uuid.uuid4().hex[:6]}")

    def _populate_ui_from_data(self, data):
        for section in self.sections: section.destroy()
        self.sections.clear()
        
        header_data = data.get("header", {})
        for key, value in header_data.items():
            if key in self.header_widgets:
                self.header_widgets[key].delete(0, tk.END)
                self.header_widgets[key].insert(0, value)
        
        for section_data in data.get("sections", []):
            self.add_section(section_data)

    def _sanitize_filename(self, name):
        return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

    def manual_save(self):
        try:
            data = self.get_data()
            header = data.get("header", {})
            main_title = self._sanitize_filename(header.get("main_title", "제목없음"))
            issue_no = self._sanitize_filename(header.get("issue_no", "호수없음"))
            issue_date = self._sanitize_filename(header.get("issue_date", "날짜없음"))
            
            base_filename = f"{main_title}_{issue_no}_{issue_date}"
            json_path = BACKUP_DIR / f"{base_filename}.json"
            html_path = BACKUP_DIR / f"{base_filename}.html"

            with json_path.open('w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            html_content = self.get_html_content(data)
            with html_path.open('w', encoding='utf-8') as f:
                f.write(html_content)
                
            messagebox.showinfo("백업 저장 완료", f"백업이 성공적으로 저장되었습니다.\n위치: {BACKUP_DIR}")
        except Exception as e:
            messagebox.showerror("백업 저장 오류", f"백업 파일 저장 중 오류가 발생했습니다:\n{e}")

    def manual_load(self):
        filepath = filedialog.askopenfilename(initialdir=BACKUP_DIR, title="백업 파일 선택", filetypes=(("JSON 파일", "*.json"), ("모든 파일", "*.*")))
        if not filepath or not messagebox.askyesno("불러오기 확인", "현재 작업 내용이 사라집니다. 정말로 불러오시겠습니까?"):
            return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self._populate_ui_from_data(data)
            messagebox.showinfo("성공", "백업 파일을 성공적으로 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 불러오는 중 오류가 발생했습니다: {e}")

    def get_html_content(self, data):
        try:
            header = data.get("header", {})
            html_title = f"{header.get('main_title', '')}_{header.get('issue_no', '')}_{header.get('issue_date', '')}"
            
            sections_html_parts = []
            for section_data in data.get("sections", []):
                contents_html_parts = []
                for content_data in section_data.get("contents", []):
                    font_weight = "bold" if content_data.get("is_bold") else "normal"
                    section_color = section_data.get("color", COLOR_INFO)
                    color = content_data.get("color") or section_color
                    
                    title_text = content_data.get("title") or "&nbsp;"
                    body_text = (content_data.get("body") or "").replace('\n', '<br>')
                    link_url = content_data.get("link")

                    if link_url:
                        link_html = f'<p style="margin-top:15px;text-align:right;font-size:0.9em;"><a href="{link_url}" target="_blank" style="color:{color};font-weight:500;text-decoration:none;border-bottom:1px solid {color};padding-bottom:2px;">바로가기 링크 &rarr;</a></p>'
                        body_text += link_html
                    
                    contents_html_parts.append(CONTENT_TEMPLATE.format(content_title=title_text, content_body=body_text, color=color, font_weight=font_weight))
                
                sections_html_parts.append(SECTION_TEMPLATE.format(
                    sidebar_title=section_data.get("sidebar_title", ""),
                    section_title=section_data.get("title") or "&nbsp;",
                    color=section_data.get("color", COLOR_INFO),
                    contents_html="".join(contents_html_parts)
                ))

            return HTML_TEMPLATE.format(
                html_title=html_title,
                main_title=header.get("main_title", ""),
                issue_no=header.get("issue_no", ""),
                issue_date=header.get("issue_date", ""),
                sections_html="".join(sections_html_parts)
            )
        except Exception as e:
            messagebox.showerror("HTML 생성 오류", f"HTML 생성 중 오류가 발생했습니다:\n{e}")
            return "<html><body><h1>HTML 생성 오류</h1></body></html>"

    def generate_html(self):
        self.save_data()
        final_html = self.get_html_content(self.get_data())
        with OUTPUT_FILE.open('w', encoding='utf-8') as f:
            f.write(final_html)
        messagebox.showinfo("성공", "HTML 파일이 성공적으로 생성되었습니다.")
        webbrowser.open_new_tab(OUTPUT_FILE.resolve().as_uri())

    def on_closing(self):
        self.save_data()
        self.root.destroy()
        
    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

# =================================================================================
# 섹션 프레임 클래스
# =================================================================================
class SectionFrame(tk.LabelFrame):
    def __init__(self, parent, remove_callback, data=None, initial_color=None):
        super().__init__(parent, text="섹션 정보", padx=15, pady=15, bd=1, relief="solid")
        self.id = str(uuid.uuid4())
        self.remove_callback = remove_callback
        self.contents = []
        
        top_frame = Frame(self)
        top_frame.pack(fill="x", pady=(0, 10))
        tk.Button(top_frame, text="섹션 삭제", command=self.destroy_frame, bg=COLOR_DANGER, fg="white").pack(side="right")
        
        fields_frame = Frame(self)
        fields_frame.pack(fill="x")
        
        tk.Label(fields_frame, text="세로 제목:").grid(row=0, column=0, sticky="w")
        self.sidebar_title_entry = tk.Entry(fields_frame)
        self.sidebar_title_entry.grid(row=0, column=1, columnspan=2, sticky="ew")
        
        tk.Label(fields_frame, text="가로 제목:").grid(row=1, column=0, sticky="w", pady=(5,0))
        self.title_entry = tk.Entry(fields_frame)
        self.title_entry.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(5,0))
        
        tk.Label(fields_frame, text="색상 코드:").grid(row=2, column=0, sticky="w", pady=(5,0))
        self.color_entry = tk.Entry(fields_frame, width=10)
        self.color_entry.grid(row=2, column=1, sticky="w", pady=(5,0))
        
        self.color_preview = Canvas(fields_frame, width=22, height=22, relief="sunken", bd=1)
        self.color_preview.grid(row=2, column=2, sticky="w", padx=5, pady=(5,0))
        self.color_entry.bind("<KeyRelease>", self._update_color_preview)
        self.color_preview.bind("<Button-1>", self.choose_color)
        
        fields_frame.columnconfigure(1, weight=1)
        
        self.contents_frame = Frame(self)
        self.contents_frame.pack(fill="x", pady=10)
        tk.Button(self, text="콘텐츠 추가 (+)", command=lambda: self.add_content()).pack(pady=5)
        
        if data:
            self.load_data(data)
        else:
            self.color_entry.insert(0, initial_color or "#74438d")
            self.add_content()
            self.add_content()
        self._update_color_preview()

    def choose_color(self, event=None):
        _, color_hex = colorchooser.askcolor(title="색상 선택", initialcolor=self.color_entry.get())
        if color_hex:
            self.color_entry.delete(0, tk.END)
            self.color_entry.insert(0, color_hex)
            self._update_color_preview()

    def _update_color_preview(self, event=None):
        try:
            self.color_preview.config(bg=self.color_entry.get())
        except tk.TclError:
            self.color_preview.config(bg="white")

    def add_content(self, data=None):
        content = ContentFrame(self.contents_frame, self.remove_content, data)
        content.pack(fill="x", pady=5, expand=True)
        self.contents.append(content)

    def remove_content(self, content_to_remove):
        content_to_remove.destroy()
        self.contents.remove(content_to_remove)

    def destroy_frame(self):
        if messagebox.askyesno("삭제 확인", "이 섹션을 정말 삭제하시겠습니까?"):
            self.remove_callback(self)

    def get_data(self):
        return {
            "sidebar_title": self.sidebar_title_entry.get(),
            "title": self.title_entry.get(),
            "color": self.color_entry.get(),
            "contents": [content.get_data() for content in self.contents]
        }

    def load_data(self, data):
        self.sidebar_title_entry.insert(0, data.get("sidebar_title", ""))
        self.title_entry.insert(0, data.get("title", ""))
        self.color_entry.insert(0, data.get("color", "#FFFFFF"))
        for content_data in data.get("contents", []):
            self.add_content(content_data)

# =================================================================================
# 콘텐츠 프레임 클래스
# =================================================================================
class ContentFrame(tk.LabelFrame):
    def __init__(self, parent_frame, remove_callback, data=None):
        super().__init__(parent_frame, text="콘텐츠", padx=10, pady=10, bd=1, relief="solid")
        self.parent_section = self.master.master
        self.id = str(uuid.uuid4())
        self.remove_callback = remove_callback
        
        tk.Label(self, text="제목:").grid(row=0, column=0, sticky="w", pady=2)
        self.title_entry = tk.Entry(self)
        self.title_entry.grid(row=0, column=1, columnspan=4, sticky="ew")

        tk.Label(self, text="내용:").grid(row=1, column=0, sticky="nw", pady=2)
        self.body_text = scrolledtext.ScrolledText(self, height=4, wrap=tk.WORD)
        self.body_text.grid(row=1, column=1, columnspan=4, sticky="ew")
        self.body_text.bind("<Tab>", lambda e: self.link_entry.focus_set() or "break")
        self.body_text.bind("<Shift-Tab>", lambda e: self.title_entry.focus_set() or "break")

        tk.Label(self, text="링크:").grid(row=2, column=0, sticky="w", pady=2)
        self.link_entry = tk.Entry(self)
        self.link_entry.grid(row=2, column=1, columnspan=4, sticky="ew")

        self.is_bold = tk.BooleanVar()
        tk.Checkbutton(self, text="굵게", variable=self.is_bold).grid(row=3, column=1, sticky="w")
        
        tk.Label(self, text="색상:").grid(row=3, column=2, sticky="e", padx=(10,0))
        self.color_entry = tk.Entry(self, width=10)
        self.color_entry.grid(row=3, column=3, sticky="w")
        
        self.color_preview = Canvas(self, width=22, height=22, relief="sunken", bd=1)
        self.color_preview.grid(row=3, column=4, sticky="w", padx=5)
        self.color_entry.bind("<KeyRelease>", self._update_color_preview)
        self.color_preview.bind("<Button-1>", self.choose_color)
        
        delete_button = tk.Button(self, text="이 콘텐츠 삭제", command=self.destroy_frame, bg=COLOR_INFO, fg="white")
        delete_button.grid(row=4, column=0, columnspan=5, pady=(10,0), sticky="e")
        
        self.columnconfigure(1, weight=1)
        if data: self.load_data(data)
        self._update_color_preview()

    def choose_color(self, event=None):
        current_color = self.color_entry.get() or self.parent_section.color_entry.get()
        _, color_hex = colorchooser.askcolor(title="색상 선택", initialcolor=current_color)
        if color_hex:
            self.color_entry.delete(0, tk.END)
            self.color_entry.insert(0, color_hex)
            self._update_color_preview()

    def _update_color_preview(self, event=None):
        color_code = self.color_entry.get() or self.parent_section.color_entry.get()
        try:
            self.color_preview.config(bg=color_code)
        except tk.TclError:
            self.color_preview.config(bg="white")

    def destroy_frame(self):
        if messagebox.askyesno("콘텐츠 삭제", "이 콘텐츠를 정말 삭제하시겠습니까?"):
            self.remove_callback(self)

    def get_data(self):
        return {
            "title": self.title_entry.get(),
            "body": self.body_text.get("1.0", tk.END).strip(),
            "link": self.link_entry.get(),
            "is_bold": self.is_bold.get(),
            "color": self.color_entry.get()
        }

    def load_data(self, data):
        self.title_entry.insert(0, data.get("title", ""))
        self.body_text.insert("1.0", data.get("body", ""))
        self.link_entry.insert(0, data.get("link", ""))
        self.is_bold.set(data.get("is_bold", False))
        self.color_entry.insert(0, data.get("color", ""))

# =================================================================================
# 애플리케이션 실행
# =================================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()