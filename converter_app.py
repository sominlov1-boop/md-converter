"""
MD Converter — macOS App
DOCX / PDF / Text → Markdown
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re

# ── 변환 로직 ──────────────────────────────────────────────

def docx_to_md(path: str) -> str:
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        raise ImportError("python-docx가 설치되어 있지 않습니다.")

    doc = Document(path)
    lines = []

    for para in doc.paragraphs:
        style = para.style.name or ""
        text = para.text.strip()
        if not text:
            lines.append("")
            continue

        if style.startswith("Heading 1") or style == "제목 1":
            lines.append(f"# {text}")
        elif style.startswith("Heading 2") or style == "제목 2":
            lines.append(f"## {text}")
        elif style.startswith("Heading 3") or style == "제목 3":
            lines.append(f"### {text}")
        elif style.startswith("List"):
            lines.append(f"- {text}")
        else:
            # 인라인 서식 처리
            md_text = ""
            for run in para.runs:
                t = run.text
                if run.bold and run.italic:
                    t = f"***{t}***"
                elif run.bold:
                    t = f"**{t}**"
                elif run.italic:
                    t = f"*{t}*"
                md_text += t
            lines.append(md_text if md_text.strip() else text)

    # 표 처리
    for table in doc.tables:
        if not table.rows:
            continue
        header = [cell.text.strip() for cell in table.rows[0].cells]
        lines.append("")
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join(["---"] * len(header)) + " |")
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
        lines.append("")

    return "\n".join(lines)


def pdf_to_md(path: str) -> str:
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber가 설치되어 있지 않습니다.")

    lines = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                lines.append(text)
            # 표 추출
            for table in page.extract_tables():
                if not table:
                    continue
                header = [str(c or "").strip() for c in table[0]]
                lines.append("")
                lines.append("| " + " | ".join(header) + " |")
                lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                for row in table[1:]:
                    cells = [str(c or "").strip() for c in row]
                    lines.append("| " + " | ".join(cells) + " |")
                lines.append("")

    return "\n".join(lines)


def text_to_md(text: str) -> str:
    """기본 텍스트 → Markdown 휴리스틱 변환"""
    lines = text.splitlines()
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            result.append("")
            continue
        # 짧고 끝에 마침표 없으면 제목 후보
        if len(stripped) < 40 and not stripped.endswith((".", ",", "?", "!")):
            result.append(f"## {stripped}")
        else:
            result.append(stripped)
    return "\n".join(result)


def clean_markdown(md: str) -> str:
    """빈 줄 3개 이상 → 2개로 정리"""
    md = re.sub(r'\n{3,}', '\n\n', md)
    return md.strip()


# ── UI ─────────────────────────────────────────────────────

DARK_BG   = "#1a1a1a"
PANEL_BG  = "#242424"
BORDER    = "#333333"
ACCENT    = "#7eceff"
TEXT_MAIN = "#e8e8e8"
TEXT_DIM  = "#888888"
GREEN     = "#5ddb8e"
RED_ERR   = "#ff6b6b"

FONT_MONO = ("JetBrains Mono", 12) if os.path.exists("/Library/Fonts/JetBrainsMono-Regular.ttf") else ("Menlo", 12)
FONT_UI   = ("SF Pro Display", 13) if os.system("fc-list | grep 'SF Pro' > /dev/null 2>&1") == 0 else ("Helvetica Neue", 13)


class MDConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MD Converter")
        self.geometry("920x680")
        self.minsize(760, 560)
        self.configure(bg=DARK_BG)
        self._result_md = ""
        self._build_ui()
        self._check_deps()

    def _build_ui(self):
        # ── 헤더
        hdr = tk.Frame(self, bg=DARK_BG, pady=16)
        hdr.pack(fill="x", padx=28)
        tk.Label(hdr, text="MD Converter", font=("Helvetica Neue", 22, "bold"),
                 fg=ACCENT, bg=DARK_BG).pack(side="left")
        tk.Label(hdr, text="DOCX · PDF · Text  →  Markdown",
                 font=("Helvetica Neue", 12), fg=TEXT_DIM, bg=DARK_BG).pack(side="left", padx=14, pady=4)

        # ── 구분선
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── 탭
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        self._style_notebook()

        self.tab_file = tk.Frame(nb, bg=DARK_BG)
        self.tab_text = tk.Frame(nb, bg=DARK_BG)
        nb.add(self.tab_file, text="  파일 변환  ")
        nb.add(self.tab_text, text="  텍스트 붙여넣기  ")

        self._build_file_tab(self.tab_file)
        self._build_text_tab(self.tab_text)

    def _style_notebook(self):
        s = ttk.Style(self)
        s.theme_use("default")
        s.configure("TNotebook", background=DARK_BG, borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab",
                    background=PANEL_BG, foreground=TEXT_DIM,
                    font=("Helvetica Neue", 12), padding=(16, 8),
                    borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", DARK_BG)],
              foreground=[("selected", ACCENT)])

    # ── 파일 탭 ─────────────────────────────────────────────

    def _build_file_tab(self, parent):
        # 드롭존
        drop_frame = tk.Frame(parent, bg=PANEL_BG, bd=0, relief="flat",
                              highlightbackground=BORDER, highlightthickness=1)
        drop_frame.pack(fill="x", padx=28, pady=(20, 0))

        self.drop_label = tk.Label(
            drop_frame,
            text="📂  파일을 드래그하거나 클릭하여 선택\n\n.docx   .pdf",
            font=("Helvetica Neue", 14), fg=TEXT_DIM, bg=PANEL_BG,
            pady=32, cursor="hand2"
        )
        self.drop_label.pack(fill="x")
        self.drop_label.bind("<Button-1>", lambda e: self._pick_file())

        # 파일명 표시
        self.file_info = tk.Label(parent, text="", font=("Menlo", 11),
                                  fg=ACCENT, bg=DARK_BG)
        self.file_info.pack(pady=(8, 0))

        # 변환 버튼
        self.btn_convert = self._btn(parent, "변환하기", self._convert_file)
        self.btn_convert.pack(pady=(12, 0))

        # 결과
        self._build_result_area(parent)

    def _build_text_tab(self, parent):
        tk.Label(parent, text="텍스트를 붙여넣으세요",
                 font=("Helvetica Neue", 12), fg=TEXT_DIM, bg=DARK_BG).pack(
            anchor="w", padx=28, pady=(20, 6))

        self.text_input = tk.Text(
            parent, height=10, bg=PANEL_BG, fg=TEXT_MAIN,
            font=("Menlo", 12), insertbackground=ACCENT,
            relief="flat", bd=0, padx=12, pady=10,
            highlightbackground=BORDER, highlightthickness=1
        )
        self.text_input.pack(fill="x", padx=28)

        self._btn(parent, "변환하기", self._convert_text).pack(pady=(12, 0))

        self._build_result_area(parent, prefix="txt")

    def _build_result_area(self, parent, prefix="file"):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=28, pady=(20, 0))

        label_row = tk.Frame(parent, bg=DARK_BG)
        label_row.pack(fill="x", padx=28, pady=(8, 4))
        tk.Label(label_row, text="Markdown 결과", font=("Helvetica Neue", 11),
                 fg=TEXT_DIM, bg=DARK_BG).pack(side="left")

        btn_row = tk.Frame(parent, bg=DARK_BG)
        btn_row.pack(fill="x", padx=28)
        self._small_btn(btn_row, "클립보드 복사", lambda: self._copy(prefix)).pack(side="left", padx=(0, 8))
        self._small_btn(btn_row, ".md 저장", lambda: self._save(prefix)).pack(side="left")

        box = tk.Text(
            parent, height=12, bg=PANEL_BG, fg=TEXT_MAIN,
            font=("Menlo", 11), insertbackground=ACCENT,
            relief="flat", bd=0, padx=12, pady=10,
            highlightbackground=BORDER, highlightthickness=1,
            state="disabled"
        )
        box.pack(fill="both", expand=True, padx=28, pady=(8, 20))

        if prefix == "txt":
            self._txt_result_box = box
        else:
            self._file_result_box = box

    # ── 위젯 팩토리 ──────────────────────────────────────────

    def _btn(self, parent, text, cmd):
        return tk.Button(
            parent, text=text, command=cmd,
            bg=ACCENT, fg="#0a0a0a", font=("Helvetica Neue", 13, "bold"),
            relief="flat", bd=0, padx=28, pady=10, cursor="hand2",
            activebackground="#a8e6ff", activeforeground="#0a0a0a"
        )

    def _small_btn(self, parent, text, cmd):
        return tk.Button(
            parent, text=text, command=cmd,
            bg=BORDER, fg=TEXT_MAIN, font=("Helvetica Neue", 11),
            relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
            activebackground="#444", activeforeground=TEXT_MAIN
        )

    # ── 이벤트 ───────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("지원 파일", "*.docx *.pdf"), ("DOCX", "*.docx"), ("PDF", "*.pdf")]
        )
        if path:
            self._selected_file = path
            self.file_info.config(text=f"✓  {os.path.basename(path)}", fg=GREEN)

    def _convert_file(self):
        path = getattr(self, "_selected_file", None)
        if not path:
            messagebox.showwarning("파일 없음", "파일을 먼저 선택하세요.")
            return
        self._run_in_thread(self._do_convert_file, path)

    def _do_convert_file(self, path):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".docx":
                md = docx_to_md(path)
            elif ext == ".pdf":
                md = pdf_to_md(path)
            else:
                raise ValueError(f"지원하지 않는 형식: {ext}")
            md = clean_markdown(md)
            self._result_md = md
            self._set_result(self._file_result_box, md)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("변환 오류", str(e)))

    def _convert_text(self):
        text = self.text_input.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("입력 없음", "텍스트를 입력하세요.")
            return
        self._run_in_thread(self._do_convert_text, text)

    def _do_convert_text(self, text):
        try:
            md = text_to_md(text)
            md = clean_markdown(md)
            self._txt_result_md = md
            self._set_result(self._txt_result_box, md)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("변환 오류", str(e)))

    def _set_result(self, box, md):
        def _update():
            box.config(state="normal")
            box.delete("1.0", "end")
            box.insert("1.0", md)
            box.config(state="disabled")
        self.after(0, _update)

    def _copy(self, prefix):
        md = getattr(self, "_txt_result_md" if prefix == "txt" else "_result_md", "")
        if not md:
            messagebox.showinfo("없음", "변환 결과가 없습니다.")
            return
        self.clipboard_clear()
        self.clipboard_append(md)
        messagebox.showinfo("복사 완료", "클립보드에 복사되었습니다.")

    def _save(self, prefix):
        md = getattr(self, "_txt_result_md" if prefix == "txt" else "_result_md", "")
        if not md:
            messagebox.showinfo("없음", "변환 결과가 없습니다.")
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt")]
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(md)
            messagebox.showinfo("저장 완료", f"저장되었습니다:\n{save_path}")

    def _run_in_thread(self, fn, *args):
        t = threading.Thread(target=fn, args=args, daemon=True)
        t.start()

    def _check_deps(self):
        missing = []
        try:
            import docx
        except ImportError:
            missing.append("python-docx")
        try:
            import pdfplumber
        except ImportError:
            missing.append("pdfplumber")
        if missing:
            msg = "다음 패키지가 없습니다:\n" + "\n".join(f"  pip install {p}" for p in missing)
            self.after(500, lambda: messagebox.showwarning("패키지 누락", msg))


if __name__ == "__main__":
    app = MDConverterApp()
    app.mainloop()
