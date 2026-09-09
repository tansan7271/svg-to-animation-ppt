#!/usr/bin/env python3
"""svg-to-animation-ppt GUI: SVG 를 드래그 앤 드롭 → 옵션 조절 → 잉크 재생 애니메이션 pptx 생성."""
from __future__ import annotations

import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageFont, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:  # pragma: no cover
    _HAS_DND = False

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import svg_to_animation_ppt as core  # noqa: E402

PREVIEW_W, PREVIEW_H = 520, 260
LIST_W = 34


class App:
    def __init__(self, root):
        self.root = root
        root.title("SVG to Animation PPT")
        root.resizable(False, False)
        self.svg_path: Path | None = None
        self.template_path: Path | None = None
        self._preview_img = None
        self._busy = False
        self.strokes = []          # 원본 획 (문서 순서)
        self.paths = {}            # order -> (name, 획 수)
        self.order = []            # 현재 재생 순서 (원래 order 번호 목록)
        self.reversed = set()      # 방향 뒤집은 path

        pad = {"padx": 12, "pady": 6}
        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        # --- 드롭 영역 ---------------------------------------------------
        self.drop = tk.Label(
            frm, text="여기에 SVG 또는 펜으로 그린 pptx 를 드래그 앤 드롭\n또는 클릭해서 파일 선택",
            relief="ridge", bd=2, width=64, height=5, bg="#F4F6F8", fg="#555", cursor="hand2",
        )
        self.drop.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 8))
        self.drop.bind("<Button-1>", lambda e: self.pick_svg())
        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)

        # --- 미리보기 ----------------------------------------------------
        # 이미지가 없는 Label 은 width/height 단위가 '글자 수' 라서 빈 이미지를 먼저 넣어 픽셀로 고정한다
        self._preview_img = ImageTk.PhotoImage(Image.new("RGB", (PREVIEW_W, PREVIEW_H), "white"))
        self.preview = tk.Label(frm, image=self._preview_img, bg="white", relief="sunken", bd=1)
        self.preview.grid(row=1, column=0, columnspan=3, pady=(0, 8), sticky="n")
        self.info = ttk.Label(frm, text="파일을 등록하면 미리보기가 표시됩니다.", foreground="#666")
        self.info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # --- 그리기 순서 목록 (드래그로 재배치) ---------------------------
        of = ttk.LabelFrame(frm, text="그리기 순서 (드래그해서 바꾸기)", padding=6)
        of.grid(row=1, column=3, rowspan=2, sticky="ns", padx=(10, 0), pady=(0, 10))
        self.order_list = tk.Listbox(of, width=LIST_W, height=12, activestyle="none",
                                     selectmode="browse", exportselection=False)
        self.order_list.grid(row=0, column=0, columnspan=4, sticky="nsew")
        sb = ttk.Scrollbar(of, orient="vertical", command=self.order_list.yview)
        sb.grid(row=0, column=4, sticky="ns")
        self.order_list.config(yscrollcommand=sb.set)
        self.order_list.bind("<ButtonPress-1>", self._drag_start)
        self.order_list.bind("<B1-Motion>", self._drag_motion)
        self.order_list.bind("<ButtonRelease-1>", self._drag_end)
        self.order_list.bind("<<ListboxSelect>>", lambda e: self.redraw_preview())
        ttk.Button(of, text="▲", width=3, command=lambda: self.move_selected(-1)).grid(row=1, column=0, pady=(6, 0))
        ttk.Button(of, text="▼", width=3, command=lambda: self.move_selected(+1)).grid(row=1, column=1, pady=(6, 0))
        ttk.Button(of, text="방향 반전", command=self.toggle_reverse).grid(row=1, column=2, pady=(6, 0), padx=(6, 0))
        ttk.Button(of, text="원래대로", command=self.reset_order).grid(row=1, column=3, pady=(6, 0), padx=(6, 0))
        ttk.Label(of, text="번호 = 시작점, ● = 끝점", foreground="#888").grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self._drag_from = None

        # --- 옵션 --------------------------------------------------------
        opts = ttk.LabelFrame(frm, text="옵션", padding=10)
        opts.grid(row=3, column=0, columnspan=4, sticky="ew")
        opts.columnconfigure(1, weight=1)

        r = 0
        ttk.Label(opts, text="그려지는 시간 (초)").grid(row=r, column=0, sticky="w")
        self.duration = tk.DoubleVar(value=3.0)
        self._scale(opts, self.duration, 0.5, 10.0, 0.5, r).grid(row=r, column=1, sticky="ew", padx=8)
        self.duration_lbl = ttk.Label(opts, width=6)
        self.duration_lbl.grid(row=r, column=2, sticky="w")
        r += 1

        ttk.Label(opts, text="부드럽게 시작 (ease-in)").grid(row=r, column=0, sticky="w")
        self.ease_in = tk.DoubleVar(value=0.2)
        self._scale(opts, self.ease_in, 0.0, 0.5, 0.05, r).grid(row=r, column=1, sticky="ew", padx=8)
        self.ease_in_lbl = ttk.Label(opts, width=6)
        self.ease_in_lbl.grid(row=r, column=2, sticky="w")
        r += 1

        ttk.Label(opts, text="부드럽게 끝 (ease-out)").grid(row=r, column=0, sticky="w")
        self.ease_out = tk.DoubleVar(value=0.2)
        self._scale(opts, self.ease_out, 0.0, 0.5, 0.05, r).grid(row=r, column=1, sticky="ew", padx=8)
        self.ease_out_lbl = ttk.Label(opts, width=6)
        self.ease_out_lbl.grid(row=r, column=2, sticky="w")
        r += 1

        # 이징 곡선 미니 그래프
        self.curve = tk.Canvas(opts, width=140, height=90, bg="white", highlightthickness=1,
                               highlightbackground="#CCC")
        self.curve.grid(row=0, column=3, rowspan=3, padx=(12, 0))

        ttk.Label(opts, text="로고 가로 크기 (cm)").grid(row=r, column=0, sticky="w", pady=(8, 0))
        self.width_cm = tk.StringVar(value="")
        e = ttk.Entry(opts, textvariable=self.width_cm, width=10)
        e.grid(row=r, column=1, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(opts, text="비우면 슬라이드 폭의 60%", foreground="#888").grid(row=r, column=2, columnspan=2, sticky="w", pady=(8, 0))
        r += 1

        ttk.Label(opts, text="펜 굵기 (cm)").grid(row=r, column=0, sticky="w")
        self.stroke_w = tk.StringVar(value="")
        ttk.Entry(opts, textvariable=self.stroke_w, width=10).grid(row=r, column=1, sticky="w", padx=8)
        ttk.Label(opts, text="비우면 SVG stroke-width 비례", foreground="#888").grid(row=r, column=2, columnspan=2, sticky="w")
        r += 1

        ttk.Label(opts, text="색 강제").grid(row=r, column=0, sticky="w")
        cf = ttk.Frame(opts)
        cf.grid(row=r, column=1, columnspan=3, sticky="w", padx=8)
        self.color = tk.StringVar(value="")
        ttk.Entry(cf, textvariable=self.color, width=10).pack(side="left")
        ttk.Button(cf, text="선택…", command=self.pick_color, width=6).pack(side="left", padx=4)
        ttk.Button(cf, text="지우기", command=lambda: self.color.set(""), width=6).pack(side="left")
        ttk.Label(cf, text="비우면 SVG 색 그대로", foreground="#888").pack(side="left", padx=8)
        r += 1

        self.keep_rhythm = tk.BooleanVar(value=False)
        self.rhythm_chk = ttk.Checkbutton(
            opts, text="손그림 리듬 유지 (pptx 잉크 입력일 때: 그린 속도 그대로, 끄면 균일 속도)",
            variable=self.keep_rhythm, command=self._reload_current, state="disabled")
        self.rhythm_chk.grid(row=r, column=0, columnspan=4, sticky="w", pady=(8, 0))
        r += 1

        self.split = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="SVG path 마다 잉크 개체를 나눠 순서대로 재생 (애니메이션 창에서 개별 조정 가능)",
                        variable=self.split).grid(row=r, column=0, columnspan=4, sticky="w", pady=(2, 0))
        r += 1

        tf = ttk.Frame(opts)
        tf.grid(row=r, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Button(tf, text="템플릿 pptx…", command=self.pick_template).pack(side="left")
        self.template_lbl = ttk.Label(tf, text="없음 (새 16:9 빈 프레젠테이션)", foreground="#666")
        self.template_lbl.pack(side="left", padx=8)
        ttk.Button(tf, text="×", width=2, command=self.clear_template).pack(side="left")
        r += 1

        self.open_after = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="완료 후 PowerPoint 로 열기", variable=self.open_after).grid(
            row=r, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # --- 버튼 --------------------------------------------------------
        bf = ttk.Frame(frm)
        bf.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        self.status = ttk.Label(bf, text="", foreground="#666")
        self.status.pack(side="left")
        self.go = ttk.Button(bf, text="완료 → pptx 만들기", command=self.run, state="disabled")
        self.go.pack(side="right")

        for v in (self.duration, self.ease_in, self.ease_out):
            v.trace_add("write", lambda *_: self.refresh_labels())
        self.refresh_labels()

    # ------------------------------------------------------------------
    def _scale(self, parent, var, lo, hi, step, row):
        s = ttk.Scale(parent, from_=lo, to=hi, variable=var, orient="horizontal", length=260,
                      command=lambda v, var=var, step=step: var.set(round(float(v) / step) * step))
        return s

    def refresh_labels(self):
        self.duration_lbl.config(text=f"{self.duration.get():.1f}s")
        self.ease_in_lbl.config(text=f"{self.ease_in.get():.2f}")
        self.ease_out_lbl.config(text=f"{self.ease_out.get():.2f}")
        self.draw_curve()

    def draw_curve(self):
        c = self.curve
        c.delete("all")
        W, H, m = 140, 90, 8
        c.create_line(m, H - m, W - m, H - m, fill="#DDD")
        c.create_line(m, H - m, m, m, fill="#DDD")
        c.create_line(m, H - m, W - m, m, fill="#EEE", dash=(2, 2))
        pts = core.ease_curve(self.ease_in.get(), self.ease_out.get(), n=40)
        coords = []
        for t, v in pts:
            coords += [m + t * (W - 2 * m), H - m - v * (H - 2 * m)]
        c.create_line(*coords, fill="#2A7BDE", width=2, smooth=False)

    # ------------------------------------------------------------------
    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith((".svg", ".pptx")):
                self.load_svg(Path(f))
                return
        messagebox.showwarning("지원하지 않는 파일", "SVG 또는 pptx 파일을 놓아주세요.")

    def pick_svg(self):
        f = filedialog.askopenfilename(title="SVG 또는 pptx 선택",
                                       filetypes=[("SVG / PowerPoint 잉크", "*.svg *.pptx"), ("모든 파일", "*")])
        if f:
            self.load_svg(Path(f))

    def pick_template(self):
        f = filedialog.askopenfilename(title="템플릿 pptx", filetypes=[("PowerPoint", "*.pptx")])
        if f:
            self.template_path = Path(f)
            self.template_lbl.config(text=self.template_path.name)

    def clear_template(self):
        self.template_path = None
        self.template_lbl.config(text="없음 (새 16:9 빈 프레젠테이션)")

    def pick_color(self):
        rgb, hexv = colorchooser.askcolor(color=self.color.get() or "#000000", title="펜 색")
        if hexv:
            self.color.set(hexv.upper())

    # ------------------------------------------------------------------
    def load_svg(self, path: Path):
        try:
            strokes = core.load_strokes_any(path, None, self.keep_rhythm.get())
        except SystemExit as e:
            messagebox.showerror("변환 실패", str(e))
            return
        except Exception as e:
            messagebox.showerror("SVG 읽기 실패", f"{type(e).__name__}: {e}")
            return
        same_file = (self.svg_path == path and set(self.order) == {st.order for st in strokes})
        self.svg_path = path
        self.strokes = strokes
        self.paths = {}
        for st in strokes:
            name, n = self.paths.get(st.order, (st.name, 0))
            self.paths[st.order] = (name, n + 1)
        if not same_file:
            self.order = sorted(self.paths)
            self.reversed = set()
        self.rhythm_chk.config(state="normal" if path.suffix.lower() == ".pptx" else "disabled")
        self.fill_list()
        self.redraw_preview()
        self.drop.config(text=f"등록됨: {path.name}", bg="#E8F4EA", fg="#1E6B32")
        self.info.config(text=f"path {len(self.paths)}개, 획 {len(strokes)}개")
        self.go.config(state="normal")
        self.status.config(text="")

    # --- 순서 목록 ---------------------------------------------------------
    def fill_list(self, select: int | None = None):
        lb = self.order_list
        lb.delete(0, "end")
        for i, o in enumerate(self.order):
            name, n = self.paths[o]
            rev = "  ↩" if o in self.reversed else ""
            lb.insert("end", f"{i+1:>2}. {name}  (획 {n}){rev}")
        if select is not None and 0 <= select < len(self.order):
            lb.selection_set(select)
            lb.see(select)

    def _selected_index(self) -> int | None:
        sel = self.order_list.curselection()
        return int(sel[0]) if sel else None

    def _drag_start(self, e):
        self._drag_from = self.order_list.nearest(e.y)

    def _drag_motion(self, e):
        if self._drag_from is None or not self.order:
            return
        to = self.order_list.nearest(e.y)
        if to != self._drag_from and 0 <= to < len(self.order):
            item = self.order.pop(self._drag_from)
            self.order.insert(to, item)
            self._drag_from = to
            self.fill_list(select=to)

    def _drag_end(self, e):
        self._drag_from = None
        self.redraw_preview()

    def move_selected(self, delta: int):
        i = self._selected_index()
        if i is None:
            return
        j = i + delta
        if 0 <= j < len(self.order):
            self.order[i], self.order[j] = self.order[j], self.order[i]
            self.fill_list(select=j)
            self.redraw_preview()

    def toggle_reverse(self):
        i = self._selected_index()
        if i is None:
            return
        o = self.order[i]
        self.reversed ^= {o}
        self.fill_list(select=i)
        self.redraw_preview()

    def reset_order(self):
        self.order = sorted(self.paths)
        self.reversed = set()
        self.fill_list()
        self.redraw_preview()

    # --- 미리보기 렌더 ------------------------------------------------------
    def redraw_preview(self):
        if not self.strokes:
            return
        strokes = core.apply_order(self.strokes, self.order, self.reversed)
        sel = self._selected_index()
        xs = [x for st in strokes for x, _ in st.points]
        ys = [y for st in strokes for _, y in st.points]
        minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
        pad = 24
        sc = min((PREVIEW_W - 2 * pad) / max(maxx - minx, 1e-9), (PREVIEW_H - 2 * pad) / max(maxy - miny, 1e-9))
        ox = (PREVIEW_W - (maxx - minx) * sc) / 2
        oy = (PREVIEW_H - (maxy - miny) * sc) / 2
        ss = 2
        img = Image.new("RGB", (PREVIEW_W * ss, PREVIEW_H * ss), "white")
        d = ImageDraw.Draw(img)

        def P(pt):
            return ((ox + (pt[0] - minx) * sc) * ss, (oy + (pt[1] - miny) * sc) * ss)

        for st in strokes:
            pts = [P(p) for p in st.points]
            w = max(2, int(round(st.width * sc * ss)))
            color = st.color
            if sel is not None and st.order != sel:
                color = "#D0D0D0"
            d.line(pts, fill=color, width=w)
            r = w / 2
            if w > 3 * ss:  # 굵은 선은 이음새가 벌어지므로 점마다 원을 찍어 메운다
                for (x, y) in pts:
                    d.ellipse([x - r, y - r, x + r, y + r], fill=color)
        # 번호(시작점) 와 끝점 표시: path 별 첫 획 시작 / 마지막 획 끝
        try:
            font = ImageFont.load_default(size=13 * ss)
        except Exception:
            font = ImageFont.load_default()
        seen = {}
        for st in strokes:
            seen.setdefault(st.order, [st, st])[1] = st
        for o, (first, last) in seen.items():
            hi = (sel is None) or (o == sel)
            fg = "#2A7BDE" if hi else "#BBB"
            ex, ey = P(last.points[-1])
            rr = 4 * ss
            d.ellipse([ex - rr, ey - rr, ex + rr, ey + rr], fill=fg, outline="white", width=ss)
            sx, sy = P(first.points[0])
            R = 10 * ss
            d.ellipse([sx - R, sy - R, sx + R, sy + R], fill=fg, outline="white", width=ss)
            label = str(o + 1)
            bb = d.textbbox((0, 0), label, font=font)
            d.text((sx - (bb[2] - bb[0]) / 2, sy - (bb[3] - bb[1]) / 2 - bb[1]), label, fill="white", font=font)
        img = img.resize((PREVIEW_W, PREVIEW_H), Image.LANCZOS)
        self._preview_img = ImageTk.PhotoImage(img)
        self.preview.config(image=self._preview_img)

    def _reload_current(self):
        if self.svg_path:
            self.load_svg(self.svg_path)

    def _float_or_none(self, var, name):
        v = var.get().strip()
        if not v:
            return None
        try:
            return float(v)
        except ValueError:
            raise ValueError(f"{name} 값이 숫자가 아니에요: {v!r}")

    def run(self):
        if self._busy or not self.svg_path:
            return
        try:
            opt = core.ConvertOptions(
                template=self.template_path,
                width_cm=self._float_or_none(self.width_cm, "가로 크기"),
                duration=float(self.duration.get()),
                stroke_width=self._float_or_none(self.stroke_w, "펜 굵기"),
                color=(self.color.get().strip() or None),
                split="path" if self.split.get() else "none",
                ease_in=float(self.ease_in.get()),
                ease_out=float(self.ease_out.get()),
                path_order=list(self.order),
                reversed_paths=set(self.reversed),
                keep_rhythm=bool(self.keep_rhythm.get()),
            )
        except ValueError as e:
            messagebox.showerror("옵션 오류", str(e))
            return
        out = filedialog.asksaveasfilename(
            title="pptx 저장", defaultextension=".pptx", initialfile=self.svg_path.stem + ".pptx",
            initialdir=str(self.svg_path.parent), filetypes=[("PowerPoint", "*.pptx")])
        if not out:
            return
        out_path = Path(out)
        self._busy = True
        self.go.config(state="disabled")
        self.status.config(text="변환 중…")

        def work():
            try:
                r = core.convert(self.svg_path, out_path, opt)
                self.root.after(0, lambda: self.done(r))
            except Exception as e:  # noqa: BLE001
                self.root.after(0, lambda: self.fail(e))

        threading.Thread(target=work, daemon=True).start()

    def done(self, r):
        self._busy = False
        self.go.config(state="normal")
        self.status.config(text=f"완료: {r.output.name}  (잉크 {r.ink_objects}개, 획 {r.strokes}개, "
                                f"{r.width_cm:.1f}×{r.height_cm:.1f} cm)")
        if self.open_after.get() and sys.platform == "darwin":
            subprocess.Popen(["open", "-a", "Microsoft PowerPoint", str(r.output)])

    def fail(self, e):
        self._busy = False
        self.go.config(state="normal")
        self.status.config(text="실패")
        messagebox.showerror("변환 실패", f"{type(e).__name__}: {e}")


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
