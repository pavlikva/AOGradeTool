from __future__ import annotations

import csv
import dataclasses
import datetime as dt
import json
import os
import sys
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BG_COLOR = "#eef2f7"
FG_LABEL = "#1f2937"
FG_COLOR = "#111827"
QUESTION_DONE_BG = "#dcfce7"
QUESTION_DONE_SELECT_BG = "#86efac"
QUESTION_ACTIVE_SELECT_BG = "#cfe3ff"
STUDENT_NONE_BG = "#ffffff"
STUDENT_PARTIAL_BG = "#fef3c7"
STUDENT_DONE_BG = "#dcfce7"

@dataclass
class Bucket:
    bid: str
    label: str
    points: float
    key: str = ""
    mode: str = "add"  # add | set


@dataclass
class QuestionConfig:
    qid: str
    max_points: float
    buckets: list[Bucket] = field(default_factory=list)


@dataclass
class Anchor:
    page_index: int
    x_ratio: float
    y_ratio: float


@dataclass
class CellValue:
    score: Optional[float] = None
    note: str = ""
    applied_bucket_ids: list[str] = field(default_factory=list)


@dataclass
class RenderedPDFBundle:
    page_images: list[Image.Image]
    thumb_images: list[Image.Image]


@dataclass
class PDFRenderRequest:
    request_id: int
    path: Path
    display_width: int
    zoom_factor: float
    source_bytes: Optional[bytes] = None

def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat()



def safe_float(text: str) -> Optional[float]:
    text = text.strip()
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_shortcut_key_event(event: tk.Event) -> Optional[str]:
    """Normalize keyboard or numpad digit input to a single shortcut key."""
    keysym = getattr(event, "keysym", "")
    char = getattr(event, "char", "")
    if char and len(char) == 1 and char.isdigit():
        return char
    if isinstance(keysym, str) and keysym.startswith("KP_") and keysym[3:].isdigit():
        return keysym[3:]
    if keysym in {"BackSpace", "Delete", "Escape"}:
        return ""
    return None



def bind_mousewheel_recursive(widget: tk.Misc, target: tk.Widget) -> None:
    def _on_mousewheel(event: tk.Event) -> str:
        delta = 0
        if getattr(event, "delta", 0):
            delta = -1 * int(event.delta / 120)
        elif getattr(event, "num", None) == 5:
            delta = 1
        elif getattr(event, "num", None) == 4:
            delta = -1
        if delta:
            try:
                target.yview_scroll(delta, "units")
            except Exception:
                pass
        return "break"

    widget.bind("<MouseWheel>", _on_mousewheel, add="+")
    widget.bind("<Button-4>", _on_mousewheel, add="+")
    widget.bind("<Button-5>", _on_mousewheel, add="+")
    for child in widget.winfo_children():
        bind_mousewheel_recursive(child, target)


class ToggleDropdownButton(ttk.Frame):
    _active = None

    def __init__(self, master: tk.Misc, label: str, items: list[tuple[str, object]]):
        super().__init__(master)
        self._label = label
        self._posted = False
        self.button = tk.Button(
            self,
            text=f"{label} ▾",
            relief="raised",
            padx=10,
            pady=3,
            command=self.toggle,
            anchor="center",
        )
        self.button.grid(row=0, column=0, sticky="ew")
        self.columnconfigure(0, weight=1)
        self.menu = tk.Menu(self, tearoff=0)
        for item_label, command in items:
            if item_label == "__separator__":
                self.menu.add_separator()
            else:
                self.menu.add_command(label=item_label, command=self._wrap(command))
        self.button.bind("<Destroy>", lambda _e: self.close(), add="+")
        self.menu.bind("<Unmap>", lambda _e: self._set_closed(), add="+")

    def _wrap(self, command):
        def _cmd():
            self.close()
            if command is not None:
                command()
        return _cmd

    def _set_closed(self) -> None:
        if ToggleDropdownButton._active is self:
            ToggleDropdownButton._active = None
        self._posted = False
        self.button.configure(text=f"{self._label} ▾")

    def close(self) -> None:
        if self._posted:
            try:
                self.menu.unpost()
            except Exception:
                pass
        self._set_closed()

    def toggle(self) -> None:
        if ToggleDropdownButton._active is not None and ToggleDropdownButton._active is not self:
            ToggleDropdownButton._active.close()
        if self._posted:
            self.close()
            return
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height()
        self.menu.post(x, y)
        try:
            self.menu.grab_release()
        except Exception:
            pass
        self._posted = True
        ToggleDropdownButton._active = self
        self.button.configure(text=f"{self._label} ▴")


class QuestionDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc, questions: list[QuestionConfig]):
        super().__init__(master)
        self.title("Questions")
        self.transient(master)
        self.grab_set()
        self.geometry("760x680")
        self.minsize(720, 620)
        self.result: Optional[list[QuestionConfig]] = None
        self.questions = [dataclasses.replace(q, buckets=[dataclasses.replace(b) for b in q.buckets]) for q in questions]
        self.selected_index: Optional[int] = None
        self._bucket_rows: list[dict[str, object]] = []

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=0)
        right.rowconfigure(4, weight=1)

        ttk.Label(left, text="Questions").grid(row=0, column=0, sticky="w")
        self.listbox = tk.Listbox(left, height=14)
        self.listbox.grid(row=1, column=0, sticky="nsew")
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        list_scroll = ttk.Scrollbar(left, orient="vertical", command=self.listbox.yview)
        list_scroll.grid(row=1, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=list_scroll.set)

        btns = ttk.Frame(left)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="Add", command=self.add_question).pack(side="left")
        ttk.Button(btns, text="Delete", command=self.delete_question).pack(side="left", padx=(6, 0))

        ttk.Label(right, text="Question ID").grid(row=0, column=0, sticky="w")
        ttk.Label(right, text="Max points").grid(row=1, column=0, sticky="w")
        ttk.Entry(right, textvariable=tk.StringVar()).grid_forget()
        self.qid_var = tk.StringVar()
        self.max_points_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.qid_var, width=20).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Entry(right, textvariable=self.max_points_var, width=10).grid(row=1, column=1, sticky="w", pady=2)

        ttk.Label(right, text="Buckets").grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        bucket_container = ttk.Frame(right)
        bucket_container.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        bucket_container.rowconfigure(0, weight=1)
        bucket_container.columnconfigure(0, weight=1)

        self.bucket_canvas = tk.Canvas(bucket_container, height=320, highlightthickness=0)
        self.bucket_canvas.grid(row=0, column=0, sticky="nsew")
        bucket_scroll = ttk.Scrollbar(bucket_container, orient="vertical", command=self.bucket_canvas.yview)
        bucket_scroll.grid(row=0, column=1, sticky="ns")
        self.bucket_canvas.configure(yscrollcommand=bucket_scroll.set)

        self.bucket_rows_frame = ttk.Frame(self.bucket_canvas)
        self.bucket_window = self.bucket_canvas.create_window((0, 0), window=self.bucket_rows_frame, anchor="nw")
        self.bucket_rows_frame.bind("<Configure>", lambda _e: self.bucket_canvas.configure(scrollregion=self.bucket_canvas.bbox("all")))
        self.bucket_canvas.bind("<Configure>", self._on_bucket_canvas_configure)
        bind_mousewheel_recursive(bucket_container, self.bucket_canvas)

        header = ttk.Frame(right)
        header.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        header.columnconfigure(0, weight=0)
        header.columnconfigure(1, weight=0)
        header.columnconfigure(2, weight=0)
        header.columnconfigure(3, weight=0)
        header.columnconfigure(4, weight=0)
        ttk.Label(header, text="Bucket description", width=24).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        ttk.Label(header, text="points", width=6).grid(row=0, column=1, sticky="w", padx=(0, 6), pady=2)
        ttk.Label(header, text="key", width=4).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        ttk.Label(header, text="mode", width=5).grid(row=0, column=3, sticky="w", padx=(0, 6), pady=2)
        ttk.Label(header, text="delete", width=3).grid(row=0, column=4, sticky="w", padx=(0, 6), pady=2)

        form_btns = ttk.Frame(right)
        form_btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(form_btns, text="Add Bucket", command=self.add_bucket_row).pack(side="left")
        ttk.Button(form_btns, text="Apply", command=self.apply_current).pack(side="left", padx=(6, 0))
        ttk.Button(form_btns, text="OK", command=self.ok).pack(side="left", padx=(6, 0))
        ttk.Button(form_btns, text="Cancel", command=self.cancel).pack(side="left", padx=(6, 0))

        self._populate_question_list()
        if self.questions:
            self.listbox.selection_set(0)
            self.on_select()

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.focus_force()
        self.bind("<KeyPress>", self._on_key_press)

    def _on_bucket_canvas_configure(self, event: tk.Event) -> None:
        self.bucket_canvas.itemconfig(self.bucket_window, width=event.width)
        try:
            self.bucket_frame.configure(width=event.width)
        except Exception:
            pass
        wrap = max(180, event.width - 28)
        for btn in self.bucket_buttons.values():
            try:
                btn.configure(wraplength=wrap)
            except Exception:
                pass
        try:
            self.bucket_rows_frame.configure(width=event.width)
        except Exception:
            pass

    def _populate_question_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for q in self.questions:
            self.listbox.insert(tk.END, f"{q.qid}  ({q.max_points:g})")

    def _clear_bucket_rows(self) -> None:
        for row in self._bucket_rows:
            frame = row.get("frame")
            if isinstance(frame, tk.Widget):
                frame.destroy()
        self._bucket_rows.clear()

    def _make_bucket_row(self, bucket: Bucket, removable: bool) -> None:
        row = ttk.Frame(self.bucket_rows_frame)
        row.grid_columnconfigure(0, weight=0)
        row.grid_columnconfigure(1, weight=0)
        row.grid_columnconfigure(2, weight=0)
        row.grid_columnconfigure(3, weight=0)
        row.grid_columnconfigure(4, weight=0)

        label_var = tk.StringVar(value=bucket.label)
        points_var = tk.StringVar(value=f"{bucket.points:g}")
        key_var = tk.StringVar(value=bucket.key)
        mode_var = tk.StringVar(value=bucket.mode if bucket.mode in {"add", "set"} else "add")

        ttk.Entry(row, textvariable=label_var, width=24).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        ttk.Entry(row, textvariable=points_var, width=6).grid(row=0, column=1, sticky="w", padx=(0, 6), pady=2)
        ttk.Entry(row, textvariable=key_var, width=4).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        ttk.Combobox(row, textvariable=mode_var, values=("add", "set"), width=5, state="readonly").grid(row=0, column=3, sticky="w", padx=(0, 6), pady=2)
        if removable:
            ttk.Button(row, text="-", width=3, command=lambda: self.remove_bucket_row(row)).grid(row=0, column=4, sticky="e", pady=2)
        else:
            ttk.Label(row, text=" ", width=3).grid(row=0, column=4, sticky="e", pady=2)

        self._bucket_rows.append({
            "frame": row,
            "id": bucket.bid,
            "label_var": label_var,
            "points_var": points_var,
            "key_var": key_var,
            "mode_var": mode_var,
        })
        row.grid(row=len(self._bucket_rows) - 1, column=0, sticky="ew")
        row.columnconfigure(0, weight=1)
        bind_mousewheel_recursive(row, self.bucket_canvas)

    def _bind_single_key_entry(self, entry: tk.Widget, var: tk.StringVar) -> None:
        def _on_keypress(event: tk.Event) -> str:
            key = normalize_shortcut_key_event(event)
            if key is None:
                return "break"
            var.set(key)
            return "break"

        entry.bind("<KeyPress>", _on_keypress, add="+")
        entry.bind("<FocusIn>", lambda _e: entry.after_idle(lambda: entry.icursor(tk.END)), add="+")

    def _reflow_bucket_rows(self) -> None:
        for idx, row in enumerate(self._bucket_rows):
            frame = row["frame"]
            if isinstance(frame, tk.Widget):
                frame.grid(row=idx, column=0, sticky="ew")
                for child in frame.winfo_children():
                    if isinstance(child, ttk.Button) and child.cget("text") == "x":
                        if idx == 0:
                            child.grid_remove()
                        else:
                            child.grid()
        self.bucket_canvas.update_idletasks()
        self.bucket_canvas.configure(scrollregion=self.bucket_canvas.bbox("all"))

    def add_bucket_row(self, bucket: Optional[Bucket] = None) -> None:
        if bucket is None:
            bucket = Bucket(bid=uuid.uuid4().hex, label="", points=0.0, key="", mode="add")
        self._make_bucket_row(bucket, removable=len(self._bucket_rows) > 0)
        self._reflow_bucket_rows()

    def remove_bucket_row(self, row_frame: tk.Widget) -> None:
        if len(self._bucket_rows) <= 1:
            return
        kept: list[dict[str, object]] = []
        for row in self._bucket_rows:
            frame = row["frame"]
            if frame is row_frame:
                if isinstance(frame, tk.Widget):
                    frame.destroy()
                continue
            kept.append(row)
        self._bucket_rows = kept
        self._reflow_bucket_rows()

    def _default_bucket(self) -> Bucket:
        return Bucket(bid=uuid.uuid4().hex, label="not answered/wrong", points=0.0, key="0", mode="set")

    def on_select(self, _event=None) -> None:
        idxs = self.listbox.curselection()
        if not idxs:
            return
        idx = idxs[0]
        self.selected_index = idx
        q = self.questions[idx]
        self.qid_var.set(q.qid)
        self.max_points_var.set(str(q.max_points))
        self._clear_bucket_rows()
        buckets = q.buckets if q.buckets else [self._default_bucket()]
        for i, bucket in enumerate(buckets):
            self._make_bucket_row(bucket, removable=i > 0)
        self._reflow_bucket_rows()

    def _focus_is_editable(self) -> bool:
        widget = self.focus_get()
        return isinstance(widget, (tk.Entry, tk.Text))

    def _on_key_press(self, event: tk.Event) -> None:
        key = getattr(event, "char", "")
        if key.isdigit() and not self._focus_is_editable():
            self._toggle_bucket_by_key(key)

    def _toggle_bucket_by_key(self, key: str) -> None:
        master = getattr(self, "master", None)
        if master is not None and hasattr(master, "toggle_bucket_by_key"):
            try:
                master.toggle_bucket_by_key(key)
            except Exception:
                pass

    def parse_buckets(self) -> list[Bucket]:
        buckets: list[Bucket] = []
        used_keys: set[str] = set()
        for row in self._bucket_rows:
            label_var = row.get("label_var")
            points_var = row.get("points_var")
            key_var = row.get("key_var")
            mode_var = row.get("mode_var")
            bid = str(row.get("id", uuid.uuid4().hex))
            if not all(isinstance(v, tk.StringVar) for v in [label_var, points_var, key_var, mode_var]):
                continue
            label = label_var.get().strip()  # type: ignore[union-attr]
            points_raw = points_var.get().strip()  # type: ignore[union-attr]
            key = key_var.get().strip()  # type: ignore[union-attr]
            mode = mode_var.get().strip()  # type: ignore[union-attr]
            if not label and not points_raw and not key:
                continue
            if not label:
                raise ValueError("Each bucket needs a description.")
            points = safe_float(points_raw)
            if points is None:
                raise ValueError(f"Invalid points for bucket {label!r}: {points_raw!r}")
            if key:
                if len(key) != 1 or not key.isdigit():
                    raise ValueError(f"Shortcut key for {label!r} must be a single digit 0-9.")
                if key in used_keys:
                    raise ValueError(f"Duplicate shortcut key {key!r} in the same question.")
                used_keys.add(key)
            if mode not in {"add", "set"}:
                mode = "add"
            buckets.append(Bucket(bid=bid, label=label, points=points, key=key, mode=mode))
        if not buckets:
            buckets.append(self._default_bucket())
        return buckets

    def apply_current(self) -> None:
        if self.selected_index is None:
            self.add_question()
            return
        qid = self.qid_var.get().strip()
        if not qid:
            messagebox.showerror("Invalid question", "Question ID cannot be blank.", parent=self)
            return
        max_points = safe_float(self.max_points_var.get())
        if max_points is None:
            messagebox.showerror("Invalid points", "Max points must be numeric.", parent=self)
            return
        try:
            buckets = self.parse_buckets()
        except Exception as exc:
            messagebox.showerror("Invalid buckets", str(exc), parent=self)
            return
        self.questions[self.selected_index] = QuestionConfig(qid=qid, max_points=max_points, buckets=buckets)
        self._populate_question_list()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(self.selected_index)
        self.result = self.questions

    def add_question(self) -> None:
        q = QuestionConfig(qid=f"Q{len(self.questions)+1}", max_points=10.0, buckets=[self._default_bucket()])
        self.questions.append(q)
        self._populate_question_list()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(tk.END)
        self.on_select()
        self.selected_index = len(self.questions) - 1
        self.result = self.questions

    def delete_question(self) -> None:
        idxs = self.listbox.curselection()
        if not idxs:
            return
        idx = idxs[0]
        if messagebox.askyesno("Delete question", "Delete the selected question?", parent=self):
            del self.questions[idx]
            self._populate_question_list()
            self.result = self.questions
            if self.questions:
                self.listbox.selection_set(max(0, min(idx, len(self.questions) - 1)))
                self.on_select()

    def ok(self) -> None:
        self.apply_current()
        if self.result is None:
            return
        self.destroy()

    def cancel(self) -> None:
        self.result = None
        self.destroy()


class OfflineGraderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Astronomy Olympiad PDF Grader")
        self.root.geometry("1500x950")

        self.submission_dir: Optional[Path] = None
        self.clean_pdf: Optional[Path] = None
        self.solution_pdf: Optional[Path] = None
        self.csv_path: Optional[Path] = None
        self.schema_path: Optional[Path] = None

        self.students: list[str] = []
        self.current_student_index: int = -1
        self.questions: list[QuestionConfig] = []
        self.current_question_index: int = -1
        self.anchors: dict[str, Anchor] = {}
        self.grades: dict[str, dict[str, CellValue]] = {}
        self.status_map: dict[str, str] = {}
        self.last_saved: dict[str, str] = {}

        self.pdf_doc: Optional[fitz.Document] = None
        self.current_pdf_path: Optional[Path] = None
        self.current_pdf_bytes: Optional[bytes] = None
        self.page_positions: list[dict[str, object]] = []
        self.page_photos: list[ImageTk.PhotoImage] = []
        self.thumb_photos: list[ImageTk.PhotoImage] = []
        self.render_after_id: Optional[str] = None
        self.anchor_mode = False
        self.pending_anchor_question: Optional[str] = None
        self.loading_pdf = False
        self.zoom_factor = 1.0
        self.min_zoom = 0.5
        self.max_zoom = 3.0
        self._prefetch_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="pdf-prefetch")
        self._prefetch_cache: dict[Path, bytes] = {}
        self._prefetch_inflight: set[Path] = set()
        self._render_prefetch_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pdf-render")
        self._render_cache: dict[tuple[str, int, float], RenderedPDFBundle] = {}
        self._render_inflight: set[tuple[str, int, float]] = set()
        self._view_render_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pdf-view-render")
        self._view_render_request_id = 0
        self._view_render_after_id: Optional[str] = None
        self._view_render_pending_key: Optional[tuple[str, int, float]] = None
        self._view_render_pending_request: Optional[PDFRenderRequest] = None
        self._preview_render_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pdf-preview-render")

        self.bucket_buttons: dict[str, tk.Button] = {}
        self.bucket_button_specs: dict[str, Bucket] = {}
        self.bucket_button_base_bg: dict[str, str] = {}

        self._configure_style()
        self._set_app_icon()
        self._build_ui()
        self._bind_shortcuts()
        self._update_status("Load a submission folder to begin.")

    def _current_display_width(self) -> int:
        return max(650, self.canvas.winfo_width() - 25)

    def _render_cache_key(self, path: Path, display_width: int, zoom_factor: float) -> tuple[str, int, float]:
        return (str(path.resolve()), int(display_width), round(float(zoom_factor), 3))

    def _render_pdf_bundle(self, path: Path, display_width: int, zoom_factor: float, source_bytes: Optional[bytes] = None) -> RenderedPDFBundle:
        doc = fitz.open(stream=source_bytes, filetype="pdf") if source_bytes is not None else fitz.open(str(path))
        try:
            page_images: list[Image.Image] = []
            thumb_images: list[Image.Image] = []
            thumb_width = 150

            for i in range(doc.page_count):
                page = doc.load_page(i)

                scale = (display_width / page.rect.width) * zoom_factor
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                page_images.append(img)

                timg = img.copy()
                timg.thumbnail((thumb_width, int(thumb_width * img.height / max(1, img.width))))
                thumb_images.append(timg)

            return RenderedPDFBundle(page_images=page_images, thumb_images=thumb_images)
        finally:
            doc.close()

    def _queue_student_pdf_render_prefetch(self, index: int) -> None:
        if index < 0 or index >= len(self.students):
            return

        pdf_path = self._student_key_to_path(self.students[index])
        if pdf_path is None:
            return

        display_width = self._current_display_width()
        key = self._render_cache_key(pdf_path, display_width, self.zoom_factor)
        if key in self._render_cache or key in self._render_inflight:
            return

        self._render_inflight.add(key)

        future = self._render_prefetch_executor.submit(
            self._render_pdf_bundle,
            pdf_path,
            display_width,
            self.zoom_factor,
            None,
        )

        def _done(fut, cache_key=key):
            self._render_inflight.discard(cache_key)
            try:
                bundle = fut.result()
            except Exception:
                return
            self._render_cache[cache_key] = bundle

        future.add_done_callback(_done)

    def _cancel_view_render(self) -> None:
        self._view_render_request_id += 1
        if self._view_render_after_id is not None:
            try:
                self.root.after_cancel(self._view_render_after_id)
            except Exception:
                pass
            self._view_render_after_id = None
        self._view_render_pending_key = None
        self._view_render_pending_request = None

    def _show_pdf_loading_placeholder(self, text: str = "Rendering PDF preview...") -> None:
        self.canvas.delete("all")
        self.page_positions = []
        self.page_photos = []
        self.thumb_photos = []
        for child in self.thumb_frame.winfo_children():
            child.destroy()
        self.canvas.configure(scrollregion=(0, 0, 700, 500))
        self.canvas.create_text(20, 20, anchor="nw", text=text, fill=FG_LABEL, font=("Segoe UI", 11))

    def _begin_view_render(self, path: Path, source_bytes: Optional[bytes], display_width: int, zoom_factor: float) -> None:
        self._cancel_view_render()
        request_id = self._view_render_request_id
        cache_key = self._render_cache_key(path, display_width, zoom_factor)
        if cache_key in self._render_cache:
            bundle = self._render_cache[cache_key]
            self._apply_render_bundle_to_canvas(bundle, display_width)
            return

        self.loading_pdf = True
        self._view_render_pending_key = cache_key
        request = PDFRenderRequest(
            request_id=request_id,
            path=path,
            display_width=display_width,
            zoom_factor=zoom_factor,
            source_bytes=source_bytes,
        )
        self._view_render_pending_request = request
        self._show_pdf_loading_placeholder()

        future = self._view_render_executor.submit(
            self._render_pdf_bundle,
            path,
            display_width,
            zoom_factor,
            source_bytes,
        )

        def _done(fut, req=request, key=cache_key):
            try:
                bundle = fut.result()
            except Exception as exc:
                def _show_error() -> None:
                    if req.request_id != self._view_render_request_id:
                        return
                    self.loading_pdf = False
                    self._view_render_pending_request = None
                    self._view_render_pending_key = None
                    messagebox.showerror("PDF render failed", str(exc))

                self.root.after(0, _show_error)
                return

            def _apply() -> None:
                if req.request_id != self._view_render_request_id:
                    return
                self.loading_pdf = False
                self._view_render_pending_request = None
                self._view_render_pending_key = None
                self._render_cache[key] = bundle
                self._apply_render_bundle_to_canvas(bundle, display_width)

            self.root.after(0, _apply)

        future.add_done_callback(_done)

    def _apply_render_bundle_to_canvas(self, bundle: RenderedPDFBundle, display_width: int) -> None:
        self.loading_pdf = False
        self._view_render_pending_key = None
        self._view_render_pending_request = None
        self.canvas.delete("all")
        self.page_positions = []
        self.page_photos = []
        self.thumb_photos = []

        for child in self.thumb_frame.winfo_children():
            child.destroy()

        self.root.update_idletasks()
        padding = 18
        y = padding
        max_content_width = 0

        for i, img in enumerate(bundle.page_images):
            photo = ImageTk.PhotoImage(img)
            self.page_photos.append(photo)
            self.canvas.create_image(10, y, anchor="nw", image=photo)
            self.page_positions.append({
                "page_index": i,
                "top": y,
                "height": photo.height(),
                "width": photo.width(),
                "scale": photo.width() / max(1, img.width),
            })
            y += photo.height() + padding
            max_content_width = max(max_content_width, photo.width() + 20)

            tphoto = ImageTk.PhotoImage(bundle.thumb_images[i])
            self.thumb_photos.append(tphoto)
            btn = ttk.Button(self.thumb_frame, image=tphoto, command=lambda p=i: self.scroll_to_page(p))
            btn.image = tphoto
            btn.grid(row=i * 2, column=0, sticky="ew", pady=(0, 2))
            ttk.Label(self.thumb_frame, text=f"Page {i + 1}").grid(row=i * 2 + 1, column=0, sticky="w", pady=(0, 8))

        total_height = y + padding
        canvas_width = max(700, max_content_width, self.canvas.winfo_width())
        self.canvas.configure(scrollregion=(0, 0, canvas_width, total_height))
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))
        self._jump_to_current_question_anchor()
        total_height = y + padding
        self.canvas.configure(scrollregion=(0, 0, max(display_width + 50, 700), total_height))
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))
        self._jump_to_current_question_anchor()
    
    def _configure_style(self) -> None:
        self.root.configure(bg="#eef2f7")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", font=("Segoe UI", 10), background="#eef2f7")
        style.configure("TFrame", background="#eef2f7")
        style.configure("TLabel", background="#eef2f7", foreground=FG_LABEL)
        style.configure("TButton", padding=(10, 5))
        style.configure("TMenubutton", padding=(10, 5))
        style.configure("TLabelframe", background="#eef2f7", padding=(10, 8))
        style.configure("TLabelframe.Label", background="#eef2f7", foreground=FG_COLOR, font=("Segoe UI", 10, "bold"))
        style.configure("Title.TLabel", background="#eef2f7", foreground=FG_COLOR, font=("Segoe UI", 12, "bold"))
        style.configure("Header.TLabel", background="#eef2f7", foreground=FG_COLOR, font=("Segoe UI", 10, "bold"))

    def _set_app_icon(self) -> None:
        base = Path(__file__).resolve().parent
        for name in ("app_icon.png", "app_icon.ico"):
            icon_path = base / name
            if not icon_path.exists():
                continue
            try:
                if icon_path.suffix.lower() == ".png":
                    img = ImageTk.PhotoImage(Image.open(icon_path))
                    self.root.iconphoto(True, img)
                    self._icon_photo = img
                elif icon_path.suffix.lower() == ".ico" and sys.platform.startswith("win"):
                    self.root.iconbitmap(default=str(icon_path))
                return
            except Exception:
                pass

    def show_readme_popup(self) -> None:
        base = Path(__file__).resolve().parent
        readme_path = base / "README.md"
        if not readme_path.exists():
            messagebox.showinfo("README", "README.md was not found next to the script.")
            return

        top = tk.Toplevel(self.root)
        top.title("README")
        top.geometry("860x760")
        top.transient(self.root)

        frame = ttk.Frame(top, padding=10)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", relief="flat", padx=10, pady=10)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        try:
            content = readme_path.read_text(encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("README", f"Could not open README.md: {exc}", parent=top)
            top.destroy()
            return

        text.insert("1.0", content)
        text.configure(state="disabled")

    def _build_ui(self) -> None:
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        outer = ttk.Frame(self.root, padding=8)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.columnconfigure(2, weight=0)

        toolbar = ttk.Frame(outer)
        toolbar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        toolbar.columnconfigure(12, weight=1)
        toolbar.columnconfigure(13, weight=0)

        submissions_btn = ToggleDropdownButton(
            toolbar,
            "Submissions",
            [
                ("Open CSV file", self.open_csv),
                ("__separator__", None),
                ("Load folder with PDFs", self.load_submissions),
                ("New CSV file", self.create_csv),
            ],
        )
        submissions_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        solution_btn = ToggleDropdownButton(
            toolbar,
            "Solution",
            [
                ("Load Solution PDF", self.load_solution_pdf),
                ("Load Answersheet PDF", self.load_clean_pdf),
                ("__separator__", None),
                ("Preview Solution", self.view_solution_pdf),
                ("Preview Answersheet", self.view_clean_pdf),
                ("__separator__", None),
                ("Questions", self.edit_questions),
                ("Set Anchor", self.toggle_anchor_mode),
            ],
        )
        solution_btn.grid(row=0, column=1, padx=(0, 6), sticky="ew")

        ttk.Button(toolbar, text="Save", command=self.save_csv).grid(row=0, column=2, padx=(0, 12))
        ttk.Button(toolbar, text="?", width=3, command=self.show_readme_popup).grid(row=0, column=13, sticky="e")

        left = ttk.Frame(outer)
        left.grid(row=1, column=0, sticky="nsw", padx=(0, 8))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        left.columnconfigure(1, weight=0)

        student_header = ttk.Frame(left)
        student_header.grid(row=0, column=0, columnspan=2, sticky="ew")
        student_header.columnconfigure(0, weight=1)
        student_header.columnconfigure(1, weight=0)
        student_header.columnconfigure(2, weight=0)
        ttk.Label(student_header, text="Students", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(student_header, text="|←", width=1.5, command=self.prev_ungraded_student).grid(row=0, column=1, padx=(0, 3))
        ttk.Button(student_header, text="←", width=1.5, command=self.prev_student).grid(row=0, column=2, padx=(0, 3))
        ttk.Button(student_header, text="→", width=1.5, command=self.next_student).grid(row=0, column=3, padx=(0, 3))
        ttk.Button(student_header, text="→|", width=1.5, command=self.next_ungraded_student).grid(row=0, column=4)

        self.student_list = tk.Listbox(left, height=16, exportselection=False, width=26)
        self.student_list.grid(row=1, column=0, sticky="nsew")
        self.student_list.bind("<<ListboxSelect>>", self.on_student_select)
        st_scroll = ttk.Scrollbar(left, orient="vertical", command=self.student_list.yview)
        st_scroll.grid(row=1, column=1, sticky="ns")
        self.student_list.configure(yscrollcommand=st_scroll.set)

        ttk.Separator(left, orient="horizontal").grid(row=2, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Label(left, text="Pages", style="Header.TLabel").grid(row=3, column=0, sticky="w")
        thumb_container = ttk.Frame(left)
        thumb_container.grid(row=4, column=0, columnspan=2, sticky="nsew")
        thumb_container.rowconfigure(0, weight=1)
        thumb_container.columnconfigure(0, weight=1)
        self.thumb_canvas = tk.Canvas(thumb_container, width=220, height=550, highlightthickness=0)
        self.thumb_canvas.grid(row=0, column=0, sticky="nsew")
        thumb_scroll = ttk.Scrollbar(thumb_container, orient="vertical", command=self.thumb_canvas.yview)
        thumb_scroll.grid(row=0, column=1, sticky="ns")
        self.thumb_canvas.configure(yscrollcommand=thumb_scroll.set)
        self.thumb_frame = ttk.Frame(self.thumb_canvas)
        self.thumb_window = self.thumb_canvas.create_window((0, 0), window=self.thumb_frame, anchor="nw")
        self.thumb_frame.bind("<Configure>", lambda _e: self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all")))
        self.thumb_canvas.bind("<Configure>", self._on_thumb_canvas_configure)
        self.thumb_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.thumb_canvas.bind("<Button-4>", self._on_mousewheel)
        self.thumb_canvas.bind("<Button-5>", self._on_mousewheel)
        bind_mousewheel_recursive(thumb_container, self.thumb_canvas)

        center = ttk.Frame(outer)
        center.grid(row=1, column=1, sticky="nsew", padx=(0, 8))
        center.rowconfigure(0, weight=1)
        center.rowconfigure(1, weight=0)
        center.columnconfigure(0, weight=1)
        center.columnconfigure(1, weight=0)
        self.canvas = tk.Canvas(center, bg="#f5f5f5", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        cscroll = ttk.Scrollbar(center, orient="vertical", command=self.canvas.yview)
        cscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(center, orient="horizontal", command=self.canvas.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=cscroll.set, xscrollcommand=hscroll.set)
        self.canvas.bind("<Configure>", self._schedule_rerender)
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)

        right = ttk.Frame(outer)
        right.grid(row=1, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(4, weight=1)

        q_header = ttk.Frame(right)
        q_header.grid(row=0, column=0, sticky="ew")
        q_header.columnconfigure(0, weight=1)
        q_header.columnconfigure(1, weight=0)
        ttk.Label(q_header, text="Current Question", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(q_header, text="Set Anchor", command=self.toggle_anchor_mode).grid(row=0, column=1, sticky="e")

        q_list_container = ttk.Frame(right)
        q_list_container.grid(row=1, column=0, sticky="nsew")
        q_list_container.columnconfigure(0, weight=1)
        q_list_container.rowconfigure(0, weight=1)
        self.question_list = tk.Listbox(
            q_list_container,
            height=8,
            exportselection=False,
            bg="#ffffff",
            fg=FG_COLOR,
            selectbackground="#cfe3ff",
            selectforeground=FG_COLOR,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            activestyle="dotbox",
        )
        self.question_list.grid(row=0, column=0, sticky="nsew")
        self.question_list.bind("<<ListboxSelect>>", self.on_question_select)
        q_scroll = ttk.Scrollbar(q_list_container, orient="vertical", command=self.question_list.yview)
        q_scroll.grid(row=0, column=1, sticky="ns")
        self.question_list.configure(yscrollcommand=q_scroll.set)

        ttk.Separator(right, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=8)
        ttk.Label(right, text="Rubric Buckets", style="Header.TLabel").grid(row=3, column=0, sticky="w")

        self.bucket_container = ttk.Frame(right)
        self.bucket_container.grid(row=4, column=0, sticky="nsew")
        self.bucket_container.rowconfigure(0, weight=1)
        self.bucket_container.columnconfigure(0, weight=1)

        self.bucket_canvas = tk.Canvas(self.bucket_container, height=260, highlightthickness=0)
        self.bucket_canvas.grid(row=0, column=0, sticky="nsew")
        bucket_scroll = ttk.Scrollbar(self.bucket_container, orient="vertical", command=self.bucket_canvas.yview)
        bucket_scroll.grid(row=0, column=1, sticky="ns")
        self.bucket_canvas.configure(yscrollcommand=bucket_scroll.set)
        self.bucket_frame = ttk.Frame(self.bucket_canvas)
        self.bucket_frame.columnconfigure(0, weight=1)
        self.bucket_window = self.bucket_canvas.create_window((0, 0), window=self.bucket_frame, anchor="nw")
        self.bucket_frame.bind("<Configure>", lambda _e: self.bucket_canvas.configure(scrollregion=self.bucket_canvas.bbox("all")))
        self.bucket_canvas.bind("<Configure>", self._on_bucket_canvas_configure)
        bind_mousewheel_recursive(self.bucket_container, self.bucket_canvas)

        custom = ttk.LabelFrame(right, text="Total score")
        custom.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        custom.columnconfigure(0, weight=0)
        custom.columnconfigure(1, weight=1)
        custom.columnconfigure(2, weight=0)

        ttk.Label(custom, text="Score").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.custom_score_var = tk.StringVar()
        ttk.Entry(custom, textvariable=self.custom_score_var, width=12).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(custom, text="Apply score", command=self.apply_custom_score).grid(row=0, column=2, sticky="e", padx=4, pady=4)

        notes = ttk.LabelFrame(right, text="Notes")
        notes.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        notes.columnconfigure(0, weight=1)
        self.note_text = tk.Text(notes, width=28, height=3, wrap="word")
        self.note_text.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ttk.Button(notes, text="Apply note", command=self.apply_note).grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))

        self.anchor_label = ttk.Label(right, text="Anchor: none")
        self.anchor_label.grid(row=7, column=0, sticky="w", pady=(8, 0))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Separator(self.root, orient="horizontal").grid(row=1, column=0, sticky="ew")
        ttk.Label(self.root, textvariable=self.status_var, anchor="w", style="Header.TLabel").grid(row=2, column=0, sticky="ew", padx=10, pady=(6, 8))

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-s>", lambda e: self.save_csv())
        self.root.bind_all("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.root.bind_all("<Control-Button-4>", self._on_ctrl_mousewheel)
        self.root.bind_all("<Control-Button-5>", self._on_ctrl_mousewheel)
        self.root.bind_all("<Page_Down>", self._on_page_down)
        self.root.bind_all("<Next>", self._on_page_down)
        self.root.bind_all("<Shift-Next>", self._on_page_down)
        self.root.bind_all("<Page_Up>", self._on_page_up)
        self.root.bind_all("<Prior>", self._on_page_up)
        self.root.bind_all("<Shift-Prior>", self._on_page_up)
        self.root.bind("<Right>", lambda e: self.next_question())
        self.root.bind("<Left>", lambda e: self.prev_question())
        self.root.bind("<Up>", lambda e: self.scroll_pdf(-3))
        self.root.bind("<Down>", lambda e: self.scroll_pdf(3))
        self.root.bind("<Return>", lambda e: self.apply_custom_score())
        self.root.bind("<KeyPress>", self._on_keypress)

    def _update_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_dirty_status(self) -> None:
        if self.csv_path:
            self._update_status(f"Modified: {self.csv_path.name}")
        else:
            self._update_status("Modified (no CSV yet)")

    def _focus_is_textlike(self) -> bool:
        widget = self.root.focus_get()
        return isinstance(widget, (tk.Entry, tk.Text))

    def _on_keypress(self, event: tk.Event) -> None:
        if self._focus_is_textlike():
            return
        key = normalize_shortcut_key_event(event)
        if key:
            self.toggle_bucket_by_key(key)

    def _on_mousewheel(self, event: tk.Event) -> None:
        delta = -1 * int(event.delta / 120) if getattr(event, "delta", 0) else (1 if event.num == 5 else -1)
        widget = event.widget
        try:
            if widget is self.thumb_canvas:
                self.thumb_canvas.yview_scroll(delta, "units")
            else:
                self.canvas.yview_scroll(delta, "units")
        except Exception:
            pass

    def _on_ctrl_mousewheel(self, event: tk.Event) -> str:
        delta = 0
        if getattr(event, "delta", 0):
            delta = 1 if event.delta > 0 else -1
        elif getattr(event, "num", None) == 4:
            delta = 1
        elif getattr(event, "num", None) == 5:
            delta = -1
        if delta > 0:
            self.zoom_in()
        elif delta < 0:
            self.zoom_out()
        return "break"

    def zoom_in(self) -> None:
        self.set_zoom(self.zoom_factor * 1.15)

    def zoom_out(self) -> None:
        self.set_zoom(self.zoom_factor / 1.15)

    def set_zoom(self, factor: float) -> None:
        factor = max(self.min_zoom, min(self.max_zoom, factor))
        if abs(factor - self.zoom_factor) < 1e-6:
            return
        self.zoom_factor = factor
        if self.current_pdf_path is not None:
            self._render_current_pdf()
            self._update_status(f"Zoom: {int(self.zoom_factor * 100)}%")

    def _on_thumb_canvas_configure(self, event: tk.Event) -> None:
        self.thumb_canvas.itemconfig(self.thumb_window, width=event.width)

    def _on_bucket_canvas_configure(self, event: tk.Event) -> None:
        self.bucket_canvas.itemconfig(self.bucket_window, width=event.width)

    def _schema_path_for_csv(self, csv_path: Path) -> Path:
        return csv_path.with_suffix(".schema.json")

    def _current_schema_path(self) -> Optional[Path]:
        if self.schema_path is not None:
            return self.schema_path
        if self.csv_path is not None:
            return self._schema_path_for_csv(self.csv_path)
        return None

    def _serialize_schema(self) -> dict:
        return {
            "submission_dir": str(self.submission_dir) if self.submission_dir else "",
            "csv_path": str(self.csv_path) if self.csv_path else "",
            "clean_pdf": str(self.clean_pdf) if self.clean_pdf else "",
            "solution_pdf": str(self.solution_pdf) if self.solution_pdf else "",
            "questions": [
                {
                    "qid": q.qid,
                    "max_points": q.max_points,
                    "buckets": [dataclasses.asdict(b) for b in q.buckets],
                }
                for q in self.questions
            ],
            "anchors": {qid: dataclasses.asdict(anchor) for qid, anchor in self.anchors.items()},
            "applied_buckets": {
                student: {
                    qid: list(cell.applied_bucket_ids)
                    for qid, cell in grades.items()
                    if isinstance(cell, CellValue) and cell.applied_bucket_ids
                }
                for student, grades in self.grades.items()
                if any(isinstance(cell, CellValue) and cell.applied_bucket_ids for cell in grades.values())
            },
            "current_question_index": self.current_question_index,
            "current_student_index": self.current_student_index,
        }

    def _save_schema(self) -> None:
        path = self._current_schema_path()
        if path is None:
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump(self._serialize_schema(), f, indent=2)
        os.replace(tmp, path)

    def _load_schema(self, path: Path) -> bool:
        if not path.exists():
            return False
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return False
        self.schema_path = path
        self._apply_schema_dict(data)
        return True

    def _apply_schema_dict(self, data: dict) -> None:
        self.questions = []
        for qraw in data.get("questions", []):
            buckets: list[Bucket] = []
            for braw in qraw.get("buckets", []):
                buckets.append(Bucket(
                    bid=braw.get("bid", uuid.uuid4().hex),
                    label=braw.get("label", ""),
                    points=float(braw.get("points", 0.0)),
                    key=str(braw.get("key", "")),
                    mode=str(braw.get("mode", "add")),
                ))
            self.questions.append(QuestionConfig(
                qid=qraw.get("qid", "Q?"),
                max_points=float(qraw.get("max_points", 0.0)),
                buckets=buckets or [Bucket(bid=uuid.uuid4().hex, label="not answered/wrong", points=0.0, key="0", mode="set")],
            ))
        self.anchors = {}
        for qid, araw in data.get("anchors", {}).items():
            try:
                self.anchors[qid] = Anchor(
                    page_index=int(araw.get("page_index", 0)),
                    x_ratio=float(araw.get("x_ratio", 0.0)),
                    y_ratio=float(araw.get("y_ratio", 0.0)),
                )
            except Exception:
                continue
        subdir = data.get("submission_dir", "")
        self.submission_dir = Path(subdir) if subdir else self.submission_dir
        clean = data.get("clean_pdf", "")
        sol = data.get("solution_pdf", "")
        self.clean_pdf = Path(clean) if clean else None
        self.solution_pdf = Path(sol) if sol else None
        self.current_question_index = int(data.get("current_question_index", -1))
        self.current_student_index = int(data.get("current_student_index", -1))
        for student, qmap in data.get("applied_buckets", {}).items():
            if not isinstance(qmap, dict):
                continue
            g = self.grades.setdefault(student, {})
            for qid, bucket_ids in qmap.items():
                cell = g.setdefault(qid, CellValue())
                if isinstance(bucket_ids, list):
                    cell.applied_bucket_ids = [str(bid) for bid in bucket_ids if str(bid)]
        self._refresh_question_list()
        self._refresh_anchor_label()
        self._refresh_scoring_panel()

    def _student_key_to_path(self, student: str) -> Optional[Path]:
        if not self.submission_dir:
            return None
        direct = self.submission_dir / f"{student}.pdf"
        if direct.exists():
            return direct
        matches = [p for p in self.submission_dir.glob("*.pdf") if p.stem.lower() == student.lower()]
        return matches[0] if matches else None

    def load_submissions(self) -> None:
        folder = filedialog.askdirectory(title="Choose submission folder")
        if not folder:
            return
        self.submission_dir = Path(folder)
        self.students = sorted(p.stem for p in self.submission_dir.glob("*.pdf"))
        if not self.students:
            messagebox.showerror("No PDFs found", "The selected folder does not contain any PDF files.")
            return
        self._refresh_student_list()
        self._init_grade_store_from_students()
        self._refresh_student_list_styles()
        self._update_status(f"Loaded {len(self.students)} submissions from {self.submission_dir}")
        if self.current_student_index < 0:
            self.select_student(0)
        self._refresh_question_list_styles()
        self._refresh_student_list_styles()
        self._save_schema()

    def load_clean_pdf(self) -> None:
        path = filedialog.askopenfilename(title="Choose Clean-answersheet PDF", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        self.clean_pdf = Path(path)
        self._update_status(f"Loaded clean PDF: {self.clean_pdf.name}")
        self._save_schema()

    def load_solution_pdf(self) -> None:
        path = filedialog.askopenfilename(title="Choose Solution PDF", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        self.solution_pdf = Path(path)
        self._update_status(f"Loaded solution PDF: {self.solution_pdf.name}")
        self._save_schema()

    def create_csv(self) -> None:
        if not self.students:
            messagebox.showwarning("Load submissions first", "Please load a submission folder first.")
            return
        if not self.questions:
            if not self.edit_questions():
                return
        path = filedialog.asksaveasfilename(
            title="Create CSV file",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="grades.csv",
        )
        if not path:
            return
        self.csv_path = Path(path)
        self.schema_path = self._schema_path_for_csv(self.csv_path)
        self._sync_grade_store_to_questions()
        self.save_csv()
        self._save_schema()
        self._update_status(f"Created CSV: {self.csv_path}")

    def open_csv(self) -> None:
        path = filedialog.askopenfilename(title="Open CSV file", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        self.csv_path = Path(path)
        self.schema_path = self._schema_path_for_csv(self.csv_path)
        loaded_schema = self._load_schema(self.schema_path)
        self._load_csv()
        if not loaded_schema:
            self._save_schema()
        self._update_status(f"Opened CSV: {self.csv_path}")

    def edit_questions(self) -> bool:
        dlg = QuestionDialog(self.root, self.questions)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return False
        self.questions = dlg.result
        self._refresh_question_list()
        self._sync_grade_store_to_questions()
        self._save_schema()
        return True

    def _preview_pdf(self, path: Optional[Path], title: str) -> None:
        if not path:
            messagebox.showinfo("PDF Preview", f"Load {title.lower()} first.")
            return

        top = tk.Toplevel(self.root)
        top.title(f"{title} Preview — {path.name}")
        top.geometry("900x800")

        outer = ttk.Frame(top, padding=8)
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(outer)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        toolbar.columnconfigure(3, weight=1)

        zoom_var = tk.DoubleVar(value=1.0)
        zoom_label = ttk.Label(toolbar, text="Zoom: 100%")
        zoom_label.grid(row=0, column=0, sticky="w")

        canvas_holder = ttk.Frame(outer)
        canvas_holder.grid(row=1, column=0, sticky="nsew")
        canvas_holder.rowconfigure(0, weight=1)
        canvas_holder.columnconfigure(0, weight=1)

        canvas = tk.Canvas(canvas_holder, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(canvas_holder, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(canvas_holder, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")

        def _sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        frame.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * int(e.delta / 120) if getattr(e, "delta", 0) else (1 if e.num == 5 else -1), "units"))
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        top._preview_token = uuid.uuid4().hex
        top._preview_photos = []
        top._preview_render_after_id = None
        top._preview_render_request_id = 0

        def _schedule_render(_event=None) -> None:
            after_id = getattr(top, "_preview_render_after_id", None)
            if after_id is not None:
                try:
                    self.root.after_cancel(after_id)
                except Exception:
                    pass
                top._preview_render_after_id = None
            top._preview_render_after_id = self.root.after(180, _render_now)

        def _set_zoom(value: float, schedule: bool = True) -> None:
            value = max(0.5, min(3.0, float(value)))
            zoom_var.set(value)
            zoom_label.configure(text=f"Zoom: {int(round(value * 100))}%")
            if schedule:
                _schedule_render()

        def _zoom_step(multiplier: float) -> None:
            _set_zoom(zoom_var.get() * multiplier)

        def _render_now() -> None:
            if not top.winfo_exists():
                return
            token = getattr(top, "_preview_token", None)
            if token is None:
                return
            req_id = getattr(top, "_preview_render_request_id", 0) + 1
            top._preview_render_request_id = req_id
            zoom = float(zoom_var.get())

            for child in frame.winfo_children():
                child.destroy()
            ttk.Label(frame, text="Rendering preview...", style="Header.TLabel").pack(anchor="w", padx=12, pady=12)

            future = self._preview_render_executor.submit(
                self._render_pdf_bundle,
                path,
                800,
                zoom,
                None,
            )

            def _finish(fut, rid=req_id, token=token):
                try:
                    bundle = fut.result()
                except Exception as exc:
                    def _show_error() -> None:
                        if not top.winfo_exists() or getattr(top, "_preview_token", None) != token:
                            return
                        if rid != getattr(top, "_preview_render_request_id", -1):
                            return
                        messagebox.showerror(f"{title} preview failed", str(exc), parent=top)

                    self.root.after(0, _show_error)
                    return

                def _apply() -> None:
                    if not top.winfo_exists() or getattr(top, "_preview_token", None) != token:
                        return
                    if rid != getattr(top, "_preview_render_request_id", -1):
                        return

                    for child in frame.winfo_children():
                        child.destroy()

                    photos: list[ImageTk.PhotoImage] = []
                    for i, img in enumerate(bundle.page_images):
                        photo = ImageTk.PhotoImage(img)
                        photos.append(photo)
                        lbl = ttk.Label(frame, image=photo)
                        lbl.image = photo
                        lbl.pack(anchor="n", pady=(8 if i else 0, 8))

                    top._preview_photos = photos
                    canvas.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    canvas.xview_moveto(0.0)
                    canvas.yview_moveto(0.0)

                self.root.after(0, _apply)

            future.add_done_callback(_finish)

        ttk.Button(toolbar, text="-", width=3, command=lambda: _zoom_step(1 / 1.15)).grid(row=0, column=1, padx=(12, 4))
        ttk.Button(toolbar, text="+", width=3, command=lambda: _zoom_step(1.15)).grid(row=0, column=2, padx=(0, 10))
        zoom_slider = ttk.Scale(toolbar, from_=0.5, to=3.0, variable=zoom_var, command=_schedule_render)
        zoom_slider.grid(row=0, column=3, sticky="ew")
        zoom_slider.bind("<ButtonRelease-1>", lambda _e: _set_zoom(zoom_var.get()))
        top.bind("<Control-MouseWheel>", lambda e: (_zoom_step(1.15) if getattr(e, "delta", 0) > 0 else _zoom_step(1 / 1.15), "break")[1])
        top.bind("<Control-Button-4>", lambda e: (_zoom_step(1.15), "break")[1])
        top.bind("<Control-Button-5>", lambda e: (_zoom_step(1 / 1.15), "break")[1])

        def _close_preview() -> None:
            after_id = getattr(top, "_preview_render_after_id", None)
            if after_id is not None:
                try:
                    self.root.after_cancel(after_id)
                except Exception:
                    pass
            top.destroy()

        top.protocol("WM_DELETE_WINDOW", _close_preview)

        _set_zoom(1.0, schedule=False)
        _render_now()

    def view_clean_pdf(self) -> None:
        self._preview_pdf(self.clean_pdf, "Clean PDF")

    def view_solution_pdf(self) -> None:
        self._preview_pdf(self.solution_pdf, "Solution PDF")

    def toggle_anchor_mode(self) -> None:
        if not self.questions:
            messagebox.showwarning("Define questions first", "Create the question list before defining anchors.")
            return
        if self.current_question_index < 0:
            messagebox.showwarning("Select a question", "Select the question you want to anchor first.")
            return
        if not self.clean_pdf:
            messagebox.showwarning("Load Clean PDF first", "Load the Clean-answersheet PDF before setting anchors.")
            return
        qid = self.questions[self.current_question_index].qid
        if not self.anchor_mode:
            self.anchor_mode = True
            self.pending_anchor_question = qid
            self._update_status(f"Anchor mode ON for {qid}: click the Clean PDF where the question starts.")
            self.anchor_label.configure(text=f"Anchor: awaiting click for {qid}")
            self.current_pdf_path = self.clean_pdf
            self._render_current_pdf()
        else:
            self.anchor_mode = False
            self.pending_anchor_question = None
            self._update_status("Anchor mode OFF")
            self._refresh_anchor_label()
            if self.current_student_index >= 0:
                self._load_current_student_pdf()

    def _refresh_student_list(self) -> None:
        self.student_list.delete(0, tk.END)
        for s in self.students:
            self.student_list.insert(tk.END, s)
        self._refresh_student_list_styles()

    def _student_completion_state(self, student: str) -> str:
        if not self.questions:
            return "none"
        graded = 0
        for q in self.questions:
            cell = self.grades.get(student, {}).get(q.qid)
            if cell is not None and cell.score is not None:
                graded += 1
        if graded == 0:
            return "none"
        if graded >= len(self.questions):
            return "done"
        return "partial"

    def _refresh_student_list_styles(self) -> None:
        if not hasattr(self, "student_list"):
            return
        for idx, student in enumerate(self.students):
            state = self._student_completion_state(student)
            bg = STUDENT_NONE_BG if state == "none" else STUDENT_PARTIAL_BG if state == "partial" else STUDENT_DONE_BG
            try:
                self.student_list.itemconfig(idx, background=bg, foreground=FG_COLOR)
            except Exception:
                pass

    def _refresh_question_list(self) -> None:
        self.question_list.delete(0, tk.END)
        for q in self.questions:
            self.question_list.insert(tk.END, f"{q.qid}  ({q.max_points:g})")
        if self.questions and self.current_question_index < 0:
            self.select_question(0)
        elif self.questions and self.current_question_index >= len(self.questions):
            self.select_question(len(self.questions) - 1)
        else:
            self._ensure_question_selection()
        self._refresh_question_list_styles()

    def _ensure_question_selection(self) -> None:
        if not self.questions or self.current_question_index < 0:
            return
        index = max(0, min(self.current_question_index, len(self.questions) - 1))
        try:
            self.question_list.selection_clear(0, tk.END)
            self.question_list.selection_set(index)
            self.question_list.activate(index)
            self.question_list.see(index)
        except Exception:
            pass

    def _question_is_complete(self, qid: str) -> bool:
        if not self.students:
            return False
        for student in self.students:
            cell = self.grades.get(student, {}).get(qid)
            if cell is None or cell.score is None:
                return False
        return True

    def _refresh_question_list_styles(self) -> None:
        if not hasattr(self, "question_list"):
            return
        current_complete = False
        for idx, q in enumerate(self.questions):
            complete = self._question_is_complete(q.qid)
            if idx == self.current_question_index:
                current_complete = complete
            try:
                self.question_list.itemconfig(idx, background=QUESTION_DONE_BG if complete else "#ffffff", foreground=FG_COLOR)
            except Exception:
                pass
        try:
            self.question_list.configure(
                selectbackground=QUESTION_DONE_SELECT_BG if current_complete else QUESTION_ACTIVE_SELECT_BG,
                selectforeground=FG_COLOR,
            )
        except Exception:
            pass
        self._ensure_question_selection()

    def on_student_select(self, _event=None) -> None:
        selection = self.student_list.curselection()
        if not selection:
            return
        self.select_student(selection[0])

    def on_question_select(self, _event=None) -> None:
        selection = self.question_list.curselection()
        if not selection:
            return
        self.select_question(selection[0])

    def _init_grade_store_from_students(self) -> None:
        for s in self.students:
            self.grades.setdefault(s, {})
            self.status_map.setdefault(s, "ungraded")
            self.last_saved.setdefault(s, "")

    def _sync_grade_store_to_questions(self) -> None:
        for s in self.students:
            g = self.grades.setdefault(s, {})
            for q in self.questions:
                g.setdefault(q.qid, CellValue())
        self._refresh_question_list()
        self._refresh_anchor_label()
        self._refresh_scoring_panel()
        self._refresh_student_list_styles()
        self._refresh_question_list_styles()
        if self.current_student_index >= 0:
            self._load_current_student_question_into_editor()

    def _refresh_anchor_label(self) -> None:
        if self.current_question_index < 0 or not self.questions:
            self.anchor_label.configure(text="Anchor: none")
            return
        qid = self.questions[self.current_question_index].qid
        anchor = self.anchors.get(qid)
        if anchor is None:
            self.anchor_label.configure(text=f"Anchor: none for {qid}")
        else:
            self.anchor_label.configure(text=f"Anchor: {qid} page {anchor.page_index + 1}, x={anchor.x_ratio:.2f}, y={anchor.y_ratio:.2f}")

    def _start_prefetch_for_next_ungraded_student(self) -> None:
        if not self.students or not self.questions or self.current_question_index < 0:
            return
        next_idx = self._find_student_index_for_question(1, require_ungraded=True)
        if next_idx is None:
            return
        self._queue_student_pdf_prefetch(next_idx)      # keeps raw bytes fallback
        self._queue_student_pdf_render_prefetch(next_idx)
    
    def _render_cached_pdf_to_canvas(self, bundle: RenderedPDFBundle, display_width: int) -> None:
        self.canvas.delete("all")
        self.page_positions = []
        self.page_photos = []
        self.thumb_photos = []

        for child in self.thumb_frame.winfo_children():
            child.destroy()

        self.root.update_idletasks()
        padding = 18
        y = padding
        thumb_width = 150

        for i, img in enumerate(bundle.page_images):
            photo = ImageTk.PhotoImage(img)
            self.page_photos.append(photo)
            self.canvas.create_image(10, y, anchor="nw", image=photo)
            self.page_positions.append({
                "page_index": i,
                "top": y,
                "height": photo.height(),
                "width": photo.width(),
                "scale": photo.width() / max(1, img.width),
            })
            y += photo.height() + padding

            tphoto = ImageTk.PhotoImage(bundle.thumb_images[i])
            self.thumb_photos.append(tphoto)
            btn = ttk.Button(self.thumb_frame, image=tphoto, command=lambda p=i: self.scroll_to_page(p))
            btn.image = tphoto
            btn.grid(row=i * 2, column=0, sticky="ew", pady=(0, 2))
            ttk.Label(self.thumb_frame, text=f"Page {i + 1}").grid(row=i * 2 + 1, column=0, sticky="w", pady=(0, 8))

        total_height = y + padding
        self.canvas.configure(scrollregion=(0, 0, max(display_width + 50, 700), total_height))
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))
        self._jump_to_current_question_anchor()

    def _queue_student_pdf_prefetch(self, index: int) -> None:
        if index < 0 or index >= len(self.students):
            return
        pdf_path = self._student_key_to_path(self.students[index])
        if pdf_path is None:
            return
        if pdf_path in self._prefetch_cache or pdf_path in self._prefetch_inflight:
            return
        self._prefetch_inflight.add(pdf_path)
        future = self._prefetch_executor.submit(pdf_path.read_bytes)

        def _done(fut, path=pdf_path):
            try:
                data = fut.result()
            except Exception:
                data = None
            self._prefetch_inflight.discard(path)
            if data is not None:
                self._prefetch_cache[path] = data

        future.add_done_callback(_done)

    def _consume_prefetched_pdf_bytes(self, path: Path) -> Optional[bytes]:
        return self._prefetch_cache.pop(path, None)

    def _find_student_index_for_question(self, direction: int, require_ungraded: bool) -> Optional[int]:
        if not self.students or not self.questions or self.current_question_index < 0:
            return None
        qid = self.questions[self.current_question_index].qid
        if self.current_student_index < 0:
            start = 0 if direction > 0 else len(self.students) - 1
        else:
            start = self.current_student_index + direction
        if direction > 0:
            rng = range(start, len(self.students))
        else:
            rng = range(start, -1, -1)
        for idx in rng:
            cell = self.grades.get(self.students[idx], {}).get(qid)
            graded = cell is not None and cell.score is not None
            if require_ungraded:
                if not graded:
                    return idx
            else:
                return idx
        return None

    def _on_page_down(self, event: tk.Event) -> str:
        if getattr(event, "state", 0) & 0x0001:
            self.next_student()
        else:
            self.next_ungraded_student()
        return "break"

    def _on_page_up(self, event: tk.Event) -> str:
        if getattr(event, "state", 0) & 0x0001:
            self.prev_student()
        else:
            self.prev_ungraded_student()
        return "break"

    def _bucket_button_text(self, bucket: Bucket, active: bool = False) -> str:
        key = bucket.key if bucket.key else "?"
        return f"[{key}] ({bucket.points:g} pts)\n{bucket.label}"

    def _refresh_scoring_panel(self) -> None:
        for child in self.bucket_frame.winfo_children():
            child.destroy()
        self.bucket_buttons.clear()
        self.bucket_button_specs.clear()
        self.bucket_button_base_bg.clear()
        if self.current_question_index < 0 or not self.questions:
            ttk.Label(self.bucket_frame, text="No question selected").grid(row=0, column=0, sticky="w")
            return
        q = self.questions[self.current_question_index]
        self.bucket_frame.columnconfigure(0, weight=1)
        ttk.Label(self.bucket_frame, text=f"{q.qid} — max {q.max_points:g}").grid(row=0, column=0, sticky="ew", pady=(0, 4))
        row = 1
        available_width = max(220, self.bucket_canvas.winfo_width() - 20)
        for b in q.buckets:
            btn = tk.Button(
                self.bucket_frame,
                text=self._bucket_button_text(b, active=False),
                command=lambda bucket=b: self.toggle_bucket(bucket),
                wraplength=max(180, available_width - 28),
                justify="left",
                anchor="w",
                relief="raised",
                bd=1,
                highlightthickness=0,
                padx=12,
                pady=10,
                font=("Segoe UI", 10),
                activeforeground=FG_COLOR,
            )
            btn.grid(row=row, column=0, sticky="ew", padx=0, pady=4)
            self.bucket_buttons[b.bid] = btn
            self.bucket_button_specs[b.bid] = b
            self.bucket_button_base_bg[b.bid] = btn.cget("bg")
            row += 1
        if not q.buckets:
            ttk.Label(self.bucket_frame, text="No buckets defined for this question.").grid(row=row, column=0, sticky="w")
        self._update_bucket_button_states()

    def _update_bucket_button_states(self) -> None:
        if self.current_student_index < 0 or self.current_question_index < 0 or not self.students or not self.questions:
            for bid, btn in self.bucket_buttons.items():
                bucket = self.bucket_button_specs.get(bid)
                if bucket is None:
                    continue
                btn.configure(text=self._bucket_button_text(bucket, active=False), relief="raised")
            return
        student = self.students[self.current_student_index]
        qid = self.questions[self.current_question_index].qid
        cell = self.grades.setdefault(student, {}).setdefault(qid, CellValue())
        applied = set(cell.applied_bucket_ids)
        for bucket in self.questions[self.current_question_index].buckets:
            btn = self.bucket_buttons.get(bucket.bid)
            if btn is None:
                continue
            base_bg = self.bucket_button_base_bg.get(bucket.bid, self.root.cget("bg"))
            active = bucket.bid in applied
            btn.configure(
                text=self._bucket_button_text(bucket, active=active),
                relief="sunken" if active else "raised",
                bg="#dbeafe" if active else base_bg,
                activebackground="#dbeafe" if active else base_bg,
            )

    def select_student(self, index: int) -> None:
        if not self.students:
            return
        index = max(0, min(index, len(self.students) - 1))
        self.current_student_index = index
        self.student_list.selection_clear(0, tk.END)
        self.student_list.selection_set(index)
        self.student_list.see(index)
        self._load_current_student_pdf()
        self._load_current_student_question_into_editor()
        self._update_bucket_button_states()
        self._refresh_question_list_styles()
        self._refresh_student_list_styles()
        self._start_prefetch_for_next_ungraded_student()

    def select_question(self, index: int) -> None:
        if not self.questions:
            return
        index = max(0, min(index, len(self.questions) - 1))
        self.current_question_index = index
        self.question_list.selection_clear(0, tk.END)
        self.question_list.selection_set(index)
        self.question_list.see(index)
        self._refresh_scoring_panel()
        self._refresh_anchor_label()
        self._load_current_student_question_into_editor()
        self._jump_to_current_question_anchor()
        self._refresh_question_list_styles()
        self._refresh_student_list_styles()
        self._start_prefetch_for_next_ungraded_student()
        self._save_schema()

    def next_student(self) -> None:
        if self.students:
            self.select_student(self.current_student_index + 1 if self.current_student_index >= 0 else 0)

    def prev_student(self) -> None:
        if self.students:
            self.select_student(self.current_student_index - 1 if self.current_student_index > 0 else 0)

    def next_ungraded_student(self) -> None:
        idx = self._find_student_index_for_question(1, require_ungraded=True)
        if idx is not None:
            self.select_student(idx)

    def prev_ungraded_student(self) -> None:
        idx = self._find_student_index_for_question(-1, require_ungraded=True)
        if idx is not None:
            self.select_student(idx)

    def next_question(self) -> None:
        if self.questions:
            self.select_question(self.current_question_index + 1 if self.current_question_index >= 0 else 0)

    def prev_question(self) -> None:
        if self.questions:
            self.select_question(self.current_question_index - 1 if self.current_question_index > 0 else 0)

    def scroll_pdf(self, delta_units: int) -> None:
        try:
            self.canvas.yview_scroll(delta_units, "units")
        except Exception:
            pass

    def _load_current_student_pdf(self) -> None:
        if self.current_student_index < 0 or not self.students:
            return
        student = self.students[self.current_student_index]
        pdf_path = self._student_key_to_path(student)
        if pdf_path is None:
            messagebox.showerror("Missing PDF", f"Could not locate PDF for student {student}")
            return
        self.current_pdf_path = pdf_path
        self.current_pdf_bytes = self._consume_prefetched_pdf_bytes(pdf_path)
        self._render_current_pdf()
        self._update_status(f"Viewing {student}")
        self._jump_to_current_question_anchor()

    def _render_current_pdf(self) -> None:
        if self.current_pdf_path is None:
            return
        display_width = max(650, self.canvas.winfo_width() - 25)
        self._begin_view_render(
            self.current_pdf_path,
            self.current_pdf_bytes,
            display_width,
            self.zoom_factor,
        )

    def scroll_to_page(self, page_index: int) -> None:
        if page_index < 0 or page_index >= len(self.page_positions):
            return
        pos = self.page_positions[page_index]
        bbox = self.canvas.bbox("all")
        total_height = max(1, bbox[3] if bbox else 1)
        target = float(pos["top"])
        self.canvas.yview_moveto(max(0.0, min(1.0, target / total_height)))

    def _schedule_rerender(self, _event=None) -> None:
        if self.render_after_id is not None:
            try:
                self.root.after_cancel(self.render_after_id)
            except Exception:
                pass
        self.render_after_id = self.root.after(300, self._rerender_if_possible)

    def _rerender_if_possible(self) -> None:
        self.render_after_id = None
        if self.current_pdf_path is not None:
            self._render_current_pdf()

    def on_canvas_click(self, event: tk.Event) -> None:
        if not self.anchor_mode or self.pending_anchor_question is None:
            return
        if not self.page_positions:
            return
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        pos = None
        for item in self.page_positions:
            top = float(item["top"])
            height = float(item["height"])
            if top <= y <= top + height:
                pos = item
                break
        if pos is None:
            return
        page_index = int(pos["page_index"])
        local_y = max(0.0, min(float(y - float(pos["top"])), float(pos["height"])))
        local_x = max(0.0, min(float(x - 10), float(pos["width"])))
        anchor = Anchor(
            page_index=page_index,
            x_ratio=local_x / max(1.0, float(pos["width"])),
            y_ratio=local_y / max(1.0, float(pos["height"])),
        )
        self.anchors[self.pending_anchor_question] = anchor
        qid = self.pending_anchor_question
        self.anchor_mode = False
        self.pending_anchor_question = None
        self._refresh_anchor_label()
        self._save_schema()
        self._update_status(f"Anchor set for {qid} on page {page_index + 1}")
        if self.current_student_index >= 0:
            self._load_current_student_pdf()

    def _jump_to_current_question_anchor(self) -> None:
        if self.current_question_index < 0 or not self.questions or not self.page_positions:
            return
        qid = self.questions[self.current_question_index].qid
        anchor = self.anchors.get(qid)
        if anchor is None or anchor.page_index >= len(self.page_positions):
            return
        pos = self.page_positions[anchor.page_index]
        target = float(pos["top"]) + anchor.y_ratio * float(pos["height"]) - 40
        bbox = self.canvas.bbox("all")
        total_height = max(1, bbox[3] if bbox else 1)
        self.canvas.yview_moveto(max(0.0, min(1.0, target / total_height)))

    def _load_current_student_question_into_editor(self) -> None:
        if self.current_student_index < 0 or self.current_question_index < 0 or not self.questions or not self.students:
            return
        student = self.students[self.current_student_index]
        qid = self.questions[self.current_question_index].qid
        cell = self.grades.setdefault(student, {}).setdefault(qid, CellValue())
        self.custom_score_var.set("" if cell.score is None else str(cell.score))
        self.note_text.delete("1.0", tk.END)
        if cell.note:
            self.note_text.insert("1.0", cell.note)
        self._update_bucket_button_states()

    def _current_student_and_question(self) -> tuple[str, str]:
        if self.current_student_index < 0 or self.current_question_index < 0:
            raise ValueError("No active student/question")
        return self.students[self.current_student_index], self.questions[self.current_question_index].qid

    def _bucket_lookup_for_current_question(self) -> dict[str, Bucket]:
        if self.current_question_index < 0 or self.current_question_index >= len(self.questions):
            return {}
        return {b.bid: b for b in self.questions[self.current_question_index].buckets}

    def toggle_bucket(self, bucket: Bucket) -> None:
        try:
            student, qid = self._current_student_and_question()
        except ValueError:
            return
        cell = self.grades.setdefault(student, {}).setdefault(qid, CellValue())
        active = list(cell.applied_bucket_ids)
        bucket_map = self._bucket_lookup_for_current_question()
        if bucket.bid not in bucket_map:
            return

        if bucket.mode == "set":
            if active == [bucket.bid]:
                active = []
            else:
                active = [bucket.bid]
        else:
            if bucket.bid in active:
                active.remove(bucket.bid)
            else:
                active = [bid for bid in active if bucket_map.get(bid, Bucket("", "", 0.0)).mode != "set"]
                active.append(bucket.bid)

        cell.applied_bucket_ids = active
        cell.score = self._compute_score_from_active_buckets(active, bucket_map)
        self.custom_score_var.set("" if cell.score is None else str(cell.score))
        self._refresh_cell_after_change(student, qid)

    def _compute_score_from_active_buckets(self, active_ids: list[str], bucket_map: dict[str, Bucket]) -> Optional[float]:
        if not active_ids:
            return None
        set_bucket = next((bucket_map.get(bid) for bid in active_ids if bucket_map.get(bid) and bucket_map[bid].mode == "set"), None)
        if set_bucket is not None:
            return float(set_bucket.points)
        total = 0.0
        found = False
        for bid in active_ids:
            bucket = bucket_map.get(bid)
            if bucket is None:
                continue
            total += float(bucket.points)
            found = True
        return total if found else None

    def toggle_bucket_by_key(self, key: str) -> None:
        if self.current_question_index < 0 or self.current_question_index >= len(self.questions):
            return
        for bucket in self.questions[self.current_question_index].buckets:
            if bucket.key == key:
                self.toggle_bucket(bucket)
                return

    def apply_custom_score(self) -> None:
        try:
            student, qid = self._current_student_and_question()
        except ValueError:
            return
        raw = self.custom_score_var.get().strip()
        if not raw:
            return
        try:
            score = float(raw)
        except ValueError:
            messagebox.showerror("Invalid score", "Custom score must be numeric.")
            return
        cell = self.grades.setdefault(student, {}).setdefault(qid, CellValue())
        cell.score = score
        cell.applied_bucket_ids = []
        self._refresh_cell_after_change(student, qid)

    def apply_note(self) -> None:
        try:
            student, qid = self._current_student_and_question()
        except ValueError:
            return
        cell = self.grades.setdefault(student, {}).setdefault(qid, CellValue())
        cell.note = self.note_text.get("1.0", "end").strip()
        self._refresh_cell_after_change(student, qid, preserve_editor=True)

    def _refresh_cell_after_change(self, student: str, qid: str, preserve_editor: bool = False) -> None:
        self._recompute_total_and_status(student)
        self._write_csv_if_ready()
        if not preserve_editor:
            self._load_current_student_question_into_editor()
        self._set_dirty_status()
        self._refresh_student_row(student)
        self._refresh_question_list_styles()
        self._refresh_student_list_styles()
        self._save_schema()
        self._update_status(f"Saved {student} / {qid}")
        self._start_prefetch_for_next_ungraded_student()

    def _refresh_student_row(self, student: str) -> None:
        pass

    def _recompute_total_and_status(self, student: str) -> None:
        total = 0.0
        filled = 0
        for q in self.questions:
            cell = self.grades.setdefault(student, {}).setdefault(q.qid, CellValue())
            if cell.score is not None:
                total += float(cell.score)
                filled += 1
        self.grades[student]["_total"] = CellValue(score=total, note="")  # type: ignore[index]
        if filled == 0:
            status = "ungraded"
        elif filled < len(self.questions):
            status = "in-progress"
        else:
            status = "done"
        self.status_map[student] = status
        self.last_saved[student] = now_iso()

    def _build_csv_headers(self) -> list[str]:
        headers = ["student"]
        headers.extend(q.qid for q in self.questions)
        headers.append("Total")
        headers.extend(f"Notes{q.qid}" for q in self.questions)
        headers.extend(f"Buckets{q.qid}" for q in self.questions)
        headers.extend(["Status", "LastSaved"])
        return headers

    def _row_to_csv(self, student: str) -> dict[str, str]:
        row: dict[str, str] = {"student": student}
        total = 0.0
        for q in self.questions:
            cell = self.grades.setdefault(student, {}).setdefault(q.qid, CellValue())
            row[q.qid] = "" if cell.score is None else self._format_number(cell.score)
            if cell.score is not None:
                total += float(cell.score)
        row["Total"] = self._format_number(total)
        for q in self.questions:
            cell = self.grades.setdefault(student, {}).setdefault(q.qid, CellValue())
            row[f"Notes{q.qid}"] = cell.note
            row[f"Buckets{q.qid}"] = ",".join(cell.applied_bucket_ids)
        row["Status"] = self.status_map.get(student, "ungraded")
        row["LastSaved"] = self.last_saved.get(student, "")
        return row

    def _format_number(self, value: float | int) -> str:
        if float(value).is_integer():
            return str(int(value))
        return f"{float(value):g}"

    def _write_csv_if_ready(self) -> None:
        if not self.csv_path:
            return
        self._write_csv()

    def save_csv(self) -> None:
        if not self.csv_path:
            messagebox.showwarning("No CSV path", "Create or open a CSV file first.")
            return
        self._write_csv()
        self._save_schema()
        self._update_status(f"Saved CSV: {self.csv_path}")

    def _write_csv(self) -> None:
        if not self.csv_path:
            return
        self.csv_path.parent.mkdir(parents=True, exist_ok=True)
        headers = self._build_csv_headers()
        tmp = self.csv_path.with_suffix(self.csv_path.suffix + ".tmp")
        with tmp.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for student in self.students:
                writer.writerow(self._row_to_csv(student))
        os.replace(tmp, self.csv_path)

    def _load_csv(self) -> None:
        if not self.csv_path or not self.csv_path.exists():
            messagebox.showerror("CSV not found", "Could not open the selected CSV file.")
            return
        existing_grades = self.grades
        with self.csv_path.open(newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames or []
            if not self.questions:
                score_cols = [h for h in headers if h not in {"student", "Total", "Status", "LastSaved"} and not h.startswith("Notes") and not h.startswith("Buckets")]
                self.questions = [QuestionConfig(qid=h, max_points=0.0, buckets=[Bucket(bid=uuid.uuid4().hex, label="not answered/wrong", points=0.0, key="0", mode="set")]) for h in score_cols]
            self.grades = {}
            self.status_map.clear()
            self.last_saved.clear()
            for row in reader:
                student = (row.get("student") or "").strip()
                if not student:
                    continue
                self.grades.setdefault(student, {})
                for q in self.questions:
                    raw = (row.get(q.qid) or "").strip()
                    score = safe_float(raw)
                    note = row.get(f"Notes{q.qid}", "")
                    bucket_raw = row.get(f"Buckets{q.qid}")
                    if bucket_raw is None or not bucket_raw.strip():
                        existing_cell = existing_grades.get(student, {}).get(q.qid)
                        bucket_ids = list(existing_cell.applied_bucket_ids) if isinstance(existing_cell, CellValue) else []
                    else:
                        bucket_ids = [bid for bid in bucket_raw.split(",") if bid]
                    self.grades[student][q.qid] = CellValue(score=score, note=note, applied_bucket_ids=bucket_ids)
                self.status_map[student] = row.get("Status", "ungraded") or "ungraded"
                self.last_saved[student] = row.get("LastSaved", "")
        if self.submission_dir:
            self.students = sorted(p.stem for p in self.submission_dir.glob("*.pdf"))
        else:
            self.students = list(self.grades.keys())
        self._init_grade_store_from_students()
        self._sync_grade_store_to_questions()
        self._refresh_student_list()
        self._refresh_student_list_styles()
        self._refresh_question_list_styles()
        if self.students:
            self.select_student(self.current_student_index if 0 <= self.current_student_index < len(self.students) else 0)


def main() -> None:
    root = tk.Tk()
    OfflineGraderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
