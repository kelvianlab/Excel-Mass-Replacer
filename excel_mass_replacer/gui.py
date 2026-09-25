"""Minimal desktop front-end: pick a folder, type two words, press a button."""

from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .core import Rule, RunSummary, discover_files, run


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Excel Mass Replacer {__version__}")
        self.minsize(720, 520)

        self.folder = tk.StringVar(value=str(Path.cwd()))
        self.find_text = tk.StringVar()
        self.replace_text = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.ignore_case = tk.BooleanVar(value=False)
        self.whole_cell = tk.BooleanVar(value=False)
        self.backup = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Ready.")

        self._events: queue.Queue = queue.Queue()
        self._busy = False

        self._build()
        self.after(100, self._drain_events)

    def _build(self) -> None:
        pad = {"padx": 8, "pady": 4}
        form = ttk.Frame(self)
        form.pack(fill="x", **pad)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Folder").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(form, textvariable=self.folder).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(form, text="Browse…", command=self._browse).grid(row=0, column=2, **pad)

        ttk.Label(form, text="Find").grid(row=1, column=0, sticky="w", **pad)
        find_entry = ttk.Entry(form, textvariable=self.find_text)
        find_entry.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        find_entry.focus_set()

        ttk.Label(form, text="Replace with").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(form, textvariable=self.replace_text).grid(
            row=2, column=1, columnspan=2, sticky="ew", **pad
        )

        options = ttk.Frame(self)
        options.pack(fill="x", **pad)
        ttk.Checkbutton(options, text="Include subfolders", variable=self.recursive).pack(
            side="left", padx=6
        )
        ttk.Checkbutton(options, text="Ignore case", variable=self.ignore_case).pack(
            side="left", padx=6
        )
        ttk.Checkbutton(options, text="Whole cell only", variable=self.whole_cell).pack(
            side="left", padx=6
        )
        ttk.Checkbutton(options, text="Keep .bak backup", variable=self.backup).pack(
            side="left", padx=6
        )

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", **pad)
        self.preview_button = ttk.Button(
            buttons, text="Preview (safe)", command=lambda: self._start(apply=False)
        )
        self.preview_button.pack(side="left", padx=6)
        self.apply_button = ttk.Button(
            buttons, text="Replace now", command=lambda: self._start(apply=True)
        )
        self.apply_button.pack(side="left", padx=6)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", **pad)

        self.log = tk.Text(self, height=16, wrap="none")
        self.log.pack(fill="both", expand=True, padx=8, pady=4)
        self.log.configure(state="disabled")

        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=8, pady=6)

    def _browse(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.folder.get() or str(Path.cwd()))
        if chosen:
            self.folder.set(chosen)

    def _write_log(self, line: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", line + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _start(self, apply: bool) -> None:
        if self._busy:
            return

        folder = Path(self.folder.get()).expanduser()
        if not folder.exists():
            messagebox.showerror("Folder not found", f"This path does not exist:\n{folder}")
            return
        if not self.find_text.get():
            messagebox.showerror("Nothing to find", "Type the text you want to replace.")
            return

        files = discover_files(folder, recursive=self.recursive.get())
        if not files:
            messagebox.showinfo(
                "No Excel files", f"No .xlsx, .xlsm or .xls files found in:\n{folder}"
            )
            return

        if apply and not messagebox.askyesno(
            "Confirm replacement",
            f"Replace text in {len(files)} Excel file(s) under:\n{folder}\n\n"
            f"Find:    {self.find_text.get()}\n"
            f"Replace: {self.replace_text.get()}\n\n"
            + (
                "A .bak backup will be kept next to each changed file."
                if self.backup.get()
                else "No backup will be kept. This cannot be undone."
            )
            + "\n\nContinue?",
        ):
            return

        rule = Rule(
            find=self.find_text.get(),
            replace=self.replace_text.get(),
            match_case=not self.ignore_case.get(),
            whole_cell=self.whole_cell.get(),
        )

        self._busy = True
        self.preview_button.state(["disabled"])
        self.apply_button.state(["disabled"])
        self._clear_log()
        self.progress.configure(maximum=len(files), value=0)
        self.status.set(
            f"{'Replacing in' if apply else 'Previewing'} {len(files)} file(s)…"
        )
        self._write_log(
            f"{'APPLY' if apply else 'PREVIEW'} — {len(files)} file(s) in {folder}"
        )

        # Tk variables may only be read from the main thread, so resolve
        # every setting here and hand the worker plain values.
        backup = self.backup.get()
        threading.Thread(
            target=self._work, args=(files, rule, apply, backup), daemon=True
        ).start()

    def _work(self, files, rule: Rule, apply: bool, backup: bool) -> None:
        started = time.perf_counter()
        try:
            summary = run(
                files,
                [rule],
                apply=apply,
                backup=backup,
                progress=lambda result: self._events.put(("file", result)),
            )
            self._events.put(("done", (summary, apply, time.perf_counter() - started)))
        except Exception as exc:
            self._events.put(("fatal", f"{type(exc).__name__}: {exc}"))

    def _drain_events(self) -> None:
        while True:
            try:
                kind, payload = self._events.get_nowait()
            except queue.Empty:
                break

            if kind == "file":
                self.progress.step(1)
                if payload.error:
                    self._write_log(f"ERROR   {payload.path.name}: {payload.error}")
                elif payload.replacements:
                    verb = "changed" if payload.written else "would change"
                    self._write_log(
                        f"{verb:14} {payload.path.name}: "
                        f"{payload.replacements} replacement(s)"
                    )
            elif kind == "done":
                self._finish(*payload)
            elif kind == "fatal":
                self._write_log(f"FAILED: {payload}")
                self.status.set("Failed — see the log above.")
                self._unlock()

        self.after(100, self._drain_events)

    def _finish(self, summary: RunSummary, apply: bool, elapsed: float) -> None:
        self._write_log("")
        self._write_log(
            f"{summary.total_replacements} replacement(s) in "
            f"{summary.total_cells} cell(s) across "
            f"{len(summary.changed_files)} file(s) — {elapsed:.2f}s"
        )
        if summary.failed_files:
            self._write_log(f"{len(summary.failed_files)} file(s) could not be processed.")
        if not apply:
            self._write_log('Preview only — press "Replace now" to save these changes.')
        self.status.set(
            f"Done in {elapsed:.2f}s — {summary.total_replacements} replacement(s)."
        )
        self._unlock()

    def _unlock(self) -> None:
        self._busy = False
        self.preview_button.state(["!disabled"])
        self.apply_button.state(["!disabled"])


def main() -> int:
    import multiprocessing

    multiprocessing.freeze_support()
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
