
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from dxf_layer_report import (
    ScanController,
    ScanOptions,
    create_temp_output_path,
    normalize_layers,
    open_in_default_spreadsheet,
    scan_dxf_files,
    status_details,
    write_results,
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        icon_candidates = [
            Path(__file__).resolve().parent / "assets" / "icon.png",
            Path(__file__).resolve().parent / "assets" / "icon.ico",
        ]
        self._icon_image = None
        for icon_path in icon_candidates:
            if not icon_path.exists():
                continue
            try:
                if icon_path.suffix.lower() == ".ico":
                    self.iconbitmap(default=str(icon_path))
                else:
                    self._icon_image = tk.PhotoImage(file=str(icon_path))
                    self.iconphoto(True, self._icon_image)
                break
            except Exception:
                pass

        self.title("DXF Layer Report")
        self.geometry("1080x760")
        self.minsize(980, 680)

        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._worker_thread: threading.Thread | None = None
        self._controller: ScanController | None = None
        self._scan_running = False
        self._scan_paused = False
        self._buffered_log_lines: list[tuple[str, str]] = []
        self._max_log_lines = 6000

        self.source_var = tk.StringVar()
        self.layers_var = tk.StringVar(value="1")
        self.source_kind_var = tk.StringVar(value="folder")
        self.recursive_var = tk.BooleanVar(value=True)
        self.workers_var = tk.IntVar(value=4)
        self.factor_var = tk.StringVar(value="1.0")
        self.keep_missing_var = tk.BooleanVar(value=True)
        self.keep_errors_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")
        self.progress_text_var = tk.StringVar(value="0 / 0")
        self.percent_var = tk.StringVar(value="0.0%")
        self.discovery_var = tk.StringVar(value="Folders: 0 | Files: 0")
        self.current_folder_var = tk.StringVar(value="Current folder: -")

        self._build_ui()
        self.after(70, self._poll_queue)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(9, weight=1)

        title = ttk.Label(root, text="DXF Layer Report", font=("Segoe UI", 13, "bold"))
        title.grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 12))

        settings = ttk.LabelFrame(root, text="Scan settings", padding=12)
        settings.grid(row=1, column=0, columnspan=5, sticky="ew")
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Source type").grid(row=0, column=0, sticky="w", pady=4)
        type_wrap = ttk.Frame(settings)
        type_wrap.grid(row=0, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(type_wrap, text="Folder", value="folder", variable=self.source_kind_var).pack(side="left")
        ttk.Radiobutton(type_wrap, text="Single file", value="file", variable=self.source_kind_var).pack(side="left", padx=(16, 0))

        ttk.Label(settings, text="Source").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(settings, textvariable=self.source_var).grid(row=1, column=1, sticky="ew", padx=(10, 10))
        ttk.Button(settings, text="Browse", command=self._browse_source, width=12).grid(row=1, column=2, sticky="e")

        ttk.Label(settings, text="Layers").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(settings, textvariable=self.layers_var).grid(row=2, column=1, sticky="ew", padx=(10, 10))
        ttk.Label(settings, text="Comma separated, e.g. 14,20,CUT", foreground="#666666").grid(row=2, column=2, sticky="w")

        ttk.Checkbutton(settings, text="Recursive", variable=self.recursive_var).grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(settings, text="Workers").grid(row=3, column=1, sticky="e")
        ttk.Spinbox(settings, from_=1, to=32, textvariable=self.workers_var, width=8).grid(row=3, column=2, sticky="w")
        ttk.Label(settings, text="Scale to mm").grid(row=4, column=1, sticky="e", pady=(8, 0))
        ttk.Entry(settings, textvariable=self.factor_var, width=8).grid(row=4, column=2, sticky="w", pady=(8, 0))

        options = ttk.Frame(root)
        options.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(8, 8))
        ttk.Checkbutton(options, text="Keep files with no requested layer", variable=self.keep_missing_var).pack(side="left")
        ttk.Checkbutton(options, text="Keep read errors", variable=self.keep_errors_var).pack(side="left", padx=(16, 0))

        actions = ttk.Frame(root)
        actions.grid(row=3, column=0, columnspan=5, sticky="ew", pady=(4, 6))
        self.start_button = ttk.Button(actions, text="Start", command=self._start_scan, width=10)
        self.pause_button = ttk.Button(actions, text="Pause", command=self._toggle_pause, width=10, state="disabled")
        self.stop_button = ttk.Button(actions, text="Stop", command=self._stop_scan, width=10, state="disabled")
        self.start_button.pack(side="left")
        self.pause_button.pack(side="left", padx=(8, 0))
        self.stop_button.pack(side="left", padx=(8, 0))
        ttk.Label(actions, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).pack(side="right")

        progress_wrap = ttk.Frame(root)
        progress_wrap.grid(row=4, column=0, columnspan=5, sticky="ew")
        progress_wrap.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_wrap, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.progress_overlay = ttk.Label(progress_wrap, textvariable=self.percent_var, anchor="center", font=("Segoe UI", 9, "bold"))
        self.progress_overlay.place(relx=0.5, rely=0.5, anchor="center")

        meta = ttk.Frame(root)
        meta.grid(row=5, column=0, columnspan=5, sticky="ew", pady=(8, 6))
        ttk.Label(meta, textvariable=self.progress_text_var).pack(side="left")
        ttk.Label(meta, textvariable=self.discovery_var).pack(side="left", padx=(20, 0))
        ttk.Label(meta, textvariable=self.current_folder_var).pack(side="left", padx=(20, 0))

        log_frame = ttk.LabelFrame(root, text="Live log", padding=0)
        log_frame.grid(row=6, column=0, columnspan=5, sticky="nsew", pady=(8, 0))
        root.rowconfigure(6, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="none", font=("Consolas", 10), relief="flat")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=y_scroll.set)

        self.log_text.tag_configure("info", font=("Segoe UI", 9, "bold"))
        self.log_text.tag_configure("ok", foreground="#2E7D32")
        self.log_text.tag_configure("missing", foreground="#B85C5C")
        self.log_text.tag_configure("error", foreground="#8D4E85")
        self.log_text.insert("end", "Ready to scan\n", "info")
        self.log_text.configure(state="disabled")

    def _browse_source(self) -> None:
        if self.source_kind_var.get() == "folder":
            value = filedialog.askdirectory(title="Select a folder containing DXF files")
        else:
            value = filedialog.askopenfilename(
                title="Select a DXF file",
                filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")],
            )
        if value:
            self.source_var.set(value)

    def _validate_options(self) -> ScanOptions:
        source_value = self.source_var.get().strip()
        if not source_value:
            raise ValueError("Please select a source path.")

        try:
            factor = float(self.factor_var.get().strip())
        except ValueError as exc:
            raise ValueError("Scale to mm must be a valid number.") from exc

        return ScanOptions(
            source_path=Path(source_value),
            target_layers=normalize_layers(self.layers_var.get().split(",")),
            unit_to_mm_factor=factor,
            recursive=self.recursive_var.get(),
            workers=max(1, int(self.workers_var.get())),
            include_missing_layers=self.keep_missing_var.get(),
            include_read_errors=self.keep_errors_var.get(),
        )

    def _start_scan(self) -> None:
        if self._scan_running:
            return

        try:
            options = self._validate_options()
        except Exception as exc:
            messagebox.showerror("Invalid settings", str(exc))
            return

        self._clear_log()
        self._append_log(f"Source: {options.source_path}", "info")
        self._append_log(f"Mode: {'Folder' if options.source_path.is_dir() else 'Single file'}", "info")
        self._append_log(f"Workers: {options.workers}", "info")
        self._append_log("Scan started", "info")

        self._controller = ScanController()
        self._scan_running = True
        self._scan_paused = False

        self.status_var.set("Running")
        self.discovery_var.set("Folders: ... | Files: ...")
        self.current_folder_var.set("Current folder: -")
        self.progress_text_var.set("0 / 0")
        self.percent_var.set("0.0%")
        self.progress.configure(value=0, maximum=100)

        self.start_button.configure(state="disabled")
        self.pause_button.configure(state="normal", text="Pause")
        self.stop_button.configure(state="normal")

        self._worker_thread = threading.Thread(target=self._run_scan, args=(options,), daemon=True)
        self._worker_thread.start()

    def _toggle_pause(self) -> None:
        if not self._controller or not self._scan_running:
            return
        if self._scan_paused:
            self._controller.resume()
            self._scan_paused = False
            self.pause_button.configure(text="Pause")
            self.status_var.set("Running")
            self._append_log("Scan resumed", "info")
        else:
            self._controller.pause()
            self._scan_paused = True
            self.pause_button.configure(text="Resume")
            self.status_var.set("Paused")
            self._append_log("Pause requested", "info")

    def _stop_scan(self) -> None:
        if not self._controller or not self._scan_running:
            return
        self._controller.stop()
        self.status_var.set("Stopping...")
        self._append_log("Stop requested", "info")
        self.stop_button.configure(state="disabled")

    def _run_scan(self, options: ScanOptions) -> None:
        try:
            results, total_files, stopped_early = scan_dxf_files(
                options=options,
                controller=self._controller,
                progress_callback=lambda done, total: self._queue.put(("progress", done, total)),
                result_callback=lambda result, done, total, folder_idx, folder_total: self._queue.put(
                    ("result", result, done, total, folder_idx, folder_total)
                ),
                state_callback=lambda state, done, total: self._queue.put(("state", state, done, total)),
                discovery_callback=lambda folders, files: self._queue.put(("discovery", folders, files)),
                folder_callback=lambda folder_rel, idx, total: self._queue.put(("folder", folder_rel, idx, total)),
            )

            output_path = create_temp_output_path()
            write_results(output_path, options.target_layers, results)

            self._queue.put(("finished", output_path, total_files, stopped_early))
        except Exception as exc:
            self._queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        processed_any = False

        try:
            while True:
                item = self._queue.get_nowait()
                processed_any = True
                kind = item[0]

                if kind == "progress":
                    _, done, total = item
                    self._update_progress(done, total)
                elif kind == "result":
                    _, result, done, total, folder_idx, folder_total = item
                    self._handle_result(result, done, total, folder_idx, folder_total)
                elif kind == "state":
                    _, state, _, _ = item
                    if state == "paused":
                        self.status_var.set("Paused")
                    elif state == "running" and self._scan_running:
                        self.status_var.set("Running")
                elif kind == "discovery":
                    _, folders, files = item
                    self.discovery_var.set(f"Folders: {folders} | Files: {files}")
                    self._append_log(f"Discovery complete: {folders} folders, {files} DXF files", "info")
                elif kind == "folder":
                    _, folder_rel, idx, total = item
                    self.current_folder_var.set(f"Current folder: {idx}/{total} - {folder_rel}")
                elif kind == "finished":
                    _, output_path, total_files, stopped_early = item
                    if self._buffered_log_lines:
                        self._flush_log_buffer()
                    self._scan_running = False
                    self._scan_paused = False
                    self.start_button.configure(state="normal")
                    self.pause_button.configure(state="disabled", text="Pause")
                    self.stop_button.configure(state="disabled")
                    self.status_var.set("Stopped" if stopped_early else "Done")

                    if stopped_early:
                        self._append_log("Scan stopped. Partial workbook generated.", "info")
                    else:
                        self._append_log(f"Scan complete: {total_files} file(s) processed", "info")

                    open_in_default_spreadsheet(output_path)
                    self._append_log(f"Workbook opened: {output_path}", "info")
                elif kind == "error":
                    _, message = item
                    self._scan_running = False
                    self._scan_paused = False
                    self.start_button.configure(state="normal")
                    self.pause_button.configure(state="disabled", text="Pause")
                    self.stop_button.configure(state="disabled")
                    self.status_var.set("Error")
                    self._append_log(message, "error")
                    messagebox.showerror("Scan error", message)
        except queue.Empty:
            pass

        if self._buffered_log_lines:
            self._flush_log_buffer()

        self.after(70 if processed_any else 100, self._poll_queue)

    def _update_progress(self, done: int, total: int) -> None:
        total = max(total, 1)
        percent = (done / total) * 100.0
        self.progress.configure(maximum=100, value=percent)
        self.progress_text_var.set(f"{done} / {total}")
        self.percent_var.set(f"{percent:.1f}%")

    def _handle_result(self, result, done: int, total: int, folder_idx: int, folder_total: int) -> None:
        details = status_details(result, normalize_layers(self.layers_var.get().split(",")))
        line = f"{result.filename:<36} {details}"
        tag = "info"

        if result.status == "OK":
            tag = "ok"
        elif result.status == "PARTIAL":
            tag = "ok"
        elif result.status == "NO LAYER FOUND":
            tag = "missing"
        elif result.status.startswith("ERROR:"):
            tag = "error"

        self._buffered_log_lines.append((line, tag))
        self.current_folder_var.set(f"Current folder: {folder_idx}/{folder_total} - {result.folder_rel}")
        self._update_progress(done, total)

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._buffered_log_lines.clear()

    def _append_log(self, text: str, tag: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n", tag)
        self._trim_log()
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _flush_log_buffer(self) -> None:
        self.log_text.configure(state="normal")
        for text, tag in self._buffered_log_lines[:250]:
            self.log_text.insert("end", text + "\n", tag)
        del self._buffered_log_lines[:250]
        self._trim_log()
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _trim_log(self) -> None:
        line_count = int(float(self.log_text.index("end-1c").split(".")[0]))
        overflow = line_count - self._max_log_lines
        if overflow > 0:
            self.log_text.delete("1.0", f"{overflow + 1}.0")


if __name__ == "__main__":
    App().mainloop()
