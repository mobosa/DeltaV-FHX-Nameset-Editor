"""
DeltaV FHX Nameset Editor — GUI + CLI entry point.
Backend logic lives in fhx_core.py.
"""

import os
import sys
import time
import threading
import queue
import subprocess
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from fhx_core import (
    compare_lib_and_export, validate_excel_for_generation, read_lib_edited_excel,
    generate_new_lib_fhx, _subprocess_generate, cli_progress,
)

import customtkinter as ctk
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Fix rounded corners & blurry text on high-DPI Windows
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Per-monitor DPI aware
except Exception:
    pass

APPLE_COLORS = {
    "bg": "#FFFFFF",
    "text_primary": "#1D1D1F",
    "text_secondary": "#86868B",
    "text_tertiary": "#AEAEB2",
    "accent": "#007AFF",
    "accent_hover": "#0063D1",
    "success": "#34C759",
    "success_hover": "#2DB84E",
    "border": "#E5E5EA",
    "card_bg": "#F5F5F7",
    "input_border": "#D2D2D7",
}


class FHX_Migrator_App:
    def __init__(self, root):
        self.root = root
        self.root.title("DeltaV FHX Nameset Editor")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.configure(fg_color=APPLE_COLORS["bg"])
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, 'exp_logo.ico')
            else:
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'exp_logo.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.lib_path = tk.StringVar()
        self.lib_setup_path = tk.StringVar()
        self.lib_gen_path = tk.StringVar()
        self.lib_gen_setup_path = tk.StringVar()
        self.lib_excel_path = tk.StringVar()
        self._progress_start_time = time.time()
        self._log_queue = queue.Queue()
        self._drain_id = None
        self._build_ui()

    def _build_ui(self):
        # ── Header ──
        header = ctk.CTkFrame(self.root, fg_color=APPLE_COLORS["bg"], corner_radius=0, height=80)
        header.pack(fill="x", padx=0, pady=(16, 0))
        header.pack_propagate(False)

        ctk.CTkLabel(header, text="FHX Nameset Editor",
                     font=("Segoe UI", 24, "bold"),
                     text_color=APPLE_COLORS["text_primary"]).pack(anchor="w", padx=28, pady=(0, 2))
        ctk.CTkLabel(header, text="DeltaV FHX 文件名称集编辑工具",
                     font=("Segoe UI", 13),
                     text_color=APPLE_COLORS["text_secondary"]).pack(anchor="w", padx=28)

        ctk.CTkFrame(self.root, fg_color=APPLE_COLORS["border"], height=1).pack(fill="x", padx=28, pady=(8, 0))

        # ── Scrollable content area ──
        content = ctk.CTkFrame(self.root, fg_color=APPLE_COLORS["bg"], corner_radius=0)
        content.pack(fill="both", expand=True, padx=28, pady=(12, 10))

        # ═══════════════════════════════════════
        # Step 1: Compare
        # ═══════════════════════════════════════
        step1 = ctk.CTkFrame(content, fg_color=APPLE_COLORS["card_bg"], corner_radius=12)
        step1.pack(fill="x", pady=(0, 12))

        s1_inner = ctk.CTkFrame(step1, fg_color="transparent")
        s1_inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(s1_inner, text="Step 1: Compare FHX with Setup",
                     font=("Segoe UI", 13, "bold"),
                     text_color=APPLE_COLORS["text_primary"]).pack(anchor="w", pady=(0, 10))

        r1 = ctk.CTkFrame(s1_inner, fg_color="transparent")
        r1.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(r1, text="FHX File", font=("Segoe UI", 12),
                     text_color=APPLE_COLORS["text_secondary"], width=120).pack(side="left")
        self._lib_entry = ctk.CTkEntry(r1, textvariable=self.lib_path, font=("Segoe UI", 11),
                                       placeholder_text="Original FHX file...",
                                       height=32, corner_radius=8, border_width=1,
                                       border_color=APPLE_COLORS["input_border"], fg_color="#FFFFFF")
        self._lib_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(r1, text="Browse", width=72, height=32, corner_radius=8,
                      font=("Segoe UI", 11), fg_color=APPLE_COLORS["accent"],
                      hover_color=APPLE_COLORS["accent_hover"],
                      command=lambda: self._browse(self.lib_path, "FHX", "*.fhx")).pack(side="right")

        r2 = ctk.CTkFrame(s1_inner, fg_color="transparent")
        r2.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(r2, text="Database.fhx", font=("Segoe UI", 12),
                     text_color=APPLE_COLORS["text_secondary"], width=120).pack(side="left")
        self._lib_setup_entry = ctk.CTkEntry(r2, textvariable=self.lib_setup_path, font=("Segoe UI", 11),
                                              placeholder_text="New Database.fhx...",
                                              height=32, corner_radius=8, border_width=1,
                                              border_color=APPLE_COLORS["input_border"], fg_color="#FFFFFF")
        self._lib_setup_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(r2, text="Browse", width=72, height=32, corner_radius=8,
                      font=("Segoe UI", 11), fg_color=APPLE_COLORS["accent"],
                      hover_color=APPLE_COLORS["accent_hover"],
                      command=lambda: self._browse(self.lib_setup_path, "FHX", "*.fhx")).pack(side="right")

        btn_row1 = ctk.CTkFrame(s1_inner, fg_color="transparent")
        btn_row1.pack(fill="x")
        self.lib_compare_btn = ctk.CTkButton(
            btn_row1, text="Compare and Export Excel", height=36, corner_radius=10,
            font=("Segoe UI", 13, "bold"), fg_color=APPLE_COLORS["accent"],
            hover_color=APPLE_COLORS["accent_hover"], command=self._do_lib_compare)
        self.lib_compare_btn.pack(side="left")

        # ═══════════════════════════════════════
        # Step 2: Generate
        # ═══════════════════════════════════════
        step2 = ctk.CTkFrame(content, fg_color=APPLE_COLORS["card_bg"], corner_radius=12)
        step2.pack(fill="x", pady=(0, 12))

        s2_inner = ctk.CTkFrame(step2, fg_color="transparent")
        s2_inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(s2_inner, text="Step 2: Generate New FHX from Edited Excel",
                     font=("Segoe UI", 13, "bold"),
                     text_color=APPLE_COLORS["text_primary"]).pack(anchor="w", pady=(0, 10))

        for label_text, var in [("FHX File", self.lib_gen_path),
                                 ("Database.fhx", self.lib_gen_setup_path),
                                 ("Edited Excel", self.lib_excel_path)]:
            row = ctk.CTkFrame(s2_inner, fg_color="transparent")
            row.pack(fill="x", pady=(0, 6))
            ctk.CTkLabel(row, text=label_text, font=("Segoe UI", 12),
                         text_color=APPLE_COLORS["text_secondary"], width=120).pack(side="left")
            entry = ctk.CTkEntry(row, textvariable=var, font=("Segoe UI", 11),
                                 placeholder_text=f"Select {label_text.lower()}...",
                                 height=32, corner_radius=8, border_width=1,
                                 border_color=APPLE_COLORS["input_border"], fg_color="#FFFFFF")
            entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ext = "*.xlsx" if "Excel" in label_text else "*.fhx"
            ctk.CTkButton(row, text="Browse", width=72, height=32, corner_radius=8,
                          font=("Segoe UI", 11), fg_color=APPLE_COLORS["accent"],
                          hover_color=APPLE_COLORS["accent_hover"],
                          command=lambda v=var, e=ext: self._browse(v, "File", e)).pack(side="right")

        btn_row2 = ctk.CTkFrame(s2_inner, fg_color="transparent")
        btn_row2.pack(fill="x", pady=(4, 0))
        self.lib_generate_btn = ctk.CTkButton(
            btn_row2, text="Generate New FHX", height=36, corner_radius=10,
            font=("Segoe UI", 13, "bold"), fg_color=APPLE_COLORS["success"],
            hover_color=APPLE_COLORS["success_hover"], command=self._do_lib_generate)
        self.lib_generate_btn.pack(side="left")
        ctk.CTkButton(btn_row2, text="Clear Log", height=32, corner_radius=8,
                      font=("Segoe UI", 11), fg_color=APPLE_COLORS["border"],
                      hover_color="#D1D1D6", text_color=APPLE_COLORS["text_primary"],
                      command=lambda: self._clear_log(self.log4)).pack(side="left", padx=(10, 0))

        # ═══════════════════════════════════════
        # Progress
        # ═══════════════════════════════════════
        prog_frame = ctk.CTkFrame(content, fg_color="transparent")
        prog_frame.pack(fill="x", pady=(0, 8))

        # ── Author footer (pack before log so it stays at bottom) ──
        ctk.CTkLabel(content, text="Author: Jared.Ji  |  Jared.Ji@emerson.com",
                     font=("Segoe UI", 10),
                     text_color=APPLE_COLORS["text_tertiary"]).pack(side="bottom", pady=(0, 2))

        self.prog_label4 = ctk.CTkLabel(prog_frame, text="", font=("Segoe UI", 11),
                                        text_color=APPLE_COLORS["text_secondary"], anchor="w")
        self.prog_label4.pack(side="left", fill="x", expand=True)

        self.progress4 = ctk.CTkProgressBar(prog_frame, width=200, height=6,
                                            corner_radius=3,
                                            progress_color=APPLE_COLORS["accent"],
                                            fg_color=APPLE_COLORS["border"])
        self.progress4.pack(side="right")
        self.progress4.set(0)

        # ═══════════════════════════════════════
        # Log
        # ═══════════════════════════════════════
        log_card = ctk.CTkFrame(content, fg_color=APPLE_COLORS["card_bg"], corner_radius=12)
        log_card.pack(fill="both", expand=True, pady=(0, 4))

        ctk.CTkLabel(log_card, text="Log", font=("Segoe UI", 12, "bold"),
                     text_color=APPLE_COLORS["text_primary"]).pack(anchor="w", padx=16, pady=(10, 4))

        self.log4 = ctk.CTkTextbox(log_card, font=("Cascadia Code", 12),
                                   fg_color="#FFFFFF", text_color=APPLE_COLORS["text_primary"],
                                   corner_radius=8, border_width=1,
                                   border_color=APPLE_COLORS["border"],
                                   activate_scrollbars=True)
        self.log4.pack(fill="both", expand=True, padx=12, pady=(0, 4))

    def _browse(self, var, file_type, ext):
        path = filedialog.askopenfilename(filetypes=[(f"{file_type} files", ext), ("All files", "*.*")])
        if path:
            var.set(path)

    def _log(self, log_widget, msg):
        log_widget.insert(tk.END, msg + "\n")
        log_widget.see(tk.END)

    def _clear_log(self, log_widget):
        log_widget.delete(1.0, tk.END)

    # Maximum number of log lines to insert per drain cycle to keep GUI responsive.
    _DRAIN_BATCH = 20

    def _drain_log_queue(self):
        self._drain_id = None
        try:
            if self._log_queue.empty():
                self._schedule_drain()
                return
            log_texts = []
            last_progress = None
            last_done = None
            info_msgs = []
            batch = 0
            while not self._log_queue.empty() and batch < self._DRAIN_BATCH:
                try:
                    msg_type, payload = self._log_queue.get_nowait()
                    batch += 1
                    if msg_type == 'log':
                        log_texts.append(payload[1])
                    elif msg_type == 'progress':
                        last_progress = payload
                    elif msg_type == 'info':
                        info_msgs.append(payload)
                    elif msg_type == 'subprocess_log':
                        log_texts.append(payload)
                    elif msg_type == 'subprocess_progress':
                        last_progress = payload
                    elif msg_type == 'subprocess_done':
                        last_done = payload
                except queue.Empty:
                    break
            if log_texts:
                self.log4.configure(state="normal")
                for txt in log_texts:
                    self.log4.insert(tk.END, txt + "\n")
                self.log4.see(tk.END)
            if last_progress:
                if isinstance(last_progress, tuple) and len(last_progress) == 4:
                    bar, label, pct, text = last_progress
                else:
                    pct, text = last_progress
                    bar, label = self.progress4, self.prog_label4
                bar.set(pct / 100)
                elapsed = time.time() - self._progress_start_time
                eta_str = ""
                if pct > 0:
                    eta = elapsed * (100 - pct) / pct
                    eta_str = f"  |  {elapsed:.1f}s elapsed, ~{eta:.0f}s remaining"
                else:
                    eta_str = f"  |  {elapsed:.1f}s elapsed"
                if pct >= 100:
                    label.configure(text=f"100% {text}  |  Completed in {elapsed:.1f}s", text_color="green")
                else:
                    label.configure(text=f"{pct}% {text}{eta_str}", text_color=APPLE_COLORS["text_primary"])
            for args in info_msgs:
                self.root.after(0, *args)
            if last_done:
                self._handle_subprocess_done(*last_done)
        except Exception:
            import logging
            logging.exception("Error in _drain_log_queue")
        finally:
            self._schedule_drain()

    def _handle_subprocess_done(self, ok, count, output_path, error):
        self.progress4.set(1.0)
        if ok:
            self.prog_label4.configure(text="Completed", text_color="green")
            self.log4.insert(tk.END, f"\n{'=' * 50}\nOutput: {output_path}\n")
            self.log4.see(tk.END)
            messagebox.showinfo("Success",
                f"New FHX generated!\n\nChanges: {count}\nOutput: {output_path}")
        else:
            self.prog_label4.configure(text="Failed", text_color="red")
            messagebox.showerror("Error", f"Generation failed:\n{error}")

    def _schedule_drain(self):
        if self._drain_id is not None:
            return
        try:
            self._drain_id = self.root.after(16, self._drain_log_queue)
        except RuntimeError:
            # Tkinter is shutting down — stop scheduling, drain is done.
            pass
        except Exception:
            import logging
            logging.exception("Failed to schedule drain")
            # Retry once after a short delay in case of transient TclError.
            try:
                self._drain_id = self.root.after(50, self._drain_log_queue)
            except Exception:
                pass

    def _update_progress(self, pct, text):
        self.progress4.set(pct / 100)
        elapsed = time.time() - self._progress_start_time
        if pct > 0:
            eta = elapsed * (100 - pct) / pct
            time_str = f"  |  {elapsed:.1f}s elapsed, ~{eta:.0f}s remaining"
        else:
            time_str = f"  |  {elapsed:.1f}s elapsed"
        if pct >= 100:
            self.prog_label4.configure(text=f"100% {text}  |  Completed in {elapsed:.1f}s", text_color="green")
        else:
            self.prog_label4.configure(text=f"{pct}% {text}{time_str}", text_color=APPLE_COLORS["text_primary"])

    def _start_bg_task(self, func, *args):
        self._schedule_drain()
        t = threading.Thread(target=func, args=args, daemon=True)
        t.start()

    # ── Step 1: Compare ──
    def _do_lib_compare(self):
        lib_path = self.lib_path.get().strip()
        setup_path = self.lib_setup_path.get().strip()
        if not lib_path or not setup_path:
            messagebox.showerror("Error", "Please select both FHX and Setup files.")
            return
        if not os.path.exists(lib_path):
            messagebox.showerror("Error", f"FHX not found: {lib_path}")
            return
        if not os.path.exists(setup_path):
            messagebox.showerror("Error", f"Setup not found: {setup_path}")
            return

        base, _ = os.path.splitext(lib_path)
        excel_out = f"{base}_library_comparison.xlsx"
        if os.path.exists(excel_out):
            if not messagebox.askyesno("Confirm", f"Excel file already exists:\n{excel_out}\n\nOverwrite?"):
                return

        self._clear_log(self.log4)
        self._log(self.log4, "DeltaV FHX Nameset Editor - Compare FHX with Setup")
        self._log(self.log4, f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._log(self.log4, "=" * 50)

        self.lib_compare_btn.configure(state="disabled", text="Comparing...")
        self.progress4.set(0)
        self._progress_start_time = time.time()
        self._start_bg_task(self._lib_compare_worker, lib_path, setup_path, excel_out)

    def _lib_compare_worker(self, lib_path, setup_path, excel_out):
        try:
            nameset_comp, sv_comp, expr_comp = compare_lib_and_export(
                lib_path, setup_path, excel_out,
                log_callback=lambda m: self._log_queue.put(('log', (self.log4, m))),
                progress_callback=lambda p, t: self._log_queue.put(('progress', (self.progress4, self.prog_label4, p, t)))
            )
            def _show_compare_result():
                summary = (f"\n{'=' * 50}\n"
                           f"Excel exported: {excel_out}\n"
                           f"  ENUMERATION_SET: {len(nameset_comp)}\n"
                           f"  STRING_VALUE: {len(sv_comp)}\n"
                           f"  Expression refs: {len(expr_comp)}\n\n"
                           f"Instructions:\n"
                           f"  1. Open the Excel file\n"
                           f"  2. Review 'Namesets' sheet\n"
                           f"  3. Review 'String Values' sheet\n"
                           f"  4. Review 'Expression Refs' sheet\n"
                           f"  5. Modify 'New Value' columns if needed\n"
                           f"  6. Go to Step 2 to generate new FHX")
                self.log4.insert(tk.END, summary + "\n")
                self.log4.see(tk.END)
                self.progress4.set(1.0)
                self.prog_label4.configure(text="Completed", text_color="green")
                msg = f"Comparison complete!\n\nENUMERATION_SET: {len(nameset_comp)}\nSTRING_VALUE: {len(sv_comp)}\nExpression refs: {len(expr_comp)}\n\nExcel: {excel_out}"
                messagebox.showinfo("Success", msg)
            self.root.after(0, _show_compare_result)
        except tk.TclError:
            pass
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Comparison failed:\n{e}"))
        finally:
            self.root.after(0, lambda: self.lib_compare_btn.configure(state="normal", text="Compare and Export Excel"))

    # ── Step 2: Generate ──
    def _do_lib_generate(self):
        lib_path = self.lib_gen_path.get().strip()
        setup_path = self.lib_gen_setup_path.get().strip()
        excel_path = self.lib_excel_path.get().strip()
        if not lib_path or not setup_path or not excel_path:
            messagebox.showerror("Error", "Please select FHX, Setup, and Excel files.")
            return
        if not os.path.exists(lib_path):
            messagebox.showerror("Error", f"FHX not found: {lib_path}")
            return
        if not os.path.exists(setup_path):
            messagebox.showerror("Error", f"Setup not found: {setup_path}")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel not found: {excel_path}")
            return

        base, ext = os.path.splitext(lib_path)
        output_path = f"{base}_NEW{ext}"
        if os.path.exists(output_path):
            if not messagebox.askyesno("Confirm", f"Output exists:\n{output_path}\n\nOverwrite?"):
                return

        self._clear_log(self.log4)
        self._log(self.log4, "DeltaV FHX Nameset Editor - Generate New FHX")
        self._log(self.log4, f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._log(self.log4, "=" * 50)

        self.lib_generate_btn.configure(state="disabled", text="Generating...")
        self.progress4.set(0)
        self._progress_start_time = time.time()
        self._start_bg_task(self._lib_generate_worker, lib_path, setup_path, excel_path, output_path)

    def _lib_generate_worker(self, lib_path, setup_path, excel_path, output_path):
        try:
            self._log_queue.put(('log', (self.log4, "Validating Excel data...")))
            is_valid, validation_errors = validate_excel_for_generation(
                excel_path,
                log_callback=lambda m: self._log_queue.put(('log', (self.log4, m))),
            )
            if not is_valid:
                error_summary = '\n'.join(
                    f"[{e['sheet']}] Row {e['row']}, {e['column']}: {e['message']}"
                    for e in validation_errors[:20]
                )
                if len(validation_errors) > 20:
                    error_summary += f"\n... and {len(validation_errors) - 20} more issues"
                self._log_queue.put(('log', (self.log4,
                    f"\nGeneration aborted: {len(validation_errors)} issue(s) found in Excel.")))
                self._log_queue.put(('info', (messagebox.showwarning, "Excel Validation Failed",
                    f"Found {len(validation_errors)} issue(s) in Excel:\n\n{error_summary}\n\n"
                    "Please fix these issues and try again.")))
                return

            nameset_changes, new_namesets, desc_changes, sv_changes, expr_changes, alarm_changes, priority_changes = read_lib_edited_excel(excel_path)
            self._log_queue.put(('log', (self.log4, f"Loaded {len(nameset_changes)} value changes, {len(new_namesets)} new namesets, {len(desc_changes)} description changes, {len(sv_changes)} STRING_VALUE changes, {len(expr_changes)} expression changes, {len(alarm_changes)} alarm changes, {len(priority_changes)} priority changes")))

            if not nameset_changes and not new_namesets and not desc_changes and not sv_changes and not expr_changes and not alarm_changes and not priority_changes:
                self._log_queue.put(('info', (messagebox.showinfo, "Info", "No changes found in Excel.")))
                return

            # P0 FIX: Use JSON instead of pickle for subprocess communication
            import tempfile
            # nameset_changes has tuple keys {(set_name, old_entry): new_entry}
            # JSON doesn't support tuple keys, so convert to list of [key, value] pairs
            nameset_serializable = [[list(k), v] for k, v in nameset_changes.items()]

            if getattr(sys, 'frozen', False):
                # Frozen (PyInstaller) exe: sys.executable is the .exe with its own
                # argparse — can't use '-c'. Run generation in a thread directly.
                def _run_gen():
                    import io, contextlib
                    class LiveStream:
                        def __init__(self):
                            self._buf = ''
                        def write(self, s):
                            self._buf += s
                            while '\n' in self._buf:
                                line, self._buf = self._buf.split('\n', 1)
                                line = line.strip()
                                if not line:
                                    continue
                                try:
                                    msg = json.loads(line)
                                except json.JSONDecodeError:
                                    self._queue.put(('subprocess_log', line))
                                    continue
                                if msg['type'] == 'log':
                                    self._queue.put(('subprocess_log', msg['msg']))
                                elif msg['type'] == 'progress':
                                    self._queue.put(('subprocess_progress', (msg['pct'], msg['text'])))
                                elif msg['type'] == 'done':
                                    self._queue.put(('subprocess_done', (
                                        msg['ok'], msg.get('count', 0), output_path, msg.get('error') or '')))
                        def flush(self):
                            pass
                    stream = LiveStream()
                    stream._queue = self._log_queue
                    try:
                        with contextlib.redirect_stdout(stream):
                            _subprocess_generate(
                                lib_path, setup_path, nameset_changes, new_namesets,
                                desc_changes, sv_changes, expr_changes, output_path,
                                alarm_changes, priority_changes)
                    except Exception as e:
                        self._log_queue.put(('subprocess_done', (False, 0, '', str(e))))
                threading.Thread(target=_run_gen, daemon=True).start()
            else:
                # Normal Python: use subprocess so heavy work can't freeze the GUI.
                args_data = json.dumps((lib_path, setup_path, nameset_serializable, new_namesets,
                                        desc_changes, sv_changes, expr_changes, output_path,
                                        alarm_changes, priority_changes))
                args_file = tempfile.NamedTemporaryFile(delete=False, suffix='.json', mode='w', encoding='utf-8')
                args_file.write(args_data)
                args_file.close()
                script_dir = os.path.dirname(os.path.abspath(__file__))
                sub_env = os.environ.copy()
                sub_env['PYTHONUNBUFFERED'] = '1'
                proc = subprocess.Popen(
                    [sys.executable, '-c',
                     f'import sys, json; sys.path.insert(0, {script_dir!r}); '
                     f'import fhx_core as fc; '
                     f'd = json.load(open(sys.argv[1], encoding="utf-8")); '
                     f'fc._subprocess_generate(*d)',
                     args_file.name],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    env=sub_env,
                    text=True, encoding='utf-8', errors='replace')

                def _reader():
                    try:
                        for line in proc.stdout:
                            line = line.strip()
                            if not line:
                                continue
                            try:
                                msg = json.loads(line)
                            except json.JSONDecodeError:
                                # P0 FIX: route through queue instead of direct tkinter access
                                self._log_queue.put(('subprocess_log', line))
                                continue
                            if msg['type'] == 'log':
                                self._log_queue.put(('subprocess_log', msg['msg']))
                            elif msg['type'] == 'progress':
                                self._log_queue.put(('subprocess_progress', (msg['pct'], msg['text'])))
                            elif msg['type'] == 'done':
                                self._log_queue.put(('subprocess_done', (
                                    msg['ok'], msg.get('count', 0), output_path, msg.get('error') or '')))
                        stderr_out = proc.stderr.read()
                        if stderr_out.strip():
                            self._log_queue.put(('subprocess_log', f"[STDERR] {stderr_out.strip()[:500]}"))
                    except Exception as e:
                        self._log_queue.put(('subprocess_done', (False, 0, '', str(e))))
                    finally:
                        try:
                            os.unlink(args_file.name)
                        except Exception:
                            pass

                threading.Thread(target=_reader, daemon=True).start()
        except tk.TclError:
            pass
        except Exception as e:
            self._log_queue.put(('info', (messagebox.showerror, "Error", f"Generation failed:\n{e}")))
        finally:
            self.root.after(0, lambda: self.lib_generate_btn.configure(state="normal", text="Generate New FHX"))


# ============================================================
# CLI
# ============================================================
def main():
    if len(sys.argv) > 1:
        import argparse
        parser = argparse.ArgumentParser(description='DeltaV FHX Nameset Editor')
        sub = parser.add_subparsers(dest='command')

        p_compare = sub.add_parser('compare', help='Compare any FHX with Setup, export Excel')
        p_compare.add_argument('fhx', help='Input FHX file (any type)')
        p_compare.add_argument('--setup', help='New Database.fhx reference file', required=True)
        p_compare.add_argument('-o', '--output', help='Output Excel path', default=None)

        p_generate = sub.add_parser('generate', help='Generate new FHX from edited Excel')
        p_generate.add_argument('fhx', help='Original FHX file')
        p_generate.add_argument('--setup', help='New Database.fhx reference file', required=True)
        p_generate.add_argument('--excel', help='Edited Excel file', required=True)
        p_generate.add_argument('-o', '--output', help='Output FHX path', default=None)

        args = parser.parse_args()

        if args.command == 'compare':
            output = args.output or os.path.splitext(args.fhx)[0] + '_comparison.xlsx'
            nameset_comp, sv_comp, expr_comp = compare_lib_and_export(
                args.fhx, args.setup, output,
                log_callback=print, progress_callback=cli_progress
            )
            print(f"\nExcel exported: {output}")
            print(f"  ENUMERATION_SET: {len(nameset_comp)}")
            print(f"  STRING_VALUE: {len(sv_comp)}")
            print(f"  Expression refs: {len(expr_comp)}")

        elif args.command == 'generate':
            output = args.output or os.path.splitext(args.fhx)[0] + '_NEW' + os.path.splitext(args.fhx)[1]
            print("Validating Excel data...")
            is_valid, validation_errors = validate_excel_for_generation(
                args.excel, log_callback=print
            )
            if not is_valid:
                print(f"\nGeneration aborted: {len(validation_errors)} issue(s) found.")
                sys.exit(1)
            nameset_changes, new_namesets, desc_changes, sv_changes, expr_changes, alarm_changes, priority_changes = read_lib_edited_excel(args.excel)
            count = generate_new_lib_fhx(
                args.fhx, args.setup, nameset_changes, new_namesets, desc_changes, sv_changes, expr_changes, output,
                alarm_changes=alarm_changes, priority_changes=priority_changes,
                log_callback=print, progress_callback=cli_progress
            )
            print(f"\nOutput: {output}")
            print(f"  Changes: {count}")

        else:
            parser.print_help()
    else:
        root = ctk.CTk()
        FHX_Migrator_App(root)
        root.mainloop()

if __name__ == '__main__':
    main()
