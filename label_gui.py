import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# Drag & drop support (optional)
# Install once:
#   py -m pip install tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    USE_DND = True
except Exception:
    USE_DND = False

ENGINE_SCRIPT = os.path.join(os.path.dirname(__file__), "auto_labels_from_workorder.py")


def run_engine(input_path: str, status_var: tk.StringVar):
    input_path = input_path.strip().strip('"')
    if not input_path:
        return

    if not os.path.exists(input_path):
        messagebox.showerror("File not found", f"Can't find:\n{input_path}")
        return

    if not os.path.exists(ENGINE_SCRIPT):
        messagebox.showerror("Missing engine script", f"Can't find:\n{ENGINE_SCRIPT}")
        return

    status_var.set("Generating labels…")
    root.update_idletasks()

    cmd = [sys.executable, ENGINE_SCRIPT, input_path]

    try:
        p = subprocess.run(cmd, capture_output=True, text=True)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        status_var.set("Idle.")
        return

    out = (p.stdout or "").strip()
    err = (p.stderr or "").strip()

    if p.returncode == 0:
        status_var.set("Done ✅")
        messagebox.showinfo("Success", out if out else "Labels generated.")
    else:
        status_var.set("Failed ❌")
        msg = "Label generation failed.\n\n"
        if out:
            msg += "Output:\n" + out + "\n\n"
        if err:
            msg += "Error:\n" + err
        messagebox.showerror("Failed", msg)


def browse_file(status_var):
    path = filedialog.askopenfilename(
        title="Select Work Order PDF (or image)",
        filetypes=[
            ("PDF files", "*.pdf"),
            ("Images", "*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp"),
            ("All files", "*.*"),
        ],
    )
    if path:
        run_engine(path, status_var)


def on_drop(event, status_var):
    data = event.data.strip()
    if data.startswith("{") and data.endswith("}"):
        data = data[1:-1]
    run_engine(data, status_var)


# ---- UI ----
if USE_DND:
    root = TkinterDnD.Tk()
else:
    root = tk.Tk()

root.title("JTI Label Maker")
root.geometry("520x320")  # increased height so buttons don't get clipped
root.resizable(False, False)

status_var = tk.StringVar(value="Drop a work order PDF here, or click Browse.")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(fill="both", expand=True)

title = tk.Label(frame, text="JTI Label Maker", font=("Segoe UI", 16, "bold"))
title.pack(anchor="w")

sub = tk.Label(frame, text="Drop a work order PDF (or image) to generate labels.", font=("Segoe UI", 10))
sub.pack(anchor="w", pady=(2, 12))

drop_box = tk.Label(
    frame,
    text="DROP FILE HERE",
    font=("Segoe UI", 14, "bold"),
    bg="#f2f2f2",
    relief="ridge",
    bd=2,
    height=6,
)
drop_box.pack(pady=(0, 12), fill="x")

if USE_DND:
    drop_box.drop_target_register(DND_FILES)
    drop_box.dnd_bind("<<Drop>>", lambda e: on_drop(e, status_var))
else:
    drop_box.configure(text="Drag & drop disabled.\nInstall: py -m pip install tkinterdnd2")

btn_row = tk.Frame(frame)
btn_row.pack(fill="x", pady=(10, 10))

browse_btn = tk.Button(
    btn_row,
    text="Browse…",
    command=lambda: browse_file(status_var),
    width=18,
    font=("Segoe UI", 11, "bold"),
    pady=8
)
browse_btn.pack(side="left", padx=8)

quit_btn = tk.Button(
    btn_row,
    text="Quit",
    command=root.destroy,
    width=18,
    font=("Segoe UI", 11),
    pady=8
)
quit_btn.pack(side="right", padx=8)

status = tk.Label(frame, textvariable=status_var, font=("Segoe UI", 10))
status.pack(anchor="w", pady=(6, 0))

note = tk.Label(
    frame,
    text="Outputs a single PDF named: JobNumber_JobName_Labels.pdf",
    font=("Segoe UI", 9),
    fg="#444",
)
note.pack(anchor="w", pady=(6, 0))

root.mainloop()