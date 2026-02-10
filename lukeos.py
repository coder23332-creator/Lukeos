import os
import json
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from PIL import Image, ImageTk
import subprocess
import threading
import webview
import time
import win32com.shell.shell as shell
import win32gui
import win32ui
from win32con import DI_NORMAL

# ================= PATHS =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE32 = os.path.join(BASE_DIR, "file32")
SHORTCUTS_DIR = os.path.join(FILE32, "shortcuts")
WALLPAPERS_DIR = os.path.join(FILE32, "wallpapers")
DATA_DIR = os.path.join(BASE_DIR, "data")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
BOOT_IMAGE = os.path.join(BASE_DIR, "boot.dir", "bootimg.png")

for d in [SHORTCUTS_DIR, WALLPAPERS_DIR, DATA_DIR]:
    os.makedirs(d, exist_ok=True)

DEFAULT_WALLPAPER = os.path.join(WALLPAPERS_DIR, "default.png")
if not os.path.exists(DEFAULT_WALLPAPER):
    Image.new("RGB", (1920,1080), color=(137,207,240)).save(DEFAULT_WALLPAPER)

# ================= DATA PERSISTENCE =================
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)

settings = load_settings()
settings.setdefault("wallpaper", DEFAULT_WALLPAPER)

# ================= ICON EXTRACTION =================
def extract_icon(path, size=32):
    try:
        large, small = shell.ExtractIconEx(path, 0)
        hicon = large[0] if large else small[0]

        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, size, size)
        hdc_mem = hdc.CreateCompatibleDC()
        hdc_mem.SelectObject(hbmp)

        win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon, size, size, 0, None, DI_NORMAL)

        bmpinfo = hbmp.GetInfo()
        bmpstr = hbmp.GetBitmapBits(True)

        img = Image.frombuffer("RGBA", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRA", 0, 1)
        return ImageTk.PhotoImage(img)
    except:
        return None

# ================= BOOT SCREEN =================
class BootScreen(tk.Toplevel):
    def __init__(self, master, delay=2.5):
        super().__init__(master)
        self.overrideredirect(True)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        if os.path.exists(BOOT_IMAGE):
            img = Image.open(BOOT_IMAGE)
            img = img.resize((sw, sh), Image.Resampling.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            tk.Label(self, image=self.photo).pack(fill="both", expand=True)
        else:
            tk.Label(self, text="LukeOS Booting...", font=("Arial", 40)).pack(expand=True)
        self.after(int(delay*1000), self.destroy)

# ================= DESKTOP =================
class Desktop:
    def __init__(self, root):
        self.root = root
        self.selected = None
        self.buttons = []
        self.wallpaper_path = settings.get("wallpaper", DEFAULT_WALLPAPER)
        self.set_wallpaper(self.wallpaper_path)
        self.load()

    def set_wallpaper(self, path):
        self.wallpaper_path = path
        settings["wallpaper"] = path
        save_settings(settings)
        try:
            self.wallpaper_image = Image.open(path)
            self.wallpaper_image = self.wallpaper_image.resize(
                (self.root.winfo_screenwidth(), self.root.winfo_screenheight()),
                Image.Resampling.LANCZOS
            )
            self.wallpaper_photo = ImageTk.PhotoImage(self.wallpaper_image)
            if hasattr(self, "wallpaper_label"):
                self.wallpaper_label.config(image=self.wallpaper_photo)
            else:
                self.wallpaper_label = tk.Label(self.root, image=self.wallpaper_photo)
                self.wallpaper_label.place(x=0, y=0, relwidth=1, relheight=1)
        except:
            pass

    def load(self):
        for b in self.buttons:
            b.destroy()
        self.buttons.clear()
        files = sorted(f for f in os.listdir(SHORTCUTS_DIR) if f.endswith(".json"))
        for i, file in enumerate(files):
            data = json.load(open(os.path.join(SHORTCUTS_DIR, file)))
            self.create_button(i, data["name"], data["path"])

    def create_button(self, index, name, path):
        icon = extract_icon(path)
        frame = tk.Frame(self.root, width=90, height=90, bg="#C0C0C0", relief="raised", bd=2)
        frame.path = path
        frame.filename = os.path.join(SHORTCUTS_DIR, f"{name}.json")
        if icon:
            lbl_i = tk.Label(frame, image=icon, bg="#C0C0C0")
            lbl_i.image = icon
            lbl_i.pack(pady=(6,0))
        lbl_t = tk.Label(frame, text=name, bg="#C0C0C0", wraplength=80, justify="center")
        lbl_t.pack(pady=(2,0))

        cols = 6
        x = 30 + (index % cols) * 110
        y = 30 + (index // cols) * 110
        frame.place(x=x, y=y)

        frame.bind("<Button-1>", lambda e,f=frame: self.select(f))
        frame.bind("<Double-Button-1>", lambda e,p=path: self.launch(p))
        frame.bind("<Button-3>", lambda e,f=frame: self.menu(e,f))

        for w in frame.winfo_children():
            w.bind("<Button-1>", lambda e,f=frame: self.select(f))
            w.bind("<Double-Button-1>", lambda e,p=path: self.launch(p))
            w.bind("<Button-3>", lambda e,f=frame: self.menu(e,f))

        self.buttons.append(frame)

    def select(self, frame):
        if self.selected:
            self.selected.config(relief="raised")
        frame.config(relief="sunken")
        self.selected = frame

    def launch(self, path):
        try:
            os.startfile(path)
        except:
            messagebox.showerror("Error", f"Cannot open {path}")

    def menu(self, event, frame):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="Rename Shortcut", command=lambda: self.rename(frame))
        m.add_command(label="Delete Shortcut", command=lambda: self.delete(frame))
        m.tk_popup(event.x_root, event.y_root)

    def rename(self, frame):
        new_name = simpledialog.askstring("Rename Shortcut", "Enter new name:", initialvalue=os.path.splitext(os.path.basename(frame.filename))[0])
        if new_name:
            old_file = frame.filename
            new_file = os.path.join(SHORTCUTS_DIR, f"{new_name}.json")
            if os.path.exists(old_file):
                data = json.load(open(old_file))
                data["name"] = new_name
                json.dump(data, open(new_file, "w"), indent=4)
                os.remove(old_file)
            self.load()

    def delete(self, frame):
        if os.path.exists(frame.filename):
            os.remove(frame.filename)
        self.load()

# ================= LUKEOS =================
class LukeOS:
    def __init__(self, root):
        self.root = root
        self.root.title("LukeOS")
        self.root.attributes("-fullscreen", True)
        self.root.configure(bg="#008080")

        self.desktop = Desktop(root)
        self.create_taskbar()
        self.create_start_menu()
        self.root.mainloop()

    def create_taskbar(self):
        self.taskbar = tk.Frame(self.root, bg="#C0C0C0", height=36)
        self.taskbar.pack(side="bottom", fill="x")
        tk.Button(self.taskbar, text="Start", command=self.toggle_start).pack(side="left")
        self.time_label = tk.Label(self.taskbar, text="", bg="#C0C0C0")
        self.time_label.pack(side="right", padx=5)
        self.update_time()

    def update_time(self):
        self.time_label.config(text=time.strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self.update_time)

    def create_start_menu(self):
        self.start = tk.Frame(self.root, bg="#D4D0C8", width=240, relief="raised", bd=2)
        self.visible = False
        tk.Button(self.start, text="File Explorer", anchor="w",
                  command=lambda: FileExplorer(self.root)).pack(fill="x")
        tk.Button(self.start, text="Add Desktop Shortcut", anchor="w",
                  command=self.add_desktop_shortcut).pack(fill="x")
        tk.Button(self.start, text="Add Start Menu Shortcut", anchor="w",
                  command=self.add_start_menu_shortcut).pack(fill="x")
        tk.Button(self.start, text="Add Taskbar Shortcut", anchor="w",
                  command=self.add_taskbar_shortcut).pack(fill="x")
        tk.Button(self.start, text="LukeOS Browser", anchor="w",
                  command=self.open_browser).pack(fill="x")
        tk.Button(self.start, text="Change Wallpaper", anchor="w",
                  command=self.select_wallpaper).pack(fill="x")
        tk.Button(self.start, text="Exit LukeOS", anchor="w",
                  command=self.root.quit).pack(fill="x")

    def toggle_start(self):
        if self.visible:
            self.start.place_forget()
        else:
            self.start.place(x=0, y=self.root.winfo_height()-self.start.winfo_reqheight()-36)
        self.visible = not self.visible

    # ---------- Shortcuts ----------
    def add_desktop_shortcut(self):
        path = filedialog.askopenfilename()
        if not path: return
        name = simpledialog.askstring("Shortcut Name", "Enter a name for this shortcut:")
        if not name:
            name = os.path.splitext(os.path.basename(path))[0]
        shortcut_file = os.path.join(SHORTCUTS_DIR, f"{name}.json")
        json.dump({"name": name, "path": path}, open(shortcut_file, "w"), indent=4)
        self.desktop.load()

    def add_start_menu_shortcut(self):
        messagebox.showinfo("Info", "Start menu shortcut system not yet implemented.")

    def add_taskbar_shortcut(self):
        messagebox.showinfo("Info", "Taskbar shortcut system not yet implemented.")

    # ---------- Browser ----------
    def open_browser(self):
        win = tk.Toplevel(self.root)
        win.title("LukeOS Browser")
        win.geometry("400x100")
        entry_frame = tk.Frame(win)
        entry_frame.pack(fill="x", padx=5, pady=5)
        url_var = tk.StringVar(value="https://www.google.com")
        url_entry = tk.Entry(entry_frame, textvariable=url_var, width=40)
        url_entry.pack(side="left", fill="x", expand=True)
        def open_webview():
            url = url_var.get()
            if not url.startswith("http"):
                url = "https://www.google.com/search?q=" + url.replace(" ","+")
            webview.create_window("LukeOS Browser", url, width=900, height=600)
            webview.start()
        go_button = tk.Button(entry_frame, text="Go", command=open_webview)
        go_button.pack(side="left", padx=5)
        url_entry.bind("<Return>", lambda e: open_webview())

    # ---------- Wallpaper Selection ----------
    def select_wallpaper(self):
        win = tk.Toplevel(self.root)
        win.title("Select Wallpaper")
        win.geometry("500x400")
        tk.Label(win, text="Select Wallpaper:", font=("Arial", 12)).pack(pady=5)
        canvas = tk.Canvas(win)
        scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        wallpapers = [f for f in os.listdir(WALLPAPERS_DIR) if f.lower().endswith((".png",".jpg",".jpeg"))]
        for wp in wallpapers:
            path = os.path.join(WALLPAPERS_DIR, wp)
            try:
                img = Image.open(path)
                img.thumbnail((150,100))
                photo = ImageTk.PhotoImage(img)
                btn = tk.Button(scrollable_frame, image=photo, text=wp, compound="top",
                                command=lambda p=path: self.apply_wallpaper(p))
                btn.image = photo
                btn.pack(padx=5,pady=5)
            except:
                continue

    def apply_wallpaper(self, path):
        self.desktop.set_wallpaper(path)

# ================= FILE EXPLORER =================
class FileExplorer(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("File Explorer")
        self.geometry("700x450")
        self.path = BASE_DIR
        self.path_label = tk.Label(self, text=self.path, anchor="w")
        self.path_label.pack(fill="x")
        self.listbox = tk.Listbox(self)
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<Double-Button-1>", self.open_item)
        tk.Button(self, text="Back", command=self.go_back).pack(fill="x")
        self.refresh()
    def refresh(self):
        self.listbox.delete(0, "end")
        self.path_label.config(text=self.path)
        for item in os.listdir(self.path):
            self.listbox.insert("end", item)
    def open_item(self, _):
        item = self.listbox.get(self.listbox.curselection())
        p = os.path.join(self.path, item)
        if os.path.isdir(p):
            self.path = p
            self.refresh()
        else:
            os.startfile(p)
    def go_back(self):
        parent = os.path.dirname(self.path)
        if os.path.exists(parent):
            self.path = parent
            self.refresh()

# ================= START =================
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    boot = BootScreen(root)
    boot.wait_window()
    root.deiconify()
    LukeOS(root)
