import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Listbox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import threading
import time
import os
import webbrowser
import json
import sys


# --- 核心修改：将配置文件路径指向系统AppData目录 ---
def get_app_data_path():
    """获取标准的系统应用数据目录，确保EXE所在目录不产生垃圾文件"""
    app_name = "Cyberpunk_Mailer_Data"  # 你的程序数据文件夹名

    # 获取 Windows 的 AppData/Roaming 目录
    # 例如: C:\Users\你的用户名\AppData\Roaming
    if os.name == 'nt':
        base_path = os.getenv('APPDATA')
    else:
        # 兼容 Mac/Linux，虽然主要跑exe是windows
        base_path = os.path.expanduser("~")

    # 拼接完整目录
    full_dir = os.path.join(base_path, app_name)

    # 如果目录不存在，自动创建
    if not os.path.exists(full_dir):
        os.makedirs(full_dir)

    # 返回 json 的完整路径
    return os.path.join(full_dir, "user_config.json")


# 初始化配置文件路径
CONFIG_FILE = get_app_data_path()


# -----------------------------------------------

class CyberpunkMailer:
    def __init__(self, root):
        self.root = root
        self.root.title("投稿小精灵————恋樱制造————仅供给希门信徒使用！")
        self.root.geometry("700x850")
        self.root.configure(bg="#050505")

        # 样式定义
        self.colors = {
            "bg": "#050505",
            "fg": "#00ff41",
            "accent": "#ff00ff",
            "cyan": "#00ffff",
            "input_bg": "#1a1a1a",
            "border": "#00ffff",
            "warning": "#ffff00"
        }
        self.font_main = ("Consolas", 10)
        self.font_header = ("Consolas", 14, "bold")
        self.font_small = ("Consolas", 8)

        self.attachment_path = None
        self.is_sending = False

        # 数据结构初始化
        self.folder_data = {"默认文件夹": ""}
        self.current_active_folder = "默认文件夹"

        # 加载配置
        self.saved_config = self.load_config_from_file()

        self._init_ui()

        # 填充加载的数据
        if self.saved_config:
            self.restore_settings()

    def _init_ui(self):
        tk.Label(self.root, text="投稿小精灵",
                 font=self.font_header, bg=self.colors["bg"], fg=self.colors["accent"]).pack(pady=10)

        main_frame = tk.Frame(self.root, bg=self.colors["bg"])
        main_frame.pack(padx=20, fill="both", expand=True)

        # 发件人设置
        self._create_label_entry(main_frame, "SENDER_EMAIL [发件箱]:", "sender_email")

        auth_frame = tk.Frame(main_frame, bg=self.colors["bg"])
        auth_frame.pack(fill="x", pady=(5, 0))
        tk.Label(auth_frame, text="AUTH_CODE [授权码]:", bg=self.colors["bg"], fg=self.colors["fg"],
                 font=self.font_main).pack(side="left")

        link_lbl = tk.Label(auth_frame, text="[?] 如何获取授权码", bg=self.colors["bg"], fg=self.colors["cyan"],
                            font=("Consolas", 9, "underline"), cursor="hand2")
        link_lbl.pack(side="left", padx=10)
        link_lbl.bind("<Button-1>", lambda e: self.open_url("https://service.mail.qq.com/detail/0/75"))

        self.auth_code_entry = tk.Entry(main_frame, bg=self.colors["input_bg"], fg="white",
                                        insertbackground=self.colors["accent"], relief="flat", show="*")
        self.auth_code_entry.pack(fill="x", pady=2)
        self._add_glow_border(self.auth_code_entry)

        self._create_label_entry(main_frame, "SMTP_SERVER:", "smtp_server", default="smtp.qq.com")

        # 文件夹管理
        tk.Label(main_frame, text="TARGET_DATABASE [编辑邮箱管理]:",
                 bg=self.colors["bg"], fg=self.colors["fg"], font=self.font_main, anchor="w").pack(fill="x",
                                                                                                   pady=(15, 2))

        folder_container = tk.Frame(main_frame, bg=self.colors["input_bg"], bd=1, relief="solid")
        folder_container.pack(fill="x", ipady=5)

        left_pane = tk.Frame(folder_container, bg=self.colors["bg"], width=150)
        left_pane.pack(side="left", fill="y", padx=2, pady=2)

        btn_bar = tk.Frame(left_pane, bg=self.colors["bg"])
        btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="[+]", command=self.add_folder, bg="#333", fg=self.colors["fg"],
                  relief="flat", width=3, cursor="hand2").pack(side="left")
        tk.Button(btn_bar, text="[-]", command=self.delete_folder, bg="#333", fg="red",
                  relief="flat", width=3, cursor="hand2").pack(side="right")

        self.folder_listbox = Listbox(left_pane, bg="#000", fg=self.colors["cyan"],
                                      selectbackground=self.colors["accent"], selectforeground="white",
                                      relief="flat", height=6, font=self.font_main, exportselection=False)
        self.folder_listbox.pack(fill="both", expand=True)
        self.folder_listbox.bind("<<ListboxSelect>>", self.on_folder_select)
        self.folder_listbox.bind("<Double-Button-1>", self.rename_folder)

        right_pane = tk.Frame(folder_container, bg=self.colors["bg"])
        right_pane.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        tk.Label(right_pane, text="EMAIL_LIST (One per line):", bg=self.colors["bg"], fg="#666",
                 font=self.font_small, anchor="w").pack(fill="x")

        self.recipients_text = tk.Text(right_pane, height=6, bg=self.colors["input_bg"],
                                       fg=self.colors["border"], insertbackground=self.colors["accent"],
                                       relief="flat", font=self.font_main)
        self.recipients_text.pack(fill="both", expand=True)
        self.recipients_text.bind("<KeyRelease>", self.save_current_text_to_memory)

        self.update_folder_listbox()

        # 邮件内容
        self._create_label_entry(main_frame, "SUBJECT [邮件标题]:", "subject")

        tk.Label(main_frame, text="BODY [正文内容]:",
                 bg=self.colors["bg"], fg=self.colors["fg"], font=self.font_main, anchor="w").pack(fill="x",
                                                                                                   pady=(10, 0))
        self.body_text = tk.Text(main_frame, height=5, bg=self.colors["input_bg"],
                                 fg="white", insertbackground=self.colors["accent"],
                                 relief="flat", font=self.font_main)
        self.body_text.pack(fill="x", pady=2)
        self._add_glow_border(self.body_text)

        # 按钮
        self.attach_btn = tk.Button(main_frame, text="[LOAD_DATA] 上传小说附件",
                                    command=self.upload_attachment,
                                    bg=self.colors["input_bg"], fg=self.colors["accent"],
                                    relief="flat", font=self.font_main, cursor="hand2")
        self.attach_btn.pack(fill="x", pady=5)
        self.file_label = tk.Label(main_frame, text="NO_DATA", bg=self.colors["bg"], fg="#555", font=self.font_small)
        self.file_label.pack()

        bottom_frame = tk.Frame(self.root, bg=self.colors["bg"])
        bottom_frame.pack(side="bottom", fill="x", pady=20, padx=20)

        tk.Button(bottom_frame, text="[SAVE_CONFIG] 保存设置", command=self.save_settings_to_file,
                  bg="#333", fg=self.colors["cyan"], relief="flat", font=("Consolas", 10), cursor="hand2").pack(
            side="left", fill="x", expand=True, padx=5)

        self.send_btn = tk.Button(bottom_frame, text=">>> INITIATE [开始投稿] <<<",
                                  command=self.start_sending_thread,
                                  bg=self.colors["fg"], fg="black",
                                  activebackground="white", relief="flat",
                                  font=("Consolas", 12, "bold"), cursor="hand2")
        self.send_btn.pack(side="left", fill="x", expand=True, padx=5)

        self.status_var = tk.StringVar(value="SYSTEM_READY")
        self.status_bar = tk.Label(self.root, textvariable=self.status_var,
                                   bg="#111", fg=self.colors["fg"], font=self.font_small, anchor="w")
        self.status_bar.pack(side="bottom", fill="x")

    def _create_label_entry(self, parent, text, attr_name, show=None, default=""):
        tk.Label(parent, text=text, bg=self.colors["bg"], fg=self.colors["fg"],
                 font=self.font_main, anchor="w").pack(fill="x", pady=(5, 0))
        entry = tk.Entry(parent, bg=self.colors["input_bg"], fg="white",
                         insertbackground=self.colors["accent"], relief="flat",
                         font=self.font_main, show=show)
        entry.pack(fill="x", pady=2)
        if default: entry.insert(0, default)
        setattr(self, f"{attr_name}_entry", entry)
        self._add_glow_border(entry)

    def _add_glow_border(self, widget):
        f = tk.Frame(widget.master, bg=self.colors["border"], height=1)
        f.pack(fill="x")

    def open_url(self, url):
        webbrowser.open(url)

    def update_folder_listbox(self):
        self.folder_listbox.delete(0, tk.END)
        for folder_name in self.folder_data.keys():
            self.folder_listbox.insert(tk.END, folder_name)

        try:
            idx = list(self.folder_data.keys()).index(self.current_active_folder)
            self.folder_listbox.select_set(idx)
            self.folder_listbox.activate(idx)
        except ValueError:
            if self.folder_data:
                first_key = list(self.folder_data.keys())[0]
                self.current_active_folder = first_key
                self.folder_listbox.select_set(0)
            else:
                self.current_active_folder = None

    def save_current_text_to_memory(self, event=None):
        if self.current_active_folder:
            content = self.recipients_text.get("1.0", tk.END).strip()
            self.folder_data[self.current_active_folder] = content

    def on_folder_select(self, event):
        selection = self.folder_listbox.curselection()
        if not selection:
            return

        self.folder_data[self.current_active_folder] = self.recipients_text.get("1.0", tk.END).strip()

        index = selection[0]
        folder_name = self.folder_listbox.get(index)
        self.current_active_folder = folder_name

        content = self.folder_data.get(folder_name, "")
        self.recipients_text.delete("1.0", tk.END)
        self.recipients_text.insert("1.0", content)

    def add_folder(self):
        new_name = simpledialog.askstring("NEW_DIR", "请输入新文件夹名称:")
        if new_name:
            if new_name in self.folder_data:
                messagebox.showerror("ERROR", "Folder already exists")
            else:
                self.folder_data[new_name] = ""
                self.update_folder_listbox()
                idx = list(self.folder_data.keys()).index(new_name)
                self.folder_listbox.select_clear(0, tk.END)
                self.folder_listbox.select_set(idx)
                self.on_folder_select(None)

    def delete_folder(self):
        if len(self.folder_data) <= 1:
            messagebox.showwarning("DENIED", "Cannot delete the last folder")
            return

        selected = self.folder_listbox.curselection()
        if not selected:
            return

        folder_name = self.folder_listbox.get(selected[0])
        if messagebox.askyesno("CONFIRM", f"Delete '{folder_name}'?"):
            del self.folder_data[folder_name]
            self.current_active_folder = list(self.folder_data.keys())[0]
            self.update_folder_listbox()
            self.on_folder_select(None)

    def rename_folder(self, event):
        selected = self.folder_listbox.curselection()
        if not selected:
            return
        old_name = self.folder_listbox.get(selected[0])

        new_name = simpledialog.askstring("RENAME", f"Rename '{old_name}' to:")
        if new_name and new_name != old_name:
            if new_name in self.folder_data:
                messagebox.showerror("ERROR", "Name exists.")
                return

            content = self.folder_data[old_name]
            del self.folder_data[old_name]
            self.folder_data[new_name] = content

            self.current_active_folder = new_name
            self.update_folder_listbox()

    # --- 保存和读取逻辑 ---
    def load_config_from_file(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"Config Load Error: {e}")
        return None

    def save_settings_to_file(self):
        self.save_current_text_to_memory()

        data = {
            "sender_email": self.sender_email_entry.get(),
            "auth_code": self.auth_code_entry.get(),
            "smtp_server": self.smtp_server_entry.get(),
            "folder_data": self.folder_data,
            "current_folder": self.current_active_folder
        }

        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.status_var.set("CONFIGURATION SAVED (INTERNAL)")
            messagebox.showinfo("SAVED", "配置已保存！下次继续使用吧~")
        except Exception as e:
            messagebox.showerror("ERROR", f"Save failed: {e}")

    def restore_settings(self):
        if not self.saved_config: return
        try:
            self.sender_email_entry.insert(0, self.saved_config.get("sender_email", ""))
            self.auth_code_entry.insert(0, self.saved_config.get("auth_code", ""))

            saved_smtp = self.saved_config.get("smtp_server", "")
            if saved_smtp:
                self.smtp_server_entry.delete(0, tk.END)
                self.smtp_server_entry.insert(0, saved_smtp)

            saved_folders = self.saved_config.get("folder_data", {})
            if saved_folders:
                self.folder_data = saved_folders

            saved_active = self.saved_config.get("current_folder", "")
            if saved_active and saved_active in self.folder_data:
                self.current_active_folder = saved_active

            self.update_folder_listbox()
            content = self.folder_data.get(self.current_active_folder, "")
            self.recipients_text.delete("1.0", tk.END)
            self.recipients_text.insert("1.0", content)

        except Exception as e:
            print(f"Restore Error: {e}")

    def upload_attachment(self):
        file_path = filedialog.askopenfilename(title="选择小说文件",
                                               filetypes=[("Documents", "*.txt *.doc *.docx *.pdf")])
        if file_path:
            self.attachment_path = file_path
            self.file_label.config(text=f"DATA: {os.path.basename(file_path)}", fg=self.colors["cyan"])

    def start_sending_thread(self):
        if self.is_sending: return

        self.save_current_text_to_memory()
        raw_recipients = self.recipients_text.get("1.0", tk.END).strip()
        recipients = [r.strip() for r in raw_recipients.split('\n') if r.strip()]

        sender = self.sender_email_entry.get().strip()
        auth = self.auth_code_entry.get().strip()
        smtp_srv = self.smtp_server_entry.get().strip()
        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", tk.END).strip()

        if not all([sender, auth, smtp_srv, recipients, subject, body]):
            messagebox.showerror("ERROR", "Incomplete Data")
            return

        if not self.attachment_path:
            if not messagebox.askyesno("WARNING", "No attachment. Proceed?"):
                return

        self.is_sending = True
        self.send_btn.config(state="disabled", text=">>> TRANSMITTING... <<<", bg="#333", fg="#999")

        thread = threading.Thread(target=self.run_batch_send,
                                  args=(sender, auth, smtp_srv, recipients, subject, body))
        thread.daemon = True
        thread.start()

    def run_batch_send(self, sender, auth, smtp_srv, recipients, subject, body):
        total = len(recipients)

        try:
            self.status_var.set(f"HANDSHAKE: {smtp_srv}...")
            server = smtplib.SMTP_SSL(smtp_srv, 465)
            server.login(sender, auth)
            server.quit()
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("LOGIN ERROR", str(e)))
            self._reset_ui_state()
            return

        for index, recipient in enumerate(recipients):
            current_num = index + 1
            self.status_var.set(f"SENDING [{current_num}/{total}] >> {recipient}")

            try:
                msg = MIMEMultipart()
                msg['From'] = sender
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))

                if self.attachment_path:
                    with open(self.attachment_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(self.attachment_path))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(self.attachment_path)}"'
                        msg.attach(part)

                with smtplib.SMTP_SSL(smtp_srv, 465) as server:
                    server.login(sender, auth)
                    server.send_message(msg)

                print(f"Sent to {recipient}")

                if current_num < total:
                    wait_time = 180
                    for t in range(wait_time, 0, -1):
                        self.status_var.set(f"COOLING DOWN: {t}s / 等待冷却")
                        time.sleep(1)
            except Exception as e:
                self.status_var.set(f"ERROR: {recipient}")
                time.sleep(2)

        self.status_var.set("ALL TASKS COMPLETED")
        self.root.after(0, lambda: messagebox.showinfo("REPORT", "Sequence Finished"))
        self._reset_ui_state()

    def _reset_ui_state(self):
        self.is_sending = False
        self.send_btn.config(state="normal", text=">>> INITIATE [开始投稿] <<<", bg=self.colors["fg"], fg="black")
        self.status_var.set("SYSTEM READY")


if __name__ == "__main__":
    root = tk.Tk()
    app = CyberpunkMailer(root)
    root.mainloop()