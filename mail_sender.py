#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Michaelsoft Mail Sender - 邮件自动发送器
版本: 1.1.0
"""

import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import json
import sys
import logging
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("mail_sender.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("MailSender")

class MailSender:
    """邮件发送器核心类"""
    
    def __init__(self, config=None):
        """初始化邮件发送器
        
        Args:
            config (dict, optional): 邮件配置信息. 默认为 None.
        """
        self.config = config or {}
        self.load_config()
    
    def load_config(self, config_file="mail_config.json"):
        """从配置文件加载配置
        
        Args:
            config_file (str, optional): 配置文件路径. 默认为 "mail_config.json".
        """
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.config.update(json.load(f))
                logger.info(f"配置已从 {config_file} 加载")
            else:
                logger.warning(f"配置文件 {config_file} 不存在，使用默认配置")
        except Exception as e:
            logger.error(f"加载配置文件失败: {str(e)}")
    
    def save_config(self, config_file="mail_config.json"):
        """保存配置到文件
        
        Args:
            config_file (str, optional): 配置文件路径. 默认为 "mail_config.json".
        """
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            logger.info(f"配置已保存到 {config_file}")
        except Exception as e:
            logger.error(f"保存配置文件失败: {str(e)}")
    
    def send_mail(self, recipients, subject, body, attachments=None, html=False):
        """发送邮件
        
        Args:
            recipients (list): 收件人列表
            subject (str): 邮件主题
            body (str): 邮件正文
            attachments (list, optional): 附件路径列表. 默认为 None.
            html (bool, optional): 是否为HTML格式. 默认为 False.
        
        Returns:
            bool: 发送成功返回True，否则返回False
        """
        if not recipients:
            logger.error("收件人列表为空")
            return False
        
        if not self.config.get("smtp_server") or not self.config.get("smtp_port"):
            logger.error("SMTP服务器配置不完整")
            return False
        
        if not self.config.get("sender_email") or not self.config.get("password"):
            logger.error("发件人信息配置不完整")
            return False
        
        # 创建邮件
        msg = MIMEMultipart()
        msg['From'] = self.config.get("sender_email")
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = subject
        
        # 添加正文
        if html:
            msg.attach(MIMEText(body, 'html', 'utf-8'))
        else:
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 添加附件
        if attachments:
            for file_path in attachments:
                try:
                    with open(file_path, 'rb') as file:
                        part = MIMEApplication(file.read(), Name=os.path.basename(file_path))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                    msg.attach(part)
                except Exception as e:
                    logger.error(f"添加附件 {file_path} 失败: {str(e)}")
        
        # 发送邮件
        try:
            context = ssl.create_default_context()
            
            with smtplib.SMTP_SSL(self.config.get("smtp_server"), 
                                 int(self.config.get("smtp_port")), 
                                 context=context) as server:
                server.login(self.config.get("sender_email"), self.config.get("password"))
                server.sendmail(self.config.get("sender_email"), recipients, msg.as_string())
            
            logger.info(f"邮件已成功发送给 {', '.join(recipients)}")
            return True
        
        except Exception as e:
            logger.error(f"发送邮件失败: {str(e)}")
            return False


class MailSenderGUI:
    """邮件发送器图形界面"""
    
    def __init__(self, root):
        """初始化图形界面
        
        Args:
            root: Tkinter根窗口
        """
        self.root = root
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        self.mail_sender = MailSender()
        self.attachments = []
        
        # 加载语言配置
        self.current_lang = "zh_CN"  # 默认中文
        self.langs = {}
        self._load_languages()
        
        self._create_widgets()
        self._load_config_to_ui()
        
    def _load_languages(self):
        """加载语言配置文件"""
        lang_files = ["lang_zh_CN.json", "lang_en_US.json"]
        for lang_file in lang_files:
            try:
                if os.path.exists(lang_file):
                    lang_code = lang_file.replace("lang_", "").replace(".json", "")
                    with open(lang_file, 'r', encoding='utf-8') as f:
                        self.langs[lang_code] = json.load(f)
                    logger.info(f"语言配置已从 {lang_file} 加载")
            except Exception as e:
                logger.error(f"加载语言配置文件失败: {str(e)}")
        
        # 如果没有加载到任何语言文件，使用默认的中文
        if not self.langs:
            logger.warning("未找到语言配置文件，使用内置中文")
            self.langs["zh_CN"] = {
                "app_title": "Michaelsoft Mail Sender v1.1.0",
                "status_ready": "就绪"
                # 其他默认值...
            }
        
        # 设置窗口标题
        self.root.title(self.langs.get(self.current_lang, {}).get("app_title", "Michaelsoft Mail Sender v1.1.0"))
    
    def _create_widgets(self):
        """创建界面组件"""
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 语言菜单
        lang_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self._get_text("language"), menu=lang_menu)
        lang_menu.add_command(label=self._get_text("lang_zh_CN"), command=lambda: self._change_language("zh_CN"))
        lang_menu.add_command(label=self._get_text("lang_en_US"), command=lambda: self._change_language("en_US"))
        
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 发送邮件标签页
        send_frame = ttk.Frame(notebook)
        notebook.add(send_frame, text=self._get_text("tab_send"))
        
        # 配置标签页
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text=self._get_text("tab_config"))
        
        # 帮助标签页
        help_frame = ttk.Frame(notebook)
        notebook.add(help_frame, text=self._get_text("tab_help"))
        
        # 保存notebook引用以便后续更新标签文本
        self.notebook = notebook
        self.notebook_frames = {
            "send": send_frame,
            "config": config_frame,
            "help": help_frame
        }
        
        # 发送邮件页面
        self._create_send_page(send_frame)
        
        # 配置页面
        self._create_config_page(config_frame)
        
        # 帮助页面
        self._create_help_page(help_frame)
        
        # 状态栏
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10)
        
        self.status_var = tk.StringVar()
        self.status_var.set(self._get_text("status_ready"))
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        version_label = ttk.Label(status_frame, text="v1.1.0")
        version_label.pack(side=tk.RIGHT)
        
    def _get_text(self, key, *args):
        """获取当前语言的文本
        
        Args:
            key: 文本键名
            *args: 格式化参数
            
        Returns:
            str: 对应语言的文本
        """
        text = self.langs.get(self.current_lang, {}).get(key, key)
        if args:
            try:
                return text.format(*args)
            except:
                return text
        return text
        
    def _change_language(self, lang_code):
        """切换语言
        
        Args:
            lang_code: 语言代码
        """
        if lang_code in self.langs and lang_code != self.current_lang:
            self.current_lang = lang_code
            self._update_ui_language()
            
    def _update_ui_language(self):
        """更新界面语言"""
        # 更新窗口标题
        self.root.title(self._get_text("app_title"))
        
        # 更新标签页文本
        self.notebook.tab(0, text=self._get_text("tab_send"))
        self.notebook.tab(1, text=self._get_text("tab_config"))
        self.notebook.tab(2, text=self._get_text("tab_help"))
        
        # 更新状态栏
        self.status_var.set(self._get_text("status_ready"))
        
        # 重新创建各页面
        for frame in self.notebook_frames.values():
            for widget in frame.winfo_children():
                widget.destroy()
                
        self._create_send_page(self.notebook_frames["send"])
        self._create_config_page(self.notebook_frames["config"])
        self._create_help_page(self.notebook_frames["help"])
        
        # 重新加载配置到UI
        self._load_config_to_ui()
    
    def _create_send_page(self, parent):
        """创建发送邮件页面
        
        Args:
            parent: 父容器
        """
        # 收件人
        ttk.Label(parent, text=self._get_text("recipients")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.recipients_entry = ttk.Entry(parent, width=70)
        self.recipients_entry.grid(row=0, column=1, sticky=tk.W+tk.E, padx=10, pady=5)
        ttk.Label(parent, text=self._get_text("recipients_hint")).grid(row=0, column=2, sticky=tk.W, padx=10, pady=5)
        
        # 主题
        ttk.Label(parent, text=self._get_text("subject")).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.subject_entry = ttk.Entry(parent, width=70)
        self.subject_entry.grid(row=1, column=1, columnspan=2, sticky=tk.W+tk.E, padx=10, pady=5)
        
        # 正文
        ttk.Label(parent, text=self._get_text("body")).grid(row=2, column=0, sticky=tk.NW, padx=10, pady=5)
        self.body_text = tk.Text(parent, height=15)
        self.body_text.grid(row=2, column=1, columnspan=2, sticky=tk.W+tk.E+tk.N+tk.S, padx=10, pady=5)
        
        # HTML选项
        self.html_var = tk.BooleanVar()
        html_check = ttk.Checkbutton(parent, text=self._get_text("html_format"), variable=self.html_var)
        html_check.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        
        # 附件
        ttk.Label(parent, text=self._get_text("attachments")).grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        
        attachments_frame = ttk.Frame(parent)
        attachments_frame.grid(row=4, column=1, sticky=tk.W+tk.E, padx=10, pady=5)
        
        self.attachments_var = tk.StringVar()
        self.attachments_var.set(self._get_text("no_attachments"))
        ttk.Label(attachments_frame, textvariable=self.attachments_var).pack(side=tk.LEFT)
        
        ttk.Button(attachments_frame, text=self._get_text("add_attachment"), command=self._add_attachment).pack(side=tk.LEFT, padx=5)
        ttk.Button(attachments_frame, text=self._get_text("clear_attachments"), command=self._clear_attachments).pack(side=tk.LEFT)
        
        # 发送按钮
        send_button = ttk.Button(parent, text=self._get_text("send_mail"), command=self._send_mail)
        send_button.grid(row=5, column=1, pady=20)
        
        # 配置网格权重
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)
    
    def _create_config_page(self, parent):
        """创建配置页面
        
        Args:
            parent: 父容器
        """
        # SMTP服务器
        ttk.Label(parent, text=self._get_text("smtp_server")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.smtp_server_entry = ttk.Entry(parent, width=40)
        self.smtp_server_entry.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        
        # SMTP端口
        ttk.Label(parent, text=self._get_text("smtp_port")).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.smtp_port_entry = ttk.Entry(parent, width=10)
        self.smtp_port_entry.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        # 发件人邮箱
        ttk.Label(parent, text=self._get_text("sender_email")).grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        self.sender_email_entry = ttk.Entry(parent, width=40)
        self.sender_email_entry.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        
        # 密码/授权码
        ttk.Label(parent, text=self._get_text("password")).grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.password_entry = ttk.Entry(parent, width=40, show="*")
        self.password_entry.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        
        # 常用服务器配置
        ttk.Label(parent, text=self._get_text("common_servers")).grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        
        servers_frame = ttk.Frame(parent)
        servers_frame.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Button(servers_frame, text="QQ邮箱", 
                  command=lambda: self._set_preset_server("smtp.qq.com", "465")).pack(side=tk.LEFT, padx=5)
        ttk.Button(servers_frame, text="163邮箱", 
                  command=lambda: self._set_preset_server("smtp.163.com", "465")).pack(side=tk.LEFT, padx=5)
        ttk.Button(servers_frame, text="Gmail", 
                  command=lambda: self._set_preset_server("smtp.gmail.com", "465")).pack(side=tk.LEFT, padx=5)
        ttk.Button(servers_frame, text="Outlook", 
                  command=lambda: self._set_preset_server("smtp.office365.com", "587")).pack(side=tk.LEFT, padx=5)
        
        # 保存配置按钮
        save_button = ttk.Button(parent, text=self._get_text("save_config"), command=self._save_config)
        save_button.grid(row=5, column=1, sticky=tk.W, padx=10, pady=20)
        
        # 测试连接按钮
        test_button = ttk.Button(parent, text=self._get_text("test_connection"), command=self._test_connection)
        test_button.grid(row=5, column=0, sticky=tk.E, padx=10, pady=20)
    
    def _create_help_page(self, parent):
        """创建帮助页面
        
        Args:
            parent: 父容器
        """
        help_text = tk.Text(parent, wrap=tk.WORD, padx=10, pady=10)
        help_text.pack(fill=tk.BOTH, expand=True)
        
        help_content = self._get_text("help_content")
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)
    
    def _load_config_to_ui(self):
        """将配置加载到界面"""
        config = self.mail_sender.config
        
        # 加载语言设置
        if "language" in config and config["language"] in self.langs:
            if self.current_lang != config["language"]:
                self.current_lang = config["language"]
                # 不在这里调用 _update_ui_language，避免递归调用
        
        if "smtp_server" in config:
            self.smtp_server_entry.delete(0, tk.END)
            self.smtp_server_entry.insert(0, config["smtp_server"])
        
        if "smtp_port" in config:
            self.smtp_port_entry.delete(0, tk.END)
            self.smtp_port_entry.insert(0, config["smtp_port"])
        
        if "sender_email" in config:
            self.sender_email_entry.delete(0, tk.END)
            self.sender_email_entry.insert(0, config["sender_email"])
        
        if "password" in config:
            self.password_entry.delete(0, tk.END)
            self.password_entry.insert(0, config["password"])
    
    def _save_config(self):
        """保存配置"""
        config = {
            "smtp_server": self.smtp_server_entry.get(),
            "smtp_port": self.smtp_port_entry.get(),
            "sender_email": self.sender_email_entry.get(),
            "password": self.password_entry.get(),
            "language": self.current_lang  # 保存当前语言设置
        }
        
        self.mail_sender.config = config
        self.mail_sender.save_config()
        
        self.status_var.set(self._get_text("config_saved"))
        messagebox.showinfo(self._get_text("success"), self._get_text("config_saved_msg"))
    
    def _set_preset_server(self, server, port):
        """设置预设服务器
        
        Args:
            server (str): SMTP服务器地址
            port (str): SMTP服务器端口
        """
        self.smtp_server_entry.delete(0, tk.END)
        self.smtp_server_entry.insert(0, server)
        
        self.smtp_port_entry.delete(0, tk.END)
        self.smtp_port_entry.insert(0, port)
        
        self.status_var.set(self._get_text("server_set", server, port))
    
    def _add_attachment(self):
        """添加附件"""
        files = filedialog.askopenfilenames(title="选择附件")
        if files:
            self.attachments.extend(files)
            self._update_attachments_label()
    
    def _clear_attachments(self):
        """清除附件"""
        self.attachments = []
        self._update_attachments_label()
    
    def _update_attachments_label(self):
        """更新附件标签"""
        if not self.attachments:
            self.attachments_var.set(self._get_text("no_attachments"))
        elif len(self.attachments) == 1:
            self.attachments_var.set(os.path.basename(self.attachments[0]))
        else:
            self.attachments_var.set(self._get_text("attachments_selected", len(self.attachments)))
    
    def _send_mail(self):
        """发送邮件"""
        recipients = [r.strip() for r in self.recipients_entry.get().split(",") if r.strip()]
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END)
        html = self.html_var.get()
        
        if not recipients:
            messagebox.showerror(self._get_text("error"), self._get_text("empty_recipients"))
            return
        
        if not subject:
            if not messagebox.askyesno(self._get_text("warning"), self._get_text("empty_subject")):
                return
        
        self.status_var.set(self._get_text("sending"))
        self.root.update()
        
        success = self.mail_sender.send_mail(recipients, subject, body, self.attachments, html)
        
        if success:
            self.status_var.set(self._get_text("send_success"))
            messagebox.showinfo(self._get_text("success"), self._get_text("send_success_msg", len(recipients)))
        else:
            self.status_var.set(self._get_text("send_fail"))
            messagebox.showerror(self._get_text("error"), self._get_text("send_fail_msg"))
    
    def _test_connection(self):
        """测试SMTP连接"""
        server = self.smtp_server_entry.get()
        port = self.smtp_port_entry.get()
        email = self.sender_email_entry.get()
        password = self.password_entry.get()
        
        if not server or not port or not email or not password:
            messagebox.showerror(self._get_text("error"), self._get_text("incomplete_smtp"))
            return
        
        self.status_var.set(self._get_text("testing_connection"))
        self.root.update()
        
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(server, int(port), context=context) as smtp_server:
                smtp_server.login(email, password)
                self.status_var.set(self._get_text("test_success"))
                messagebox.showinfo(self._get_text("success"), self._get_text("test_success_msg"))
        except Exception as e:
            self.status_var.set(self._get_text("test_fail"))
            messagebox.showerror(self._get_text("error"), self._get_text("test_fail_msg", str(e)))


def main():
    """主函数"""
    root = tk.Tk()
    app = MailSenderGUI(root)
    
    # 如果配置中有语言设置，则应用它
    if "language" in app.mail_sender.config:
        app._change_language(app.mail_sender.config["language"])
    
    root.mainloop()


if __name__ == "__main__":
    main()
