#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
朗润播放器客户端 - 独立版
功能：Excel导入、文件下载、媒体播放
特点：纯Python标准库实现，无需外部依赖
作者：AI助手
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import urllib.request
import urllib.parse
import os
import threading
import time
import json
import subprocess
import sys
import csv
import re
from pathlib import Path
import ssl
import socket

# 禁用SSL验证（处理某些下载链接的SSL问题）
ssl._create_default_https_context = ssl._create_unverified_context

class SimpleExcelReader:
    """简化的Excel读取器（纯Python实现）"""
    
    @staticmethod
    def read_csv(file_path):
        """读取CSV文件"""
        try:
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                data = list(reader)
                columns = reader.fieldnames
            return data, columns
        except UnicodeDecodeError:
            # 尝试其他编码
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    reader = csv.DictReader(f)
                    data = list(reader)
                    columns = reader.fieldnames
                return data, columns
            except:
                raise Exception("无法读取文件，请确保文件编码为UTF-8或GBK")
    
    @staticmethod
    def read_excel_simple(file_path):
        """简单的Excel读取（转换为CSV后读取）"""
        # 提示用户将Excel转换为CSV
        raise Exception("请将Excel文件另存为CSV格式后重新导入，或安装pandas库支持Excel文件")
    
    @staticmethod
    def read_file(file_path):
        """统一的文件读取接口"""
        if file_path.lower().endswith('.csv'):
            return SimpleExcelReader.read_csv(file_path)
        elif file_path.lower().endswith(('.xlsx', '.xls')):
            return SimpleExcelReader.read_excel_simple(file_path)
        else:
            raise Exception("不支持的文件格式，请使用CSV或Excel文件")

class MediaDownloader:
    """媒体文件下载器（纯Python实现）"""
    
    def __init__(self, progress_callback=None, log_callback=None):
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.download_dir = "downloaded_media"
        self.downloaded_files = {}
        self.load_download_history()
        
    def load_download_history(self):
        """加载下载历史"""
        history_file = os.path.join(self.download_dir, "download_history.json")
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    self.downloaded_files = json.load(f)
            except Exception as e:
                self.log(f"加载下载历史失败: {e}")
                
    def save_download_history(self):
        """保存下载历史"""
        os.makedirs(self.download_dir, exist_ok=True)
        history_file = os.path.join(self.download_dir, "download_history.json")
        try:
            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(self.downloaded_files, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存下载历史失败: {e}")
            
    def log(self, message):
        """记录日志"""
        print(f"[{time.strftime('%H:%M:%S')}] {message}")
        if self.log_callback:
            self.log_callback(message)
            
    def get_safe_filename(self, url, original_name=""):
        """生成安全的文件名"""
        parsed = urllib.parse.urlparse(url)
        filename = os.path.basename(parsed.path)
        
        if not filename or '.' not in filename:
            if original_name:
                # 清理原始名称并添加扩展名
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', original_name)
                # 尝试从URL或Content-Type推断扩展名
                if not '.' in safe_name:
                    safe_name += '.mp4'  # 默认扩展名
                filename = safe_name
            else:
                filename = f"media_{int(time.time())}.mp4"
                
        # 清理文件名
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        return filename
        
    def download_file(self, url, display_name, performance_number):
        """下载单个文件（使用urllib）"""
        try:
            # 检查是否已下载
            if url in self.downloaded_files:
                local_path = self.downloaded_files[url]
                if os.path.exists(local_path):
                    self.log(f"文件已存在，跳过下载: {display_name}")
                    return local_path
                    
            filename = self.get_safe_filename(url, display_name)
            local_path = os.path.join(self.download_dir, filename)
            
            # 确保目录存在
            os.makedirs(self.download_dir, exist_ok=True)
            
            self.log(f"开始下载: {display_name}")
            
            # 创建请求
            req = urllib.request.Request(url)
            req.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
            
            # 下载文件
            with urllib.request.urlopen(req, timeout=30) as response:
                total_size = int(response.headers.get('Content-Length', 0))
                downloaded = 0
                
                with open(local_path, 'wb') as f:
                    while True:
                        chunk = response.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        if self.progress_callback and total_size > 0:
                            progress = (downloaded / total_size) * 100
                            self.progress_callback(progress)
                            
            # 记录下载成功
            self.downloaded_files[url] = local_path
            self.save_download_history()
            
            self.log(f"下载完成: {display_name}")
            return local_path
            
        except Exception as e:
            self.log(f"下载失败 {display_name}: {e}")
            return None

class SimpleMediaPlayer:
    """简化的媒体播放器（使用系统默认播放器）"""
    
    def __init__(self, log_callback=None):
        self.log_callback = log_callback
        self.current_file = None
        
    def log(self, message):
        """记录日志"""
        print(f"[播放器] {message}")
        if self.log_callback:
            self.log_callback(f"[播放器] {message}")
            
    def play_file(self, file_path):
        """播放文件"""
        try:
            if not os.path.exists(file_path):
                self.log(f"文件不存在: {file_path}")
                return False
                
            self.current_file = file_path
            
            # 使用系统默认播放器
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):
                subprocess.run(['open', file_path])
            else:
                subprocess.run(['xdg-open', file_path])
                
            self.log(f"使用系统播放器打开: {os.path.basename(file_path)}")
            return True
                
        except Exception as e:
            self.log(f"播放失败: {e}")
            return False
            
    def stop(self):
        """停止播放（提示信息）"""
        self.log("请在播放器中手动停止播放")
        
    def pause(self):
        """暂停播放（提示信息）"""
        self.log("请在播放器中手动暂停播放")
        
    def resume(self):
        """恢复播放（提示信息）"""
        self.log("请在播放器中手动恢复播放")

class LangrunPlayerApp:
    """朗润播放器主应用程序"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("朗润播放器客户端 v1.0 (独立版)")
        self.root.geometry("1200x800")
        self.root.configure(bg='#F5F5F7')
        
        # 数据存储
        self.data = []
        self.columns = []
        self.media_data = {}
        
        # 初始化组件
        self.downloader = MediaDownloader(
            progress_callback=self.update_progress,
            log_callback=self.add_log
        )
        self.player = SimpleMediaPlayer(log_callback=self.add_log)
        
        # 创建界面
        self.create_ui()
        
    def create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="朗润播放器客户端 (独立版)", 
                               font=('Microsoft YaHei', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 左侧控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # 文件导入
        ttk.Button(control_frame, text="导入CSV文件", 
                  command=self.import_file).grid(row=0, column=0, sticky=tk.W+tk.E, pady=2)
        
        # 提示信息
        tip_label = ttk.Label(control_frame, text="提示：Excel文件请另存为CSV格式", 
                             font=('Microsoft YaHei', 8), foreground='gray')
        tip_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # 下载控制
        ttk.Button(control_frame, text="开始下载", 
                  command=self.start_download).grid(row=2, column=0, sticky=tk.W+tk.E, pady=2)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, 
                                          maximum=100)
        self.progress_bar.grid(row=3, column=0, sticky=tk.W+tk.E, pady=5)
        
        # 状态标签
        self.status_label = ttk.Label(control_frame, text="就绪")
        self.status_label.grid(row=4, column=0, sticky=tk.W, pady=2)
        
        # 搜索框
        search_frame = ttk.Frame(control_frame)
        search_frame.grid(row=5, column=0, sticky=tk.W+tk.E, pady=10)
        
        ttk.Label(search_frame, text="展演号码:").grid(row=0, column=0, sticky=tk.W)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.grid(row=1, column=0, sticky=tk.W+tk.E, pady=2)
        self.search_entry.bind('<Return>', self.search_and_play)
        
        ttk.Button(search_frame, text="搜索播放", 
                  command=self.search_and_play).grid(row=2, column=0, sticky=tk.W+tk.E, pady=2)
        
        search_frame.columnconfigure(0, weight=1)
        
        # 播放控制提示
        play_frame = ttk.LabelFrame(control_frame, text="播放控制", padding="5")
        play_frame.grid(row=6, column=0, sticky=tk.W+tk.E, pady=10)
        
        ttk.Label(play_frame, text="播放控制请在播放器中操作", 
                 font=('Microsoft YaHei', 8)).grid(row=0, column=0, columnspan=3)
        
        # 工具按钮
        tools_frame = ttk.LabelFrame(control_frame, text="工具", padding="5")
        tools_frame.grid(row=7, column=0, sticky=tk.W+tk.E, pady=10)
        
        ttk.Button(tools_frame, text="打开下载目录", 
                  command=self.open_download_dir).grid(row=0, column=0, sticky=tk.W+tk.E, pady=2)
        ttk.Button(tools_frame, text="清空下载历史", 
                  command=self.clear_download_history).grid(row=1, column=0, sticky=tk.W+tk.E, pady=2)
        
        # 中间数据列表
        list_frame = ttk.LabelFrame(main_frame, text="作品列表", padding="10")
        list_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview
        columns = ('展演号码', '姓名', '作品名称', '状态', '文件路径')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        # 设置列标题和宽度
        for col in columns:
            self.tree.heading(col, text=col)
            if col == '展演号码':
                self.tree.column(col, width=100)
            elif col == '姓名':
                self.tree.column(col, width=100)
            elif col == '作品名称':
                self.tree.column(col, width=200)
            elif col == '状态':
                self.tree.column(col, width=80)
            else:
                self.tree.column(col, width=300)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 双击播放
        self.tree.bind('<Double-1>', self.play_selected)
        
        # 右键菜单
        self.create_context_menu()
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        # 右侧日志面板
        log_frame = ttk.LabelFrame(main_frame, text="日志信息", padding="10")
        log_frame.grid(row=1, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=40, height=25)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 底部状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.bottom_status = ttk.Label(status_frame, text="朗润播放器客户端 (独立版) - 无需外部依赖")
        self.bottom_status.grid(row=0, column=0, sticky=tk.W)
        
    def create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="播放", command=self.play_selected)
        self.context_menu.add_command(label="打开文件位置", command=self.open_file_location)
        self.context_menu.add_command(label="重新下载", command=self.redownload_selected)
        
    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
            
    def open_file_location(self):
        """打开文件位置"""
        selection = self.tree.selection()
        if not selection:
            return
            
        item = self.tree.item(selection[0])
        values = item['values']
        
        if len(values) >= 5 and values[4]:
            file_path = values[4]
            if os.path.exists(file_path):
                # 打开文件所在目录
                if sys.platform.startswith('win'):
                    subprocess.run(['explorer', '/select,', file_path])
                elif sys.platform.startswith('darwin'):
                    subprocess.run(['open', '-R', file_path])
                else:
                    subprocess.run(['xdg-open', os.path.dirname(file_path)])
            else:
                messagebox.showinfo("提示", "文件不存在")
        else:
            messagebox.showinfo("提示", "文件未下载")
            
    def redownload_selected(self):
        """重新下载选中的文件"""
        selection = self.tree.selection()
        if not selection:
            return
            
        item = self.tree.item(selection[0])
        values = item['values']
        performance_number = values[0]
        
        if performance_number in self.media_data:
            data = self.media_data[performance_number]
            if data['url']:
                # 从下载历史中移除
                if data['url'] in self.downloader.downloaded_files:
                    del self.downloader.downloaded_files[data['url']]
                    self.downloader.save_download_history()
                
                # 重新下载
                threading.Thread(target=self._redownload_single, args=(data,), daemon=True).start()
                
    def _redownload_single(self, data):
        """重新下载单个文件"""
        local_path = self.downloader.download_file(
            data['url'], 
            data['work_name'], 
            data.get('performance_number', '')
        )
        
        if local_path:
            data['local_path'] = local_path
            self.root.after(0, self.update_file_list)
            
    def open_download_dir(self):
        """打开下载目录"""
        download_dir = os.path.abspath(self.downloader.download_dir)
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
            
        if sys.platform.startswith('win'):
            os.startfile(download_dir)
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', download_dir])
        else:
            subprocess.run(['xdg-open', download_dir])
            
    def clear_download_history(self):
        """清空下载历史"""
        result = messagebox.askyesno("确认", "确定要清空下载历史吗？这不会删除已下载的文件。")
        if result:
            self.downloader.downloaded_files = {}
            self.downloader.save_download_history()
            self.update_file_list()
            self.add_log("下载历史已清空")
        
    def add_log(self, message):
        """添加日志信息"""
        timestamp = time.strftime('%H:%M:%S')
        log_message = f"[{timestamp}] {message}\n"
        
        # 在主线程中更新GUI
        self.root.after(0, lambda: self._update_log_text(log_message))
        
    def _update_log_text(self, message):
        """更新日志文本（在主线程中执行）"""
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        
    def update_progress(self, value):
        """更新进度条"""
        self.root.after(0, lambda: self.progress_var.set(value))
        
    def update_status(self, message):
        """更新状态"""
        self.root.after(0, lambda: self.status_label.config(text=message))
        
    def import_file(self):
        """导入CSV文件"""
        file_path = filedialog.askopenfilename(
            title="选择CSV文件",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
            
        try:
            self.add_log(f"正在读取文件: {os.path.basename(file_path)}")
            
            # 使用简化的文件读取器
            self.data, self.columns = SimpleExcelReader.read_file(file_path)
            
            # 检查必要的列
            required_columns = ['展演号码', '姓名', '作品名称']
            missing_columns = [col for col in required_columns if col not in self.columns]
            
            if missing_columns:
                messagebox.showerror("错误", f"文件缺少必要的列: {', '.join(missing_columns)}")
                return
                
            # 查找媒体文件列
            media_columns = []
            for col in self.columns:
                if any(keyword in col.lower() for keyword in ['链接', 'url', 'link', '地址']):
                    media_columns.append(col)
                    
            if not media_columns:
                messagebox.showwarning("警告", "未找到媒体文件链接列，请确保文件中包含文件链接")
                
            # 更新列表
            self.update_file_list()
            
            self.add_log(f"成功读取 {len(self.data)} 条记录")
            self.update_status(f"已加载 {len(self.data)} 条记录")
            
        except Exception as e:
            error_msg = f"读取文件失败: {e}"
            self.add_log(error_msg)
            messagebox.showerror("错误", error_msg)
            
    def update_file_list(self):
        """更新文件列表"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not self.data:
            return
            
        # 添加数据到列表
        for row in self.data:
            performance_number = str(row.get('展演号码', ''))
            name = str(row.get('姓名', ''))
            work_name = str(row.get('作品名称', ''))
            
            # 查找媒体链接
            media_url = ""
            for col in self.columns:
                if any(keyword in col.lower() for keyword in ['链接', 'url', 'link', '地址']):
                    url_value = str(row.get(col, ''))
                    if url_value and url_value != 'nan' and url_value.startswith('http'):
                        media_url = url_value
                        break
                        
            # 检查下载状态
            status = "未下载"
            file_path = ""
            if media_url in self.downloader.downloaded_files:
                local_path = self.downloader.downloaded_files[media_url]
                if os.path.exists(local_path):
                    status = "已下载"
                    file_path = local_path
                    
            # 存储媒体数据
            self.media_data[performance_number] = {
                'name': name,
                'work_name': work_name,
                'url': media_url,
                'local_path': file_path,
                'performance_number': performance_number
            }
            
            # 添加到树视图
            self.tree.insert('', tk.END, values=(
                performance_number, name, work_name, status, file_path
            ))
            
    def start_download(self):
        """开始下载"""
        if not self.data:
            messagebox.showwarning("警告", "请先导入CSV文件")
            return
            
        # 在新线程中执行下载
        threading.Thread(target=self._download_thread, daemon=True).start()
        
    def _download_thread(self):
        """下载线程"""
        try:
            self.update_status("正在下载...")
            download_count = 0
            
            for performance_number, data in self.media_data.items():
                if not data['url']:
                    continue
                    
                # 检查是否已下载
                if data['local_path'] and os.path.exists(data['local_path']):
                    self.add_log(f"跳过已下载文件: {data['work_name']}")
                    continue
                    
                # 下载文件
                local_path = self.downloader.download_file(
                    data['url'], 
                    data['work_name'], 
                    performance_number
                )
                
                if local_path:
                    data['local_path'] = local_path
                    download_count += 1
                    
                    # 更新列表显示
                    self.root.after(0, self.update_file_list)
                    
            self.add_log(f"下载完成，共下载 {download_count} 个文件")
            self.update_status(f"下载完成 ({download_count} 个文件)")
            
        except Exception as e:
            error_msg = f"下载过程出错: {e}"
            self.add_log(error_msg)
            self.update_status("下载失败")
            
    def search_and_play(self, event=None):
        """搜索并播放"""
        search_number = self.search_var.get().strip()
        if not search_number:
            messagebox.showwarning("警告", "请输入展演号码")
            return
            
        if search_number in self.media_data:
            data = self.media_data[search_number]
            if data['local_path'] and os.path.exists(data['local_path']):
                self.player.play_file(data['local_path'])
                self.add_log(f"播放作品: {data['work_name']} ({data['name']})")
            else:
                messagebox.showinfo("提示", f"文件未下载: {data['work_name']}")
        else:
            messagebox.showinfo("提示", f"未找到展演号码: {search_number}")
            
    def play_selected(self, event=None):
        """播放选中的文件"""
        selection = self.tree.selection()
        if not selection:
            return
            
        item = self.tree.item(selection[0])
        values = item['values']
        
        if len(values) >= 5 and values[4]:  # 文件路径
            file_path = values[4]
            if os.path.exists(file_path):
                self.player.play_file(file_path)
                self.add_log(f"播放: {values[2]} ({values[1]})")
            else:
                messagebox.showinfo("提示", "文件不存在，请先下载")
        else:
            messagebox.showinfo("提示", "文件未下载")
            
    def run(self):
        """运行应用程序"""
        self.add_log("朗润播放器客户端 (独立版) 启动成功")
        self.add_log("提示: 独立版无需外部依赖，使用系统默认播放器")
        self.add_log("Excel文件请先另存为CSV格式后导入")
        self.root.mainloop()

if __name__ == "__main__":
    app = LangrunPlayerApp()
    app.run() 