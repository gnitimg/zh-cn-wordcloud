import jieba
from collections import Counter
import re
import os
import threading
import time
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, font
from matplotlib.font_manager import FontProperties
import matplotlib
matplotlib.use('Agg')  # 避免与Tkinter冲突

# 设置中文字体 - 更强大的字体查找逻辑
def find_chinese_font():
    """查找系统中可用的中文字体"""
    # 常见中文字体路径
    font_paths = [
        'simhei.ttf',  # 黑体
        'simsun.ttc',  # 宋体
        'msyh.ttc',    # 微软雅黑
        'simkai.ttf',  # 楷体
        'C:/Windows/Fonts/simhei.ttf',
        'C:/Windows/Fonts/simsun.ttc',
        'C:/Windows/Fonts/msyh.ttc',
        'C:/Windows/Fonts/simkai.ttf',
        'C:/Windows/Fonts/STKAITI.TTF',  # 华文楷体
        '/System/Library/Fonts/PingFang.ttc',  # macOS
        '/System/Library/Fonts/STHeiti Medium.ttc',  # macOS
        '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',  # Linux
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'  # Linux文泉驿
    ]
    
    # 尝试查找存在的字体
    for path in font_paths:
        if os.path.exists(path):
            return path
    
    # 尝试通过matplotlib查找
    try:
        from matplotlib.font_manager import findfont, FontProperties
        chinese_font_prop = FontProperties(family=['sans-serif'], style='normal', size=12)
        return findfont(chinese_font_prop)
    except:
        return None

# 获取中文字体路径
FONT_PATH = find_chinese_font()

# 分词函数
def segment_text(text, use_hmm=True):
    """对文本进行分词处理"""
    words = jieba.lcut(text, HMM=use_hmm)
    
    # 过滤掉空白字符和标点符号
    filtered_words = []
    for word in words:
        word = word.strip()
        if word and not re.match(r'^\W+$', word):  # 排除纯标点符号
            filtered_words.append(word)
    
    return filtered_words

# 词频统计函数
def count_word_frequency(words):
    """统计词频并排序"""
    word_counts = Counter(words)
    sorted_words = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)
    return sorted_words

# 生成词云函数 - 确保使用正确字体
def generate_word_cloud(word_freq, output_path):
    """根据词频生成词云"""
    # 将词频转换为字典
    freq_dict = {word: freq for word, freq in word_freq}
    
    # 创建词云对象
    wc = WordCloud(
        font_path=FONT_PATH,  # 关键：使用找到的字体
        width=1200,
        height=800,
        background_color='white',
        max_words=200,
        colormap='viridis',
        prefer_horizontal=0.9  # 增加水平文本比例
    )
    
    # 生成词云
    wc.generate_from_frequencies(freq_dict)
    
    # 保存词云
    wc.to_file(output_path)
    return output_path

class WordSegmenterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("中文分词分析工具")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        # 初始化变量
        self.word_freq = []
        self.progress_value = tk.IntVar(value=0)
        self.analyzing = False
        
        # 创建界面
        self.create_widgets()
        
        # 显示字体状态
        self.update_font_status()
        
    def update_font_status(self):
        """更新字体状态显示"""
        if FONT_PATH:
            font_name = os.path.basename(FONT_PATH)
            self.status_label.config(text=f"就绪 | 使用字体: {font_name}")
        else:
            self.status_label.config(text="就绪 | 警告: 未找到中文字体!")
    
    def create_widgets(self):
        # 创建主框架
        main_frame = tk.Frame(self.root, bg='#f0f0f0', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_font = font.Font(family="Microsoft YaHei", size=16, weight="bold")
        title_label = tk.Label(main_frame, 
                              text="中文分词与词频分析工具", 
                              font=title_font,
                              bg='#f0f0f0',
                              fg='#2c3e50')
        title_label.pack(pady=(0, 10))
        
        # 副标题
        subtitle_font = font.Font(family="Microsoft YaHei", size=10)
        subtitle_label = tk.Label(main_frame, 
                                 text="支持大文本分词、词频统计、词云生成", 
                                 font=subtitle_font,
                                 bg='#f0f0f0',
                                 fg='#7f8c8d')
        subtitle_label.pack(pady=(0, 20))
        
        # 输入区域
        input_frame = tk.LabelFrame(main_frame, 
                                   text="输入文本", 
                                   font=("Microsoft YaHei", 10),
                                   bg='#f0f0f0',
                                   padx=10,
                                   pady=10)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.input_text = scrolledtext.ScrolledText(
            input_frame, 
            height=8, 
            font=('Microsoft YaHei', 10),
            wrap=tk.WORD,
            bg='white',
            fg='#333333'
        )
        self.input_text.pack(fill=tk.BOTH, expand=True)
        
        # 分析按钮区域
        btn_frame = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 开始分析按钮 - 使用显式颜色设置
        self.analyze_btn = tk.Button(
            btn_frame, 
            text="开始分析", 
            command=self.start_analysis,
            font=('Microsoft YaHei', 10, 'bold'),
            bg='#3498db',
            fg='white',
            padx=10,
            pady=5,
            relief=tk.FLAT
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        style = ttk.Style()
        style.configure("TProgressbar", thickness=20, background='#3498db')
        
        self.progress_bar = ttk.Progressbar(
            btn_frame, 
            variable=self.progress_value,
            style="TProgressbar",
            length=300
        )
        self.progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # 状态标签 - 使用深色文字
        self.status_label = tk.Label(
            btn_frame, 
            text="就绪", 
            font=('Microsoft YaHei', 9),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        self.status_label.pack(side=tk.LEFT, padx=10)
        
        # 结果区域
        result_frame = tk.LabelFrame(
            main_frame, 
            text="分词结果", 
            font=("Microsoft YaHei", 10, "bold"),
            bg='#f0f0f0',
            padx=10,
            pady=10
        )
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建表格
        self.result_tree = ttk.Treeview(
            result_frame, 
            columns=('word', 'freq'), 
            show='headings',
            selectmode='browse'
        )
        
        # 设置列
        self.result_tree.heading('word', text='词语', anchor=tk.W)
        self.result_tree.heading('freq', text='频次', anchor=tk.W)
        self.result_tree.column('word', width=300, anchor=tk.W)
        self.result_tree.column('freq', width=100, anchor=tk.W)
        
        # 设置表格样式
        style.configure("Treeview", 
                       font=('Microsoft YaHei', 10), 
                       rowheight=25,
                       background='white',
                       fieldbackground='white',
                       foreground='#333333')
        style.configure("Treeview.Heading", 
                       font=('Microsoft YaHei', 10, 'bold'),
                       background='#3498db',
                       foreground='white')
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(
            result_frame, 
            orient=tk.VERTICAL, 
            command=self.result_tree.yview
        )
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮
        btn_frame2 = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame2.pack(fill=tk.X, padx=5, pady=10)
        
        # 保存结果按钮 - 使用显式颜色设置
        save_btn = tk.Button(
            btn_frame2, 
            text="保存结果", 
            command=self.save_results,
            font=('Microsoft YaHei', 10),
            bg='#2ecc71',
            fg='white',
            padx=10,
            pady=5,
            relief=tk.GROOVE
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # 生成词云按钮 - 使用显式颜色设置
        wc_btn = tk.Button(
            btn_frame2, 
            text="生成词云", 
            command=self.generate_wordcloud,
            font=('Microsoft YaHei', 10),
            bg='#9b59b6',
            fg='white',
            padx=10,
            pady=5,
            relief=tk.GROOVE
        )
        wc_btn.pack(side=tk.LEFT, padx=5)
        
        # 退出按钮 - 使用显式颜色设置
        exit_btn = tk.Button(
            btn_frame2, 
            text="退出", 
            command=self.root.destroy,
            font=('Microsoft YaHei', 10),
            bg='#e74c3c',
            fg='white',
            padx=10,
            pady=5,
            relief=tk.GROOVE
        )
        exit_btn.pack(side=tk.RIGHT, padx=5)
    
    def start_analysis(self):
        if self.analyzing:
            return
            
        text = self.input_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("输入错误", "请输入要分析的文本！")
            return
            
        if len(text) < 100:
            if not messagebox.askyesno("警告", "文本较短（少于100字），确定要继续吗？"):
                return
        
        # 更新状态
        self.analyzing = True
        self.analyze_btn.config(state=tk.DISABLED, bg='#95a5a6')
        self.status_label.config(text="分析中...")
        self.progress_value.set(0)
        
        # 清空表格
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # 在后台线程中执行分析
        threading.Thread(target=self.analyze_text, args=(text,), daemon=True).start()
    
    def analyze_text(self, text):
        try:
            # 分词
            self.status_label.config(text="分词中...")
            words = segment_text(text)
            
            # 更新进度
            self.progress_value.set(30)
            self.status_label.config(text="分词完成，正在统计词频...")
            
            # 词频统计
            self.word_freq = count_word_frequency(words)
            
            # 更新进度
            self.progress_value.set(80)
            self.status_label.config(text="正在显示结果...")
            
            # 更新表格
            for word, freq in self.word_freq[:500]:  # 最多显示500个
                self.result_tree.insert('', tk.END, values=(word, freq))
            
            # 完成
            self.progress_value.set(100)
            self.status_label.config(text=f"分析完成! 共发现 {len(self.word_freq)} 个词语")
            
        except Exception as e:
            self.status_label.config(text=f"分析错误: {str(e)}")
            messagebox.showerror("分析错误", f"分词过程中发生错误:\n{str(e)}")
        
        finally:
            self.analyzing = False
            self.analyze_btn.config(state=tk.NORMAL, bg='#3498db')
            self.update_font_status()
    
    def save_results(self):
        if not self.word_freq:
            messagebox.showwarning("错误", "请先进行分析！")
            return
            
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
            title="保存结果"
        )
        
        if not save_path:
            return
            
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write('词语\t频次\n')
                f.write('----------------\n')
                for word, freq in self.word_freq:
                    f.write(f'{word}\t{freq}\n')
            messagebox.showinfo("保存成功", f"结果已保存到:\n{save_path}")
        except Exception as e:
            messagebox.showerror("保存失败", f"保存文件时出错:\n{str(e)}")
    
    def generate_wordcloud(self):
        if not self.word_freq:
            messagebox.showwarning("错误", "请先进行分析！")
            return
            
        if not FONT_PATH:
            messagebox.showwarning("字体缺失", "未找到中文字体，无法生成词云！")
            return
            
        save_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Files", "*.png")],
            title="保存词云"
        )
        
        if not save_path:
            return
            
        try:
            self.status_label.config(text="生成词云中...")
            self.analyze_btn.config(state=tk.DISABLED, bg='#95a5a6')
            
            # 在后台生成词云
            threading.Thread(target=self._generate_wordcloud, args=(save_path,), daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("词云生成失败", f"启动词云生成时出错:\n{str(e)}")
            self.status_label.config(text="就绪")
            self.analyze_btn.config(state=tk.NORMAL, bg='#3498db')
    
    def _generate_wordcloud(self, save_path):
        try:
            generate_word_cloud(self.word_freq, save_path)
            self.status_label.config(text="词云生成完成!")
            messagebox.showinfo("词云生成成功", f"词云已保存到:\n{save_path}")
        except Exception as e:
            self.status_label.config(text="词云生成失败")
            messagebox.showerror("词云生成失败", f"生成词云时出错:\n{str(e)}")
        finally:
            self.analyze_btn.config(state=tk.NORMAL, bg='#3498db')
            self.update_font_status()

if __name__ == "__main__":
    # 初始化jieba分词器
    jieba.initialize()
    
    # 创建主窗口
    root = tk.Tk()
    
    # 设置窗口图标（可选）
    try:
        root.iconbitmap(default='')  # 清除默认图标
    except:
        pass
    
    # 启动应用
    app = WordSegmenterApp(root)
    
    # 启动主循环
    root.mainloop()
