import sys
import os

# --- 核心修复：防止 .pyw 无控制台模式下第三方库打印 Warning 导致程序静默崩溃 ---
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')
# --------------------------------------------------------------------

import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import difflib
import io
import re
import unicodedata
import urllib.parse
import posixpath
import subprocess

# --- 依赖项检查 ---
try:
    from PIL import Image, ImageTk
except ImportError:
    root_err = tk.Tk()
    root_err.withdraw()
    messagebox.showerror("缺少依赖", "未能检测到 Pillow 库。\n\n请在终端运行以下命令安装：\npip install Pillow")
    root_err.destroy()
    exit()

try:
    from send2trash import send2trash
except ImportError:
    root_err = tk.Tk()
    root_err.withdraw()
    messagebox.showerror("缺少依赖", "未能检测到 send2trash 库。\n\n请在终端运行以下命令安装：\npip install send2trash")
    root_err.destroy()
    exit()
# ------------------

NAMESPACES = {
    'opf': 'http://www.idpf.org/2007/opf',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'container': 'urn:oasis:names:tc:opendocument:xmlns:container'
}

PREVIEW_HEIGHT = 280

def get_epub_info(filepath):
    title = ""
    size_mb = os.path.getsize(filepath) / (1024 * 1024)
    try:
        with zipfile.ZipFile(filepath, 'r') as archive:
            container_xml = archive.read('META-INF/container.xml')
            container_tree = ET.fromstring(container_xml)
            rootfile_path = container_tree.find('.//container:rootfile', NAMESPACES).attrib['full-path']

            opf_content = archive.read(rootfile_path)
            opf_tree = ET.fromstring(opf_content)
            title_elem = opf_tree.find('.//dc:title', NAMESPACES)
            
            if title_elem is not None and title_elem.text:
                title = title_elem.text.strip()
    except Exception:
        pass 
        
    if not title:
        title = os.path.splitext(os.path.basename(filepath))[0]
        
    return {"path": filepath, "title": title, "size": size_mb, "filename": os.path.basename(filepath)}

def clean_filename(name):
    """超级净化器：剥离所有可能的污染项，还原最干净的书名主干"""
    name = os.path.splitext(name)[0]
    name = unicodedata.normalize('NFKC', str(name))
    name = re.sub(r'1080[pP]?|720[pP]?|2160[pP]?|4[kK]', '', name)
    name = re.sub(r'[ _\-\(\[]v\d\b', '', name, flags=re.IGNORECASE)
    
    # 清理头部前缀
    name = re.sub(r'^\[.*?\]\s*', '', name)
    name = re.sub(r'^【.*?】\s*', '', name)
    
    # --- 核心修复：防止正则贪婪匹配吃掉无辜的括号 ---
    # 使用 [^\)\]】）>〉」』]* 替代 .*?，确保每次只检查单个括号内部的内容
    name = re.sub(r'[\(\[【（<〈「『][^\)\]】）>〉」』]*(掃图|扫图|录入|制作|修图|译|特装|特典|電子|电子|限定|SS|ss|描き下ろし|書き下ろし|まんがで読破|角川|文庫|ブックス|ノベル|タイガ|電撃|ガガガ|ファミ通)[^\)\]】）>〉」』]*[\)\]】）>〉」』]', '', name, flags=re.IGNORECASE)
    
    # 清理末尾作者名，防止误杀导致书名变为空字符串
    name_no_author = re.sub(r'\s*[-－/／~～]\s*[^ -~]+(\s*&\s*[^ -~]+)?$', '', name)
    if name_no_author.strip(): 
        name = name_no_author
        
    return name.strip()

def get_volume_fingerprint(text):
    """阵列化特征提取器：将文件中的全部卷号特征提取为一个元组"""
    if not text: return None
    
    text = unicodedata.normalize('NFKC'， str(text)).lower()
    text = re.sub(r'1080p?|720p?|2160p?|4k|h264|h265|x264|x265|mp4|mp3|aac|flac|201\d|202\d|203\d', ' ', text, flags=re.I)
    
    # --- 核心修复：保护类似“道中04”这样的词汇不被误删 ---
    # (?:^|[^\u4e00-\u9fa5]) 确保前面不能是汉字。所以 "高中2" 会被删，但 "道中04" 会被安全保留！
    text = re.sub(r'(?:^|[^\u4e00-\u9fa5])(高|中|小|大)\d+', ' ', text)
    text = re.sub(r'\d+(周目|度目|年|歳|人|回|軍|話|ヶ?月|ヶ?国|日|時間|次元|次|姉|妹|男|女|万|千|百)', ' ', text)
    text = re.sub(r'第\d+(の|回|次|章|世代)', ' ', text)
    
    cn_num_map = {
        '十一':'11', '十二':'12', '十三':'13', '十四':'14', '十五':'15',
        '十六':'16', '十七':'17', '十八':'18', '十九':'19', '二十':'20',
        '一':'1', '二':'2', '三':'3', '四':'4', '五':'5', 
        '六':'6', '七':'7', '八':'8', '九':'9', '十':'10',
        '零':'0', '壱':'1', '弐':'2', '参':'3', '肆':'4', '伍':'5',
        '陸':'6', '漆':'7', '捌':'8', '玖':'9'
    }
    for k, v in cn_num_map.items():
        # --- 核心修复：追加 ~ 符号，应对 NFKC 标准化后的波浪号 ---
        text = text.replace(f"第{k}", f"第{v}").replace(f"卷{k}", f"卷{v}").replace(f"巻{k}", f"巻{v}").replace(f"～{k}", f"～{v}").replace(f"~{k}", f"~{v}")
        text = re.sub(rf'[\(\[【（<〈「『≪《]{k}[\)\]】）>〉」』≫》]', f'({v})', text)

    markers = []

    # 1. 提取罗马数字
    roman_map = {
        'xxx':30, 'xxix':29, 'xxviii':28, 'xxvii':27, 'xxvi':26, 'xxv':25, 'xxiv':24, 'xxiii':23, 'xxii':22, 'xxi':21,
        'xx':20, 'xix':19, 'xviii':18, 'xvii':17, 'xvi':16, 'xv':15, 'xiv':14, 'xiii':13,
        'xii':12, 'xi':11, 'x':10, 'ix':9, 'viii':8, 'vii':7, 'vi':6, 'iv':4, 'iii':3, 'ii':2, 'i':1, 'v':5
    }
    roman_keys = '|'.join(roman_map.keys())
    roman_pattern = r'(?<![a-z])(' + roman_keys + r')(?![a-z])'
    for m in re.finditer(roman_pattern, text):
        markers.append(f"R{roman_map[m.group(1)]}")

    # 2. 提取特殊篇章标记
    clean_text = text.replace(' ', '').replace('　', '')
    special_map = {
        'dlc2': 'W_DLC2', 'dlc': 'W_DLC', 'ss': 'W_SS', 'fd2': 'W_FD2', 'fd': 'W_FD', 
        'ex': 'W_EX', 'discex': 'W_DISCEX', 'sp': 'W_SP', '外伝': 'W_外伝', 
        '短編集': 'W_短編集', '番外編': 'W_番外編', '特別篇': 'W_特別篇', 'eachstories': 'W_EACH',
        'φ': 'W_Φ', 'σ': 'W_Σ', 'α': 'W_α', 'β': 'W_β', 'γ': 'W_γ', 'δ': 'W_Δ', 'Δ': 'W_Δ'
    }
    for k, v in special_map.items():
        if k in clean_text: markers.append(v)

    # 3. 提取文字卷号
    for m in re.findall(r'(上巻|中巻|下巻|前編|後編)', text):
        markers.append(f"W_{m}")
    for m in re.findall(r'(?:^|[^\w\u4e00-\u9fa5ぁ-んァ-ン])(上|中|下|前|後|春|夏|秋|冬)(?:[^\w\u4e00-\u9fa5ぁ-んァ-ン]|$)', text):
        markers.append(f"W_{m}")
    for m in re.finditer(r'[-_ 　](a|b|c)(?![a-z])', text):
        markers.append(f"PART_{m.group(1).upper()}")

    # 4. 提取阿拉伯数字卷号
    patterns = [
        r'(?:第|vol\.?|v|ep\.?|sp\.?|ex\.?|part|卷|巻|幕|章|層|節|話|番外|特别篇|外伝|特典)\s*(\d+(?:\.\d+)?)',
        r'(\d+(?:\.\d+)?)\s*[卷巻册話幕章層節]',
        r'[\(\[【（<〈「『≪《]\s*(\d+(?:\.\d+)?)\s*[\)\]】）>〉」』≫》]'
    ]
    for p in patterns:
        for m in re.findall(p, text):
            try:
                if float(m) < 500: markers.append(f"N{float(m)}")
            except: pass
            
    for m in re.finditer(r'(?<!\d)(\d+(?:\.\d+)?)(?!\d)', text):
        try:
            val = float(m.group(1))
            if val < 500:
                markers.append(f"N{val}")
        except: pass
        
    return tuple(sorted(list(set(markers)))) if markers else None

def is_definitely_not_vol_1(v_tuple):
    """防第一卷隐身补丁"""
    if not v_tuple: return False
    for v in v_tuple:
        if v.startswith('N') or v.startswith('R'):
            if abs(float(v[1:]) - 1.0) > 0.1: return True
        if v.startswith('W_') or v.startswith('PART_'):
            return True
    return False

def get_core_title(name):
    """书名号内核抽取器"""
    m = re.search(r'[『「](.*?)[』」]', name)
    if m: return m.group(1)
    return name

def analyze_diff(name1, name2):
    """差异探伤核心：精准分析字符串的不同之处"""
    if name1 == name2: return 'MATCH'
    s = difflib.SequenceMatcher(None, name1, name2)
    diff1, diff2 = "", ""
    for tag, i1, i2, j1, j2 in s.get_opcodes():
        if tag in ('replace', 'insert', 'delete'):
            diff1 += name1[i1:i2]
            diff2 += name2[j1:j2]
            
    has_text1 = bool(re.search(r'[^\W\d_]', diff1))
    has_text2 = bool(re.search(r'[^\W\d_]', diff2))
    
    if has_text1 and has_text2:
        return 'NO_MATCH' 
        
    if has_text1 or has_text2:
        extra_text = diff1 if has_text1 else diff2
        
        # --- 核心修复：把 壱弐参肆 等汉字大写数字加入绝对阻断名单 ---
        if re.search(r'(上|中|下|前|後|特典|外伝|外传|番外|特装|限定|ss|ex|disc|dlc|episodes?|part|vol|巻|卷|第|[一二三四五六七八九十壱弐参肆伍陸漆捌玖拾]|I|V|X|S|A|B|C|Φ|Σ|Δ|α|β|γ)', extra_text, re.I):
            return 'NO_MATCH'
            
        fluff_stripped = re.sub(r'(電子書籍|電子版|電子|あとがき|イラスト|挿絵|ふりがな|文庫|完結|小説|版|付き?|書き下ろし|描き下ろし)', '', extra_text)
        jp_chars = re.findall(r'[ぁ-んァ-ン一-龥]', fluff_stripped)
        if len(jp_chars) >= 2:
            return 'NO_MATCH'
            
        return 'MANUAL' 
        
    nums1 = set(re.findall(r'\d+', diff1))
    nums2 = set(re.findall(r'\d+', diff2))
    if nums1 != nums2:
        return 'MANUAL' 
        
    return 'MATCH'
        
    nums1 = set(re.findall(r'\d+', diff1))
    nums2 = set(re.findall(r'\d+', diff2))
    if nums1 != nums2:
        return 'MANUAL' 
        
    return 'MATCH'

def extract_epub_cover(filepath):
    try:
        with zipfile.ZipFile(filepath, 'r') as archive:
            container_xml = archive.read('META-INF/container.xml')
            container_tree = ET.fromstring(container_xml)
            rootfile_path = container_tree.find('.//container:rootfile', NAMESPACES).attrib['full-path']
            
            opf_dir = os.path.dirname(rootfile_path)
            opf_content = archive.read(rootfile_path)
            opf_tree = ET.fromstring(opf_content)

            cover_href = 无
            manifest = opf_tree.find('.//opf:manifest', NAMESPACES)
            if manifest is not None:
                for item in manifest.findall('.//opf:item', NAMESPACES):
                    if 'properties' in item.attrib and 'cover-image' in item.attrib['properties']:
                        cover_href = item.attrib['href']
                        break
            
            if not cover_href:
                metadata = opf_tree.find('.//opf:metadata', NAMESPACES)
                if metadata is not None:
                    cover_meta = metadata.find('.//opf:meta[@name="cover"]', NAMESPACES)
                    if cover_meta is not None:
                        cover_id = cover_meta.attrib['content']
                        item = opf_tree.find(f'.//opf:item[@id="{cover_id}"]', NAMESPACES)
                        if item is not None:
                            cover_href = item.attrib['href']

            if cover_href:
                cover_href = urllib.parse.unquote(cover_href)
                img_path = posixpath.normpath(posixpath.join(opf_dir, cover_href))
                img_path = img_path.replace('\\', '/')
                try:
                    img_data = archive.read(img_path)
                    return io.BytesIO(img_data)
                except KeyError:
                    pass 
                
    except Exception:
        pass
    return None

class EpubDupeFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EPUB 查重助手 v3.4")
        self.root.geometry("1200x850") 
        
        self.folder1_path = tk.StringVar()
        self.folder2_path = tk.StringVar()
        self.duplicates = [] 
        
        self.current_b1 = None
        self.current_b2 = None
        self._selection_after_id = None 

        self.setup_ui()

    def setup_ui(self):
        frame_top = ttk.Frame(self.root, padding=10)
        frame_top.pack(side=tk.TOP, fill=tk.X)

        self.frame_bottom = ttk.Frame(self.root, padding=5)
        self.frame_bottom.pack(side=tk.BOTTOM, fill=tk.X)

        self.frame_covers = ttk.LabelFrame(self.root, text="封面AB对比与详情 (单选时显示)", padding=5)
        self.frame_covers.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 5))

        self.frame_action = ttk.Frame(self.root, padding=(10, 5, 10, 5))
        self.frame_action.pack(side=tk.BOTTOM, fill=tk.X)

        frame_mid = ttk.Frame(self.root, padding=(10, 0, 10, 5))
        frame_mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        ttk.Label(frame_top, text="主文件夹 (必填):").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame_top, textvariable=self.folder1_path, width=85).grid(row=0, column=1, padx=5)
        ttk.Button(frame_top, text="浏览", command=lambda: self.browse_folder(self.folder1_path)).grid(row=0, column=2)

        ttk.Label(frame_top, text="次文件夹 (留空则在主文件夹内查重):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame_top, textvariable=self.folder2_path, width=85).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_top, text="浏览", command=lambda: self.browse_folder(self.folder2_path)).grid(row=1, column=2, pady=5)

        self.btn_scan = ttk.Button(frame_top, text="开始扫描与比对", command=self.scan_and_compare, style='Accent.TButton')
        self.btn_scan.grid(row=2, column=1, pady=10)

        self.progress_var = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(frame_top, variable=self.progress_var, maximum=100)
        self.progressbar.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=(0, 5))
        self.progressbar.grid_remove() 

        columns = ("title", "fileA", "sizeA", "fileB", "sizeB", "action")
        self.tree = ttk.Treeview(frame_mid, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("title", text="判定书名")
        self.tree.heading("fileA", text="文件A (文件名)")
        self.tree.heading("sizeA", text="大小 (MB)")
        self.tree.heading("fileB", text="文件B (文件名)")
        self.tree.heading("sizeB", text="大小 (MB)")
        self.tree.heading("action", text="计划操作") 
        
        self.tree.column("title", width=280)
        self.tree.column("fileA", width=280)
        self.tree.column("sizeA", width=50, anchor=tk.CENTER)
        self.tree.column("fileB", width=280)
        self.tree.column("sizeB", width=50, anchor=tk.CENTER)
        self.tree.column("action", width=80, anchor=tk.CENTER)
        
        vsb = ttk.Scrollbar(frame_mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Control-A>', self.select_all)

        btn_frame = ttk.Frame(self.frame_action)
        btn_frame.pack(fill=tk.X)
        
        lbl_mark = ttk.Label(btn_frame, text="修改所选行的计划:")
        lbl_mark.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_mark_a = ttk.Button(btn_frame, text="🔴 标记删 文件A", state=tk.DISABLED, command=lambda: self.change_plan('A'))
        self.btn_mark_a.pack(side=tk.LEFT, padx=5)
        
        self.btn_mark_b = ttk.Button(btn_frame, text="🔵 标记删 文件B", state=tk.DISABLED, command=lambda: self.change_plan('B'))
        self.btn_mark_b.pack(side=tk.LEFT, padx=5)
        
        self.btn_mark_none = ttk.Button(btn_frame, text="⚪ 取消标记", state=tk.DISABLED, command=lambda: self.change_plan('NONE'))
        self.btn_mark_none.pack(side=tk.LEFT, padx=5)
        
        self.btn_execute = ttk.Button(btn_frame, text="🚀 执行列表中的删除计划", state=tk.DISABLED, style='Accent.TButton', command=self.execute_all_plans)
        self.btn_execute.pack(side=tk.RIGHT, padx=5)
        
        self.cover_container = ttk.Frame(self.frame_covers, height=290) 
        self.cover_container.pack(fill=tk.X, expand=True)
        self.cover_container.grid_propagate(False)

        self.cover_container.columnconfigure(0, weight=1, uniform="half")
        self.cover_container.columnconfigure(1, weight=0)
        self.cover_container.columnconfigure(2, weight=1, uniform="half")
        self.cover_container.rowconfigure(0, weight=1)

        frame_a_inner = ttk.Frame(self.cover_container)
        frame_a_inner.grid(row=0, column=0, sticky="nsew", padx=5)
        
        self.lbl_cover_a = ttk.Label(frame_a_inner, text="等待加载...", anchor=tk.CENTER)
        self.lbl_cover_a.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        frame_right_a = ttk.Frame(frame_a_inner)
        frame_right_a.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        self.lbl_info_a = ttk.Label(frame_right_a, text="", justify=tk.LEFT, font=("Microsoft YaHei UI", 9))
        self.lbl_info_a.pack(side=tk.TOP, pady=(20, 10))
        self.btn_open_a = ttk.Button(frame_right_a, text="打开文件夹", state=tk.DISABLED, command=lambda: self.open_folder(1))
        self.btn_open_a.pack(side=tk.TOP)

        ttk.Separator(self.cover_container, orient=tk.VERTICAL).grid(row=0, column=1, sticky="ns")

        frame_b_inner = ttk.Frame(self.cover_container)
        frame_b_inner.grid(row=0, column=2, sticky="nsew", padx=5)
        
        self.lbl_cover_b = ttk.Label(frame_b_inner, text="等待加载...", anchor=tk.CENTER)
        self.lbl_cover_b.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        frame_right_b = ttk.Frame(frame_b_inner)
        frame_right_b.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        self.lbl_info_b = ttk.Label(frame_right_b, text="", justify=tk.LEFT, font=("Microsoft YaHei UI", 9))
        self.lbl_info_b.pack(side=tk.TOP, pady=(20, 10))
        self.btn_open_b = ttk.Button(frame_right_b, text="打开文件夹", state=tk.DISABLED, command=lambda: self.open_folder(2))
        self.btn_open_b.pack(side=tk.TOP)

        self.lbl_selected = ttk.Label(self.frame_bottom, text="就绪。请选择文件夹并开始扫描。", font=("", 10, "bold"))
        self.lbl_selected.pack(anchor=tk.W, padx=10)

        style = ttk.Style()
        style.configure('Accent.TButton', font=('Helvetica', 10, 'bold'))

    def open_folder(self, target):
        filepath = self.current_b1['path'] if target == 1 else self.current_b2['path']
        if filepath and os.path.exists(filepath):
            try:
                subprocess.run(['explorer', '/select,', os.path.normpath(filepath)])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件夹: {e}")

    def browse_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)

    def select_all(self, event):
        self.tree.selection_set(self.tree.get_children())
        self.on_select(None)
        return 'break'

    def get_all_epubs(self, folder):
        epubs = []
        for root, _, files in os.walk(folder):
            for file in files:
                if file.lower().endswith('.epub'):
                    epubs.append(os.path.join(root, file))
        return epubs

    def scan_and_compare(self):
        f1 = self.folder1_path.get()
        f2 = self.folder2_path.get()
        if not f1:
            messagebox.showwarning("警告", "请至少选择大文件夹 (文件夹 1)")
            return

        is_single_mode = (not f2) or (f1 == f2)

        self.btn_scan.config(state=tk.DISABLED, text="正在处理...")
        self.progressbar.grid() 
        self.progress_var.set(0)
        self.root.update()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.duplicates.clear()
        self.clear_covers()

        epubs1 = self.get_all_epubs(f1)
        if not epubs1:
            messagebox.showinfo("提示", "主文件夹中没有找到 EPUB 文件。")
            self.reset_scan_ui()
            return

        info1 = []
        total_a = len(epubs1)
        for i, p in enumerate(epubs1):
            info1.append(get_epub_info(p))
            if i % 5 == 0 or i == total_a - 1:
                self.lbl_selected.config(text=f"解析文件夹: {i+1}/{total_a}")
                self.progress_var.set((i + 1) / total_a * (100 if is_single_mode else 40))
                self.root.update()

        info2 = []
        if not is_single_mode:
            epubs2 = self.get_all_epubs(f2)
            if not epubs2:
                messagebox.showinfo("提示", "次文件夹中没有找到 EPUB 文件。")
                self.reset_scan_ui()
                return
            total_b = len(epubs2)
            for i, p in enumerate(epubs2):
                info2.append(get_epub_info(p))
                if i % 5 == 0 or i == total_b - 1:
                    self.lbl_selected.config(text=f"解析次文件夹: {i+1}/{total_b}")
                    self.progress_var.set(40 + (i + 1) / total_b * 40)
                    self.root.update()

        comp_count = 0
        match_count = 0
        
        if is_single_mode:
            total_comps = len(info1) * (len(info1) - 1) // 2
            for i in range(len(info1)):
                for j in range(i + 1, len(info1)):
                    b1 = info1[i]
                    b2 = info1[j]
                    comp_count += 1
                    
                    match_res = self.compare_logic(b1, b2)
                    if match_res != 'NO_MATCH':
                        self.add_to_tree(b1, b2, match_res)
                        match_count += 1
                        
                    if comp_count % 100 == 0 or comp_count == total_comps:
                        self.lbl_selected.config(text=f"主文件夹内部比对中: {comp_count}/{total_comps}")
                        self.progress_var.set((comp_count / total_comps) * 100)
                        self.root.update()
        else:
            total_comps = len(info1) * len(info2)
            for b1 in info1:
                for b2 in info2:
                    comp_count += 1
                    
                    match_res = self.compare_logic(b1, b2)
                    if match_res != 'NO_MATCH':
                        self.add_to_tree(b1, b2, match_res)
                        match_count += 1
                        
                    if comp_count % 100 == 0 or comp_count == total_comps:
                        self.lbl_selected.config(text=f"跨文件夹交叉比对: {comp_count}/{total_comps}")
                        self.progress_var.set(80 + (comp_count / total_comps) * 20)
                        self.root.update()
        
        mode_text = "单文件夹内部清洗" if is_single_mode else "双文件夹对照"
        self.lbl_selected.config(text=f"扫描完成 ({mode_text})，找到 {match_count} 对疑似重复项。请校对。")
        self.reset_scan_ui()
        
        if match_count > 0:
            self.btn_execute.config(state=tk.NORMAL)
        else:
            self.btn_execute.config(state=tk.DISABLED)

    def compare_logic(self, b1, b2):
        # 1. 书名号内核提取拦截（专门对付被垃圾特典词汇包裹的书名）
        core1 = get_core_title(b1['filename'])
        core2 = get_core_title(b2['filename'])
        if core1 != b1['filename'] and core2 != b2['filename']:
            if difflib.SequenceMatcher(None, core1, core2).ratio() < 0.6:
                return 'NO_MATCH'

        # 2. 从原始名字提取完整卷号特征阵列
        vol1 = get_volume_fingerprint(b1['filename']) or get_volume_fingerprint(b1['title'])
        vol2 = get_volume_fingerprint(b2['filename']) or get_volume_fingerprint(b2['title'])

        if vol1 is not None and vol2 is not None and vol1 != vol2:
            return 'NO_MATCH'

        # 3. 隐式第一卷阻击补丁
        if (vol1 is not None and vol2 is None and is_definitely_not_vol_1(vol1)) or \
           (vol2 is not None and vol1 is None and is_definitely_not_vol_1(vol2)):
            return 'NO_MATCH'

        # 4. 文本清洗与相似度测算
        name1_clean = clean_filename(b1['filename'])
        name2_clean = clean_filename(b2['filename'])

        is_match = False
        if b1['title'] and b2['title'] and b1['title'] == b2['title']:
            is_match = True
        else:
            similarity = difflib.SequenceMatcher(None, name1_clean, name2_clean).ratio()
            if similarity > 0.82: 
                is_match = True

        if is_match:
            # 5. 深度差异探伤
            return analyze_diff(name1_clean, name2_clean)
            
        return 'NO_MATCH'

    def add_to_tree(self, b1, b2, match_type):
        self.duplicates.append((b1, b2))
        curr_idx = len(self.duplicates) - 1
        
        if match_type == 'MANUAL':
            plan_text = "⚠️ 需手动(存疑)"
        else:
            plan_text = "⚠️ 需手动"
            if b1['size'] < b2['size']:
                plan_text = "🔴 删 文件A"
            elif b2['size'] < b1['size']:
                plan_text = "🔵 删 文件B"

        self.tree.insert("", tk.END, iid=str(curr_idx), values=(
            b1['title'], 
            b1['filename'], f"{b1['size']:.2f}", 
            b2['filename'], f"{b2['size']:.2f}",
            plan_text
        ))

    def reset_scan_ui(self):
        self.btn_scan.config(state=tk.NORMAL, text="重新扫描")
        self.progressbar.grid_remove()
        self.btn_mark_a.config(state=tk.DISABLED)
        self.btn_mark_b.config(state=tk.DISABLED)
        self.btn_mark_none.config(state=tk.DISABLED)

    def load_image_to_label(self, img_data, label_img, label_info):
        if not img_data:
            label_img.config(image='', text="（无封面）")
            label_info.config(text="")
            return None
            
        try:
            file_size_bytes = len(img_data.getvalue())
            if file_size_bytes > 1024 * 1024:
                size_str = f"{file_size_bytes / (1024 * 1024):.2f} MB"
            else:
                size_str = f"{file_size_bytes / 1024:.1f} KB"

            pil_image = Image.open(img_data)
            width, height = pil_image.size
            img_format = pil_image.format or "未知"

            info_text = f"【封面信息】\n\n"
            info_text += f"图片格式:\n{img_format}\n\n"
            info_text += f"分辨率:\n{width}x{height}\n\n"
            info_text += f"体积:\n{size_str}"

            if height > PREVIEW_HEIGHT:
                scale = PREVIEW_HEIGHT / height
                new_w = int(width * scale)
                try:
                    resample_filter = Image.Resampling.LANCZOS
                except AttributeError:
                    resample_filter = Image.LANCZOS
                pil_image = pil_image.resize((new_w, PREVIEW_HEIGHT), resample_filter)
                
            tk_image = ImageTk.PhotoImage(pil_image)
            label_img.config(image=tk_image, text="")
            label_info.config(text=info_text)
            return tk_image 
            
        except Exception:
            label_img.config(image='', text="（图片无法读取）")
            label_info.config(text="")
            return None

    def clear_covers(self):
        self.lbl_cover_a.config(image='', text="等待加载...")
        self.lbl_info_a.config(text="")
        self.btn_open_a.config(state=tk.DISABLED)
        
        self.lbl_cover_b.config(image='', text="等待加载...")
        self.lbl_info_b.config(text="")
        self.btn_open_b.config(state=tk.DISABLED)
        
        self.img_left_photo = None
        self.img_right_photo = None
        self.current_b1 = None
        self.current_b2 = None

    def on_select(self, event):
        if self._selection_after_id:
            self.root.after_cancel(self._selection_after_id)
            self._selection_after_id = None
            
        selected = self.tree.selection()
        if not selected:
            self.btn_mark_a.config(state=tk.DISABLED)
            self.btn_mark_b.config(state=tk.DISABLED)
            self.btn_mark_none.config(state=tk.DISABLED)
            self.clear_covers()
            return
            
        self.btn_mark_a.config(state=tk.NORMAL)
        self.btn_mark_b.config(state=tk.NORMAL)
        self.btn_mark_none.config(state=tk.NORMAL)

        if len(selected) > 1:
            self.clear_covers()
            self.lbl_cover_a.config(text="[多选模式]\n封面与详情已隐藏")
            self.lbl_cover_b.config(text="[多选模式]\n封面与详情已隐藏")
            self.lbl_selected.config(text=f"已选中 {len(selected)} 项，你可以统一修改它们的计划操作。")
            return

        self.lbl_cover_a.config(image='', text="等待加载 A...")
        self.lbl_cover_b.config(image='', text="等待加载 B...")
        self.lbl_info_a.config(text="")
        self.lbl_info_b.config(text="")
        self.btn_open_a.config(state=tk.DISABLED)
        self.btn_open_b.config(state=tk.DISABLED)
        
        self._selection_after_id = self.root.after(300, lambda: self.process_single_selection(selected[0]))

    def process_single_selection(self, item_id):
        self.lbl_cover_a.config(text="提取中 文件A...")
        self.lbl_cover_b.config(text="提取中 文件B...")
        self.root.update()
        
        try:
            idx = int(item_id) 
            self.current_b1, self.current_b2 = self.duplicates[idx]
            current_plan = self.tree.set(item_id, "action")
            
            cover_a_data = extract_epub_cover(self.current_b1['path'])
            cover_b_data = extract_epub_cover(self.current_b2['path'])
            
            self.img_left_photo = self.load_image_to_label(cover_a_data, self.lbl_cover_a, self.lbl_info_a)
            self.img_right_photo = self.load_image_to_label(cover_b_data, self.lbl_cover_b, self.lbl_info_b)

            self.btn_open_a.config(state=tk.NORMAL)
            self.btn_open_b.config(state=tk.NORMAL)

            self.lbl_selected.config(text=f"文件A: {self.current_b1['size']:.2f}MB | 文件B: {self.current_b2['size']:.2f}MB。当前计划: {current_plan}")
        except Exception:
            self.lbl_cover_a.config(text="提取遇到异常错误")
            self.lbl_cover_b.config(text="提取遇到异常错误")

    def change_plan(self, target):
        selected = self.tree.selection()
        if not selected: return
        
        new_val = ""
        if target == 'A': new_val = "🔴 删 文件A"
        elif target == 'B': new_val = "🔵 删 文件B"
        elif target == 'NONE': new_val = "⚠️ 需手动"
            
        for item in selected:
            self.tree.set(item, column="action", value=new_val)
            
        if len(selected) == 1:
            current_plan = self.tree.set(selected[0], "action")
            current_text = self.lbl_selected.cget("text")
            base_text = current_text.split("。当前计划:")[0]
            self.lbl_selected.config(text=f"{base_text}。当前计划: {current_plan}")

    def execute_all_plans(self):
        items_to_process = self.tree.get_children()
        if not items_to_process: return
        
        to_delete_list = []
        for item_id in items_to_process:
            plan = self.tree.set(item_id, "action")
            idx = int(item_id)
            b1, b2 = self.duplicates[idx]
            
            if "删 文件A" in plan:
                to_delete_list.append((item_id, b1['path'], "文件A"))
            elif "删 文件B" in plan:
                to_delete_list.append((item_id, b2['path'], "文件B"))

        if not to_delete_list:
            messagebox.showinfo("提示", "列表中没有任何被标记为“删文件A”或“删文件B”的项目。\n请先进行标记。")
            return

        msg = f"准备就绪！\n\n即将把列表中标记好的 {len(to_delete_list)} 个文件移入回收站。\n未标记或标记为“需手动”的项目将被跳过。\n\n确定执行吗？"
        if not messagebox.askyesno("执行确认", msg):
            return

        deleted_count = 0
        errors = []

        self.btn_execute.config(state=tk.DISABLED, text="正在执行...")
        self.root.update()

        for item_id, filepath, target in to_delete_list:
            try:
                if os.path.exists(filepath):
                    send2trash(os.path.normpath(filepath))
                self.tree.delete(item_id)
                deleted_count += 1
            except Exception as e:
                errors.append(f"{os.path.basename(filepath)}: {e}")
                self.tree.set(item_id, column="action", value="❌ 失败")

        self.clear_covers()
        self.on_select(None) 
        
        if not self.tree.get_children():
            self.btn_execute.config(state=tk.DISABLED, text="🚀 执行列表中所有的删除计划")
        else:
            self.btn_execute.config(state=tk.NORMAL, text="🚀 执行列表中所有的删除计划")

        result_msg = f"执行完毕！\n成功将 {deleted_count} 个文件移入回收站。"
        
        remaining = len(self.tree.get_children())
        if remaining > 0:
            result_msg += f"\n\n列表中还剩 {remaining} 项未处理（可能是未标记或执行失败）。"
            
        if errors:
            messagebox.showwarning("部分完成", result_msg + f"\n\n警告：有 {len(errors)} 个文件处理失败:\n" + "\n".join(errors[:5]))
        else:
            messagebox.showinfo("清理结果", result_msg)

if __name__ == "__main__":
    root = tk.Tk()
    default_font = ("Microsoft YaHei UI", 9)
    root.option_add("*Font", default_font)
    app = EpubDupeFinderApp(root)
    root.mainloop()
