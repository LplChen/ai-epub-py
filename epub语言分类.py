import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
import sys
import traceback
import time

# 定义语言代码到文件夹名称的映射
LANG_MAP = {
    'zh': '中文',
    'chi': '中文',
    'zho': '中文',
    
    'ja': '日文',
    'jpn': '日文',
    
    'en': '英文',
    'eng': '英文',
    
    'fr': '法文',
    'fra': '法文',
    'fre': '法文',
    
    'de': '德文',
    'deu': '德文',
    'ger': '德文',
    
    'ru': '俄文',
    'rus': '俄文',
    
    'ko': '韩文',
    'kor': '韩文'
}

DEFAULT_FOLDER = "未知语言"
LOG_FILE = "run_log.txt"

def log(message, to_console=True):
    """
    写日志并打印到控制台，带有异常处理防止因字符编码导致的闪退
    """
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    full_msg = f"[{timestamp}] {message}"
    
    # 写入文件（使用 utf-8，最安全）
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(full_msg + "\n")
    except Exception:
        pass # 写入日志失败不应该影响主程序

    # 打印到控制台（容错处理）
    if to_console:
        try:
            print(full_msg)
        except UnicodeEncodeError:
            # 如果文件名包含生僻字导致控制台打印失败，尝试去除无法编码的字符后打印
            try:
                safe_msg = full_msg.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding)
                print(safe_msg)
            except:
                print(f"[{timestamp}] [显示错误] 某条信息包含无法显示的字符")

def get_epub_language(epub_path):
    """
    读取 EPUB 文件的元数据并提取语言信息。
    """
    try:
        if not zipfile.is_zipfile(epub_path):
            return None

        with zipfile.ZipFile(epub_path, 'r') as zf:
            try:
                container_xml = zf.read('META-INF/container.xml')
            except KeyError:
                return None 

            root = ET.fromstring(container_xml)
            opf_path = None
            for elem in root.iter():
                if elem.tag.endswith('rootfile'):
                    opf_path = elem.attrib.get('full-path')
                    break
            
            if not opf_path:
                return None

            opf_data = zf.read(opf_path)
            opf_root = ET.fromstring(opf_data)
            
            for elem in opf_root.iter():
                if elem.tag.endswith('language'):
                    # 核心修复：增加 .strip() 去除前后的换行符和空格
                    if elem.text and elem.text.strip():
                        return elem.text.strip()
            
            return None

    except Exception as e:
        # 不在这里打印错误，由外部捕获或忽略，避免刷屏
        return None

def normalize_lang_folder(lang_code):
    if not lang_code:
        return DEFAULT_FOLDER
    
    # 双重保险：再次去除空格和换行符
    clean_code = lang_code.strip()
    
    # 转换为小写并取第一部分（例如 zh-CN -> zh）
    code_base = clean_code.lower().split('-')[0]
    return LANG_MAP.get(code_base, code_base)

def main():
    try:
        # 初始化日志文件
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("开始运行任务...\n")

        current_dir = os.getcwd()
        log(f"正在扫描目录: {current_dir}")
        
        # 使用 scandir 提高大量文件时的扫描性能
        files = []
        with os.scandir(current_dir) as it:
            for entry in it:
                if entry.is_file() and entry.name.lower().endswith('.epub'):
                    files.append(entry.name)

        total_files = len(files)
        log(f"发现 {total_files} 个 EPUB 文件，准备处理...")
        print("-" * 30)

        count = 0
        error_count = 0
        
        for index, filename in enumerate(files, 1):
            # 每处理100个给个心跳提示，防止用户以为卡死了
            if index % 100 == 0:
                print(f"进度: {index}/{total_files} ...")

            try:
                file_path = os.path.join(current_dir, filename)
                
                raw_lang = get_epub_language(file_path)
                target_folder_name = normalize_lang_folder(raw_lang)
                
                # 安全检查：确保文件夹名不包含非法字符（Windows）
                # 比如某些元数据可能真的写错了变成含冒号等，这里简单替换掉
                invalid_chars = '<>:"/\\|?*\n\r\t'
                for char in invalid_chars:
                    target_folder_name = target_folder_name.replace(char, '')
                
                if not target_folder_name:
                    target_folder_name = DEFAULT_FOLDER

                target_dir = os.path.join(current_dir, target_folder_name)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)
                    
                target_path = os.path.join(target_dir, filename)
                
                if os.path.exists(target_path):
                    log(f"[跳过] {filename} -> 目标已存在")
                    continue
                    
                shutil.move(file_path, target_path)
                
                # 仅在日志中记录详细信息，控制台减少输出防止刷屏过快
                # log(f"[移动] {filename} -> {target_folder_name}") 
                count += 1
                
            except Exception as e:
                error_count += 1
                log(f"[错误] 处理文件 '{filename}' 失败: {str(e)}")

        print("-" * 30)
        log(f"处理完成！成功移动: {count}，失败/跳过: {total_files - count}")
        log(f"详细运行记录请查看当前目录下的 {LOG_FILE}")

    except Exception as e:
        # 这是为了捕获致命错误，防止窗口直接关闭
        print("\n" + "!"*30)
        print("程序发生了严重的未预料错误：")
        traceback.print_exc()
        print("!"*30 + "\n")
        
        # 尝试写入致命错误到日志
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write("\n致命错误:\n")
                f.write(traceback.format_exc())
        except:
            pass

    # 无论是否出错，都暂停等待用户确认
    input("\n按回车键退出程序...")

if __name__ == "__main__":
    main()