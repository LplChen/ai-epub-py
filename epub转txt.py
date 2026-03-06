import zipfile
import os
import sys
import urllib.parse
from bs4 import BeautifulSoup

def process_epub(epub_path):
    """
    处理单个EPUB文件，提取文本并保存为TXT
    """
    # 确定输出TXT的文件名（与EPUB同名）
    txt_path = os.path.splitext(epub_path)[0] + '.txt'
    
    full_text = []
    
    try:
        with zipfile.ZipFile(epub_path, 'r') as z:
            # 1. 读取 META-INF/container.xml 找到 OPF 文件的路径
            try:
                container = z.read('META-INF/container.xml')
                container_soup = BeautifulSoup(container, 'xml')
                rootfile = container_soup.find('rootfile')
                if not rootfile:
                    raise ValueError("无法找到 rootfile 节点")
                opf_path = urllib.parse.unquote(rootfile.get('full-path'))
            except Exception as e:
                raise ValueError(f"解析 container.xml 失败: {e}")

            # 2. 读取 OPF 文件获取阅读顺序 (spine) 和文件清单 (manifest)
            opf_content = z.read(opf_path)
            opf_soup = BeautifulSoup(opf_content, 'xml')
            
            manifest = {}
            for item in opf_soup.find_all('item'):
                manifest[item.get('id')] = urllib.parse.unquote(item.get('href'))
                
            spine = []
            for itemref in opf_soup.find_all('itemref'):
                spine.append(itemref.get('idref'))
                
            opf_dir = os.path.dirname(opf_path)
            
            # 3. 按照阅读顺序提取并解析每个 HTML/XHTML 文件
            for item_id in spine:
                if item_id in manifest:
                    href = manifest[item_id]
                    
                    # 过滤掉导航目录文件 nav.xhtml
                    if 'nav.xhtml' in href.lower():
                        continue
                        
                    # 拼接内部路径
                    file_path = href if not opf_dir else f"{opf_dir}/{href}"
                    
                    try:
                        html_content = z.read(file_path)
                        soup = BeautifulSoup(html_content, 'html.parser')
                        
                        # 清理不可见元素：移除 script 和 style 标签避免提出多余代码
                        for script_or_style in soup.find_all(['script', 'style']):
                            script_or_style.decompose()
                        
                        # --- 核心逻辑：去除上标小字 ---
                        # 查找并彻底删除所有 <sup> 标签（通常用于脚注上标）
                        for sup in soup.find_all('sup'):
                            sup.decompose()  # decompose() 会将标签及其内部的文本完全销毁
                            
                        # 如果有其他特定类名的上标，比如 <span class="footnote">，可取消下面代码的注释并修改：
                        # for span in soup.find_all('span', class_='footnote'):
                        #     span.decompose()

                        # --- 修复断行问题，提取连贯的正文 ---
                        # 1. 移除源码自带的回车换行符，防止将一句话从中间切断
                        for text_node in soup.find_all(string=True):
                            if '\n' in text_node:
                                text_node.replace_with(text_node.replace('\n', ''))
                                
                        # 2. 将 <br> 标签转换为真实的换行符
                        for br in soup.find_all(['br', 'br/']):
                            br.replace_with('\n')
                            
                        # 3. 在所有块级元素后手动追加换行标志，保证段落之间正常换行
                        for block in soup.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li']):
                            block.insert_after('\n')

                        # 4. 提取纯文本。此时不使用默认的 separator 分隔符，确保剔除 <sup> 后的两侧文本无缝拼接
                        body = soup.find('body')
                        raw_text = body.get_text(separator='') if body else soup.get_text(separator='')
                        
                        # 5. 按行分割，去除多余的空白并过滤空行
                        lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
                        text = '\n'.join(lines)
                            
                        if text:
                            full_text.append(text)
                    except KeyError:
                        print(f"  [警告] 在压缩包中找不到文件: {file_path}")
                    except Exception as e:
                        print(f"  [警告] 解析 {file_path} 时出错: {e}")

        # 4. 写入TXT文件
        with open(txt_path, 'w', encoding='utf-8') as f:
            # 章节之间使用两个换行符隔开
            f.write('\n\n'.join(full_text))
            
        return True
        
    except Exception as e:
        print(f"  [错误] 处理文件时发生严重错误: {e}")
        return False

def main():
    print("=== EPUB 转 TXT 提取工具 ===")
    print("特点：自动去除正文中的上标小字\n")
    
    files_to_process = []
    
    # 检查是否通过拖放文件或命令行参数传入了文件
    if len(sys.argv) > 1:
        # 获取传入的所有 .epub 文件
        files_to_process = [f for f in sys.argv[1:] if f.lower().endswith('.epub')]
    else:
        # 如果没有拖放文件，则扫描当前目录下的所有 .epub 文件
        print("未检测到拖放文件，正在扫描当前目录...")
        files_to_process = [f for f in os.listdir('.') if f.lower().endswith('.epub')]
        
    if not files_to_process:
        print("没有找到 EPUB 文件！")
        print("使用方法：")
        print("1. 将一个或多个 .epub 文件直接拖放到此脚本图标上。")
        print("2. 或者将此脚本放在包含 .epub 文件的文件夹中双击运行。")
    else:
        success_count = 0
        for epub_file in files_to_process:
            print(f"正在转换: {os.path.basename(epub_file)}")
            if process_epub(epub_file):
                print(f"  -> 成功: 已保存为 TXT")
                success_count += 1
            else:
                print(f"  -> 失败: 转换未完成")
                
        print(f"\n处理完毕！共成功转换 {success_count}/{len(files_to_process)} 个文件。")

    # 保持窗口打开，以便用户查看结果（特别是在拖放文件时）
    print("\n")
    os.system('pause' if os.name == 'nt' else 'read -p "按回车键退出..." -n 1')

if __name__ == "__main__":
    main()
