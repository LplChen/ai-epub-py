import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

def get_epub_title(epub_path):
    try:
        with zipfile.ZipFile(epub_path) as z:
            # 解析容器文件找到OPF路径
            container = z.read('META-INF/container.xml')
            root = ET.fromstring(container)
            ns = {'n': 'urn:oasis:names:tc:opendocument:xmlns:container'}
            rootfile = root.find('.//n:rootfile', ns)
            opf_path = rootfile.get('full-path')

            # 解析OPF文件获取标题
            opf_data = z.read(opf_path)
            opf_tree = ET.fromstring(opf_data)
            ns = {
                'opf': 'http://www.idpf.org/2007/opf',
                'dc': 'http://purl.org/dc/elements/1.1/'
            }
            title = opf_tree.find('.//dc:title', ns).text
            return title.strip()
    except Exception as e:
        print(f"读取 {epub_path} 元数据失败: {str(e)}")
        return None

def safe_filename(title):
    # 移除非法字符，替换为下划线
    cleaned = re.sub(r'[\\/*?:"<>|]', '_', title)
    # 移除首尾空白和点
    cleaned = cleaned.strip().strip('.')
    # 限制文件名长度
    return cleaned[:200]

def rename_epub_files(directory='.'):
    path = Path(directory)
    for epub in path.glob('*.epub'):
        title = get_epub_title(epub)
        if not title:
            continue

        new_name = safe_filename(title) + '.epub'
        new_path = epub.with_name(new_name)
        
        # 处理文件名冲突
        counter = 1
        while new_path.exists():
            new_name = f"{safe_filename(title)}_{counter}.epub"
            new_path = epub.with_name(new_name)
            counter += 1

        try:
            epub.rename(new_path)
            print(f"重命名成功: {epub.name} -> {new_name}")
        except Exception as e:
            print(f"重命名失败 {epub.name}: {str(e)}")

if __name__ == '__main__':
    # 使用当前目录，或指定其他目录：rename_epub_files('/path/to/epubs')
    rename_epub_files()