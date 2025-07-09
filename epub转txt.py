import os
import sys
from epub2txt import epub2txt

def extract_text_from_epub(epub_file, output_txt_file):
    # 提取 EPUB 文件的文本
    text = epub2txt(epub_file)
    # 将提取的文本写入 TXT 文件
    with open(output_txt_file, 'w', encoding='utf-8') as f:
        f.write(text)

def process_epub(epub_file):
    output_txt_file = f"{os.path.splitext(os.path.basename(epub_file))[0]}.txt"  # 输出 TXT 文件名
    extract_text_from_epub(epub_file, output_txt_file)
    print(f"文本提取完成，保存为: {output_txt_file}")

def main():
    # 检查是否通过拖放的方式提供 EPUB 文件
    if len(sys.argv) > 1:
        for epub_file in sys.argv[1:]:
            epub_file = epub_file.strip('"')  # 去掉可能的引号
            epub_file = epub_file.strip('{}')  # 去掉可能的花括号
            if epub_file.lower().endswith('.epub') and os.path.isfile(epub_file):
                process_epub(epub_file)
            else:
                print(f"无效文件：{epub_file}，请提供有效的 EPUB 文件。")
        return

    # 如果没有拖放文件，检查当前目录中的 EPUB 文件
    current_directory = os.getcwd()
    epub_files = [f for f in os.listdir(current_directory) if f.endswith('.epub')]
    
    if not epub_files:
        print("当前目录没有找到 EPUB 文件。")
        return
    
    for epub_file in epub_files:
        process_epub(os.path.join(current_directory, epub_file))

if __name__ == "__main__":
    main()