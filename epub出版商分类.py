import os
import shutil
import re
from ebooklib import epub
from tqdm import tqdm

def clean_publisher_name(publisher):
    """清理出版商名称：删除'株式会社'及其前后空格，处理非法字符"""
    if not publisher:
        return "未知出版商"
    
    # 删除'株式会社'及其前后空格
    cleaned = re.sub(r'\s*株式会社\s*', '', publisher, flags=re.IGNORECASE)
    
    # 处理非法文件名字符（Windows/Linux通用）
    cleaned = re.sub(r'[\\/*?:"<>|]', '', cleaned).strip()
    return cleaned if cleaned else "未知出版商"

def process_epub_files():
    # 初始化计数器
    stats = {"total": 0, "moved": 0, "skipped_exist": 0, "skipped_error": 0}
    
    # 获取当前目录下所有EPUB文件
    current_dir = os.getcwd()
    epub_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.epub')]
    stats["total"] = len(epub_files)
    
    if not epub_files:
        print("⚠️ 未发现EPUB文件")
        return stats
    
    print(f"📚 发现 {stats['total']} 个EPUB文件，开始处理...")
    
    for filename in tqdm(epub_files, desc="处理进度"):
        file_path = os.path.join(current_dir, filename)
        
        try:
            # 读取EPUB元数据
            book = epub.read_epub(file_path)
            publishers = book.get_metadata('DC', 'publisher')
            publisher = publishers[0][0] if publishers else "未知出版商"
            
            # 清理出版商名称
            publisher_cleaned = clean_publisher_name(publisher)
            target_dir = os.path.join(current_dir, publisher_cleaned)
            
            # 创建出版商文件夹（如果不存在）
            os.makedirs(target_dir, exist_ok=True)
            
            # 检查目标文件是否已存在
            target_path = os.path.join(target_dir, filename)
            if os.path.exists(target_path):
                stats["skipped_exist"] += 1
                continue
                
            # 移动文件
            shutil.move(file_path, target_path)
            stats["moved"] += 1
            
        except Exception as e:
            stats["skipped_error"] += 1
            tqdm.write(f"❌ 错误处理 '{filename}': {str(e)}")
    
    # 输出统计结果
    print("\n处理结果:")
    print(f"✅ 成功移动: {stats['moved']} 文件")
    print(f"⏩ 跳过已存在: {stats['skipped_exist']} 文件")
    print(f"⚠️ 处理失败: {stats['skipped_error']} 文件")
    print(f"📊 总计处理: {stats['total']} 文件")
    return stats

if __name__ == "__main__":
    process_epub_files()