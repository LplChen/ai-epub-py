# 用ai写的批量处理epub脚本
需要安装`Python`，安装时勾选`Add Python 3.x to PATH`

## epub出版商分类.py
根据epub元数据中的`“出版商”`名称创建文件夹并移动文件，无出版社数据则归类到`“未知出版商”`
 - 需安装第三方库`ebooklib`和`tqdm`：在终端（PowerShell 或 CMD）输入`pip install ebooklib tqdm`并回车安装

<br/>

## epub文件名重命名.py
将epub文件名重命名为元数据中的`“标题”`

<br/>

## epub转txt.py
将`epub`转换为`txt`

<br/>

## epub语言分类.py
根据epub元数据中的`“语言”`名称创建文件夹并移动文件，无`“语言”`数据则归类到`“未知语言”`

<br/>

## epub查重助手.pyw
对单个或两个文件夹内（包含子文件夹）的epub文件进行扫描比对查重，手动选择删除重复的文件
- 需安装第三方库`Pillow`和`send2trash`：在终端（PowerShell 或 CMD）输入`pip install Pillow send2trash`并回车安装
- 双击打开`epub查重助手.pyw`
