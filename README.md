# 自动化文档处理系统

## 功能简介

本系统用于自动化生成视频授权书文档。主要功能包括：

1. 从 Excel 文件读取数据。
2. 解析 Word 模板并填充数据。
3. 自动转换日期格式。
4. 自动从企业微信云盘下载签名图片并插入文档。
5. 按“达人昵称”创建文件夹并分类存储生成的文档。

## 环境准备

1. 确保已安装 Python 3。
2. 运行安装脚本配置环境：
   ```bash
   bash setup.sh
   ```
   该脚本会创建虚拟环境并安装所需的 Python 库（pandas, python-docx, openpyxl, playwright 等），并下载 Playwright 浏览器驱动。

## 使用方法

### 1. 运行主程序

在终端中执行以下命令启动处理程序：

```bash
./venv/bin/python document_processor.py
```

### 2. 登录企业微信云盘

程序启动后会打开一个浏览器窗口并访问企业微信云盘链接。
**请在浏览器中扫码登录。**
登录成功后，程序会自动开始处理 Excel 中的每一行数据，下载对应的签名图片并生成 Word 文档。

### 3. 查看结果

生成的文档将保存在 `output` 目录下。
目录结构如下：

```
output/
  [达人昵称1]x卡赫视频授权/
    达人昵称1_卡赫视频授权书.docx
    signature.png (下载的签名图片)
  [达人昵称2]x卡赫视频授权/
    ...
```

## 调试模式

如果需要测试文档生成功能但跳过网络下载（例如网络不通或不想登录），可以使用调试脚本：

```bash
./venv/bin/python debug_processor.py
```

该脚本会使用模拟数据代替网络请求，生成带有测试图片的文档。

## 注意事项

- Excel 文件路径：`/Users/kang.song/projects/own/document-autofill/卡赫视频授权书.xlsx`
- Word 模板路径：`/Users/kang.song/projects/own/document-autofill/模板文档2.docx`
- 请确保 Excel 文件中的“请签本名”列包含有效的云盘链接。
- 如果遇到登录超时，请重新运行程序并尽快扫码。
