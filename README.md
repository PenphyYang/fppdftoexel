# 发票分析工具

这是一个带图形界面的 Python 工具，用来批量提取发票 PDF 内容，并导出为 Excel。

## 功能

- 支持一次选择多个 PDF 发票
- 自动提取发票基础信息
- 自动提取明细行
- 导出 `发票汇总` 和 `发票明细` 两个工作表
- 支持打包成单个 EXE

## 运行

```bat
E:\pyenv\pyenv-win\versions\3.8.10\python.exe invoice_pdf_to_excel_gui.py
```

## 命令行模式

```bat
E:\pyenv\pyenv-win\versions\3.8.10\python.exe invoice_pdf_to_excel_gui.py --input a.pdf b.pdf --output 结果.xlsx
```

## 打包 EXE

先确保已经安装依赖：

```bat
E:\pyenv\pyenv-win\versions\3.8.10\python.exe -m pip install -r requirements.txt
```

然后执行：

```bat
build_exe.bat
```

打包完成后的文件位置：

```text
dist\invoice_analysis_tool.exe
```

## 导出结果说明

Excel 中会生成两个工作表：

- `发票汇总`：每个 PDF 一行，包含发票号、日期、购销双方、金额合计等
- `发票明细`：提取到的明细项目

## 注意

- 当前解析规则针对常见电子发票版式做了适配
- 如果 PDF 是扫描件图片，后续可以再补 OCR 方案
- 少数复杂换行版式下，明细描述可能需要人工校对
