# LLM Extraction System

本项目是一个用于从网页、Excel 和 PDF 文件中提取信息的系统，使用 Python 编写，并集成了对网页解析、文本提取和表格处理的功能，适用于基于 LLM 的内容预处理任务。

## 功能简介

- 从指定网页中抓取并解析 HTML 内容。
- 从 PDF 文件中提取纯文本。
- 从 Excel 文件中读取数据。
- 整合提取的数据用于进一步处理。

## 环境依赖

请使用以下命令安装项目依赖项：

```bash
pip install -r requirements.txt
```

## 文件说明

- `llm-extraction-system.py`：主程序文件，包含数据提取的全部功能。
- `requirements.txt`：项目依赖项清单。

## 使用方法

1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 运行主程序：

```bash
python llm-extraction-system.py
```

3. 根据程序提示输入网页链接、PDF 文件路径或 Excel 文件路径，程序将自动完成信息提取。

## 注意事项

- 请确保 PDF 文件文本可选中，避免扫描版图像型文档。
- 需要网络连接以访问网页内容。
- 推荐使用 Python 3.8 或以上版本。

## 联系方式

如有问题，请联系项目维护者。
