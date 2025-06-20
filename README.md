# Word文档自动缩小排版工具（支持WPS/Word）

本工具用于自动压缩 `.docx` 格式的Word文档，通过调整页面方向（横向A4）、分栏数（最多4栏）、字体大小（最小8号）、行距（紧凑行距）及图片尺寸，实现文档页数的自动缩减，目标是尽量在不改变内容的前提下，将文档缩小至用户指定的页数。

> 支持使用WPS或微软Word进行PDF转换，以检测文档实际页数。

---

## 功能特点

- 自动调整页面为横向A4纸张  
- 允许最多4栏排版，自动计算每栏宽度  
- 动态压缩图片宽度，防止撑满页面  
- 字体最小8号，自动删除空白行  
- 文字段落和含图段落行距分别设定，保证图片显示  
- 反复尝试多种压缩策略组合，直到达到目标页数  
- 通过 `docx2pdf` 转换为PDF后使用 `pypdf` 统计页数  
- 适配WPS和微软Word环境  

---

## 环境依赖

- Python 3.7+  
- 需要安装Python库：

```bash
pip install python-docx docx2pdf pypdf
```
## 使用方法
- 将文件放入根目录
- 运行以下命令

```bash
python word_shrinker.py [输入文件.docx] [目标页数]
```

- 例子
```bash
python word_shrinker.py sample.docx 10
