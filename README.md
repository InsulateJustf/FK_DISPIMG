# FK_DISPIMG - WPS 图片转换工具

## 简介
FK_DISPIMG 是一个用于解决 WPS Office 创建的 XLSX 文件中图片显示问题的工具。

### 问题背景
WPS Office 在插入图片时使用专有的 `=@_xlfn.DISPIMG` 函数，将图片嵌入单元格中。虽然图片数据包含在 XLSX 文件中，但其他电子表格软件（如 Microsoft Excel、LibreOffice Calc、OnlyOffice 等）无法正常显示这些图片。

### 解决方案
本工具可以将 WPS 创建的 XLSX 文件中的图片转换为标准的 Excel 嵌入图片格式，确保在其他电子表格软件中正常显示。

## 功能特性
- 支持批量转换多个 XLSX 文件
- 提供图形用户界面 (GUI) 和命令行界面
- 支持拖放文件操作
- 保持原始文件结构，仅修改图片嵌入方式
- 转换后的图片可以调整大小和位置

## 快速开始（推荐）

### 下载预编译版本
无需安装 Python，直接下载对应操作系统的可执行文件：

1. 访问 [GitHub Releases](https://github.com/yourusername/FK_DISPIMG/releases) 页面
2. 下载最新版本：
   - **Windows**: 下载 `FK_DISPIMG-Windows.exe`
   - **macOS**: 下载 `FK_DISPIMG-macOS.app`
3. 直接运行可执行文件

### 从源码运行
如果需要从源码运行，请按照以下步骤：

## 安装要求
- Python 3.7+
- 依赖包：`tkinterdnd2>=0.6.0`

## 安装步骤
1. 克隆或下载本项目
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 运行程序：
   ```bash
   python gui.py
   ```

## 使用方法
### 图形界面模式 (GUI)
1. 启动程序：
   - 如果使用预编译版本：直接双击可执行文件
   - 如果从源码运行：`python gui.py`
2. 通过以下方式添加文件：
   - **拖放操作**：直接将 XLSX 文件拖放到程序窗口
   - **文件选择**：点击"添加文件"按钮选择文件
   - **批量处理**：可以同时添加多个文件
3. 点击"开始转换"按钮
4. 转换完成后，文件将保存在原文件同目录，文件名添加 `_converted` 后缀

### GUI 界面说明
- **文件列表区域**：显示待转换的文件
- **进度条**：显示转换进度
- **状态栏**：显示当前操作状态和完成信息
- **转换按钮**：开始批量转换所有文件

### 命令行模式
```bash
python converter.py 输入文件.xlsx
```

### 批量转换示例
```bash
# 转换单个文件
python converter.py "文档.xlsx"

# 转换多个文件（使用通配符）
python converter.py *.xlsx
```

## 文件说明
- `gui.py` - 图形用户界面主程序
- `converter.py` - 核心转换逻辑
- `requirements.txt` - Python 依赖包列表
- `build.py` - 打包构建脚本
- `FK_DISPIMG.spec` - PyInstaller 打包配置文件

## 注意事项
- 转换过程不会修改原始文件，会生成新的转换后文件
- 转换后的图片可能需要手动调整大小和位置以达到最佳显示效果
- 建议在转换前备份原始文件
- macOS 用户首次运行 .app 文件可能需要在"系统偏好设置 > 安全性与隐私"中允许运行(macOS版本未经测试)
