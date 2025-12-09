# 燃料管理系统 (Fuel Management System)

## 简介
这是一个入厂煤化验与称重管理系统，旨在自动化处理和汇总燃料管理过程中的化验数据与称重数据。系统包含一个基于 Flask 的后端服务和一个配套的微信小程序前端。

## 功能特性
- **化验月报汇总**: 自动读取 Excel 化验报告，生成月度统计、加权平均分析及分类汇总。
- **称重月报汇总**: 自动读取 Excel 称重记录，结合化验数据进行关联分析，生成供应商供货统计。
- **微信小程序**: 提供移动端访问接口，方便随时查看数据。
- **数据可视化**: 支持生成各类统计报表（Excel 格式）。

## 技术栈
- **后端**: Python, Flask, Pandas, OpenPyxl, XlsxWriter
- **前端**: 微信小程序 (Miniprogram)

## 快速开始

### 1. 环境准备
确保已安装 Python 3.8+。

```bash
pip install -r requirements.txt
```

### 2. 运行应用
```bash
python app.py
```
服务将在 `http://0.0.0.0:5000` 启动。

## 目录结构
- `app.py`: Flask 应用主入口
- `hy.py`: 化验数据处理逻辑
- `cz.py`: 称重数据处理逻辑
- `miniprogram/`: 微信小程序源码
- `templates/` & `static/`: Web 前端资源
- `无人值守化验月报/`: 化验数据输入目录
- `无人值守称重月报/`: 称重数据输入目录

## 许可证
[MIT License](LICENSE)
