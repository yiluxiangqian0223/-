# Exam Image Generator

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![Status](https://img.shields.io/badge/status-stable-green.svg)](https://github.com/yourusername/exam-image-generator/releases)

考场图像信息核对单生成工具 - 用于创建标准化考试管理文档

## 📋 目录

- [简介](#-简介)
- [功能特性](#-功能特性)
- [安装](#-安装)
- [使用方法](#-使用方法)
- [配置](#-配置)
- [示例](#-示例)
- [API接口](#-api接口)
- [贡献指南](#-贡献指南)
- [许可证](#-许可证)

## 💡 简介

**Exam Image Generator** 是一个专业的Python工具，用于自动生成标准化的考场图像信息核对单。该工具专为教育考试场景设计，能够从Excel数据源读取考生信息，结合证件照片，批量生成符合考试管理规范的PDF文档。

## ✨ 功能特性

- 📊 **数据驱动**: 从Excel文件自动读取考生信息
- 📸 **照片集成**: 自动匹配考生证件照片
- 📄 **标准化输出**: 生成符合考试管理规范的PDF文档
- 🎨 **定制布局**: 支持竖排文字和多种信息展示方式
- 📈 **批量处理**: 一次性生成多个考场的核对单
- 🌐 **多语言支持**: 内置中文字符集支持
- 🛡️ **数据安全**: 本地处理，数据不上传云端
- 🔧 **高度可配置**: 灵活的布局和样式设置

## 🛠️ 安装

### 系统要求
- Python 3.8 或更高版本
- 至少 100MB 可用磁盘空间

### 安装步骤

1. **克隆仓库**
git clone https://github.com/yiluxiangqian0223/exam-image-generator.git
cd exam-image-generator

2. **创建虚拟环境（推荐）**
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate     # Windows

3.**安装依赖**
pip install -r requirements.txt
🚀 使用方法
基本使用
准备数据文件
将考生信息保存为 2026kctx.xlsx，包含以下列：
考场号 - 考场编号
座位号 - 座位编号
姓名 - 考生姓名
身份证号 - 身份证号码
