# PyConsole Manager - Python脚本管理器

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)](https://www.microsoft.com/windows)]

除了这段话,其他都是AI生成.不知道有没有人和我有相同的痛点,每运行一个python脚本,就会多出一个cmd窗口,如果运行多个就不太好看,有些时候cmd是一点都不需要看的,因为已经有UI界面,有时候要看LOG的时候,哪个窗口的log对的是哪个脚本又有可能分不清.
所以叫AI用VB6搞了这个.开始调试的时候一直对隐藏CMD窗口都搞不好,另外对脚本读取配置文件"config.json"的目录有时候也会出错,调试还是花了挺长一段时间的,最终是用deepseek的网页版完成.

一个用VB6开发的Python脚本管理工具，可以同时管理多个Python脚本，控制CMD窗口的显示和隐藏。

## ✨ 功能特点

- 📁 **脚本管理** - 添加、删除Python脚本，列表显示运行状态
- 🚀 **单脚本控制** - 启动、停止、隐藏窗口、显示窗口、查看属性
- 📦 **批量操作** - 全部启动、全部停止、全部隐藏、全部显示
- ⚡ **快捷操作** - 双击列表项快速启动/停止
- 💾 **状态保持** - 自动保存脚本列表，关闭时自动停止所有脚本

## 🖥️ 系统要求

- Windows XP / 7 / 8 / 10 / 11
- Python 3.x（已添加到系统PATH）
- 无额外依赖，单个exe文件即可运行

## 📥 下载安装

### 方式一：下载编译好的exe
1. 前往 [Releases](https://github.com/cxocto/PyConsoleManager/releases) 页面
2. 下载最新的 `PyConsoleManager.exe`
3. 双击运行，无需安装

### 方式二：从源码编译
1. 安装 Visual Basic 6.0
2. 克隆本仓库
3. 打开 `Project1.vbp` 工程文件
4. 按 `Ctrl+F5` 运行或生成exe

## 🚀 快速上手

### 1. 添加脚本
点击「添加脚本」按钮，选择要管理的 `.py` 文件

### 2. 启动脚本
- 双击脚本列表项
- 或选中后点击「启动所选」

### 3. 隐藏/显示窗口
选中运行中的脚本，点击「隐藏窗口」或「显示窗口」

### 4. 停止脚本
- 双击运行中的脚本
- 或选中后点击「停止所选」

### 5. 批量控制
使用顶部按钮可一键控制所有脚本

## 📸 界面预览
