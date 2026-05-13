# MsOfficeHub

一个现代化的 Microsoft Office 集合启动器与文档管理工具，为 Windows 用户提供统一的 Office 应用快速启动和多源文档聚合管理。

## 核心特性

✨ **Office 应用快速启动**
- 一键访问所有已安装的 Microsoft Office 应用（Word、Excel、PowerPoint、Outlook、OneNote、Access、Publisher、Project、Visio）
- 自动检测 Office 安装路径
- 智能优先级：自动识别 WindowsApps 版 Outlook

📋 **多源文档聚合管理**
- 统一汇聚以下来源的最近文档：
  - Office 应用注册表 MRU
  - 本地缓存文件
  - Microsoft Edge 浏览历史（PDF 文档）
- 自动去重与文件存在性验证
- 本地文件与云文件自适应图标显示

🎨 **现代化用户界面**
- 基于 WinUI 3 的原生 Windows 11 设计
- 可调节左右栏宽度（隐形拖拽柄，响应式反馈）
- 系统主题自适应（深色/浅色模式）
- 文件类型图标自动识别（本地文件 + 云存储）
- 悬停提示显示完整文件路径

💾 **智能状态持久化**
- 窗口位置、大小、布局自动保存和恢复
- 栏宽偏好保存至 Windows 注册表
- 配置存储位置：`HKEY_CURRENT_USER\Software\EricSoft\MsOfficeHub`

## 系统要求

- Windows 10 21H2 或更新版本
- .NET 10.0 运行时
- Microsoft Office（任何支持 16.0 注册表的版本）

## 快速开始

### 下载预构建版

1. 从 [Releases](https://github.com/EricZhang233/MsOfficeHub/releases) 下载最新的 `MsOfficeHub_x.x.x.x.zip`
2. 解压到任意位置
3. 运行 `MsOfficeHub.exe`

### 从源代码构建

```bash
git clone https://github.com/EricZhang233/MsOfficeHub.git
cd MsOfficeHub
dotnet build -c Release
```

构建输出将自动打包到 `output/MsOfficeHub_x.x.x.x.zip`

## 使用指南

### 启动 Office 应用

点击左侧面板中的应用图标即可启动。未安装的应用将显示为禁用状态。

### 打开最近文档

- 点击右侧列表中的任意文档，将自动用关联应用打开
- 单击操作已防抖（1秒），避免误触双击打开
- 悬停显示文件完整路径

### 调整布局

- 拖拽左右面板间的隐形分割线调整宽度
- 偏好设置自动保存

### Cloud 文件与 Edge 历史

Edge 浏览历史中的 PDF 文档自动聚合到最近列表，便捷重新打开。

## 配置说明

### 调试模式

在应用目录创建 `debug.ini`：

```ini
ForceShowAllApps=true
```

显示所有 Office 应用（无论是否安装）。

### 自定义配置

在应用目录创建 `config.json` 以支持扩展设置（预留接口）。

## 技术架构

**核心模块：**

- `MainWindow.xaml/.cs` - UI 界面与事件处理
- `Recent.cs` - 多源文档聚合引擎
- `OfficeDetector.cs` - Office 安装检测
- `ConfigService.cs` - 配置管理
- `Version.cs` - 版本管理

**数据来源：**

1. **注册表 MRU** - `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\{App}\File MRU`
2. **本地缓存** - `%APPDATA%\Microsoft\Office\Recent`
3. **Edge 历史** - `%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\History`（SQLite）
4. **应用配置** - `HKEY_CURRENT_USER\Software\EricSoft\MsOfficeHub`

## 构建与发布

### 本地构建与打包

```bash
dotnet build -c Release
```

自动执行：
1. Release 配置编译
2. 生成 `output/MsOfficeHub_{版本号}.zip`

### GitHub 自动化流程

- **自动发布** - 推送到 `main` 分支时创建 Prerelease
- **后处理** - 发布时上传 ZIP 资产并清理 output 目录
- **版本转正** - 日均检查，48小时以上 Prerelease 自动转为正式版

## 贡献指南

欢迎提交 Issue 或 Pull Request！请遵循：

1. 代码无注释（项目编码规范）
2. 持久化配置使用注册表
3. 云文件 URI 需进行空值检查
4. 测试涵盖本地文件和云文件场景

## 许可证

详见 [LICENSE](LICENSE) 文件。

## 作者

[@EricZhang233](https://github.com/EricZhang233)
