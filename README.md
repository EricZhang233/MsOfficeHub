# MsOfficeHub

一个现代化的 Microsoft Office 集合启动器，为 Windows 用户提供一个统一的 Office 应用快速访问中心和智能文档管理工具。

## ✨ 主要特性

### 📱 一站式 Office 应用启动
- 所有 Office 应用一键启动：Word、Excel、PowerPoint、Outlook、OneNote、Access、Publisher、Project、Visio
- 自动识别已安装应用，未安装的应用显示为禁用状态
- 无需复杂菜单，直观的图标界面

### 📂 智能文档聚合管理
- **自动整合三类文档来源：**
  - Office 最近使用文件列表
  - 本地系统缓存
  - Microsoft Edge 浏览历史中的 PDF
- **按访问时间自动排序**，最常用的文档总在前面
- **智能去重**，同一文档不会重复显示
- **自动清理**，已删除文件自动移除列表

### 🎨 现代简洁的界面
- 深色/浅色主题自动适应 Windows 系统设置
- 清晰的左右两栏布局：Office 应用 + 最近文档
- **可调节布局**，拖动分割线自定义两栏宽度
- **文件类型识别**，本地文件和云存储文件显示对应图标
- **悬停提示**，快速查看完整文件路径

### 💾 记住你的偏好
- 窗口大小和位置下次打开时自动恢复
- 列宽设置自动保存
- 无需手动配置，开箱即用

### ⚡ 人性化交互
- 单次点击打开文档（已防止误触双击）
- 支持本地和云端文件无缝访问
- Edge PDF 文档直接从历史记录打开

## 📋 系统需求

- **Windows 10 21H2** 或更新版本
- **.NET 10.0 运行时**（首次运行会提示下载）
- **Microsoft Office** 任意版本

## 🚀 快速开始

### 下载安装（推荐）

1. 进入 [Releases](https://github.com/EricZhang233/MsOfficeHub/releases) 页面
2. 下载最新版本的 `MsOfficeHub_x.x.x.x.zip`
3. 解压到任意文件夹
4. 双击运行 `MsOfficeHub.exe`

**提示：** 无需安装向导，无需注册表修改，解压即用。

### 创建快捷方式（可选）

- 右键点击 `MsOfficeHub.exe` → 发送到 → 桌面（创建快捷方式）
- 或固定到任务栏以便快速访问

## 💡 使用技巧

### 快速启动 Office 应用

在左侧面板看到应用图标？直接点击即可启动！

### 快速打开最近文档

1. 在右侧浏览最近使用的所有文档
2. 点击任意文档名称，自动用对应应用打开
3. 列表按访问时间倒序排列，最常用的在最上面

### 调整窗口布局

- 在左右两栏之间拖动分割线调整宽度
- 关闭应用时自动保存你的布局偏好
- 下次打开时会恢复相同的窗口大小和分割位置

### Edge PDF 管理

- 用 Edge 浏览过的 PDF 文档会自动出现在最近列表
- 需要重新打开之前看过的 PDF？直接点击即可，无需再搜索链接

## 🤝 贡献

欢迎提交 Issue 或 Pull Request！

**贡献指南：**
- 配置持久化使用 Windows Registry
- 测试包含本地和云存储文件场景

## 📄 许可证

详见 [LICENSE](LICENSE) 文件。

---

**作者** [@EricZhang233](https://github.com/EricZhang233)

**反馈与支持** 如有问题或建议，欢迎在 [GitHub Issues](https://github.com/EricZhang233/MsOfficeHub/issues) 提出。
