# 部署与打包

## 普通用户

1. 在 GitHub 项目页进入 **Releases**，下载维护者发布的 Windows 压缩包。
2. 解压后双击 `Word转PDF.exe`。
3. 本机需已安装 WPS 文字或 Microsoft Word。工具不会安装或替代这些软件。
4. 选择 Word 文件与 PDF 保存目录，再点击“转换为 PDF”。

不需要 Python、账号或网络服务。旧窗口不会自动刷新成新版，更新后请关闭旧窗口并重新启动 EXE。

## 开发环境

以下命令均在项目根目录执行。使用 Windows 64 位 Python 3.13。

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\word_pdf_app.py
```

`requirements.txt` 固定了本次验证使用的主要包版本；未锁定所有平台相关的间接依赖，因此不承诺不同机器上的二进制逐字节一致。

## 构建独立程序

自动构建：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1
```

或手动运行：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm WordToPDF.spec
```

输出为 `dist\Word转PDF.exe`。打包配置按仓库位置解析路径，不需要修改作者电脑的绝对路径。

生成发布压缩包：

```powershell
.\.venv\Scripts\python.exe scripts\package_release.py --exe "dist\Word转PDF.exe" --output "release-assets\word-to-pdf-windows.zip"
```

发布包包含 EXE、项目说明、对应源码、构建脚本与许可文本。

## 验证

```powershell
.\.venv\Scripts\python.exe scripts\check_project.py
.\.venv\Scripts\python.exe -m compileall -q src
```

项目检查只验证文件、链接和代码语法，不会调用 WPS / Word。请在实际 Windows 电脑上另外验证：

- 普通文字和表格的分页。
- 带内嵌 SVG 的文档，以及图片尺寸和裁切。
- 公式上下标、求和符号和目标设备上的显示。
- 小窗口、最大化、150% / 200% 系统显示缩放。
- 同名结果另存、无效输入路径和没有可用 Office 软件时的提示。

仓库中的 GitHub Actions 提供源码检查，以及手动触发的 Windows 构建。构建任务不包含真实文档转换测试。

## 常见问题

**打不开文字处理程序或提示正在忙**

先确认 WPS / Word 可以手动打开输入文件，并关闭密码、激活、恢复文档等弹窗。稍后重试；本工具已有部分忙碌重试处理。

**PDF 中的截图仍然有像素**

截图本身是位图。只有文档中仍保存原始 SVG 等矢量信息时，才能尝试恢复矢量清晰度。

**恢复数量为零**

文档可能没有内嵌 SVG、图片无法安全定位，或原导出内容已经是矢量。请检查输出与日志，不要仅依据计数判断清晰度。

**公式无法正常复制**

兼容模式会将数学字形转成轮廓。优先需要复制或搜索时，可以取消兼容模式，并在最终阅读器中检查字体显示。

**默认保存位置**

默认保存到当前用户的“下载”目录。也可以在界面中把保存位置更改为其他可写目录，例如自己的“文档”文件夹。
