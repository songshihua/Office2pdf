# 第三方组件

本项目源码采用 GNU AGPL v3。依赖组件保留自己的版权与许可。免费使用与开源分发均应遵守对应条款，不将本项目误标为不受限制的 MIT 项目。

| 组件 | 本次固定版本 | 主要许可 | 上游源码与说明 |
| --- | --- | --- | --- |
| PyMuPDF / MuPDF | 1.27.2.3 | AGPL-3.0 或另行取得的商业许可 | [源码](https://github.com/pymupdf/PyMuPDF)、[许可说明](https://pymupdf.readthedocs.io/en/latest/about.html) |
| PyQt5 | 5.15.10 | GPL v3 或另行取得的商业许可 | [上游与许可说明](https://www.riverbankcomputing.com/software/pyqt/) |
| Qt | 随 PyQt5 安装 | 依所用模块适用 LGPL / GPL / 商业许可 | [Qt 源码与许可](https://www.qt.io/licensing/open-source-lgpl-obligations) |
| fontTools | 4.55.3 | MIT | [源码](https://github.com/fonttools/fonttools)、[许可文本](licenses/fonttools-LICENSE.txt) |
| pypdf | 6.21.4 | BSD-3-Clause | [源码](https://github.com/py-pdf/pypdf)、[许可文本](licenses/pypdf-LICENSE.txt) |
| PyInstaller | 6.22.3，构建时使用 | GPL-2.0 或更高版本，带引导程序例外 | [源码与说明](https://github.com/pyinstaller/pyinstaller) |
| CPython | 3.13，EXE 运行时 | PSF License | [源码](https://github.com/python/cpython)、[许可](https://docs.python.org/3/license.html) |

本目录附带 [AGPL-3.0](LICENSE)、[GPL-3.0](licenses/GPL-3.0.txt) 与主要小型依赖的许可文本。重新打包时还需保留所使用发行包中的第三方声明及适用许可；具体间接依赖和 Qt 组件取决于构建环境。

WPS、Microsoft Word 和 Windows 字体由用户自行安装，均不包含在源码或发布包中，其权利与使用条件由各自供应商规定。

页面中的对比截图由项目发起者提供，用于说明实际转换效果。截图示例不构成对所有文档无损转换的保证。
