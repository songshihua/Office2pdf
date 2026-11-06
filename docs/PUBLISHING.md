# 上传源码与发布程序

这个目录已经整理为 GitHub 项目结构。**本次整理未执行 `git init`，没有创建或推送远程仓库**；若本地目录已有 `.git`，保持原样。

## 第一次上传

1. 在 GitHub 新建你自己的仓库，名称可使用 `word-to-pdf`。
2. 将干净源码包解压，上传其中的 README、源码、脚本、文档、图片、许可和 `.github` 等目录。
3. 不要把本机的 EXE、转换日志、完整报告、生成 PDF、缓存和旧版备份作为源码上传。
4. 若用本地 Git 提交，现成的 `.gitignore` 会排除这些本地文件；**浏览器手动上传不会自动替你执行 `.gitignore` 的筛选**。

源码包中不包含 `.git`。这里也不要求先执行仓库初始化命令。

## 为什么 EXE 单独发布

当前 EXE 大约 84 MiB。GitHub 浏览器上传单文件上限为 25 MiB，普通 Git 对超过 50 MiB 的文件会提示警告，超过 100 MiB 的文件会阻止提交。程序二进制放到 Releases，更适合下载和更新。

参考：[GitHub 大文件说明](https://docs.github.com/en/repositories/working-with-files/managing-large-files/about-large-files-on-github)。

## 创建一个 Release

1. 上传源码后，进入仓库的 **Releases → Draft a new release**。
2. 选择版本标签，例如 `v0.1.0`，填写更新说明。
3. 在本机运行 `scripts/package_release.py` 准备发布压缩包。
4. 将 `release-assets` 中的发布 ZIP 作为附件上传。对应源码、构建脚本和许可文本应与这份 EXE 一起提供。
5. 确认目标设备上的测试结果后，再自行点击发布。

参考：[GitHub Release 管理说明](https://docs.github.com/en/repositories/releasing-projects-on-github/managing-releases-in-a-repository)。

本次整理没有替你创建 GitHub 仓库、上传文件或发布 Release。

## 启用问题反馈

- 在仓库 **Settings → General → Features** 中保留 **Issues** 开启。
- 提交 `.github/ISSUE_TEMPLATE` 后，新问题页面会显示“转换问题”和“功能建议”表单。
- README 使用仓库内相对链接，不需要先填入作者的 GitHub 用户名。
- 如果你之后调整了 README 所在位置，记得检查 Issues 相对链接。

## 手动构建工作流

仓库 **Actions → Build Windows executable → Run workflow** 可以触发打包。运行成功后在该次任务的 Artifacts 中下载发布 ZIP。

此工作流只构建文件，不自动发布 Release，不会安装 WPS / Word，也不能证明所有文档都能正确转换。
