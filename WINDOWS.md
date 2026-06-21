# Windows 使用说明

## 一、准备

1. 把整个工具文件夹完整复制或解压到 Windows 电脑。
2. 不要在压缩包预览窗口里直接双击运行，请先解压。
3. 保留这些文件和文件夹：`app.py`、`static/`、`scripts/`、所有 `.pptx` 模板、`start.bat`、`install-windows.bat`。

## 二、首次安装

双击 `install-windows.bat`。

它会自动做这些事：

1. 检查 Python 3.9 或更高版本。
2. 如果没有 Python，尝试用 Windows 自带的 `winget` 安装 Python 3.12。
3. 如果公司电脑不允许自动安装，会打开 Python 官方下载页。

手动安装 Python 时，一定勾选 `Add Python to PATH`。安装完成后重新双击 `start.bat`。

## 三、启动工具

已安装 Python 后，双击 `start.bat`。

启动成功后浏览器会自动打开，例如：

```text
http://127.0.0.1:8765
```

如果端口被占用，工具会自动顺延到后面的端口。

## 四、生成图片

Windows 上：

- 生成 PPTX：只需要 Python。
- 自动导出 PNG：需要桌面版 Microsoft PowerPoint。

如果没有安装 PowerPoint，页面会提示 PNG 暂未导出，但 PPTX 会保留，可以用 PowerPoint 手动打开并导出图片。

## 五、共享给其他人使用

在一台已能正常启动的 Windows 电脑上双击 `start-share.bat`。

窗口会显示类似：

```text
http://192.168.1.20:8765
```

把这个地址发给同一 Wi-Fi 或公司内网的同事即可。共享模式下，生成任务由这台电脑统一完成。

## 六、启动失败排查

双击 `windows-diagnose.bat`。

它会检查：

- Python 是否可用
- Python 版本是否满足 3.9+
- 模板是否能读取
- Windows PowerShell 是否可用
- PowerPoint 导出脚本是否存在

诊断和启动日志保存在：

```text
windows-start.log
```

如果仍然失败，把窗口报错和 `windows-start.log` 发给维护人员。
