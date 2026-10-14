# 转换内核目录

`libreoffice.zip` 是构建离线单文件版时嵌入的转换内核，因为体积较大，不放入 Git 历史。

在仓库根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\prepare-engine.ps1
```

脚本会从 LibreOffice 官方镜像下载 26.2.6 x64 安装包、校验 SHA-256、提取便携运行文件并生成 `engine/libreoffice.zip`。
