@echo off
chcp 65001 >nul
echo ========================================
echo  打包问卷数据处理工具为 EXE 文件
echo ========================================
echo.

REM 检查 PyInstaller 是否已安装
echo [1/3] 检查 PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller -i https://mirrors.cloud.tencent.com/pypi/simple --trusted-host mirrors.cloud.tencent.com
)
echo ✓ PyInstaller 已就绪
echo.

REM 清理旧的打包文件
echo [2/4] 清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo ✓ 清理完成
echo.

REM 打包 EXE
echo [3/4] 开始打包（需要 2-5 分钟）...
pyinstaller --onefile ^
    --add-data "templates;templates" ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import flask ^
    --hidden-import werkzeug ^
    --noconsole ^
    --name "问卷数据处理工具" ^
    app.py

if %errorlevel% neq 0 (
    echo ❌ 打包失败
    pause
    exit /b 1
)

echo ✓ 打包完成
echo.

REM 创建使用说明
echo [4/4] 创建使用说明...
(
echo 问卷数据处理工具 使用说明
echo ========================================
echo.
echo 使用方法：
echo 1. 双击运行 "问卷数据处理工具.exe"
echo 2. 等待程序启动（首次启动可能需要几秒钟）
echo 3. 浏览器会自动打开 http://localhost:5000
echo    如果未自动打开，请手动访问该地址
echo 4. 上传 Excel 文件进行处理
echo 5. 处理完成后会自动下载结果文件
echo.
echo 注意事项：
echo - 首次启动可能需要几秒钟，请耐心等待
echo - 使用时请勿关闭黑色窗口（Flask服务）
echo - 如需停止服务，关闭黑色窗口即可
echo - 支持的文件格式：.xlsx, .xls
echo - 自动删除"跳过"值，整理问卷数据为固定17列
echo.
echo 常见问题：
echo Q: 杀毒软件报警怎么办？
echo A: 这是正常现象，添加白名单即可继续使用
echo.
echo Q: 打包后的文件很大？
echo A: 这是正常的，因为包含了 Python 运行时和所有依赖库
echo.
echo Q: Mac/Linux 能用吗？
echo A: 这个 EXE 只能在 Windows 上运行
echo.
echo 技术信息：
echo - 基于 Flask + pandas + openpyxl 开发
echo - 版本：v1.0
echo - 无需安装 Python 环境
echo.
echo ========================================
) > dist\使用说明.txt

echo ✓ 使用说明已创建
echo.

echo ========================================
echo  ✓ 打包成功！
echo ========================================
echo.
echo 输出文件位于: dist\问卷数据处理工具.exe
echo.
echo 你可以将 dist 文件夹中的所有内容
echo 分发给其他人使用。
echo 对方无需安装 Python，直接双击 exe 文件即可运行。
echo.
pause
