@echo off
chcp 65001
echo ==========================================
echo   Windows字体管理工具 - 打包脚本
echo ==========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未检测到Python，请先安装Python
    pause
    exit /b 1
)

echo [1/4] 安装依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

if errorlevel 1 (
    echo 错误: 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [2/4] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist font_manager.spec del /q font_manager.spec

echo.
echo [3/4] 开始打包...
pyinstaller --onefile --windowed --name="字体管理工具" --icon=icon.ico ^
    --add-data "README.md;." ^
    --add-data "LOGO.jpg;." ^
    --clean ^
    --noupx ^
    font_manager.py

if errorlevel 1 (
    echo 错误: 打包失败
    pause
    exit /b 1
)

echo.
echo [4/4] 清理临时文件...
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist font_manager.spec del /q font_manager.spec

echo.
echo ==========================================
echo   打包完成！
echo   可执行文件位置: dist\字体管理工具.exe
echo ==========================================
echo.
echo 注意: 运行程序时需要管理员权限才能安装/卸载字体
echo.
pause

