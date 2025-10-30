@echo off
chcp 65001
title Windows字体管理工具

echo ==========================================
echo   Windows字体管理工具
echo ==========================================
echo.
echo 正在启动程序...
echo.
echo 注意: 需要管理员权限才能安装/卸载字体
echo 如果没有权限，请右键此脚本选择"以管理员身份运行"
echo.

python font_manager.py

if errorlevel 1 (
    echo.
    echo 程序运行出错！
    pause
)

