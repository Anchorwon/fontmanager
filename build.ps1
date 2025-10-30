# Windows字体管理工具 - PowerShell打包脚本
# 编码设置
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Windows字体管理工具 - 打包脚本" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 检查Python
Write-Host "[1/4] 检查Python环境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "检测到: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "错误: 未检测到Python，请先安装Python" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

# 安装依赖
Write-Host "`n[2/4] 安装依赖包..." -ForegroundColor Yellow
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 依赖安装失败" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

# 清理旧文件
Write-Host "`n[3/4] 清理旧的构建文件..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "__pycache__") { Remove-Item -Recurse -Force "__pycache__" }
if (Test-Path "font_manager.spec") { Remove-Item -Force "font_manager.spec" }

# 打包
Write-Host "`n[4/4] 开始打包..." -ForegroundColor Yellow
pyinstaller --onefile --windowed --name="字体管理工具" --icon=NONE `
    --add-data "README.md;." `
    --clean `
    --noupx `
    font_manager.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 打包失败" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

# 清理临时文件
Write-Host "`n清理临时文件..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "__pycache__") { Remove-Item -Recurse -Force "__pycache__" }
if (Test-Path "font_manager.spec") { Remove-Item -Force "font_manager.spec" }

Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "  打包完成！" -ForegroundColor Green
Write-Host "  可执行文件位置: dist\字体管理工具.exe" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "`n注意: 运行程序时需要管理员权限才能安装/卸载字体" -ForegroundColor Yellow
Write-Host ""
Read-Host "按任意键退出"

