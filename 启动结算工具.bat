@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist node_modules (
  echo 正在安装依赖，请稍候...
  call npm.cmd install
)
start http://localhost:3000
call npm.cmd run dev
