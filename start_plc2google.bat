@echo off
REM 啟動 plc2google Docker 容器腳本
REM 將此腳本添加到 Windows 任務計劃程序的開機啟動項目

echo Starting plc2google container...

cd /d "d:\掌門事業股份有限公司\汐止廠 - 文件\釀造生產部\SourceCode\Other\plc2google"

REM 檢查容器是否存在
docker ps -a --format "{{.Names}}" | findstr "plc2google" >nul 2>&1
if %errorlevel% equ 0 (
    echo Container exists, checking status...
    docker ps --format "{{.Names}}" | findstr "plc2google" >nul 2>&1
    if %errorlevel% equ 0 (
        echo Container is already running.
        exit /b 0
    ) else (
        echo Starting stopped container...
        docker start plc2google
        exit /b 0
    )
)

REM 如果容器不存在，使用 docker-compose 啟動
echo Container not found, starting with docker-compose...
docker-compose up -d

echo plc2google container started successfully.
exit /b 0
