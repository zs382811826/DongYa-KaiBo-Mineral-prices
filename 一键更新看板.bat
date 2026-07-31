@echo off
cd /d D:\Excel处理
echo ========================================
echo    矿产行情看板 - 一键更新并发布
echo ========================================
echo.
echo [1/3] 生成看板...
python build_dashboard.py
echo.
echo [2/3] 提交到 Git...
git add index.html build_dashboard.py .gitignore
git commit -m "update dashboard"
echo.
echo [3/3] 推送到 GitHub...
git push origin main
echo.
echo 完成！稍后刷新网页即可看到更新
pause
