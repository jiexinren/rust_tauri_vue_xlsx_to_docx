@echo off
echo 将命令行窗口的字符编码设置为UTF-8,避免中文等字符显示乱码问题
chcp 65001 >nul
echo 正在添加所有更改到暂存区...
git add .
echo 正在提交更改...
git commit -m "日常更新"
echo 正在推送到远程仓库...
git push origin master
echo 推送完成。
pause