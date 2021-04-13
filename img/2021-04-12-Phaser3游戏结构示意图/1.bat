:: 修复乱码问题
CHCP 65001
title Office 2019 Retail 转换 Vol 版
echo 米特修改版本 V1.0
echo 该工具用于测试使用！请勿用于商业用途！

:: 判断安装目录
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cls

echo 正在重置 Office 2019 零售激活...
cscript ospp.vbs /rearm

echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 KMS 密钥...
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP

:: 这里填入企业管理员提供给你的 KMS 地址
:: 并清除掉前面的注释符号。默认不处理。

:: echo 正在设置 KMS 服务器...
:: cscript ospp.vbs /sethst:企业管理员提供给你的KMS地址

:: echo 正在联系KMS服务器...
:: cscript ospp.vbs /act

echo 转化完成，按任意键退出！

pause >nul

exit