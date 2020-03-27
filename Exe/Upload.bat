@ echo off
Call Home
echo Ready to Upload Data!
echo **********************************************  
echo.
pause
echo Uploading Data...
echo.
mysql -u root -psamsung ak_inv < C:\Backup.sql
echo.
echo **********************************************
echo Data uploaded Sucessfully, Thank you.
pause
pause