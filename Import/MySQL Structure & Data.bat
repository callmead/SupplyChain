@ echo off
Call Home
echo Processing Backup Request
echo **********************************************  
echo.
echo Preparing Create File...
mysqldump -u root -p --no-data accounts > Create.sql
echo.
echo Table Structures Copied in file [Create.sql]
echo.
echo **********************************************  
echo.
echo Preparing Data File...
echo **********************************************  
echo.
mysqldump -u root -p --no-create-info accounts > Data.sql
echo.
echo Data Copied in File [Data.sql]
echo.
echo **********************************************  
echo.
pause