@ Echo Off
Call Home
echo Processing Request Please Wait...
echo **********************************************  
echo.
mysqldump -u root -psamsung ak_inv > Backup.sql
echo.
echo Database Backup File [Backup.sql] Created...
echo.
echo **********************************************
echo Data Downloaded Sucessfully, Thank you...
pause