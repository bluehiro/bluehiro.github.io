---
layout:     post
title:      SQL Server Drive Space SuperQuery
date:       2016-02-19
summary:    Query drivespace securely from SQL Server
categories: SQL MSSQL
---

After months of tinkering I finally found a secure and effective way to query SQL Server for drive space INCLUDING drive capacity, while preserving compatibility from 2005 up to 2014! I will be using this code embedded in an SSIS package to retrieve drive space and capacity twice day from the SQL Servers I manage at work.

{% highlight sql lineanchors %}
IF OBJECT_ID('tempdb..#OleBefore') IS NOT NULL
	DROP TABLE #OleBefore

IF OBJECT_ID('tempdb..#OleAfter') IS NOT NULL
	DROP TABLE #OleAfter

IF OBJECT_ID('tempdb..#drives') IS NOT NULL
	DROP TABLE #drives

CREATE TABLE #OleBefore 
(
name VARCHAR(32) NULL,
minimum INT NULL,
maximum INT NULL,
config_value INT NULL,
run_value INT NULL
)

CREATE TABLE #OleAfter 
(
name VARCHAR(32) NULL,
minimum INT NULL,
maximum INT NULL,
config_value INT NULL,
run_value INT NULL
)

--Record status of 'show advanced options'
INSERT INTO #OleBefore (name, minimum, maximum, config_value, run_value)
EXEC sp_configure 'show advanced options'

DECLARE @AdvOptBefore INT
SET @AdvOptBefore = (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'show advanced options')
PRINT 'Ole Adv Opt Before: ' + CONVERT(VARCHAR(1),@AdvOptBefore)
GO

--Enable 'show advanced options'
IF (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'show advanced options') = 0 BEGIN 
	EXEC sp_configure 'show advanced options', 1;
	RECONFIGURE;
	END

--Record status of 'Ole Automation Procedures'
INSERT INTO #OleBefore (name, minimum, maximum, config_value, run_value)
EXEC sp_configure 'Ole Automation Procedures'

DECLARE @OleBefore INT
SET @OleBefore = (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'Ole Automation Procedures')
PRINT 'Ole Auto Before: ' + CONVERT(VARCHAR(1),@OleBefore)

--If Ole Automation is not already enabled, enable it
IF (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'Ole Automation Procedures') = 0 BEGIN 
	EXEC sp_configure 'Ole Automation Procedures', 1;
	RECONFIGURE;
	END

PRINT '*** Retrieve drive space and size ***'
SET NOCOUNT ON
DECLARE @hr INT
DECLARE @fso INT
DECLARE @drive CHAR(1)
DECLARE @odrive INT
DECLARE @TotalSize VARCHAR(20)
DECLARE @MB NUMERIC;
SET @MB = 1048576
CREATE TABLE #drives (
	drive CHAR(1) PRIMARY KEY,
	FreeSpace INT NULL,
	TotalSize INT NULL
)
INSERT #drives (drive, FreeSpace) EXEC
master.dbo.xp_fixeddrives
EXEC @hr = sp_OACreate	'Scripting.FileSystemObject',
						@fso OUT
IF @hr <> 0
	EXEC sp_OAGetErrorInfo @fso
DECLARE dcur CURSOR LOCAL FAST_FORWARD FOR
SELECT
	drive
FROM #drives
ORDER BY drive
OPEN dcur
FETCH NEXT FROM dcur INTO @drive
WHILE @@fetch_status = 0
BEGIN
	EXEC @hr = sp_OAMethod	@fso,
							'GetDrive',
							@odrive OUT,
							@drive
	IF @hr <> 0
		EXEC sp_OAGetErrorInfo @fso
	EXEC @hr =
	sp_OAGetProperty	@odrive,
						'TotalSize',
						@TotalSize OUT
	IF @hr <> 0
		EXEC sp_OAGetErrorInfo @odrive
	UPDATE #drives
	SET TotalSize = @TotalSize / @MB
	WHERE drive = @drive
	FETCH NEXT FROM dcur INTO @drive
END
CLOSE dcur
DEALLOCATE dcur
EXEC @hr = sp_OADestroy @fso
IF @hr <> 0
	EXEC sp_OAGetErrorInfo @fso
SELECT
	@@servername AS servername,
	UPPER(drive) AS drive,
	FreeSpace AS 'Free_MB',
	TotalSize AS 'Drive_MB'
FROM #drives
ORDER BY drive

--If Ole Automation was NOT enabled at the start, disable it
IF (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'Ole Automation Procedures') = 0  BEGIN
	EXEC sp_configure 'Ole Automation Procedures', 0;
	RECONFIGURE;
	END

--Record and display final status for both values
INSERT INTO #OleAfter (name, minimum, maximum, config_value, run_value)
EXEC sp_configure 'Ole Automation Procedures'

--If show advanced options was NOT enabled at the start, disable it
IF (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'show advanced options') = 0  BEGIN
	EXEC sp_configure 'show advanced options', 0;
	RECONFIGURE;
	END

	INSERT INTO #OleAfter (name, minimum, maximum, config_value, run_value)
	EXEC sp_configure 'show advanced options'

DECLARE @OleAfter INT
SET @OleAfter = (SELECT o.config_value FROM #OleAfter o WHERE o.name = 'Ole Automation Procedures')
PRINT 'Ole Auto After: ' + CONVERT(VARCHAR(1),@OleAfter)

DECLARE @AdvOptAfter INT
SET @AdvOptAfter = (SELECT o.config_value FROM #OleBefore o WHERE o.name = 'show advanced options')
PRINT 'Ole Adv Opt After: ' + CONVERT(VARCHAR(1),@AdvOptAfter)
{% endhighlight %}


The secret sauce is really the first and last bits that turn Ole Automation on and off as needed. Most importantly, it preserves whatever setting already existed on the server it is running on. The amount of time spent with Ole Automation enabled is kept to the absolutely minimum, roughly 1 second or less in my testing. Using the xp_drive command doesn't give capacity, and os_volumes only shows drives with databases on them. Both of these are deal breakers for me.