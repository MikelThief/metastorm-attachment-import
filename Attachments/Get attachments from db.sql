DECLARE @outPutPath varchar(100)
, @i bigint
, @econtents varbinary(max) 
, @fPath varchar(max)  
, @efolderid nvarchar(max)
, @filename nvarchar(max)
, @ekey nvarchar(max)
, @clientName nvarchar(max)
, @tblname nvarchar(max)
, @debug int
, @sql1 nvarchar(max)
, @sql1results nvarchar(max)
, @howmany int
, @procedureName nvarchar(31)
, @result nvarchar(255)
, @efoldername nvarchar(31)

--Set important variables
set @debug = 1 --set to 0 to turn off debug messages
SET @outPutPath = 'C:\Metastorm' --Location to store files
SET @howmany = 10 --how many documents to try and save

--Create temp table and insert the document records
DECLARE @Doctable TABLE (id bigint identity(1,1), ekey  nvarchar(250) , esize int, econtents image)
INSERT INTO @Doctable([ekey], [esize], [econtents])
SELECT TOP (@howmany) ekey, esize, econtents  FROM eattachment where ekey not like '3%' and ekey not like '2%' and ekey not like '1%' and ekey not like '%.js'  and ekey not like '%.xml'  ORDER BY newid();
SELECT @i = 1
WHILE @i <= @howmany
BEGIN 

	SET @ekey = (SELECT ekey from @doctable where id = @i)
    SET @efolderid = (SELECT replace(replace(replace(RTRIM(LTRIM(STUFF(LEFT(@ekey,33),1,1,''))), CHAR(13), ''), CHAR(10), ''),char(9),'') from @doctable where id = @i)
    SET @econtents = (select econtents from @doctable where id = @i)
    SET @procedureName = (SELECT [eMapName] from efolder where efolderid = @efolderid)
    SET @efoldername = (SELECT [efoldername] from efolder where efolderid = @efolderid)

	IF @debug = 1 PRINT N'Loop #' + CONVERT(nchar(8), @i) + CHAR(13)
	+ 'folder: ' + @efolderid + CHAR(13)
	+ 'procedure: ' + @procedurename + CHAR(13)
	+ 'foldername: ' + @efoldername + CHAR(13)
    IF EXISTS (SELECT efolderid from [MetastormAudit].[dbo].[auditLog] where efolderid = @efolderid)
        BEGIN
        --Found this efolderid in the audit table. Decide what to do
        --The next part will look at the audit table and check what the outcome was last time, if there was one
            DECLARE @alreadySaved nvarchar(max)
            SET @alreadySaved = (SELECT [outcome] FROM [MetastormAudit].[dbo].[auditLog] where efolderid = @efolderid)
            --File already saved
            IF @alreadySaved = 'Wrote file'
                BEGIN
                    IF @debug = 1 PRINT 'Already saved, skipping. Loop number ' + CONVERT(nchar(8), @i)
                    SET @result = 'Already saved'
                    GOTO SAVEDSKIP
                END
            --Not client opening procedure
            IF @alreadySaved = 'Skipped - wrong procedure'
                BEGIN
                    IF @debug = 1 PRINT 'Not a client opening, seen before.. skipping. Loop number ' + CONVERT(nchar(8), @i)
                    SET @result = 'Wrong procedure (seen before)'
                    GOTO SAVEDSKIP
                END
            --Exception happened last time
            IF @alreadySaved like '%exception%'
                BEGIN  
                    IF @debug = 1 PRINT 'Error last time we tried to save. Skipping, already logged in audit table. Loop number ' + CONVERT(nchar(8), @i)
                    SET @result = 'Skipping due to previous error or exception'
                    GOTO SAVEDSKIP
                END
        END
        ELSE 
            GOTO NOTLOGGEDYET
    NOTLOGGEDYET: --skip to here if there is no audit log, ignores all the checks above

	IF @debug = 1 PRINT '@procedureName: ' + @procedureName + CHAR(13)
	IF @procedureName not like 'TK_ClientOpening' or @efoldername not like 'CI%'
		BEGIN
			IF @debug = 1 PRINT 'Skipping number ' + CONVERT(nchar(8), @i) + ' - not client opening procedure!'
            SET @result = 'Skipped - wrong procedure'
			GOTO WRONGPROCSKIP
		END
	SET @filename = RTRIM(LTRIM(SUBSTRING(@ekey, 35, LEN(@ekey))))
    SELECT @filename = REPLACE(@filename, '''', '')
        IF @debug = 1 PRINT '@ekey: ' + @ekey + CHAR(13) + '@efolderid: ' + @efolderid + CHAR(13) + '@filename: ' + @filename + CHAR(13) + '@efoldername: ' + @efoldername + CHAR(13)
    
    --Get the relevant process table so we can pull out the client name
    SET @tblname = (select emapname from [Metastorm].[dbo].[eFolder] where efolderid=@efolderid)
        IF @debug = 1 PRINT '@tblname is ' + @tblname + CHAR(13)
        
    SELECT @sql1 =' SELECT @clientname = ([txtclientname]) FROM ' + @tblname + ' WHERE efolderid = ''' + @efolderid + ''''
    --SELECT @sql1 = REPLACE(@sql1, '''', '') --incase the client name has an apostrophe. 
    --Removed, because it causes an arithmetic overflow exception for some reason!!!

    --Set the clientname into @clientname, from the dynamically generated SQL above
    EXECUTE sp_executesql @sql1, N'@efolderid nvarchar(max),@clientname nvarchar(max) OUTPUT', @efolderid = @efolderid, @clientname=@clientname OUTPUT
        IF @debug = 1 PRINT '@clientname: ' + @clientname + CHAR(13)
 
    --Create folder paths from @outputpath set above
	SELECT 
    	 @fPath = @outPutPath + '\'+ @clientname + '\' + @efolderName + '\' + @filename
	FROM @Doctable WHERE id = @i
	IF @debug = 1 PRINT '@fpath: ' + @fpath + CHAR(13) 

    --Attempt to write file
    IF @debug = 1 print 'Attempting to create: ' + @fpath
    select @result = dbo.WriteToFile(@fPath, @efolderid, @filename)
    SELECT CASE @result
        WHEN 'SUCCESS' THEN 'Document Generated at - ' +  @fPath  
    END
    IF @debug = 1 PRINT '@result: ' + @result + CHAR(13) 

    WRONGPROCSKIP: --skipped to here because it's not a client opening.

    --Noticed some duplicate PK errors, not sure why/how that could happen but this block should stop it
    IF NOT EXISTS  (SELECT EFOLDERID from [MetastormAudit].[dbo].[auditLog] where efolderid = @efolderid)
        BEGIN
            --Update log table
            INSERT INTO [MetastormAudit].[dbo].[auditLog]
                ([efolderid] 
                ,[eKey]
                ,[dateadded]
                ,[eProcedureName]
                ,[eFolderName]
                ,[filePath]
                ,[fileName]
                ,[clientName]
                ,[outcome]
                ,[econtents])
            VALUES
                (@efolderID,
                @ekey,
                getdate(),
                @procedureName,
                @efoldername,
                @fpath,
                @filename,
                @clientName,
                @result,
                @econtents)
        END
    ELSE
        BEGIN
            UPDATE [MetastormAudit].[dbo].[auditLog]
            SET 
            [outcome] = @result
            ,[dateadded] = GETDATE() 
            WHERE [efolderid] = @efolderid
        END

--skip to here if already saved or if it's already been identified as one that is the wrong procedure, or if previously had an error/exception. 
--This is below the insert into the audit log so that we don't get duplicates
SAVEDSKIP:            
    --Reset the variables for next use
    SELECT @econtents = NULL  
            ,@fPath = NULL  
            ,@sql1 = NULL
            ,@clientName = NULL
            ,@filename = NULL
            ,@efoldername = NULL
            ,@tblname = NULL
            ,@procedureName = NULL
            ,@result = NULL
            ,@sql1results = NULL
            ,@efolderid = NULL
            ,@ekey = NULL
    SET @i = @i + 1
--End of loop

END 