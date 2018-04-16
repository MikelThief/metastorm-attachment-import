USE [Metastorm]
GO

/****** Object:  UserDefinedFunction [dbo].[WriteToFile]    Script Date: 19/03/2018 12:49:26 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE FUNCTION [dbo].[WriteToFile](@path [nvarchar](max), @efolderid [nvarchar](max), @filename [nvarchar](max))
RETURNS [nvarchar](max) WITH EXECUTE AS CALLER
AS 
EXTERNAL NAME [ClassLibrary1].[CLR.UserDefinedFunctions].[WriteToFile]
GO

