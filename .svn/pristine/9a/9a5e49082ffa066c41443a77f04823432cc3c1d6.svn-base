USE [H3_Systematic]
GO
/****** Object:  StoredProcedure [dbo].[H3_ProductuinReport_Hist]    Script Date: 06/09/2023 11:48:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		LeoXu
-- Create date: 2023/06/06
-- Description:	現場生產報表查詢
-- =============================================
ALTER PROCEDURE [dbo].[H3_ProductuinReport_Hist]
	-- Add the parameters for the stored procedure here
	@QueryMode int, -- 1: Query by lotnum. 2: Query by time, only one section
	@LotNum varchar(500),
	@EQID varchar(500),
	@StartTime datetime,
	@EndTime datetime
AS
DECLARE @cols VARCHAR(MAX)= N''
		SELECT @cols = @cols + iif(@cols = N'',QUOTENAME([ParameterName]),N',' + QUOTENAME([ParameterName]))
		FROM 
		(
			  SELECT b.[ParameterName] FROM (
							   SELECT TOP 100000000 a.[ParameterName], [MDFK], ROW_NUMBER()OVER (PARTITION BY a.[ParameterName] ORDER BY a.[MDFK]) AS FKY
							   FROM (
							   SELECT DISTINCT TOP 100000000 [ParameterName],SUBSTRING([ProcName],1,3) + SUBSTRING([ProcName],5,3) AS MDFK FROM
							   [H3_Systematic].[dbo].[H3_ProductionParameter] WITH(NOLOCK)) a
							   ORDER BY CHARINDEX(MDFK,'WPRWPRMDLPOSMDLMDLMDLDBRPTHECUPTHPCUPLGANLPLGPTRPLGSPSPLGVCPPLGSTNPLGPOSPLSRBF')) b
							   WHERE b.[FKY] < 2
		) t
DECLARE @sql NVARCHAR(MAX)
SET @sql = N'SELECT a.[SystemTime]
				   ,a.[Class]
				   ,a.[LastEndTime]
				   ,a.[StartTime]
				   ,a.[EndTime]
				   ,a.[Partnum]
				   ,a.[Layer]
				   ,CASE WHEN b.[PressCount] = 0 THEN ''Core'' ELSE ''BU'' + CONVERT(nvarchar,[PressCount]) END AS B你媽
				   ,a.[Lotnum]
				   ,a.[ProcName]
				   ,a.[EQID]
				   ,a.[IType]
				   ,a.[Inputpcs]
				   ,a.[Outputpcs]
				   ,a.[WID]
				   ,a.[User]
				   ,a.[Remark] AS 備註
				   ,l.*
		FROM [H3_Systematic].[dbo].[H3_ProductionLog] a WITH(NOLOCK) 
		LEFT JOIN [utchfacmrpt].[acme].[dbo].[PartInfo_ACME] b WITH(NOLOCK) ON a.[Partnum] = RTRIM(b.[PartNum]) + RTRIM(b.[Revision]) AND RTRIM(a.[Layer]) = RTRIM(b.[LayerName])
		LEFT JOIN (SELECT *
		FROM (
			SELECT [ParameterName], [ParameterVaules],[PID]
			FROM [H3_Systematic].[dbo].[H3_ProductionParameter] WITH(NOLOCK) 
		) t 
		PIVOT (
				MAX([ParameterVaules]) 
				FOR [ParameterName] IN ('+@cols+')
		) p) l ON a.[Pkey] = l.[PID] '

BEGIN
	IF @QueryMode = 1
		BEGIN		
		SET @sql = @sql +'WHERE a.[Lotnum] IN ('+ @LotNum +') AND
								SUBSTRING(a.[ProcName],1,3) + SUBSTRING(a.[ProcName],5,3) IN (' + @EQID + ') 
						  ORDER BY CHARINDEX(SUBSTRING(a.[ProcName],1,3) + SUBSTRING(a.[ProcName],5,3),''WPRWPRMDLPOSMDLMDLMDLDBRPTHECUPTHPCUPLGANLPLGPTRPLGSPSPLGVCPPLGSTNPLGPOSPLSRBF''), a.[SystemTime]' 	
		END
	ELSE IF @QueryMode = 2
		BEGIN
		SET @sql = @sql +'WHERE a.[SystemTime] BETWEEN ''' + CONVERT(varchar,@StartTime) + ''' AND ''' + CONVERT(varchar,@EndTime) + '''
						  ORDER BY CHARINDEX(SUBSTRING(a.[ProcName],1,3) + SUBSTRING(a.[ProcName],5,3),''WPRWPRMDLPOSMDLMDLMDLDBRPTHECUPTHPCUPLGANLPLGPTRPLGSPSPLGVCPPLGSTNPLGPOSPLSRBF''), a.[SystemTime]'
		END
	--SELECT @sql
	EXEC sp_executesql @sql
END
