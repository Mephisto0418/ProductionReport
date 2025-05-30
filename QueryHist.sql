USE [H3_Systematic]
GO
/****** Object:  StoredProcedure [dbo].[H3_ProductuinReport_Hist]    Script Date: 09/22/2023 11:50:44 ******/
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

DECLARE @EQIDOneString varchar(500)
SET @EQIDOneString = REPLACE(@EQID,'''','')

DECLARE @cols VARCHAR(MAX)
SET @cols = N''							   
		SELECT @cols = @cols + iif(@cols = N'',QUOTENAME([ParameterName]),N',' + QUOTENAME([ParameterName]))
		FROM 
		(
			  SELECT b.[ParameterName] FROM (
							   SELECT TOP 100000000 a.[ParameterName], [MDFK], ROW_NUMBER()OVER (PARTITION BY a.[ParameterName] ORDER BY a.[MDFK]) AS FKY
							   FROM (
							   SELECT DISTINCT TOP 100000000 w.[ParameterName],SUBSTRING(p.[ProcName],1,3) + SUBSTRING(p.[ProcName],4,3) AS MDFK FROM
							   [H3_Systematic].[dbo].[H3_Production_ProcParameter] w WITH(NOLOCK)
							   LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] p WITH(NOLOCK) ON w.[AreaID] = P.[Pkey]
							   WHERE p.[ProcName] IN (SELECT [Value] FROM dbo.SplitString(@EQIDOneString,','))
							   ) a
							   ORDER BY CHARINDEX(MDFK,'RLSGRIMDLMDRWPRWPRMDLDBRPTHECUPTHPCUPLGPTRPLGTAPPLGSTNPLGVCPPLSRBFLTHRAPLTHCXPLTHDESLTHHESABFCLNABFBZOABFMECABFABFLDLCOLIMTDSCLTHACDLTHEXPLTHUDILTHDFVPTHFCUPTHSACPTHANLSMKCLNSMKMECSMKPTRSMKHDFSMKEXPSMKUDISMKDEVSMKLUVSMKLAMSMKVPFIMTIMTIMTPREDFMLIDRUTCCD')
							   ) b
							   WHERE b.[FKY] < 2
		) t

DECLARE @sql NVARCHAR(MAX)
SET @sql = N'
SELECT * 
FROM
(SELECT c.[Area] AS 報表名稱
				   ,a.[SystemTime] AS [Movie in時間]
				   ,a.[Class] AS 班別
				   ,a.[LastEndTime] AS 最後編輯時間
				   ,a.[StartTime] AS [Check in時間]
				   ,a.[EndTime] AS [Check out時間]
				   ,a.[Partnum]
				   ,a.[Layer]
				   ,CASE WHEN b.[PressCount] = 0 THEN ''Core'' ELSE ''BU'' + CONVERT(nvarchar,[PressCount]) END AS BU
				   ,a.[Lotnum]
				   ,a.[ProcName] AS 站點
				   ,a.[EQID] AS 機台
				   ,a.[IType]
				   ,a.[Inputpcs]
				   ,a.[Outputpcs]
				   ,a.[WID] AS 工號
				   ,a.[User] AS 姓名
				   ,a.[Remark] AS 備註
				   ,l.[ParameterName], l.[ParameterVaules],l.[PID]
		FROM [H3_Systematic].[dbo].[H3_ProductionLog] a WITH(NOLOCK) 
		LEFT JOIN [utchfacmrpt].[acme].[dbo].[PartInfo_ACME] b WITH(NOLOCK) ON a.[Partnum] = RTRIM(b.[PartNum]) + RTRIM(b.[Revision]) AND RTRIM(a.[Layer]) = RTRIM(b.[LayerName])
		LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] c WITH(NOLOCK) ON a.[AreaID] = c.[Pkey]
		LEFT JOIN (
			SELECT [ParameterName], [ParameterVaules],[PID],[AreaID]
			FROM [H3_Systematic].[dbo].[H3_ProductionParameter] WITH(NOLOCK) 
		) l ON a.[Pkey] = l.[PID] AND a.[AreaID] = l.[AreaID] 
		WHERE [Flag] = 1) r
		PIVOT (
				MAX([ParameterVaules]) 
				FOR [ParameterName] IN ('+@cols+')
		) p '



BEGIN
	IF @QueryMode = 1
		BEGIN		
		SET @sql = @sql +'WHERE [Lotnum] IN ('+ @LotNum +') AND
								SUBSTRING([站點],1,3) + SUBSTRING([站點],5,3) IN (' + @EQID + ') 
						  ORDER BY CHARINDEX(SUBSTRING([站點],1,3) + SUBSTRING([站點],5,3),''RLSGRIMDLMDRWPRWPRMDLDBRPTHECUPTHPCUPLGPTRPLGTAPPLGSTNPLGVCPPLSRBFLTHRAPLTHCXPLTHDESLTHHESABFCLNABFBZOABFMECABFABFLDLCOLIMTDSCLTHACDLTHEXPLTHUDILTHDFVPTHFCUPTHSACPTHANLSMKCLNSMKMECSMKPTRSMKHDFSMKEXPSMKUDISMKDEVSMKLUVSMKLAMSMKVPFIMTIMTIMTPREDFMLIDRUTCCD''), [Movie in時間]' 	
		END
	ELSE IF @QueryMode = 2
		BEGIN
		SET @sql = @sql +'WHERE [Movie in時間] BETWEEN ''' + CONVERT(varchar,@StartTime) + ''' AND ''' + CONVERT(varchar,@EndTime) + ''' 
						  ORDER BY CHARINDEX(SUBSTRING([站點],1,3) + SUBSTRING([站點],5,3),''RLSGRIMDLMDRWPRWPRMDLDBRPTHECUPTHPCUPLGPTRPLGTAPPLGSTNPLGVCPPLSRBFLTHRAPLTHCXPLTHDESLTHHESABFCLNABFBZOABFMECABFABFLDLCOLIMTDSCLTHACDLTHEXPLTHUDILTHDFVPTHFCUPTHSACPTHANLSMKCLNSMKMECSMKPTRSMKHDFSMKEXPSMKUDISMKDEVSMKLUVSMKLAMSMKVPFIMTIMTIMTPREDFMLIDRUTCCD''), [Movie in時間]'
		END
	--SELECT @sql
	EXEC sp_executesql @sql

END
