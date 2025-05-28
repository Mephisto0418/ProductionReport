USE [H3_Systematic]
GO

/****** Object:  StoredProcedure [dbo].[ProductionQuery]    Script Date: 05/28/2025 19:54:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		Boris Li
-- Create date: 2023/07/28
-- Update Date: 2023/08/01
-- Description:	生產報表UI 2.0 查詢使用
-- 08/01 修改為現行資料庫
-- 10/20 修改欄位資料判定由欄位名稱改為欄位ID
-- 11/03 修改新增面次
-- =============================================
CREATE PROCEDURE [dbo].[ProductionQuery]
	-- Add the parameters for the stored procedure here
	@AreaID nvarchar(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here

DECLARE @Temp AS TABLE(
[Sort] nvarchar(MAX),
[ParameterName]nvarchar(MAX)
)

INSERT INTO @Temp
SELECT [PID] AS [Sort]
,[ParameterName]
FROM [H3_Systematic].[dbo].[H3_Production_ProcParameter] WITH (NOLOCK)
WHERE [AreaID] = @AreaID
ORDER BY [Sort]


DECLARE @cols NVARCHAR(MAX)= N''
SELECT @cols = @cols + iif(@cols = N'',QUOTENAME([ParameterName]),N',' + QUOTENAME([ParameterName]))
FROM(SELECT [ParameterName] FROM @Temp) as r


IF @@ROWCOUNT >0
BEGIN
    DECLARE @sql NVARCHAR(MAX)
    SET @sql =N'
     DECLARE @rTEMP TABLE(
 [LogID] int,
 [完成時間] datetime,
 [日期] datetime,
 [班別] nvarchar(5),
 [前站結束時間] datetime,
 [開始時間] datetime,
 [結束時間] datetime,
 [料號] nvarchar(20),
 [批號] nvarchar(20),
 [層別] nvarchar(10),
 [站點] nvarchar(10),
 [機台] nvarchar(50),
 [產品類型] nvarchar(10),
 [入料片數] int,
 [出料片數] int,
 [過帳工號] nvarchar(10),
 [過帳人員] nvarchar(20),
 [操作員] nvarchar(50),
 [備註] nvarchar(100),
 [ParameterName] nvarchar(50),
 [ParameterVaules] nvarchar(MAX),
 [已上傳] bit,
 [面次] nvarchar(10),
 [Count] int
 )
 INSERT INTO @rTEMP													
 SELECT l.[Pkey] AS [LogID],ISNULL(l.[UploadTime],'''') AS [完成時間],l.[SystemTime] AS [日期],[Class] AS [班別],ISNULL([LastEndTime],''1900-01-02'') AS [前站結束時間],ISNULL([StartTime],'''') AS [開始時間],[EndTime] AS [結束時間],[Partnum] AS [料號],l.[Lotnum] AS [批號],[Layer] AS [層別]
    ,l.[ProcName] AS [站點],[EQID] AS [機台],[IType] AS [產品類型],[Inputpcs] AS [入料片數],[Outputpcs] AS [出料片數],[WID] AS [過帳工號],[User] AS [過帳人員],ISNULL([OP],'''') AS [操作員],ISNULL([Remark],'''') AS [備註]
    ,pa.[ParameterName],ISNULL(p.[ParameterVaules],'''') AS [ParameterVaules],l.[Flag] AS [已上傳],CASE WHEN pr.[hasFace] = 0 THEN ''N/A'' WHEN p.[Face] = 1 THEN ''PF'' WHEN p.[Face] = 2 THEN ''PB'' ELSE p.[Face] END AS [面次],ISNULL(l.[Count],0) AS [Count]
	FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
    LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID]
    INNER JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[PID] = l.[Pkey]
	LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter] AS pa WITH(NOLOCK) ON p.[CID] = pa.[Pkey]
    WHERE pr.[Pkey] = ' + @AreaID + ' AND (l.[MachineNo] = '''' OR ISNULL(pr.[MachineNo],'''') LIKE ''%'' + l.[MachineNo] + ''%'' OR ISNULL(pr.[MachineNo],'''') = '''') AND (l.[Flag] = 0 
    OR ([Flag] = 1 AND DATEDIFF(HOUR,ISNULL(l.[UploadTime],GETDATE()) ,GETDATE())<12))

Select *
    FROM @rTEMP
    PIVOT(
    	MAX([ParameterVaules])
    	FOR [ParameterName] IN(' + @cols + ')
    	
    	)p
	ORDER BY [已上傳],[LogID]'
    EXEC SP_EXECUTESQL @sql
 
END
ELSE
BEGIN
    SELECT l.[Pkey] AS [LogID],ISNULL(l.[UploadTime],'') AS [完成時間],l.[SystemTime] AS [日期],[Class] AS [班別],ISNULL([LastEndTime],'1900-01-02') AS [前站結束時間],ISNULL([StartTime],'') AS [開始時間],[EndTime] AS [結束時間],[Partnum] AS [料號],l.[Lotnum] AS [批號],[Layer] AS [層別]
    ,l.[ProcName] AS [站點],[EQID] AS [機台],[IType] AS [產品類型],[Inputpcs] AS [入料片數],[Outputpcs] AS [出料片數],[WID] AS [過帳工號],[User] AS [過帳人員],ISNULL([OP],'') AS [操作員],ISNULL([Remark],'') AS [備註]
    ,p.[ParameterName],ISNULL (p.[ParameterVaules],'') AS [ParameterVaules],l.[Flag] AS [已上傳],'N/A' AS [面次],ISNULL(l.[Count],0) AS [Count]
    FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
    LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID]
    LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[PID] = l.[Pkey]
    WHERE pr.[Pkey] = @AreaID AND (l.[MachineNo] = '' OR ISNULL(pr.[MachineNo],'') LIKE '%' + l.[MachineNo] + '%' OR ISNULL(pr.[MachineNo],'') = '') AND (l.[Flag] = 0 
    OR ([Flag] = 1 AND DATEDIFF(HOUR,ISNULL(l.[UploadTime],GETDATE()) ,GETDATE())<12))
	ORDER BY [已上傳],[LogID]
END
END
GO


