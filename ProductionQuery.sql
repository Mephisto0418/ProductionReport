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
-- Description:	�Ͳ�����UI 2.0 �d�ߨϥ�
-- 08/01 �קאּ�{���Ʈw
-- 10/20 �ק�����ƧP�w�����W�٧אּ���ID
-- 11/03 �ק�s�W����
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
 [�����ɶ�] datetime,
 [���] datetime,
 [�Z�O] nvarchar(5),
 [�e�������ɶ�] datetime,
 [�}�l�ɶ�] datetime,
 [�����ɶ�] datetime,
 [�Ƹ�] nvarchar(20),
 [�帹] nvarchar(20),
 [�h�O] nvarchar(10),
 [���I] nvarchar(10),
 [���x] nvarchar(50),
 [���~����] nvarchar(10),
 [�J�Ƥ���] int,
 [�X�Ƥ���] int,
 [�L�b�u��] nvarchar(10),
 [�L�b�H��] nvarchar(20),
 [�ާ@��] nvarchar(50),
 [�Ƶ�] nvarchar(100),
 [ParameterName] nvarchar(50),
 [ParameterVaules] nvarchar(MAX),
 [�w�W��] bit,
 [����] nvarchar(10),
 [Count] int
 )
 INSERT INTO @rTEMP													
 SELECT l.[Pkey] AS [LogID],ISNULL(l.[UploadTime],'''') AS [�����ɶ�],l.[SystemTime] AS [���],[Class] AS [�Z�O],ISNULL([LastEndTime],''1900-01-02'') AS [�e�������ɶ�],ISNULL([StartTime],'''') AS [�}�l�ɶ�],[EndTime] AS [�����ɶ�],[Partnum] AS [�Ƹ�],l.[Lotnum] AS [�帹],[Layer] AS [�h�O]
    ,l.[ProcName] AS [���I],[EQID] AS [���x],[IType] AS [���~����],[Inputpcs] AS [�J�Ƥ���],[Outputpcs] AS [�X�Ƥ���],[WID] AS [�L�b�u��],[User] AS [�L�b�H��],ISNULL([OP],'''') AS [�ާ@��],ISNULL([Remark],'''') AS [�Ƶ�]
    ,pa.[ParameterName],ISNULL(p.[ParameterVaules],'''') AS [ParameterVaules],l.[Flag] AS [�w�W��],CASE WHEN pr.[hasFace] = 0 THEN ''N/A'' WHEN p.[Face] = 1 THEN ''PF'' WHEN p.[Face] = 2 THEN ''PB'' ELSE p.[Face] END AS [����],ISNULL(l.[Count],0) AS [Count]
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
	ORDER BY [�w�W��],[LogID]'
    EXEC SP_EXECUTESQL @sql
 
END
ELSE
BEGIN
    SELECT l.[Pkey] AS [LogID],ISNULL(l.[UploadTime],'') AS [�����ɶ�],l.[SystemTime] AS [���],[Class] AS [�Z�O],ISNULL([LastEndTime],'1900-01-02') AS [�e�������ɶ�],ISNULL([StartTime],'') AS [�}�l�ɶ�],[EndTime] AS [�����ɶ�],[Partnum] AS [�Ƹ�],l.[Lotnum] AS [�帹],[Layer] AS [�h�O]
    ,l.[ProcName] AS [���I],[EQID] AS [���x],[IType] AS [���~����],[Inputpcs] AS [�J�Ƥ���],[Outputpcs] AS [�X�Ƥ���],[WID] AS [�L�b�u��],[User] AS [�L�b�H��],ISNULL([OP],'') AS [�ާ@��],ISNULL([Remark],'') AS [�Ƶ�]
    ,p.[ParameterName],ISNULL (p.[ParameterVaules],'') AS [ParameterVaules],l.[Flag] AS [�w�W��],'N/A' AS [����],ISNULL(l.[Count],0) AS [Count]
    FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
    LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID]
    LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[PID] = l.[Pkey]
    WHERE pr.[Pkey] = @AreaID AND (l.[MachineNo] = '' OR ISNULL(pr.[MachineNo],'') LIKE '%' + l.[MachineNo] + '%' OR ISNULL(pr.[MachineNo],'') = '') AND (l.[Flag] = 0 
    OR ([Flag] = 1 AND DATEDIFF(HOUR,ISNULL(l.[UploadTime],GETDATE()) ,GETDATE())<12))
	ORDER BY [�w�W��],[LogID]
END
END
GO


