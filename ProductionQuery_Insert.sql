USE [H3_Systematic]
GO

/****** Object:  StoredProcedure [dbo].[ProductionQuery_Insert_New]    Script Date: 05/28/2025 19:55:21 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		Boris Li
-- Create date: 2023/05/02
-- Update Date: 2024/02/22
-- Description:	生產報表UI查詢使用
-- 07/26 修改符合新版本
-- 08/01 將資料庫改為現行資料庫
-- 10/12 修改更新語法
-- 10/30 修改參數需求，改為只需輸入AreaID
-- 02/22 修改機台卡控 : 機台名稱 --> 機台編號
-- =============================================
CREATE PROCEDURE [dbo].[ProductionQuery_Insert_New]
	-- Add the parameters for the stored procedure here
	@AreaID int

	WITH RECOMPILE
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
DECLARE @Proc TABLE (
[ProcName] nvarchar(50),
[Location] nvarchar(50),
[Machine] nvarchar(MAX),
[MachineNo] NVARCHAR(MAX)
)

INSERT INTO @Proc
SELECT [ProcName],ISNULL([Location],''),ISNULL([Machine],''),ISNULL([MachineNo],'')
FROM [H3_Systematic].[dbo].[H3_Proc] WITH(NOLOCK)
WHERE [Pkey] = @AreaID


DECLARE @Temp AS TABLE (
[日期] nvarchar(MAX),
[班別] nvarchar(MAX),
[前站結束時間] nvarchar(MAX),
[開始時間] nvarchar(MAX),
[結束時間] nvarchar(MAX),
[料號] nvarchar(MAX),
[批號] nvarchar(MAX),
[層別] nvarchar(MAX),
[站點] nvarchar(MAX),
[機台] nvarchar(MAX),
[機台編號] nvarchar(MAX),
[產品類型] nvarchar(MAX),
[入料片數] nvarchar(MAX),
[出料片數] nvarchar(MAX),
[過帳人員工號] nvarchar(MAX),
[過帳人員] nvarchar(MAX)
)
INSERT INTO @Temp
SELECT
w.[MoveInTime]
,CASE WHEN CAST(w.[CheckInTime] AS TIME) > '07:20' AND CAST(w.[CheckInTime] AS TIME) < '19:20' THEN 'D'ELSE 'N' END AS [班別]
,w.[前站結束時間]
,w.[CheckInTime]
,w.[CheckOutTime]
,w.[PartNum]
,w.[LotNum]
,w.[LayerName]
,w.[ProcName]
,w.[MachineName]
,w.[MachineNo]
,w.[ITypeName]
,w.[Qnty_In]
,w.[Qnty_Out]
,w.[UserID]
,ISNULL((SELECT TOP(1) RTRIM([EmpName]) FROM [UTCHFACMRPT_REAL].[acme].[dbo].[EmpBas] as id WITH (NOLOCK) WHERE id.[EmpId] = [UserID]),'') AS [過帳人員]
FROM [UTCHFACMRPT_REAL].[DM].[dbo].[ProductReportWip] AS w WITH (NOLOCK)
INNER JOIN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable]((SELECT [ProcName] FROM @Proc),',')) AS pr ON w.[ProcName] LIKE (SUBSTRING(pr.[data],1,3) + '%' + SUBSTRING(pr.[data],4,3) + '%'),
@Proc AS p 
WHERE w.[LineId] = 29
AND (w.[CurrLocation] = p.[Location] or p.[Location] = '')
AND (w.[MachineNo] = '' OR p.[MachineNo] = '' OR w.[MachineNo] IN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable](p.[MachineNo],',')) )

--更新已存在Log的資料
UPDATE [H3_Systematic].[dbo].[H3_ProductionLog]
SET [SystemTime] = [日期],[Class] = [班別],[LastEndTime] = [前站結束時間],[StartTime] = [開始時間],[EndTime] = [結束時間],[Partnum] = [料號],[Lotnum] = [批號],[Layer] = [層別],[ProcName] = [站點],[EQID] = [機台],[MachineNo] = [機台編號],[IType] = [產品類型],[Inputpcs] = [入料片數],[Outputpcs] = [出料片數],[WID] = [過帳人員工號],[User] = [過帳人員]
FROM @Temp AS r
LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionLog] AS T ON T.[AreaID] = @AreaID AND r.[站點] = T.[ProcName] AND r.[批號] = T.[Lotnum] AND r.[層別] = T.[Layer]
WHERE T.[Pkey] IS NOT NULL AND (T.[SystemTime] = '' OR T.[LastEndTime] = '' OR T.[StartTime] = '' OR T.[EndTime] = '' OR T.[Inputpcs] = 0 OR T.[Outputpcs] = 0) AND (T.[Class] = 'D' OR T.[Class] = 'N')
--新增還未有Log的資料
INSERT INTO [H3_Systematic].[dbo].[H3_ProductionLog]
([AreaID],[SystemTime],[Class],[LastEndTime],[StartTime],[EndTime],[Partnum],[Lotnum],[Layer],[ProcName],[EQID],[MachineNo],[IType],[Inputpcs],[Outputpcs],[WID],[User],[Count])
SELECT @AreaID AS [AreaID],[日期],CASE WHEN CAST([開始時間] AS TIME) > '07:20' AND CAST([開始時間] AS TIME) < '19:20' THEN 'D'ELSE 'N' END AS [班別],[前站結束時間],[開始時間],[結束時間],[料號],[批號],[層別],[站點],[機台],[機台編號],[產品類型],[入料片數],[出料片數],[過帳人員工號],[過帳人員],0
FROM
@Temp AS r
LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionLog] AS T ON T.[AreaID] = @AreaID AND r.[站點] = T.[ProcName] AND r.[批號] = T.[Lotnum] AND r.[層別] = T.[Layer]
WHERE T.[Pkey] IS NULL
ORDER BY [日期] 

DECLARE @NonWIP AS TABLE (
[ProcName] nvarchar(MAX),[Lotnum] nvarchar(MAX),[layer] nvarchar(MAX),[Count] nvarchar(MAX)
)
INSERT INTO @NonWIP
SELECT [ProcName],[Lotnum],[layer],[Count]
FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS T WITH (NOLOCK)
LEFT JOIN @Temp AS r ON T.[AreaID] = @AreaID AND r.[站點] = T.[ProcName] AND r.[批號] = T.[Lotnum] AND r.[層別] = T.[Layer]
INNER JOIN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable]((SELECT [ProcName] FROM @Proc),',')) AS pr ON T.[ProcName] LIKE (SUBSTRING(pr.[data],1,3) + '%' + SUBSTRING(pr.[data],5,3) + '%')
WHERE T.[Flag] = 0 AND T.[AreaID] = @AreaID AND r.[日期] IS NULL
AND (T.[SystemTime] = '' OR T.[LastEndTime] = '' OR T.[StartTime] = '' OR T.[EndTime] = '' OR T.[Inputpcs] = 0 OR T.[Outputpcs] = 0) AND (T.[Class] = 'D' OR T.[Class] = 'N')

UPDATE L
SET 
[SystemTime] = [日期],
[Class] = CASE WHEN CAST([開始時間] AS TIME) > '07:20' AND CAST([開始時間] AS TIME) < '19:20' THEN 'D'ELSE 'N' END,
[StartTime] = [開始時間],
[EndTime] = [結束時間],
[EQID] = [機台],
[MachineNo] = [機台編號],
[Inputpcs] = [入料片數],
[Outputpcs] = [出料片數],
[WID] = [過帳人員工號],
[User] = [過帳人員] 
FROM
(
	SELECT 
	MAX(CASE WHEN h.[AftStatus] = 'MoveIn' THEN h.[ChangeTime] ELSE '' END) AS [日期],
	MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN h.[ChangeTime] ELSE '' END) AS [開始時間],
	MAX(CASE WHEN h.[AftStatus] = 'CheckOut' THEN h.[ChangeTime] ELSE '' END) AS [結束時間],
	MAX(ISNULL(m.[MachineName],'')) AS [機台],
    MAX(ISNULL(m.[MachineNo],'')) AS [機台編號],
	MAX(CASE WHEN h.[AftStatus] = 'MoveIn' THEN h.[qnty] ELSE '' END) AS [入料片數],
	MAX(CASE WHEN h.[AftStatus] = 'CheckOut' THEN h.[qnty] ELSE '' END) AS [出料片數],
	MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN h.[UserID] ELSE '' END) AS [過帳人員工號],
    MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN id.[EmpName] ELSE '' END) AS [過帳人員],
	n.[ProcName],
	n.[Lotnum],
	n.[layer]
	FROM @NonWIP AS n
	LEFT JOIN [utchfacmrpt].[report].[dbo].[view_HIST] AS h WITH (NOLOCK) ON h.[ProcName] = n.[ProcName] AND h.[Lotnum] = n.[Lotnum] AND h.[LayerName] = n.[layer]
	LEFT JOIN [utchfacmrpt].[acme].[dbo].[PDL_Machine] AS m WITH (NOLOCK) ON h.[Machine] = m.[MachineId] AND h.[AftStatus] = 'CheckIn'
	LEFT JOIN [UTCHFACMRPT_REAL].[acme].[dbo].[EmpBas] AS id WITH (NOLOCK) ON id.[EmpId] = h.[UserID]
	WHERE h.[Linename] = 'H3'
	GROUP BY n.[Lotnum],n.[ProcName],n.[layer]
) AS r,
[H3_Systematic].[dbo].[H3_ProductionLog] AS L WITH (NOLOCK)
WHERE L.[AreaID] = @AreaID AND r.[ProcName] = L.[ProcName] AND r.[Lotnum] = L.[Lotnum] AND r.[layer] = L.[Layer] AND (L.[Class] = 'D' OR L.[Class] = 'N')

END
GO


