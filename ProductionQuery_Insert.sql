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
-- Description:	�Ͳ�����UI�d�ߨϥ�
-- 07/26 �ק�ŦX�s����
-- 08/01 �N��Ʈw�אּ�{���Ʈw
-- 10/12 �ק��s�y�k
-- 10/30 �ק�ѼƻݨD�A�אּ�u�ݿ�JAreaID
-- 02/22 �ק���x�d�� : ���x�W�� --> ���x�s��
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
[���] nvarchar(MAX),
[�Z�O] nvarchar(MAX),
[�e�������ɶ�] nvarchar(MAX),
[�}�l�ɶ�] nvarchar(MAX),
[�����ɶ�] nvarchar(MAX),
[�Ƹ�] nvarchar(MAX),
[�帹] nvarchar(MAX),
[�h�O] nvarchar(MAX),
[���I] nvarchar(MAX),
[���x] nvarchar(MAX),
[���x�s��] nvarchar(MAX),
[���~����] nvarchar(MAX),
[�J�Ƥ���] nvarchar(MAX),
[�X�Ƥ���] nvarchar(MAX),
[�L�b�H���u��] nvarchar(MAX),
[�L�b�H��] nvarchar(MAX)
)
INSERT INTO @Temp
SELECT
w.[MoveInTime]
,CASE WHEN CAST(w.[CheckInTime] AS TIME) > '07:20' AND CAST(w.[CheckInTime] AS TIME) < '19:20' THEN 'D'ELSE 'N' END AS [�Z�O]
,w.[�e�������ɶ�]
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
,ISNULL((SELECT TOP(1) RTRIM([EmpName]) FROM [UTCHFACMRPT_REAL].[acme].[dbo].[EmpBas] as id WITH (NOLOCK) WHERE id.[EmpId] = [UserID]),'') AS [�L�b�H��]
FROM [UTCHFACMRPT_REAL].[DM].[dbo].[ProductReportWip] AS w WITH (NOLOCK)
INNER JOIN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable]((SELECT [ProcName] FROM @Proc),',')) AS pr ON w.[ProcName] LIKE (SUBSTRING(pr.[data],1,3) + '%' + SUBSTRING(pr.[data],4,3) + '%'),
@Proc AS p 
WHERE w.[LineId] = 29
AND (w.[CurrLocation] = p.[Location] or p.[Location] = '')
AND (w.[MachineNo] = '' OR p.[MachineNo] = '' OR w.[MachineNo] IN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable](p.[MachineNo],',')) )

--��s�w�s�bLog�����
UPDATE [H3_Systematic].[dbo].[H3_ProductionLog]
SET [SystemTime] = [���],[Class] = [�Z�O],[LastEndTime] = [�e�������ɶ�],[StartTime] = [�}�l�ɶ�],[EndTime] = [�����ɶ�],[Partnum] = [�Ƹ�],[Lotnum] = [�帹],[Layer] = [�h�O],[ProcName] = [���I],[EQID] = [���x],[MachineNo] = [���x�s��],[IType] = [���~����],[Inputpcs] = [�J�Ƥ���],[Outputpcs] = [�X�Ƥ���],[WID] = [�L�b�H���u��],[User] = [�L�b�H��]
FROM @Temp AS r
LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionLog] AS T ON T.[AreaID] = @AreaID AND r.[���I] = T.[ProcName] AND r.[�帹] = T.[Lotnum] AND r.[�h�O] = T.[Layer]
WHERE T.[Pkey] IS NOT NULL AND (T.[SystemTime] = '' OR T.[LastEndTime] = '' OR T.[StartTime] = '' OR T.[EndTime] = '' OR T.[Inputpcs] = 0 OR T.[Outputpcs] = 0) AND (T.[Class] = 'D' OR T.[Class] = 'N')
--�s�W�٥���Log�����
INSERT INTO [H3_Systematic].[dbo].[H3_ProductionLog]
([AreaID],[SystemTime],[Class],[LastEndTime],[StartTime],[EndTime],[Partnum],[Lotnum],[Layer],[ProcName],[EQID],[MachineNo],[IType],[Inputpcs],[Outputpcs],[WID],[User],[Count])
SELECT @AreaID AS [AreaID],[���],CASE WHEN CAST([�}�l�ɶ�] AS TIME) > '07:20' AND CAST([�}�l�ɶ�] AS TIME) < '19:20' THEN 'D'ELSE 'N' END AS [�Z�O],[�e�������ɶ�],[�}�l�ɶ�],[�����ɶ�],[�Ƹ�],[�帹],[�h�O],[���I],[���x],[���x�s��],[���~����],[�J�Ƥ���],[�X�Ƥ���],[�L�b�H���u��],[�L�b�H��],0
FROM
@Temp AS r
LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionLog] AS T ON T.[AreaID] = @AreaID AND r.[���I] = T.[ProcName] AND r.[�帹] = T.[Lotnum] AND r.[�h�O] = T.[Layer]
WHERE T.[Pkey] IS NULL
ORDER BY [���] 

DECLARE @NonWIP AS TABLE (
[ProcName] nvarchar(MAX),[Lotnum] nvarchar(MAX),[layer] nvarchar(MAX),[Count] nvarchar(MAX)
)
INSERT INTO @NonWIP
SELECT [ProcName],[Lotnum],[layer],[Count]
FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS T WITH (NOLOCK)
LEFT JOIN @Temp AS r ON T.[AreaID] = @AreaID AND r.[���I] = T.[ProcName] AND r.[�帹] = T.[Lotnum] AND r.[�h�O] = T.[Layer]
INNER JOIN (SELECT [data] FROM [Datamation_H3].[dbo].[my_SplitToTable]((SELECT [ProcName] FROM @Proc),',')) AS pr ON T.[ProcName] LIKE (SUBSTRING(pr.[data],1,3) + '%' + SUBSTRING(pr.[data],5,3) + '%')
WHERE T.[Flag] = 0 AND T.[AreaID] = @AreaID AND r.[���] IS NULL
AND (T.[SystemTime] = '' OR T.[LastEndTime] = '' OR T.[StartTime] = '' OR T.[EndTime] = '' OR T.[Inputpcs] = 0 OR T.[Outputpcs] = 0) AND (T.[Class] = 'D' OR T.[Class] = 'N')

UPDATE L
SET 
[SystemTime] = [���],
[Class] = CASE WHEN CAST([�}�l�ɶ�] AS TIME) > '07:20' AND CAST([�}�l�ɶ�] AS TIME) < '19:20' THEN 'D'ELSE 'N' END,
[StartTime] = [�}�l�ɶ�],
[EndTime] = [�����ɶ�],
[EQID] = [���x],
[MachineNo] = [���x�s��],
[Inputpcs] = [�J�Ƥ���],
[Outputpcs] = [�X�Ƥ���],
[WID] = [�L�b�H���u��],
[User] = [�L�b�H��] 
FROM
(
	SELECT 
	MAX(CASE WHEN h.[AftStatus] = 'MoveIn' THEN h.[ChangeTime] ELSE '' END) AS [���],
	MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN h.[ChangeTime] ELSE '' END) AS [�}�l�ɶ�],
	MAX(CASE WHEN h.[AftStatus] = 'CheckOut' THEN h.[ChangeTime] ELSE '' END) AS [�����ɶ�],
	MAX(ISNULL(m.[MachineName],'')) AS [���x],
    MAX(ISNULL(m.[MachineNo],'')) AS [���x�s��],
	MAX(CASE WHEN h.[AftStatus] = 'MoveIn' THEN h.[qnty] ELSE '' END) AS [�J�Ƥ���],
	MAX(CASE WHEN h.[AftStatus] = 'CheckOut' THEN h.[qnty] ELSE '' END) AS [�X�Ƥ���],
	MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN h.[UserID] ELSE '' END) AS [�L�b�H���u��],
    MAX(CASE WHEN h.[AftStatus] = 'CheckIn' THEN id.[EmpName] ELSE '' END) AS [�L�b�H��],
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


