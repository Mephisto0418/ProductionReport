USE [H3_Systematic]
GO

/****** Object:  StoredProcedure [dbo].[ProductionQuery_Columns_Insert_New]    Script Date: 05/28/2025 19:55:05 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		Boris Li
-- Create date: 2023/07/28
-- Update Date: 2023/08/01
-- Description:	�Ͳ�����UI�B�����إ�
-- 08/01 �קאּ�{���Ʈw
-- 10/17 �s�W�w�]��&�����P�_�y�k
-- 11/03 �ק�s�W����
-- 02/22 �ק���x�d�� : ���x�W�� --> ���x�s��
-- =============================================
CREATE PROCEDURE [dbo].[ProductionQuery_Columns_Insert_New]
	-- Add the parameters for the stored procedure here
	@AreaID nvarchar(50)

	WITH RECOMPILE
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
DECLARE @Temp AS TABLE (
[PID] int,
[CID] int,
[AreaID] int,
[���I] nvarchar(MAX),
[�Ƹ�] nvarchar(MAX),
[�帹] nvarchar(MAX),
[�h�O] nvarchar(MAX),
[Sort] int,
[ParameterName] nvarchar(MAX),
[isQuery] bit,
[QueryCommand] nvarchar(MAX),
[QueryType] nvarchar(MAX),
[FormulaColumn] nvarchar(MAX),
[Machine] nvarchar(MAX),
[�������W��] nvarchar(MAX),
[������] nvarchar(MAX),
[Count] int,
[Face] nvarchar(50),
[isDummy] bit
)

INSERT INTO @Temp
SELECT l.[Pkey],pa.[pkey],l.[AreaID],l.[ProcName] AS [���I],l.[Partnum] AS [�Ƹ�],l.[Lotnum] AS [�帹],l.[Layer] AS [�h�O],pa.[PID] AS [Sort]
,pa.[ParameterName]
,ISNULL(pa.[isQuery],'') AS [isQuery]
,ISNULL(r.[QueryCommand],'') AS [QueryCommand]
,ISNULL(r.[QueryType],'') AS [QueryType]
,CASE WHEN [QueryType] = '��춡�p��' THEN ISNULL(r.[Filter1],'') ELSE '' END AS [FormulaColumn]
,ISNULL(pr.[Machine],'') AS [Machine]
,ISNULL(p.[ParameterName],'') AS [�������W��]
,CASE WHEN pa.isQuery = 0 AND pa.[DefaultValues] <> '' AND (p.[ParameterVaules] = '' OR p.[ParameterVaules] IS NULL) THEN pa.[DefaultValues] ELSE ISNULL(p.[ParameterVaules],'') END AS [������]
,l.[Count],face.[FNo] AS [Face],CASE WHEN l.[Class] = 'Dummy' THEN 1 ELSE 0 END AS [isDummy]
FROM [H3_Systematic].[dbo].[H3_ProductionLog] AS l WITH (NOLOCK)
LEFT JOIN [H3_Systematic].[dbo].[H3_Proc] AS pr WITH (NOLOCK) ON pr.[Pkey] = l.[AreaID] 
LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter] AS pa WITH (NOLOCK) ON pr.[Pkey] = pa.[AreaID]
LEFT JOIN [H3_Systematic].[dbo].[H3_Production_ProcParameter_Rule] AS r WITH (NOLOCK) ON pa.[QID] = r.[QID]
LEFT JOIN [H3_Systematic].[dbo].[H3_ProductionParameter] AS p WITH (NOLOCK) ON p.[PID] = l.[Pkey] AND p.[CID] = pa.[Pkey] AND l.[Count] = p.[Count]
LEFT JOIN (SELECT [Pkey] AS [AreaID] , [FNo]
FROM(
SELECT [Pkey], CASE WHEN [hasFace] = 1 THEN 1 ELSE 1 END AS [PF],CASE WHEN [hasFace] = 1 THEN 2 ELSE 1 END AS [PB] FROM [H3_Systematic].[dbo].[H3_Proc] WITH (NOLOCK) WHERE [Pkey] = @AreaID
) s

UNPIVOT
(
[FNo] FOR [Face]IN([PF],[PB])
)upv
GROUP BY [Pkey],[FNo]) AS face ON face.[AreaID] = pa.[AreaID] AND (face.[FNo] = p.[Face] OR p.[Face] IS NULL)
WHERE pr.[Pkey] = @AreaID AND (l.[MachineNo] = '' OR ISNULL(pr.[MachineNo],'') LIKE '%' + l.[MachineNo] + '%' OR ISNULL(pr.[MachineNo],'') = '')   
AND l.[Flag] = 0
ORDER BY l.[Pkey],[Sort]

IF @@ROWCOUNT >0
BEGIN
INSERT [H3_Systematic].[dbo].[H3_ProductionParameter]
([PID],[CID],[AreaID],[Lotnum],[ProcName],[ParameterName],[ParameterVaules],[SystemTime],[LayerName],[Count],[Face])
SELECT [PID],[CID],@AreaID AS [AreaID],[�帹],[���I],[ParameterName],[������],GETDATE() AS [SystemTime],[�h�O],[Count],[Face]
FROM @Temp 
WHERE [�������W��] = '' AND [ParameterName] IS NOT NULL ORDER BY [PID],[Sort]

DECLARE @PID int
DECLARE @Proc nvarchar(50)
DECLARE	@Lot nvarchar(50)
DECLARE	@Layer nvarchar(50)
DECLARE	@CID nvarchar(MAX)
DECLARE	@PatameterLog nvarchar(MAX)
DECLARE @Part nvarchar(50)
DECLARE @isQ bit
DECLARE @QueryType nvarchar(50)
DECLARE @QueryCommand nvarchar(MAX)
DECLARE @Count int
DECLARE @Face nvarchar(50)

DECLARE Cursor_Var CURSOR FOR
SELECT [PID],[���I],[�Ƹ�],[�帹],[�h�O],[CID],[������],[isQuery],[QueryType],[QueryCommand],[COUNT],[Face]
FROM @Temp WHERE [isQuery] = 1 AND [������] = '' AND [isDummy] = 0 ORDER BY [PID],[Sort]

OPEN Cursor_Var

FETCH NEXT FROM Cursor_Var INTO @PID, @Proc, @Part, @Lot, @Layer, @CID, @PatameterLog, @isQ, @QueryType, @QueryCommand, @Count, @Face

WHILE @@FETCH_STATUS = 0
BEGIN
-- ����w�s�{�Ǩöǻ� @lot �M @proc ����

IF @isQ = 1
    BEGIN
	IF @QueryType <> '' AND @QueryType <> '��춡�p��' AND @PatameterLog  = ''
	    BEGIN
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarPID',@PID)
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarArea',@AreaID)
        SET @QueryCommand = REPLACE(@QueryCommand,'VarLot',@lot)
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarPart',SUBSTRING(@Part,1,8))
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarRev',SUBSTRING(@Part,9,2))
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarProc',@Proc)
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarStation',SUBSTRING(@Proc,1,3) + SUBSTRING(@Proc,5,3))
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarLayer',@Layer)
	    SET @QueryCommand = REPLACE(@QueryCommand,'VarCount',@Count)
		SET @QueryCommand = REPLACE(@QueryCommand,'VarFace',@Face)
	    
	    DECLARE @TempValue AS TABLE  (val nvarchar(MAX))
	    INSERT INTO @TempValue EXEC SP_EXECUTESQL @QueryCommand 
	    IF @@ROWCOUNT >0
	        BEGIN
	        UPDATE [H3_Systematic].[dbo].[H3_ProductionParameter]
	        SET [ParameterVaules] = (SELECT TOP 1 ISNULL([val],'') FROM @TempValue),
	            [UploadTime] = GETDATE()
	        WHERE [PID] = @PID AND [CID] = @CID AND [Count] = @Count AND [Face] = @Face
	        END
	    
	    DELETE @TempValue
	    END
    END

FETCH NEXT FROM Cursor_Var INTO @PID, @Proc, @Part, @Lot, @Layer, @CID, @PatameterLog, @isQ, @QueryType, @QueryCommand, @Count, @Face
END

CLOSE Cursor_Var;
DEALLOCATE Cursor_Var;
END
END
GO


