﻿--Function tach chuoi thanh Table
ALTER FUNCTION [dbo].[fnSplitString] 
( 
    @string NVARCHAR(MAX), 
    @delimiter CHAR(1) 
) 
RETURNS @output TABLE(splitdata NVARCHAR(MAX) 
) 
BEGIN 
    DECLARE @start INT, @end INT 
    SELECT @start = 1, @end = CHARINDEX(@delimiter, @string) 
    WHILE @start < LEN(@string) + 1 BEGIN 
        IF @end = 0  
            SET @end = LEN(@string) + 1
       
        INSERT INTO @output (splitdata)  
        VALUES(SUBSTRING(@string, @start, @end - @start)) 
        SET @start = @end + 1 
        SET @end = CHARINDEX(@delimiter, @string, @start)
        
    END 
    RETURN 
END
GO

--Function chuyen Purchase Type
ALTER FUNCTION [dbo].[fnPUType_Convert] 
( 
    @PUType_Origin varchar(50), 
    @BpCode varchar(250) 
) 
RETURNS varchar(50)
AS
BEGIN 
	DECLARE @Series as int
	DECLARE @PUType as varchar(50)
	Select @Series = Series from OCRD where CardCode = @BpCode;
	Select @PUType = case 
	  --NCC
	  when @Series in (70,71) and @PUType_Origin in ('PUT01','PUT03','PUT04','PUT05','PUT06','PUT07','PUT08') then 'PUT01'
	  when @Series in (72,73) and @PUType_Origin in ('PUT01','PUT03','PUT04','PUT07','PUT08') then 'PUT01'
	  when @Series in (78) and @PUType_Origin in ('PUT01','PUT03','PUT04','PUT07','PUT08') then 'PUT01'
	  --NTP
	  when @Series in (70,71) and @PUType_Origin in ('PUT02') then 'PUT02'
	  when @Series in (72,73) and @PUType_Origin in ('PUT02','PUT05','PUT06') then 'PUT02'
	  when @Series in (78) and @PUType_Origin in ('PUT02') then 'PUT02'
	  --DTC
	  when @Series in (70,71) and @PUType_Origin in ('PUT09') then 'PUT09'
	  when @Series in (72,73) and @PUType_Origin in ('PUT09') then 'PUT09'
	  when @Series in (78) and @PUType_Origin in ('PUT09','PUT05','PUT06') then 'PUT09'
	  else @PUType_Origin
 end;
    RETURN @PUType
END
GO

--Lay MenuUID tu Report Name
ALTER PROCEDURE [dbo].[AC_BS_GET_MENUUID_SCT]
	-- Add the parameters for the stored procedure here
	@ReportName as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT top 1 MenuUID FROM OCMN where [Name]=@ReportName;
END;
GO

--Lay danh sach tai khoan ke toan mapping
ALTER PROCEDURE [dbo].[AC_MAP_TABLE_COA]
	@Account as varchar(50)
AS
BEGIN
	Select U_BKCK 
	from [@COA] 
	where Code = @Account;
END
GO

--Bao cao du tru
--Danh sach du an
ALTER PROCEDURE [dbo].[MM_CE_GET_FPROJECT]
	-- Add the parameters for the stored procedure here
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2017' and T0.[Active] = 'Y'
END;
GO

--Doanh thu du an
ALTER PROCEDURE [dbo].[GET_DATA_BCDT_A]
	 @FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @ProjectID as int;
	DECLARE @DocEntry as int;
    -- Insert statements for procedure here
	--SELECT top 1 @ProjectID = AbsEntry from OPMG where FIPROJECT = @FinancialProject;
	IF (@GoiThauKey = '')
		BEGIN
		Select SUM(z.GTHD) as 'GTHD'
			,SUM(z.GGTM) as 'GGTM'
			,SUM(z.PA) as 'PA'
			,SUM(z.PhiQL) as 'PhiQL'
			,SUM(z.PLHD) as 'PLHD'
			,SUM(z.KHAC) as 'KHAC'
		from (
				Select SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
				,SUM(a.U_GGTM) as 'GGTM'
				,SUM(a.U_PADXTK) as 'PA'
				,SUM(a.U_PQL) as 'PhiQL'
				,'0' as 'PLHD'
				,'0' as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 47
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'

				union all

				Select '0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,SUM(t1.PLHD) as PLHD
				,'0' as 'KHAC'
				from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 142
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y') t1

				union all

				Select 
				'0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,'0' as 'PLHD'
				,SUM(t2.KHAC) as KHAC from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 203
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y') t2
				) z
		END
	ELSE
		BEGIN
			Select SUM(z.GTHD) as 'GTHD'
			,SUM(z.GGTM) as 'GGTM'
			,SUM(z.PA) as 'PA'
			,SUM(z.PhiQL) as 'PhiQL'
			,SUM(z.PLHD) as 'PLHD'
			,SUM(z.KHAC) as 'KHAC'
		from (
				Select SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
				,SUM(a.U_GGTM) as 'GGTM'
				,SUM(a.U_PADXTK) as 'PA'
				,SUM(a.U_PQL) as 'PhiQL'
				,'0' as 'PLHD'
				,'0' as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 47
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))

				union all

				Select '0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,SUM(t1.PLHD) as PLHD
				,'0' as 'KHAC'
				from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and a.Series = 142
				and a.BpType = 'C'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t1

				union all

				Select 
				'0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,'0' as 'PLHD'
				,SUM(t2.KHAC) as KHAC from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and a.Series = 203
				and a.BpType = 'C'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t2
				) z
		END
END
GO

--Du tru tong
ALTER PROCEDURE [dbo].[MM_CE_GETDATA_SUM_NEW]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
    -- Insert statements for procedure here
	if (@GoiThauKey = '')
			Select * from 
			(
			Select * 
			FROM [@DUTRUA] 
			where DocEntry in 
			(Select DocEntry
				from [@DUTRU] 
				where U_DUTRU_TYPE = 1
				and U_CTG_Key in (
					Select a.CTG_KEY 
					from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
							from [@CTG] 
							where U_PrjCode = @FinancialProject
							group by U_GoiThauKey) a)
					)) T0 left join 
			(
			Select U_001,SUM(U_TTHD) as 'TTHD' 
			from OPHA 
			where ProjectID in (Select AbsEntry from OPMG where FIPROJECT = @FinancialProject)
			and [Level] = 2
			group by U_001
			) T1 on T0.U_SubProjectCode = T1.U_001;
	else
		Select * from 
			(
			Select * 
			FROM [@DUTRUA] 
			where DocEntry in 
			(Select DocEntry
				from [@DUTRU] 
				where U_DUTRU_TYPE = 1
				and U_CTG_Key in (
					Select a.CTG_KEY 
					from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
							from [@CTG] 
							where U_PrjCode = @FinancialProject
							and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
							group by U_GoiThauKey) a)
					)) T0 left join 
			(
			Select U_001,SUM(U_TTHD) as 'TTHD' 
			from OPHA 
			where ProjectID in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
			and [Level] = 2
			group by U_001
			) T1 on T0.U_SubProjectCode = T1.U_001;
END
GO

--Du tru chi tiet
ALTER PROCEDURE [dbo].[MM_CE_GETDATA_DETAILS_NEW]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @CTG_DocEntry as int;
	DECLARE @DocEntry as int;
	DECLARE @DUTRU_DocEntry as int;
	DECLARE @TableTmp_KLTT TABLE(
		U_BPCode varchar(250) NOT NULL,
		U_BPName nvarchar(254),
		U_BPCode2 varchar(250) ,
		U_Sub3Name nvarchar(254),
		U_GoiThauKey int,
		U_PUTYPE varchar(50),
		KL_HD decimal(18,0),
		KL_TT decimal(18,0),
		KL_TT_DD decimal(18,0)
	);
		DECLARE @TableTmp_KLTT_APPROVE TABLE(
		U_BPCode varchar(250) NOT NULL,
		U_BPName nvarchar(254),
		U_BPCode2 varchar(250) ,
		U_Sub3Name nvarchar(254),
		U_GoiThauKey int,
		U_PUTYPE varchar(50),
		KL_HD decimal(18,0),
		KL_TT decimal(18,0),
		KL_TT_DD decimal(18,0)
	);
    -- Insert statements for procedure here
	if (@GoiThauKey = '')
	begin
	--Get Data KLTT All Project
	INSERT INTO @TableTmp_KLTT(U_BPCode,U_BPName,U_BPCode2,U_Sub3Name,U_GoiThauKey,U_PUTYPE,KL_HD,KL_TT,KL_TT_DD)
	Select a.U_BPCOde
		,a.U_BPName
		,a.U_BPCode2
		,b.U_Sub3Name
		,'' as 'U_GoiThauKey'
		,a.U_PUTYPE
		,SUM(SUM_PL) as 'KL_HD'
		,SUM(SUM_CA) as 'KL_TT' 
		,SUM(case a.Status when 'C' then b.SUM_CA else 0 end) as 'KL_TT_DD'
	from [@KLTT] a inner join
		(
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTA] 
		union all
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTB] 
		union all
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTK] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [@KLTTC] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [@KLTTD] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTE] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTF] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTG])b on a.DocEntry = b.DocEntry
	where a.DocEntry in 
		(Select --y.U_BPCode,
		DocEntry from [@KLTT] x inner join (
		Select U_BPCode,MAx(U_Dateto) as Dateto from [@KLTT] where U_FIPROJECT = @FinancialProject and U_BType = 2 and Canceled not in ('Y','C') group by U_BPCode) y
		on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto)
	and a.Canceled not in ('Y','C')
	and a.Status = 'C'
	and (Select GroupCode from OCRD where CardCode=a.U_BPCode) <> 112
	group by a.U_BPCOde,a.U_BPName,a.U_BPCode2--,b.U_GoiThauKey
	,b.U_Sub3Name,a.U_PUTYPE;

	Select ISNULL(T0.U_BPCode,T1.U_BPCode) as 'U_BPCode'
	   ,ISNULL(T0.U_BPName,T1.U_BPName) as 'U_BPName'
	   ,(Select CardName from OCRD where CardCode=T1.U_BPCode2) as 'CTQL'
	   ,ISNULL(T0.U_SubProjectDesc,T1.U_Sub3Name) as 'U_SubProjectDesc'
	   ,ISNULL(T0.[U_DTT_LineID],0) as 'U_DTT_LineID'
	   ,T1.U_PUTYPE
	   ,ISNULL(T0.U_CP_NCC,0) as 'U_CP_NCC'
	   ,ISNULL(T0.U_CP_CN,0) as 'U_CP_CN'
	   ,ISNULL(T0.U_CP_DP,0) as 'U_CP_DP'
	   ,ISNULL(T0.U_CP_DP2,0) as 'U_CP_DP2'
	   ,ISNULL(T0.U_CP_PRELIMS,0) as 'U_CP_PRELIMS'
	   ,ISNULL(T0.U_CP_TB,0) as 'U_CP_TB'
	   ,ISNULL(T0.U_CP_K,0) as 'U_CP_K'
	   ,ISNULL(T0.U_CP_NTP,0) as 'U_CP_NTP'
	   ,ISNULL(T0.U_CP_DTC,0) as 'U_CP_DTC'
	   ,ISNULL(T0.U_CP_VTP,0) as 'U_CP_VTP'
	   ,ISNULL(T0.U_CP_VC,0) as 'U_CP_VC'
	   ,ISNULL(T0.U_CP_MB,0) as 'U_CP_MB'
	   ,ISNULL(T0.U_CP_T,0) as 'U_CP_T'
	   ,ISNULL(T0.U_CP_VH,0) as 'U_CP_VH'
	   ,ISNULL(T1.KL_HD,0) as 'KL_HD'
	   ,ISNULL(T1.KL_TT,0) as 'KL_TT'
	   ,ISNULL(T1.KL_TT_DD,0) as 'KL_TT_DD'
	    from 
		--Du TRU
		(Select [U_BPCode]
				,[U_BPName]
				,[U_SubProjectDesc]
				,[U_DTT_LineID]
				,SUM([U_CP_NCC]) as 'U_CP_NCC'
				,SUM([U_CP_CN]) as 'U_CP_CN'
				,SUM([U_CP_DP]) as 'U_CP_DP'
				,SUM([U_CP_DP2]) as 'U_CP_DP2'
				,SUM([U_CP_Prelims]) as 'U_CP_PRELIMS'
				,SUM([U_CP_TB]) as 'U_CP_TB'
				,SUM([U_CP_K]) as 'U_CP_K'
				,SUM([U_CP_NTP]) as 'U_CP_NTP'
				,SUM([U_CP_DTC]) as 'U_CP_DTC'
				,SUM([U_CP_VTP]) as 'U_CP_VTP'
				,SUM([U_CP_VC]) as 'U_CP_VC'
				,SUM([U_CP_MB]) as 'U_CP_MB'
				,SUM([U_CP_T]) as 'U_CP_T'
				,SUM([U_CP_VH]) as 'U_CP_VH' 
				FROM [@DUTRUB] 
				where DocEntry in 
					(Select DocEntry
					from [@DUTRU] 
					where U_DUTRU_TYPE = 1
					and U_CTG_Key in (
						Select a.CTG_KEY 
						from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [@CTG] 
								where U_PrjCode = @FinancialProject
								group by U_GoiThauKey) a)
					)
				group by [U_BPCode],[U_BPName],[U_SubProjectDesc],[U_DTT_LineID]) T0
			FULL JOIN
			@TableTmp_KLTT T1 on T0.U_BPCode = T1.U_BPCode and T0.U_SubProjectDesc = T1.U_Sub3Name;
	end
	
	else
	begin
	--Get Data KLTT All Project
	INSERT INTO @TableTmp_KLTT(U_BPCode,U_BPName,U_BPCode2,U_Sub3Name,U_GoiThauKey,U_PUTYPE,KL_HD,KL_TT,KL_TT_DD)
	Select a.U_BPCOde
		,a.U_BPName
		,a.U_BPCode2
		,b.U_Sub3Name
		,'' as 'U_GoiThauKey'
		,a.U_PUTYPE
		,SUM(SUM_PL) as 'KL_HD'
		,SUM(SUM_CA) as 'KL_TT' 
		,SUM(case a.Status when 'C' then b.SUM_CA else 0 end) as 'KL_TT_DD'
	from [@KLTT] a inner join
		(
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTA] 
		union all
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTB] 
		union all
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTK] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [@KLTTC] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [@KLTTD] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTE] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTF] 
		union all
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [@KLTTG])b on a.DocEntry = b.DocEntry
	where a.DocEntry in 
		(Select --y.U_BPCode,
		DocEntry from [@KLTT] x inner join (
		Select U_BPCode,MAx(U_Dateto) as Dateto from [@KLTT] where U_FIPROJECT = @FinancialProject and U_BType = 2 and Canceled not in ('Y','C') group by U_BPCode) y
		on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto)
	and a.Canceled not in ('Y','C')
	and b.U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
	and a.[Status] = 'C'
	and (Select GroupCode from OCRD where CardCode=a.U_BPCode) <> 112
	group by a.U_BPCOde,a.U_BPName,a.U_BPCode2--,b.U_GoiThauKey
	,b.U_Sub3Name,a.U_PUTYPE;

	Select ISNULL(T0.U_BPCode,T1.U_BPCode) as 'U_BPCode'
	   ,ISNULL(T0.U_BPName,T1.U_BPName) as 'U_BPName'
	   ,(Select CardName from OCRD where CardCode=T1.U_BPCode2) as 'CTQL'
	   ,ISNULL(T0.U_SubProjectDesc,T1.U_Sub3Name) as 'U_SubProjectDesc'
	   ,ISNULL(T0.[U_DTT_LineID],0) as 'U_DTT_LineID'
	   ,T1.U_PUTYPE
	   ,ISNULL(T0.U_CP_NCC,0) as 'U_CP_NCC'
	   ,ISNULL(T0.U_CP_CN,0) as 'U_CP_CN'
	   ,ISNULL(T0.U_CP_DP,0) as 'U_CP_DP'
	   ,ISNULL(T0.U_CP_DP2,0) as 'U_CP_DP2'
	   ,ISNULL(T0.U_CP_PRELIMS,0) as 'U_CP_PRELIMS'
	   ,ISNULL(T0.U_CP_TB,0) as 'U_CP_TB'
	   ,ISNULL(T0.U_CP_K,0) as 'U_CP_K'
	   ,ISNULL(T0.U_CP_NTP,0) as 'U_CP_NTP'
	   ,ISNULL(T0.U_CP_DTC,0) as 'U_CP_DTC'
	   ,ISNULL(T0.U_CP_VTP,0) as 'U_CP_VTP'
	   ,ISNULL(T0.U_CP_VC,0) as 'U_CP_VC'
	   ,ISNULL(T0.U_CP_MB,0) as 'U_CP_MB'
	   ,ISNULL(T0.U_CP_T,0) as 'U_CP_T'
	   ,ISNULL(T0.U_CP_VH,0) as 'U_CP_VH'
	   ,ISNULL(T1.KL_HD,0) as 'KL_HD'
	   ,ISNULL(T1.KL_TT,0) as 'KL_TT'
	   ,ISNULL(T1.KL_TT_DD,0) as 'KL_TT_DD'
	    from 
		--Du TRU
		(Select [U_BPCode]
				,[U_BPName]
				,[U_SubProjectDesc]
				,[U_DTT_LineID]
				,SUM([U_CP_NCC]) as 'U_CP_NCC'
				,SUM([U_CP_CN]) as 'U_CP_CN'
				,SUM([U_CP_DP]) as 'U_CP_DP'
				,SUM([U_CP_DP2]) as 'U_CP_DP2'
				,SUM([U_CP_Prelims]) as 'U_CP_PRELIMS'
				,SUM([U_CP_TB]) as 'U_CP_TB'
				,SUM([U_CP_K]) as 'U_CP_K'
				,SUM([U_CP_NTP]) as 'U_CP_NTP'
				,SUM([U_CP_DTC]) as 'U_CP_DTC'
				,SUM([U_CP_VTP]) as 'U_CP_VTP'
				,SUM([U_CP_VC]) as 'U_CP_VC'
				,SUM([U_CP_MB]) as 'U_CP_MB'
				,SUM([U_CP_T]) as 'U_CP_T'
				,SUM([U_CP_VH]) as 'U_CP_VH' 
				FROM [@DUTRUB] 
				where DocEntry in 
					(Select DocEntry
					from [@DUTRU] 
					where U_DUTRU_TYPE = 1
					and U_CTG_Key in (
						Select a.CTG_KEY 
						from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [@CTG] 
								where U_PrjCode = @FinancialProject
								and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
								group by U_GoiThauKey) a)
					)
				group by [U_BPCode],[U_BPName],[U_SubProjectDesc],[U_DTT_LineID]) T0
			FULL JOIN
			@TableTmp_KLTT T1 on T0.U_BPCode = T1.U_BPCode and T0.U_SubProjectDesc = T1.U_Sub3Name;
	end
END
GO

--Chi phi BCH
ALTER PROCEDURE [dbo].[MM_CE_GET_DATA_BCH_NEW]
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DocEntry as int;
	if (@GoiThauKey = '')
		Select * from
		(
		Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' 
		FROM [@CTG4] d
		where DocEntry in 
			(
				 Select DocEntry from [@CTG] T0 inner join
				 (Select U_GoiThauKey,MAX(U_Date) as 'U_DATE'
				 From [@CTG] 
				 where U_PrjCode = @FinancialProject
				 group by U_GoiThauKey) T1 on T0.U_GoiThauKey = T1.U_GoiThauKey and T0.U_Date = T1.U_DATE
			)
		group by U_TKKT,U_TTKKT
		) a
		left join 
		(Select case SUBSTRING( b.Account,1,4) when '3341' then '3341' else b.Account end as 'Account'
		, SUM(b.Debit) as TOTAL_BCH
		From OJDT a inner join JDT1 b on a.TransID=b.TransId
		where b.Project = @FinancialProject
		group by case SUBSTRING( b.Account,1,4) when '3341' then '3341' else b.Account end) b on a.U_TKKT=b.Account;
	else
		Select * from
		(	Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' FROM [@CTG4] 
			where DocEntry in 
			(
					Select DocEntry from [@CTG] T0 inner join
					(Select U_GoiThauKey,MAX(U_Date) as 'U_DATE'
					From [@CTG] 
					where U_PrjCode = @FinancialProject
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
					group by U_GoiThauKey) T1 on T0.U_GoiThauKey = T1.U_GoiThauKey and T0.U_Date = T1.U_DATE
			)
			group by U_TTKKT,left(U_TKKT + '00000000',8)
		) a
		left join 
			(Select case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end as 'Account'
			, SUM(b.Debit) as TOTAL_BCH
			From OJDT a inner join JDT1 b on a.TransID=b.TransId
			where b.Project = @FinancialProject
			group by case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end) b 
		on a.U_TKKT=b.Account ;
END;
GO

--Bao cao tai chinh theo doi tuong
--Danh sach du an
ALTER PROCEDURE [dbo].[MM_FI_GET_FPROJECT]
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2017' and T0.[Active] = 'Y'
END;
GO

--Doanh thu
ALTER PROCEDURE [dbo].[MM_FI_GET_DATA_A]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
	IF (@GoiThauKey = '')
		BEGIN
			Select SUM(z.GTHD) as 'GTHD'
			,SUM(z.GGTM) as 'GGTM'
			,SUM(z.PA) as 'PA'
			,SUM(z.PhiQL) as 'PhiQL'
			,SUM(z.PLHD) as 'PLHD'
			,SUM(z.KHAC) as 'KHAC'
			from 
			(
				Select SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
				,SUM(a.U_GGTM) as 'GGTM'
				,SUM(a.U_PADXTK) as 'PA'
				,SUM(a.U_PQL) as 'PhiQL'
				,'0' as 'PLHD'
				,'0' as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 47
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				union all
				Select '0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,SUM(t1.PLHD) as PLHD
				,'0' as 'KHAC'
				from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 142
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y') t1
				union all
				Select 
				'0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,'0' as 'PLHD'
				,SUM(t2.KHAC) as KHAC from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 203
				and a.[Status] ='A'
				and a.BpType = 'C'
				and a.Cancelled <> 'Y') t2
			) z
		END
	ELSE
		BEGIN
			Select SUM(z.GTHD) as 'GTHD'
			,SUM(z.GGTM) as 'GGTM'
			,SUM(z.PA) as 'PA'
			,SUM(z.PhiQL) as 'PhiQL'
			,SUM(z.PLHD) as 'PLHD'
			,SUM(z.KHAC) as 'KHAC'
			from 
			(
				Select SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
				,SUM(a.U_GGTM) as 'GGTM'
				,SUM(a.U_PADXTK) as 'PA'
				,SUM(a.U_PQL) as 'PhiQL'
				,'0' as 'PLHD'
				,'0' as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 47
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
				
				union all
				
				Select '0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,SUM(t1.PLHD) as PLHD
				,'0' as 'KHAC'
				from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 142
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t1
				union all
				Select 
				'0' as 'GTHD'
				,'0' as 'GGTM'
				,'0' as 'PA'
				,'0' as 'PhiQL'
				,'0' as 'PLHD'
				,SUM(t2.KHAC) as KHAC from (
				Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
				from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
				where a.U_PRJ = @FinancialProject
				and a.Series = 203
				and a.BpType = 'C'
				and a.[Status] ='A'
				and a.Cancelled <> 'Y'
				and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t2
			) z
	END
END
GO

--Chi phi
ALTER PROCEDURE [dbo].[MM_FI_GET_DATA_B_NEW]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@Goithau_Key as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	if (@Goithau_Key = '')
	begin
		Select T0.U_BPCode
		,T0.U_BPName
		,(Select CardName from OCRD where CardCode = T1.U_BPCode2) as 'U_BPCode2'
		,T0.U_TYPE
		,T0.U_CP_NCC
		,T0.U_CP_NTP
		,T0.U_CP_DTC
		,T0.U_CP_DP2
		,T1.U_BGroup
		,T1.BP
		,T1.U_PUType
		,case when (Select GroupCode from OCRD where CardCode=T0.U_BPCode) <> 112 then T1.KL_HD
			else T1.KL_HD  * (U_PTQuanly/100) end as 'KL_HD'
		,T1.KL_TT
		,case when (Select GroupCode from OCRD where CardCode=T0.U_BPCode) <> 112 then T1.KL_TT_DD
			else T1.KL_TT_DD  * (U_PTQuanly/100) end as 'KL_TT_DD' --T1.KL_TT_DD
		,T2.GTHD
		,T3.TOTAL as 'TOTAL_AP_INVOICE' 
		from 
			(
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,case when b.Series in (70,71) then SUM([U_CP_NCC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])  
					else SUM([U_CP_NCC]) end
				  as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,0 as 'U_CP_NTP'
				  ,0 as 'U_CP_DTC'
				  ,'PUT01' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) a))
				 --and U_CP_NCC <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]
			Union ALL
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,0 as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,case when b.Series in (72,73) then SUM([U_CP_NTP]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				   else SUM([U_CP_NTP]) end
				   as 'U_CP_NTP'
				  ,0 as 'U_CP_DTC'
				  ,'PUT02' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) a))
				--and U_CP_NTP <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]
			Union ALL
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,0 as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,0 as 'U_CP_NTP'
				  ,case when b.Series in (78) then SUM([U_CP_DTC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				  else SUM([U_CP_DTC]) end
				  as 'U_CP_DTC'
				  ,'PUT09' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) a))
				--and U_CP_DTC <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]) T0
			LEFT JOIN
			(Select a.U_BPCode
				,a.U_BPName
				,a.U_BGroup
				,case (Select dept from OHEM where Userid = a.UserSign) when 1 then 'CCM'
					when -2 then 'KT'
					when 3 then 'DA'
					else ''
				 end as 'BP'
				,dbo.fnPUType_Convert(a.U_PUType,a.U_BPCode) as 'U_PUType'
				,ISNULL(a.U_BPCode2,'') as 'U_BPCode2'
				,ISNULL(a.U_PTQuanly,0) as 'U_PTQuanly'
				,SUM(b.U_SUM) as 'KL_HD'
				,SUM(b.U_CompleteAmount) as 'KL_TT'
				,SUM(case a.Status when 'C' then (b.U_CompleteAmount) else 0 end) as 'KL_TT_DD'
				from [@KLTT] a inner join (
				Select z1.DocEntry,SUM(z1.Sum_PL) as 'U_SUM', SUM(z1.SUM_CA) as 'U_CompleteAmount'
				from (
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTA] 
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTB] 
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTK] 
					union all
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [@KLTTC] 
					union all
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [@KLTTD] 
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTE] 
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTF] 
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTG]
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTK]) z1
					group by z1.DocEntry
				) b on a.DocEntry = b.DocEntry
				where a.DocEntry in (
				Select --y.U_BPCode,
				DocEntry from [@KLTT] x inner join (
				Select U_BPCode,U_BGroup,U_PUType,MAx(U_Dateto) as Dateto 
				from [@KLTT] 
				where U_FIPROJECT = @FinancialProject 
				and [Status] = 'C'
				and Canceled <>  'Y'
				group by U_BPCode,U_BGroup,U_PUType) y
				on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto and x.U_BGroup = y.U_BGroup and x.U_PUType = y.U_PUType
				and x.U_FIPROJECT= @FinancialProject and x.Canceled <> 'Y' and x.[Status] = 'C')
				group by  a.U_BPCode,a.U_BPName,a.U_BGroup,a.UserSign
				,dbo.fnPUType_Convert(a.U_PUType,a.U_BPCode)
				,ISNULL(a.U_BPCode2,''),ISNULL(a.U_PTQuanly,0)
				) T1
				on T0.U_BPCode = T1.U_BPCode and T0.U_TYPE = T1.U_BGroup and T0.PUType = T1.U_PUType
				and (T0.U_CP_NCC <> 0 or T0.U_CP_DTC <> 0 or T0.U_CP_NTP <>0 )
			LEFT JOIN
			(Select BpCode,U_CGroup,U_PUType,SUM(GTHD) as 'GTHD' from
				(Select x.BpCode
					,x.U_CGroup
					,x.U_PUType
					,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC))
					from OAT1 b
					where  b.AgrNo = AbsID) as 'GTHD'
					from OOAT x inner join 
					(Select BpCode
					,U_CGroup
					,U_PUType
					,Max(StartDate) as 'Last_Date'
					from OOAT
					where 
					Status ='A'
					and Cancelled <> 'Y'
					and Series =48
					and U_PRJ = @FinancialProject
					group by BpCode	,U_CGroup,U_PUType) y on x.BpCode = y.BpCode and x.U_CGroup = y.U_CGroup and x.U_PUTYPE = y.U_PUTYPE and x.StartDate = y.Last_Date
					where Series =48
					and Status = 'A'
					and Cancelled <> 'Y') T 
			where T.GTHD <> 1
			group by BpCode,U_CGroup,U_PUType) T2
			on T0.U_BPCode = T2.BpCode 
			and T1.U_BGroup = T2.U_CGroup
			and T1.U_PUType = T2.U_PUType
			LEFT JOIN 
			(Select a.CardCode,a.U_RECTYPE,a.U_PUTYPE
				,(Select SUM(LineTotal) from PCH1 where DocEntry=a.DocEntry) as 'TOTAL'
			from OPCH a
			where a.Project = @FinancialProject
			and a.CANCELED not in ('Y','C'))T3
			on T2.BpCode = T3.CardCode
			and T2.U_CGroup = T3.U_RECTYPE
			and T2.U_PUTYPE = T3.U_PUTYPE
	end
	else
	begin
		Select T0.U_BPCode
		,T0.U_BPName
		,(Select CardName from OCRD where CardCode = T1.U_BPCode2) as 'U_BPCode2'
		,T0.U_TYPE
		,T0.U_CP_NCC
		,T0.U_CP_NTP
		,T0.U_CP_DTC
		,T0.U_CP_DP2
		,T1.U_BGroup
		,T1.BP
		,T1.U_PUType
		,case when (Select GroupCode from OCRD where CardCode=T0.U_BPCode) <> 112 then T1.KL_HD
			else T1.KL_HD  * (U_PTQuanly/100) end as 'KL_HD'
		,T1.KL_TT
		,case when (Select GroupCode from OCRD where CardCode=T0.U_BPCode) <> 112 then T1.KL_TT_DD
			else T1.KL_TT_DD  * (U_PTQuanly/100) end as 'KL_TT_DD' --T1.KL_TT_DD
		,T2.GTHD
		,T3.TOTAL as 'TOTAL_AP_INVOICE' 
		from 
			(
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,case when b.Series in (70,71) then SUM([U_CP_NCC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])  
					else SUM([U_CP_NCC]) end
				  as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,0 as 'U_CP_NTP'
				  ,0 as 'U_CP_DTC'
				  ,'PUT01' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
									from [@CTG] 
									where U_PrjCode = @FinancialProject 
									and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
									group by U_GoiThauKey) a))
				 --and U_CP_NCC <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]
			Union ALL
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,0 as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,case when b.Series in (72,73) then SUM([U_CP_NTP]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				   else SUM([U_CP_NTP]) end
				   as 'U_CP_NTP'
				  ,0 as 'U_CP_DTC'
				  ,'PUT02' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
									from [@CTG] 
									where U_PrjCode = @FinancialProject 
									and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
									group by U_GoiThauKey) a))
				--and U_CP_NTP <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]
			Union ALL
			Select [U_BPCode]
				  ,[U_BPName]
				  ,a.[U_TYPE]
				  ,0 as 'U_CP_NCC'
				  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
				  ,0 as 'U_CP_NTP'
				  ,case when b.Series in (78) then SUM([U_CP_DTC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				  else SUM([U_CP_DTC]) end
				  as 'U_CP_DTC'
				  ,'PUT09' as PUType
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [@CTG] 
								where U_PrjCode = @FinancialProject 
								and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
								group by U_GoiThauKey) a))
				--and U_CP_DTC <> 0
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]) T0
			LEFT JOIN
			(Select a.U_BPCode
				,a.U_BPName
				,a.U_BGroup
				,case (Select dept from OHEM where Userid = a.UserSign) when 1 then 'CCM'
					when -2 then 'KT'
					when 3 then 'DA'
					else ''
				 end as 'BP'
				,dbo.fnPUType_Convert(a.U_PUType,a.U_BPCode) as 'U_PUType'
				,ISNULL(a.U_BPCode2,'') as 'U_BPCode2'
				,ISNULL(a.U_PTQuanly,0) as 'U_PTQuanly'
				,SUM(b.U_SUM) as 'KL_HD'
				,SUM(b.U_CompleteAmount) as 'KL_TT'
				,SUM(case a.Status when 'C' then (b.U_CompleteAmount) else 0 end) as 'KL_TT_DD'
				from [@KLTT] a inner join (
				Select z1.DocEntry,SUM(z1.Sum_PL) as 'U_SUM', SUM(z1.SUM_CA) as 'U_CompleteAmount'
				from (
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTA]
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTB] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTK] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [@KLTTC] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [@KLTTD] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTE] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTF] 
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTG]
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union all
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [@KLTTK]
					where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))) z1
					group by z1.DocEntry
				) b on a.DocEntry = b.DocEntry
				where a.DocEntry in (
				Select --y.U_BPCode,
				DocEntry from [@KLTT] x inner join (
				Select U_BPCode,U_BGroup,U_PUType,MAx(U_Dateto) as Dateto 
				from [@KLTT] 
				where U_FIPROJECT = @FinancialProject 
				and [Status] = 'C'
				and Canceled <>  'Y'
				group by U_BPCode,U_BGroup,U_PUType) y
				on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto and x.U_BGroup = y.U_BGroup and x.U_PUType = y.U_PUType)
				group by  a.U_BPCode,a.U_BPName,a.U_BGroup,a.UserSign
				,dbo.fnPUType_Convert(a.U_PUType,a.U_BPCode)
				,ISNULL(a.U_BPCode2,''),ISNULL(a.U_PTQuanly,0)
				) T1
				on T0.U_BPCode = T1.U_BPCode and T0.U_TYPE = T1.U_BGroup and T0.PUType = T1.U_PUType
				and (T0.U_CP_NCC <> 0 or T0.U_CP_DTC <> 0 or T0.U_CP_NTP <>0 )
			LEFT JOIN
			(Select BpCode,U_CGroup,U_PUType,SUM(GTHD) as 'GTHD' from
				(Select x.BpCode
					,x.U_CGroup
					,x.U_PUType
					,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC))
					from OAT1 b
					where  b.AgrNo = AbsID) as 'GTHD'
					from OOAT x inner join 
					(Select BpCode
					,U_CGroup
					,U_PUType
					,Max(StartDate) as 'Last_Date'
					from OOAT
					where 
					Status ='A'
					and Series =48
					and U_PRJ = @FinancialProject
					--and U_Goithau in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					group by BpCode	,U_CGroup,U_PUType) y on x.BpCode = y.BpCode and x.U_CGroup = y.U_CGroup and x.U_PUTYPE = y.U_PUTYPE and x.StartDate = y.Last_Date
					where Series =48
					and Status = 'A') T 
			where T.GTHD <> 1
			group by BpCode,U_CGroup,U_PUType) T2
			on T0.U_BPCode = T2.BpCode 
			and T1.U_BGroup = T2.U_CGroup
			and T1.U_PUType = T2.U_PUType
			LEFT JOIN 
			(Select a.CardCode,a.U_RECTYPE,a.U_PUTYPE
				,(Select SUM(LineTotal) from PCH1 where DocEntry=a.DocEntry) as 'TOTAL'
			from OPCH a
			where a.Project = @FinancialProject
			and a.CANCELED not in ('Y','C'))T3
			on T2.BpCode = T3.CardCode
			and T2.U_CGroup = T3.U_RECTYPE
			and T2.U_PUTYPE = T3.U_PUTYPE
	end
END
GO

--Chi phi BCH
ALTER PROCEDURE [dbo].[MM_FI_GET_DATA_BCH_NEW]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(50)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DocEntry as int;
	if (@GoiThauKey = '')
		Select * from
		(
		Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' 
		FROM [@CTG4] 
		where DocEntry in 
			(
				 Select DocEntry from [@CTG] T0 inner join
				 (Select U_GoiThauKey,MAX(U_Date) as 'U_DATE'
				 From [@CTG] 
				 where U_PrjCode = @FinancialProject
				 group by U_GoiThauKey) T1 on T0.U_GoiThauKey = T1.U_GoiThauKey and T0.U_Date = T1.U_DATE
			)
		group by U_TKKT,U_TTKKT) a
		left join 
		(Select case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end as 'Account'
		, SUM(b.Debit) as TOTAL_BCH
		From OJDT a inner join JDT1 b on a.TransID=b.TransId
		where b.Project = @FinancialProject
		group by case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end) b on a.U_TKKT=b.Account;
	else
		Select * from
		(Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' FROM [@CTG4] 
			where DocEntry in 
			(
				 Select DocEntry from [@CTG] T0 inner join
				 (Select U_GoiThauKey,MAX(U_Date) as 'U_DATE'
				 From [@CTG] 
				 where U_PrjCode = @FinancialProject
				 and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
				 group by U_GoiThauKey) T1 on T0.U_GoiThauKey = T1.U_GoiThauKey and T0.U_Date = T1.U_DATE
			)
		 group by U_TTKKT,left(U_TKKT + '00000000',8)
		) a
		left join 
			(Select case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end as 'Account'
			,SUM(b.Debit) as TOTAL_BCH
			From OJDT a inner join JDT1 b on a.TransID=b.TransId
			where b.Project = @FinancialProject
			group by case SUBSTRING( b.Account,1,4) when '3341' then '33410000' else b.Account end) b 
		on a.U_TKKT=b.Account ;
END;
GO

--Chi phi du tru, du phong, bao hanh, ... ( dung chung cho Bao cao du tru)
ALTER PROCEDURE [dbo].[MM_FI_GET_DATA_VII]
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
	IF (@GoiThauKey = '')
		BEGIN
			Select
			  z.U_GOITHAU
			  ,(Select U_CPHT1 from OPMG where DocNum=z.U_GOITHAU) as 'HT1'
			  ,(Select U_CPHT2 from OPMG where DocNum=z.U_GOITHAU) as 'HT2'
			  ,(Select U_CPNG from OPMG where DocNum=z.U_GOITHAU) as 'CPNG'
			  ,(Select U_DPCP from OPMG where DocNum=z.U_GOITHAU) as 'DPCP'
			  ,(Select U_DPBH from OPMG where DocNum=z.U_GOITHAU) as 'DPBH'
			  ,(Select U_CPQLCT from OPMG where DocNum=z.U_GOITHAU) as 'CPQLCT'
			  ,SUM(z.GTHD) as 'GTHD'
			  ,SUM(z.GGTM) as 'GGTM'
			  ,SUM(z.PA) as 'PA'
			  ,SUM(z.PhiQL) as 'PhiQL'
			  ,SUM(z.PLHD) as 'PLHD'
			  ,SUM(z.KHAC) as 'KHAC'
			  ,SUM(z.GTHD) + SUM(z.GGTM) + SUM(z.PA) + SUM(z.PhiQL) + SUM(z.PLHD) + SUM(z.KHAC) as 'Total'
					from 
					(
						Select 
						a.U_GOITHAU
						,SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
						,SUM(a.U_GGTM) as 'GGTM'
						,SUM(a.U_PADXTK) as 'PA'
						,SUM(a.U_PQL) as 'PhiQL'
						,'0' as 'PLHD'
						,'0' as 'KHAC'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 47
						and a.BpType = 'C'
						and a.[Status] ='A'
						and a.Cancelled <> 'Y'
						group by a.U_GOITHAU

						union all

						Select 
						t1.U_GOITHAU
						,'0' as 'GTHD'
						,'0' as 'GGTM'
						,'0' as 'PA'
						,'0' as 'PhiQL'
						,SUM(t1.PLHD) as PLHD
						,'0' as 'KHAC'
						from (
						Select a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 142
						and a.BpType = 'C'
						and a.[Status] ='A'
						and a.Cancelled <> 'Y'
						) t1
						group by t1.U_GOITHAU

						union all

						Select 
						t2.U_GOITHAU
						,'0' as 'GTHD'
						,'0' as 'GGTM'
						,'0' as 'PA'
						,'0' as 'PhiQL'
						,'0' as 'PLHD'
						,SUM(t2.KHAC) as KHAC from (
						Select a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 203
						and a.[Status] ='A'
						and a.BpType = 'C'
						and a.Cancelled <> 'Y') t2
						group by U_GOITHAU

					) z
					group by z.U_GOITHAU
					order by z.U_GOITHAU
		END
	ELSE
		BEGIN
		Select
			  z.U_GOITHAU
			  ,(Select U_CPHT1 from OPMG where DocNum=z.U_GOITHAU) as 'HT1'
			  ,(Select U_CPHT2 from OPMG where DocNum=z.U_GOITHAU) as 'HT2'
			  ,(Select U_CPNG from OPMG where DocNum=z.U_GOITHAU) as 'CPNG'
			  ,(Select U_DPCP from OPMG where DocNum=z.U_GOITHAU) as 'DPCP'
			  ,(Select U_DPBH from OPMG where DocNum=z.U_GOITHAU) as 'DPBH'
			  ,(Select U_CPQLCT from OPMG where DocNum=z.U_GOITHAU) as 'CPQLCT'
			  ,SUM(z.GTHD) as 'GTHD'
			  ,SUM(z.GGTM) as 'GGTM'
			  ,SUM(z.PA) as 'PA'
			  ,SUM(z.PhiQL) as 'PhiQL'
			  ,SUM(z.PLHD) as 'PLHD'
			  ,SUM(z.KHAC) as 'KHAC'
			  ,SUM(z.GTHD) + SUM(z.GGTM) + SUM(z.PA) + SUM(z.PhiQL) + SUM(z.PLHD) + SUM(z.KHAC) as 'Total'
					from 
					(
						Select 
						a.U_GOITHAU
						,SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
						,SUM(a.U_GGTM) as 'GGTM'
						,SUM(a.U_PADXTK) as 'PA'
						,SUM(a.U_PQL) as 'PhiQL'
						,'0' as 'PLHD'
						,'0' as 'KHAC'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 47
						and a.BpType = 'C'
						and a.[Status] ='A'
						and a.Cancelled <> 'Y'
						and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
						group by a.U_GOITHAU

						union all

						Select 
						t1.U_GOITHAU
						,'0' as 'GTHD'
						,'0' as 'GGTM'
						,'0' as 'PA'
						,'0' as 'PhiQL'
						,SUM(t1.PLHD) as PLHD
						,'0' as 'KHAC'
						from (
						Select a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 142
						and a.BpType = 'C'
						and a.[Status] ='A'
						and a.Cancelled <> 'Y'
						and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
						) t1
						group by t1.U_GOITHAU

						union all

						Select 
						t2.U_GOITHAU
						,'0' as 'GTHD'
						,'0' as 'GGTM'
						,'0' as 'PA'
						,'0' as 'PhiQL'
						,'0' as 'PLHD'
						,SUM(t2.KHAC) as KHAC from (
						Select a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
						from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
						where a.U_PRJ = @FinancialProject
						and a.Series = 203
						and a.[Status] ='A'
						and a.BpType = 'C'
						and a.Cancelled <> 'Y'
						and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t2
						group by U_GOITHAU

					) z
					group by z.U_GOITHAU
					order by z.U_GOITHAU
			--Select SUM(z.GTHD) as 'GTHD'
			--,SUM(z.GGTM) as 'GGTM'
			--,SUM(z.PA) as 'PA'
			--,SUM(z.PhiQL) as 'PhiQL'
			--,SUM(z.PLHD) as 'PLHD'
			--,SUM(z.KHAC) as 'KHAC'
			--from 
			--(
			--	Select SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
			--	,SUM(a.U_GGTM) as 'GGTM'
			--	,SUM(a.U_PADXTK) as 'PA'
			--	,SUM(a.U_PQL) as 'PhiQL'
			--	,'0' as 'PLHD'
			--	,'0' as 'KHAC'
			--	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
			--	where a.U_PRJ = @FinancialProject
			--	and a.Series = 47
			--	and a.BpType = 'C'
			--	and a.[Status] ='A'
			--	and a.Cancelled <> 'Y'
			--	and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
				
			--	union all
				
			--	Select '0' as 'GTHD'
			--	,'0' as 'GGTM'
			--	,'0' as 'PA'
			--	,'0' as 'PhiQL'
			--	,SUM(t1.PLHD) as PLHD
			--	,'0' as 'KHAC'
			--	from (
			--	Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
			--	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
			--	where a.U_PRJ = @FinancialProject
			--	and a.Series = 142
			--	and a.BpType = 'C'
			--	and a.[Status] ='A'
			--	and a.Cancelled <> 'Y'
			--	and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t1
			--	union all
			--	Select 
			--	'0' as 'GTHD'
			--	,'0' as 'GGTM'
			--	,'0' as 'PA'
			--	,'0' as 'PhiQL'
			--	,'0' as 'PLHD'
			--	,SUM(t2.KHAC) as KHAC from (
			--	Select case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
			--	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
			--	where a.U_PRJ = @FinancialProject
			--	and a.Series = 203
			--	and a.BpType = 'C'
			--	and a.[Status] ='A'
			--	and a.Cancelled <> 'Y'
			--	and (Select AbsEntry from OPMG where DocNum = a.U_GOITHAU) in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) t2
			--) z
	END
END

GO

--Bao cao tai chinh thiet bi
--Danh sach du an - Bao cao theo doi tuong
ALTER PROCEDURE [dbo].[EQ_GET_FPROJECT]
	-- Add the parameters for the stored procedure here
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2017' and T0.[Active] = 'Y'
END;
GO

--Du tru theo doi tuong
ALTER PROCEDURE [dbo].[EQ_CE_O_GET_DATA_DUTRU]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as int
	,@CTG_Entry as int
	,@Type as varchar
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DUTRU_DocEntry as int;
	Select top 1 @DUTRU_DocEntry = DocEntry 
	from [@DUTRU] 
	where U_CTG_Key = 
		(Select top 1 DocEntry 
		from [@CTG] 
		where U_PrjCode = @FinancialProject
		and U_GoiThauKey =  @GoiThauKey
		order by DocEntry desc)
	and U_DUTRU_TYPE=2 
	order by DocEntry desc;
	if(@Type = 'S')
		Select * from [@DUTRUA] where DocEntry = @DUTRU_DocEntry;
	else if (@Type = 'D')
		Select * from [@DUTRUB] where DocEntry = @DUTRU_DocEntry;
END
GO

--Danh sach du an  - Bao cao theo hang muc
ALTER PROCEDURE [dbo].[EQ_CE_GET_FPROJECT]
	-- Add the parameters for the stored procedure here
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2017' and T0.[Active] = 'Y'
END;
GO

--Du tru theo hang muc
ALTER PROCEDURE [dbo].[EQ_CE_GET_DATA_DUTRU]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as int
	,@CTG_Entry as int
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @CTG_DocEntry as int;
    -- Insert statements for procedure here
	Select top 1 @CTG_DocEntry = DocEntry from [@CTG] where U_PrjCode = @FinancialProject and U_GoiThauKey =  @GoiThauKey order by DocEntry desc;
	if (@GoiThauKey = -1)
		Select b.*,c.U_TLTB
		from [@CTG] a inner join [@CTG2] b on a.DocEntry = b.DocEntry
		inner join OITM c on b.U_MATHIETBI = c.ItemCode
		where a.U_PrjCode = @FinancialProject
		and b.DocEntry = @CTG_DocEntry;
	else
		Select b.* ,c.U_TLTB
		from [@CTG] a inner join [@CTG2] b on a.DocEntry = b.DocEntry
		inner join OITM c on b.U_MATHIETBI = c.ItemCode
		where a.U_PrjCode = @FinancialProject
		and b.DocEntry = @CTG_DocEntry;
END

GO

--Bao cao tai chinh theo doi tuong Thiet bi
ALTER PROCEDURE [dbo].[EQ_FR_GET_DATA]
	@FinancialProject as varchar(100)
	,@GoiThauKey as int
	,@CTG_Entry as int
	,@Type as varchar
	,@ToDate as date
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DUTRU_DocEntry as int;
	Select top 1 @DUTRU_DocEntry = DocEntry 
	from [@DUTRU] 
	where U_CTG_Key = 
		(Select top 1 DocEntry 
		from [@CTG] 
		where U_PrjCode = @FinancialProject
		and U_GoiThauKey =  @GoiThauKey
		order by DocEntry desc)
	and U_DUTRU_TYPE=2 
	order by DocEntry desc;
	if(@Type = 'S')
		Select * from [@DUTRUA] where DocEntry = @DUTRU_DocEntry;
	else if (@Type = 'D')
		Select a.U_BPCode,a.U_BPName,a.U_CP_VC+a.U_CP_MB+a.U_CP_T+a.U_CP_VH as 'CP',ISNULL(c.Total,0)  as 'GT'
		,case when b.Series in (70,71) then 'NCC' 
		  when b.Series in (72,73) then 'NTP' 
		  else 'UNKNOW'
		end as 'TYPE'
		from [@DUTRUB] a inner join [OCRD] b on a.U_BPCode = b.CardCode
		left join 
		(Select b.CardCode,SUM(ISNULL(a.LineTotal,0))  as Total from WTR1 a inner join OWTR b on a.DocEntry = b.DocEntry where b.Project=@FinancialProject group by b.CardCode) c on a.U_BPCode = c.CardCode
		where a.DocEntry = @DUTRU_DocEntry
END
GO

--Kiem soat khoi luong theo vat tu
ALTER PROCEDURE [dbo].[MM_QC_ITEM]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(100)
	,@ToDate as date
AS
BEGIN
if (@GoiThauKey = '')
begin
Select * from (
	Select x.*,ISNULL(y.KL_DN,0) as 'KL_DN',x.U_DVT as 'DVT_NCC'
	from 
	(Select U_ITEMNO
	,U_ITEMNAME
	,U_DVT
	--,SUM(d.U_KLDT)
	--,SUM(a.U_DinhMuc)
	,SUM(d.U_KLDT * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BOQ'
	--,d.U_003
	,SUM(d.U_003 * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BV'
	From [@CTG1] a inner join [@CTG] b on a.DocEntry = b.DocEntry 
	and b.U_PrjCode= @FinancialProject 
	and b.DocEntry in (Select x.CTG_KEY from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) x)
	inner join OPMG c on b.U_PrjCode = c.FIPROJECT and c.STATUS <> 'T'
	inner join OPHA d on d.ProjectID = c.AbsEntry and a.U_001 = d.U_001
	where U_ITEMNO is not null
	group by U_ITEMNO,U_ITEMNAME,U_DVT) x
	left join
	(Select a.ItemCode,a.Dscription,Unitmsr,SUM(a.Quantity) as 'KL_DN' from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
	where a.Project = @FinancialProject
	and b.DocDate <= @ToDate
	group by ItemCode,Dscription,Unitmsr) y
	on x.U_ITEMNO = y.ItemCode and x.U_DVT = y.unitMsr
	
	Union 

	Select y.ItemCode as 'U_ITEMNO',y.Dscription as 'U_ITEMNAME', x.U_DVT as 'U_DVT' , 0 as 'KL_BOQ',0 as 'KL_BV'
	,ISNULL(y.KL_DN,0) as 'KL_DN', y.unitMsr as 'DVT_NCC' 
	from 
	(Select U_ITEMNO
	,U_ITEMNAME
	,U_DVT
	--,SUM(d.U_KLDT)
	--,SUM(a.U_DinhMuc)
	,SUM(d.U_KLDT * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BOQ'
	--,d.U_003
	,SUM(d.U_003 * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BV'
	From [@CTG1] a inner join [@CTG] b on a.DocEntry = b.DocEntry 
	and b.U_PrjCode= @FinancialProject
	and b.DocEntry in (Select x.CTG_KEY from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) x)
	inner join OPMG c on b.U_PrjCode = c.FIPROJECT and c.STATUS <> 'T'
	inner join OPHA d on d.ProjectID = c.AbsEntry and a.U_001 = d.U_001
	where U_ITEMNO is not null
	group by U_ITEMNO,U_ITEMNAME,U_DVT) x
	inner join
	(Select a.ItemCode,a.Dscription,a.Unitmsr, SUM(a.Quantity) as 'KL_DN' from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
	where a.Project = @FinancialProject
	and b.DocDate <= @ToDate
	group by ItemCode,Dscription,Unitmsr) y
	on x.U_ITEMNO = y.ItemCode and x.U_DVT <> y.unitMsr) T0
	WHERE SUBSTRING(T0.U_ITEMNO,1,2) <> 'NC'
	order by T0.U_ITEMNO
end
else
begin
	Select x.*,ISNULL(y.KL_DN,0) as 'KL_DN',x.U_DVT as 'DVT_NCC'
	from 
	(Select U_ITEMNO
	,U_ITEMNAME
	,U_DVT
	--,SUM(d.U_KLDT)
	--,SUM(a.U_DinhMuc)
	,SUM(d.U_KLDT * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BOQ'
	--,d.U_003
	,SUM(d.U_003 * a.U_DinhMuc * a.U_HAOHUT) as 'KL_BV'
	From [@CTG1] a inner join [@CTG] b on a.DocEntry = b.DocEntry and b.U_PrjCode= @FinancialProject 
	and b.DocEntry in (Select z.CTG_KEY from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,',')) group by U_GoiThauKey) z)
	inner join OPMG c on b.U_PrjCode = c.FIPROJECT and c.STATUS <> 'T'
	inner join OPHA d on d.ProjectID = c.AbsEntry and a.U_001 = d.U_001
	where U_ITEMNO is not null
	group by U_ITEMNO,U_ITEMNAME,U_DVT) x
	left join
	(Select a.ItemCode,a.Dscription,SUM(a.Quantity) as 'KL_DN' from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
	where a.Project = @FinancialProject
	and b.DocDate <= @ToDate
	and a.U_ParentID1 in (Select AbsEntry from OPHA where ProjectID in (Select splitdata from dbo.fnSplitString(@GoiThauKey,',')) and Level = 0)
	group by ItemCode,Dscription) y
	on x.U_ITEMNO = y.ItemCode
	order by x.U_ITEMNO;
end
END
GO

--Kiem soat khoi luong theo hang muc
ALTER PROCEDURE [dbo].[MM_QC_ITEM_HM]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@GoiThauKey as varchar(100)
	,@ToDate as date
AS
BEGIN
if (@GoiThauKey = '')
begin
	Select T0.HM_Key,T0.HM_CODE,T0.HM_NAME,T0.CT_Key,T0.CT_CODE,T0.CT_NAME,T0.CT_KLDT,T0.CT_KLBV,T0.CT_DVT
	,T0.CV_Key,T0.CV_CODE,T0.CV_NAME,T0.CV_KLDT,T0.CV_KLBV,T0.CV_DVT 
	,ISNULL(T1.NTP,0) as NTP,T1.DV_NTP
	,ISNULL(T2.DTC,0) as DTC,T2.DV_DTC
	from
	(
	Select a.HM_Key,a.HM_CODE,a.HM_NAME,b.CT_Key,b.CT_CODE,b.CT_NAME,b.CT_KLDT,b.CT_KLBV,b.CT_DVT,c.CV_Key,c.CV_CODE,c.CV_NAME,c.CV_KLDT,c.CV_KLBV,c.CV_DVT from
	(Select AbsEntry as 'HM_Key',U_001 as 'HM_CODE',NAME as 'HM_NAME'
	From OPHA 
	where Level =1 
	and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FinancialProject)) a

	inner join

	(Select AbsEntry as 'CT_Key',U_001 as 'CT_CODE',NAME as 'CT_NAME',ParentID as 'CT_ParentID',U_KLDT as 'CT_KLDT',U_003 as 'CT_KLBV',U_002 as 'CT_DVT'
	From OPHA where Level =2 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FinancialProject)) b
	on a.HM_Key = b.CT_ParentID and a.HM_CODE not in ('HT','PRELIM','TB','BPTC')
	left join 

	(Select AbsEntry as 'CV_Key',U_001 as 'CV_CODE',NAME as 'CV_NAME',ParentID as 'CV_ParentID',U_KLDT as 'CV_KLDT',U_003 as 'CV_KLBV',U_002 as 'CV_DVT'
	From OPHA where Level =3 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FinancialProject)) c
	on b.CT_Key = c.CV_ParentID) T0
	left join
	(
	Select U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5, SUM(Quantity) as 'NTP',unitMsr as 'DV_NTP', 0 as 'DTC' ,'' as 'DV_DTC'
	from PDN1 a inner join OPDN b on a.DocEntry=b.DocEntry
	where a.Project=@FinancialProject
	and b.U_PUTYPE='PUT02'
	and U_ParentID1 is not null
	and b.CANCELED not in ('Y','C')
	and b.DocDate < @ToDate
	group by U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5,unitMsr
	) T1
	on T0.HM_Key= T1.U_ParentID2 and T0.CT_Key = T1.U_ParentID3 and T0.CV_Key = T1.U_ParentID4 and T0.CV_DVT = T1.DV_NTP
	left join 
	(
	Select U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5, SUM(Quantity) as 'DTC',unitMsr as 'DV_DTC'
	from PDN1 a inner join OPDN b on a.DocEntry=b.DocEntry
	where a.Project=@FinancialProject
	and b.U_PUTYPE='PUT09'
	and U_ParentID1 is not null
	and b.CANCELED not in ('Y','C')
	and b.DocDate < @ToDate
	group by U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5,unitMsr
	) T2
	on T0.HM_Key= T2.U_ParentID2 and T0.CT_Key = T2.U_ParentID3 and T0.CV_Key = T2.U_ParentID4 and T0.CV_DVT = T2.DV_DTC
end
else
begin
	Select T0.HM_Key,T0.HM_CODE,T0.HM_NAME,T0.CT_Key,T0.CT_CODE,T0.CT_NAME,T0.CT_KLDT,T0.CT_KLBV,T0.CT_DVT
	,T0.CV_Key,T0.CV_CODE,T0.CV_NAME,T0.CV_KLDT,T0.CV_KLBV,T0.CV_DVT 
	,ISNULL(T1.NTP,0) as NTP,T1.DV_NTP
	,ISNULL(T2.DTC,0) as DTC,T2.DV_DTC
	from
	(
	Select a.HM_Key,a.HM_CODE,a.HM_NAME,b.CT_Key,b.CT_CODE,b.CT_NAME,b.CT_KLDT,b.CT_KLBV,b.CT_DVT,c.CV_Key,c.CV_CODE,c.CV_NAME,c.CV_KLDT,c.CV_KLBV,c.CV_DVT from
	(Select AbsEntry as 'HM_Key',U_001 as 'HM_CODE',NAME as 'HM_NAME'
	From OPHA 
	where Level =1 
	and ProjectId in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))) a

	inner join

	(Select AbsEntry as 'CT_Key',U_001 as 'CT_CODE',NAME as 'CT_NAME',ParentID as 'CT_ParentID',U_KLDT as 'CT_KLDT',U_003 as 'CT_KLBV',U_002 as 'CT_DVT'
	From OPHA where Level =2 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FinancialProject)) b
	on a.HM_Key = b.CT_ParentID and a.HM_CODE not in ('HT','PRELIM','TB','BPTC')
	left join 

	(Select AbsEntry as 'CV_Key',U_001 as 'CV_CODE',NAME as 'CV_NAME',ParentID as 'CV_ParentID',U_KLDT as 'CV_KLDT',U_003 as 'CV_KLBV',U_002 as 'CV_DVT'
	From OPHA where Level =3 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FinancialProject)) c
	on b.CT_Key = c.CV_ParentID) T0
	left join
	(
	Select U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5, SUM(Quantity) as 'NTP',unitMsr as 'DV_NTP', 0 as 'DTC' ,'' as 'DV_DTC'
	from PDN1 a inner join OPDN b on a.DocEntry=b.DocEntry
	where a.Project=@FinancialProject
	and b.U_PUTYPE='PUT02'
	and U_ParentID1 is not null
	and b.CANCELED not in ('Y','C')
	and b.DocDate < @ToDate
	group by U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5,unitMsr
	) T1
	on T0.HM_Key= T1.U_ParentID2 and T0.CT_Key = T1.U_ParentID3 and T0.CV_Key = T1.U_ParentID4 and T0.CV_DVT = T1.DV_NTP
	left join 
	(
	Select U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5, SUM(Quantity) as 'DTC',unitMsr as 'DV_DTC'
	from PDN1 a inner join OPDN b on a.DocEntry=b.DocEntry
	where a.Project=@FinancialProject
	and b.U_PUTYPE='PUT09'
	and U_ParentID1 is not null
	and b.CANCELED not in ('Y','C')
	and b.DocDate < @ToDate
	group by U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5,unitMsr
	) T2
	on T0.HM_Key= T2.U_ParentID2 and T0.CT_Key = T2.U_ParentID3 and T0.CV_Key = T2.U_ParentID4 and T0.CV_DVT = T2.DV_DTC

	--Select x.*,ISNULL(y.KL_DN,0) as 'KL_DN'
	--from 
	--(Select U_ITEMNO
	--,U_ITEMNAME
	--,U_DVT
	----,SUM(d.U_KLDT)
	----,SUM(a.U_DinhMuc)
	--,SUM(d.U_KLDT) * SUM(a.U_DinhMuc) as 'KL_BOQ'
	----,d.U_003
	--,SUM(d.U_003) * SUM(a.U_DinhMuc) as 'KL_BV'
	--From [@CTG1] a inner join [@CTG] b on a.DocEntry = b.DocEntry and b.U_PrjCode= @FinancialProject 
	--and b.DocEntry in (Select z.CTG_KEY from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,',')) group by U_GoiThauKey) z)
	--inner join OPMG c on b.U_PrjCode = c.FIPROJECT and c.STATUS <> 'T'
	--inner join OPHA d on d.ProjectID = c.AbsEntry and a.U_001 = d.U_001
	--where U_ITEMNO is not null
	--group by U_ITEMNO,U_ITEMNAME,U_DVT) x
	--left join
	--(Select a.ItemCode,a.Dscription,SUM(a.Quantity) as 'KL_DN' from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
	--where a.Project = @FinancialProject
	--and b.DocDate <= @ToDate
	--and a.U_ParentID1 in (Select AbsEntry from OPHA where ProjectID in (Select splitdata from dbo.fnSplitString(@GoiThauKey,',')) and Level = 0)
	--group by ItemCode,Dscription) y
	--on x.U_ITEMNO = y.ItemCode
	--order by x.U_ITEMNO;
end
END

GO

/*Danh sach du an: 
	Bao cao Tong hop thanh toan ky
  , Liet ke thanh toan ky
  , Thong ke duyet hop dong
  , Hao hut thep
*/
ALTER PROCEDURE [dbo].[CCM_SUMMARY_GET_FPROJECT]
	-- Add the parameters for the stored procedure here
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2017' and T0.[Active] = 'Y'
	and T0.[PrjCode] in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
END
GO

--Tong hop thanh toan ky: Data Bill
ALTER PROCEDURE [dbo].[CCM_SUMMARY_BILL_GET_DATA]
	@FProject as varchar(200)
	,@Period as int
AS
BEGIN
Select
X.DocEntry
,X.U_BPCode
,X.U_BPName
,case X.U_BGroup when 'CDXD' then 'CD' else X.U_BGroup end as 'U_BGroup'
,X.U_PUType_Origin
,X.U_Period
,X.U_BType
,X.U_DATETO
,X.Canceled
,dbo.fnPUType_Convert(X.U_PUType_Origin,X.U_BPCode) as 'U_PUType'
,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
	where userId = (Select UserId from OUSR where User_Code = 
(Select top 1 U_Usr from [@KLTT_APPROVE] c where c.DocEntry=X.DocEntry and c.U_Status is not null order by  c.LineId desc))) as 'Last Approved by'
,(Select top 1 U_Time from [@KLTT_APPROVE] c where c.DocEntry=X.DocEntry and c.U_Status is not null order by  c.LineId desc) as 'Last Approved on'
,(Select top 1 U_Status from [@KLTT_APPROVE] c where c.DocEntry=X.DocEntry and c.U_Level= 1 and c.U_Position = 1 order by  c.LineId desc) as 'CCM Approve'
,(Select top 1 U_Status from [@KLTT_APPROVE] c where c.DocEntry=X.DocEntry and c.U_Level= -2 and c.U_Position = 1 order by  c.LineId desc) as 'KT Approve'
from
(Select T1.DocEntry,T0.U_BPCode,T1.U_BPName,T0.U_BGroup,T0.U_PUType as 'U_PUType_Origin',T1.U_Period,T1.U_BType,T1.U_DATETO,T1.Canceled from 
	(Select  U_BPCODE,U_BGroup ,U_PUTYPE, Max(U_Period) as Max_Period
	from [@KLTT]
	where U_FIPROJECT = @FProject
	and U_Period <= @Period
	group by U_BPCODE, U_BGroup, U_PUTYPE) T0
	inner join 
	[@KLTT] T1 on T0.U_BPCode=T1.U_BPCode 
			and T0.U_BGroup = T1.U_BGroup 
			and T0.U_PUType = T1.U_PUType 
			and T0.Max_Period = T1.U_Period
			and T1.U_FIPROJECT = @FProject
) X;
END
GO

--Tong hop thanh toan ky: Data Ban chi huy
ALTER PROCEDURE [dbo].[CCM_SUMMARY_BILL_GET_DATA_BCH]
	@Period as int
	,@FProject as varchar(200)
AS
BEGIN
Select ISNULL(T0.U_MACP,T1.U_MACP) as 'MA_CP', ISNULL(T0.U_TENCP,T1.U_TENCP) as 'TEN_CP'
,T0.CP as 'CP_KY_NAY'
,T1.CP as 'CP_KY_TRUOC' 
,T0.[Last Approved by]
,T0.[Last Approved on]
,T0.[CCM Approve]
,T0.[KT Approve]
from
	(
	Select X.*
	,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
	where userId = (Select UserId from OUSR where User_Code = 
	(Select top 1 U_Usr from [@JV_APROVE_D] where DocEntry=Y.DocEntry and U_Status is not null order by  LineId desc))) as 'Last Approved by'
	,(Select top 1 U_Time from [@JV_APROVE_D] where DocEntry=Y.DocEntry and U_Status is not null order by  LineId desc) as 'Last Approved on'
	,(Select top 1 U_Status from [@JV_APROVE_D] where DocEntry=Y.DocEntry and U_Level= 1 and U_Position = 1 order by  LineId desc) as 'CCM Approve'
	,(Select top 1 U_Status from [@JV_APROVE_D] where DocEntry=Y.DocEntry and U_Level= -2 and U_Position = 1 order by  LineId desc) as 'KT Approve'
	from
	(Select a.BatchNum, a.U_LCP, a.U_KTT
	,a.Project,b.U_MACP,b.U_TENCP,SUM(b.Debit) as 'CP'
	from OBTF a inner join BTF1 b on a.BatchNum = b.BatchNum
	where a.Project= @FProject
	and a.U_LCP ='BCH'
	and a.U_KTT = @Period
	group by a.BatchNum, a.U_LCP, a.U_KTT ,a.Project,b.U_MACP,b.U_TENCP) X

	left join 

	[@JV_APPROVE] Y on X.BatchNUM = Y.U_JVBatchNum
	) T0
full join
	(
	Select b.Project,b.U_MACP,b.U_TENCP,SUM(b.Debit) as 'CP'
	from OBTF a inner join BTF1 b on a.BatchNum = b.BatchNum
	where b.Project= @FProject
	--and a.U_LCP ='BCH'
	and a.U_KTT < @Period
	group by b.Project,b.U_MACP,b.U_TENCP
	) T1
on T0.U_MACP = T1.U_MACP
order by MA_CP;
END
GO

--Thong ke Duyet bill KLTT: Data
ALTER PROCEDURE [dbo].[CCM_DUYET_BILL_GET_DATA]
	-- Add the parameters for the stored procedure here
	@FProject as varchar(200)
	,@Fr_Period as int
	,@To_Period as int
AS
BEGIN
Select T0.DocEntry,T0.U_FIPROJECT,T0.U_BGroup
,dbo.fnPUType_Convert(T0.U_PUTYPE,T0.U_BPCode) 
as 'U_PUTYPE_Parse'
,T0.U_PUType as 'U_PUType',T0.U_Period,T0.U_BPCode,T0.U_BPName,T0.U_BType,T0.U_DATETO
,Case T0.Canceled when 'Y' then 'Rejected' else '' end as 'Rejected'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = 3 ) as 'CHT'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = 2 and U_Position = 1) as 'TB'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = 5 and U_Position = 1) as 'CD'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = 1 and U_Position = 1) as 'CCM'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = 6 and U_Position = 3) as 'GDDA'
,(Select top 1 Convert(Datetime,U_Time,103) from [@KLTT_APPROVE] where DocEntry = T0.DocEntry and U_Level = -2 and U_Position = 1) as 'KT'
 from [@KLTT] T0 inner join OCRD T1 on T0.U_BPCode = T1.CardCode
 where T0.U_Period >= @Fr_Period
 and T0.U_Period <= @To_Period
 and T0.U_FIPROJECT = @FProject
 order by U_Period,U_BPCode asc ;
END
GO

--Thong ke duyet Hop dong: Data
ALTER PROCEDURE [dbo].[CCM_SUMMARY_HD_GET_DATA]
	-- Add the parameters for the stored procedure here
	@FProject as varchar(200)
	,@FrDate as Date
	,@ToDate as Date
AS
BEGIN
Select Z.AbsID,Z.[Agreement No],Z.Project,Z.BpCode,Z.BpName,Z.Descript,Z.Status,Z.[Contract Group],Z.[Purchase Type]
,Z.[GTHD]
,ISNULL(Z.Date_Appr1 ,Z.Date_Appr2) as 'CHT'
,Z.Date_Appr4 as 'PC'
,case when Z.[Contract Group] = 'CD' or Z.[Contract Group] = 'CDXD' then  Z.Date_Appr6 end as 'ME'
,case when Z.[Contract Group] = 'TBXD' then  Z.Date_Appr6 end as 'TB'
,Z.Date_Appr8 as 'KT'
,Z.Date_Appr10 as 'CCM'
,Z.Date_Appr11 as 'PGD'
from
(Select AbsId , Number as 'Agreement No',U_PRJ as 'Project',BpCode,BpName,Descript,a.Status,U_CGroup as 'Contract Group',U_PUTYPE as 'Purchase Type' 
,(Select Format( (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) ,'N0','en-US' ) from OAT1 b where b.AgrNo = AbsId) as N'GTHD'
,Convert(Datetime,U_DTApprv1,103) as 'Date_Appr1'
,Convert(Datetime,U_DTApprv2,103) as 'Date_Appr2'
,Convert(Datetime,U_DTApprv4,103) as 'Date_Appr4'
,Convert(Datetime,U_DTApprv6,103) as 'Date_Appr6'
,Convert(Datetime,U_DTApprv8,103) as 'Date_Appr8'
,Convert(Datetime,U_DTApprv10,103) as 'Date_Appr10'
,Convert(Datetime,U_DTApprv11,103) as 'Date_Appr11'
,b.dept 
from OOAT a left join OHEM b on a.UserSign = b.userId
where U_PRJ = @FProject
and a.StartDate >= @FrDate
and a.StartDate <= @ToDate) Z;
END
GO

--Hao hut thep: Data
ALTER PROCEDURE [dbo].[CCM_HAOHUT_THEP_GET_DATA]
	 @FProject as varchar(200)
	,@ToDate as date
AS
BEGIN
Select T0.HM_Key,T0.HM_CODE,T0.HM_NAME,T0.CT_Key,T0.CT_CODE,T0.CT_NAME,T0.CT_KLDT,T0.CT_KLBV,T0.CT_DVT
	,T0.CV_Key,T0.CV_CODE,T0.CV_NAME,T0.CV_KLDT,T0.CV_KLBV,T0.CV_DVT 
	,ISNULL(T1.KLNV,0) as KLNV,T1.DVT
	,T2.U_pthoanthanh/100 as 'PT_HOANTHANH'
	,T2.U_klnguyen as 'KL_NGUYEN'
	,T2.U_klvun as 'KL_VUN'
	from
	(
	Select a.HM_Key,a.HM_CODE,a.HM_NAME,b.CT_Key,b.CT_CODE,b.CT_NAME,b.CT_KLDT,b.CT_KLBV,b.CT_DVT,c.CV_Key,c.CV_CODE,c.CV_NAME,c.CV_KLDT,c.CV_KLBV,c.CV_DVT from
	(Select AbsEntry as 'HM_Key',U_001 as 'HM_CODE',NAME as 'HM_NAME'
	From OPHA 
	where Level =1 
	and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FProject)) a

	inner join

	(Select AbsEntry as 'CT_Key',U_001 as 'CT_CODE',NAME as 'CT_NAME',ParentID as 'CT_ParentID',U_KLDT as 'CT_KLDT',U_003 as 'CT_KLBV',U_002 as 'CT_DVT'
	From OPHA where Level =2 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FProject)) b
	on a.HM_Key = b.CT_ParentID 
	and a.HM_CODE not in ('HT','PRELIM','TB','BPTC') 
	and b.CT_CODE='Re'
	left join 

	(Select AbsEntry as 'CV_Key',U_001 as 'CV_CODE',NAME as 'CV_NAME',ParentID as 'CV_ParentID',U_KLDT as 'CV_KLDT',U_003 as 'CV_KLBV',U_002 as 'CV_DVT'
	From OPHA where Level =3 and ProjectId in (Select AbsEntry from OPMG where FIPROJECT=@FProject)
	and UPPER(U_002)='KG' and substring(U_001,1,2)='Re') c
	on b.CT_Key = c.CV_ParentID) T0
	left join
	(
	Select U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5, SUM(Quantity) as 'KLNV',unitMsr as 'DVT'
	from PDN1 a inner join OPDN b on a.DocEntry=b.DocEntry
	where a.Project=@FProject
	and U_ParentID1 is not null
	and b.CANCELED not in ('Y','C')
	and (Select Series from OITM where ItemCode = a.Itemcode) = 81
	and b.DocDate <= @ToDate
	group by U_ParentID1,U_ParentID2,U_ParentID3,U_ParentID4,U_ParentID5,unitMsr
	) T1
	on T0.HM_Key= T1.U_ParentID2 and T0.CT_Key = T1.U_ParentID3 and T0.CV_Key = T1.U_ParentID4 and T0.CV_DVT = T1.DVT

	left join
	(
	Select a.U_Project,b.U_Makhuvuc as 'HM_CODE',b.U_tenkhuvuc as 'HM_NAME'
	,b.U_macongtac as 'CT_CODE',b.U_tencongtac as 'CT_NAME'
	,b.U_mahangmuc as 'CV_CODE',b.U_tenhangmuc as 'CV_NAME'
	,b.U_dvt,b.U_pthoanthanh,b.U_klnguyen,b.U_klvun--,b.U_haohut
	from [@DGPTHT1] a inner join [@DGPTHT2] b on a.DocEntry=b.DocEntry
	where a.U_Project=@FProject) T2
	on T0.HM_CODE = T2.HM_CODE and T0.CT_CODE = T2.CT_CODE and T0.CV_CODE = T2.CV_CODE and T0.CV_DVT = T2.U_dvt
	order by T0.HM_Key,T0.CT_Key,T0.CV_Key;
END
GO

--Theo doi chi phi Van phong: Danh sach phong ban
ALTER PROCEDURE [dbo].[CCM_DISTRIBUTION_RULE_LIST]
AS
BEGIN
	Select OcrCode,OcrName from OOCR;
END
GO

--Theo doi chi phi Van phong: Data
ALTER PROCEDURE [dbo].[CCM_THEODOI_VP_DETAILS_DATA]
	@VPCODE varchar(200)
   ,@Period as int
AS
BEGIN
	Select 
	(Select U_NCP from [@CPVP] where Code=Z.MA_CP) as 'MA_NHOM_CP'
	,(Select U_TNCP from [@CPVP] where Code=Z.MA_CP) as 'TEN_NHOM_CP'
	,UPPER(Z.MA_CP) as 'MA_CP'
	,(Select top 1 U_TCP from [@CPVP] where [Code] = Z.MA_CP) as 'TEN_CP'
	--,Z.TEN_CP
	,Z.TEN_NCC
	,SUM(Z.GT_VAT) as 'GT_VAT'
	,SUM(Z.GT_NO_VAT) as 'GT_NO_VAT'
	,SUM(Z.KT_VAT) as 'KT_VAT'
	,SUM(Z.KT_NO_VAT) as 'KT_NO_VAT'
	from(
	Select T0.U_MACP as 'MA_CP'
	,'' as 'TEN_CP'
	,T0.TENNCC as 'TEN_NCC'
	,SUM(T0.GT_VAT) as 'GT_VAT'
	,SUM(T1.GT_NO_VAT) as 'GT_NO_VAT'
	,0 as 'KT_VAT'
	,0 as 'KT_NO_VAT'
	,'JV' as 'Doc_Type'
	from
	--CCM with VAT
	(Select x.U_MACP,isnull(U_TENNCC,'') as 'TENNCC',SUM(x.Debit) as GT_VAT
	from (
	Select a.BatchNum,a.TransID,a.Memo,a.U_LCP,b.ProfitCode ,b.U_MACP,b.U_TENCP,a.CreateDate
	,b.U_TENNCC
	,b.Debit
	from OBTF a inner join BTF1 b on a.BatchNum = b.BatchNum
	where b.Project = 'VPCTY' 
	and b.ProfitCode = @VPCODE
	and a.U_KTT <= @Period) x
	group by x.U_MACP,U_TENNCC ) T0 left join
	--CCM No VAT
	(Select y.U_MACP,isnull(U_TENNCC,'') as 'TENNCC',SUM(y.Debit) as GT_NO_VAT
	from (
	Select a.BatchNum,a.TransID,a.Memo,a.U_LCP,b.ProfitCode ,b.U_MACP,b.U_TENCP,a.CreateDate
	,b.U_TENNCC
	,b.Debit
	from OBTF a inner join BTF1 b on a.BatchNum = b.BatchNum
	where b.Project = 'VPCTY' 
	and b.ProfitCode = @VPCODE
	and substring(b.Account,1,4) <> '1331'
	and a.U_KTT <= @Period) y
	group by y.U_MACP,y.U_TENCP,U_TENNCC) T1 
	on T0.U_MACP = T1.U_MACP and T0.TENNCC = T1.TENNCC
	group by T0.U_MACP,T0.TENNCC

	UNION ALL
	--BILLVP
	Select T2.U_MaCP as 'MA_CP'
	,T2.U_TenCP as 'TEN_CP'
	--,T0.U_BPCode as 'MANCC'
	,(Select CardName from OCRD where CardCode = T0.U_BPCode)as 'TENNCC'
	,SUM(T2.U_GrossTotal) as 'GT_VAT'
	,SUM(T2.U_Total) as 'GT_NO_VAT'
	,SUM(T3.KT_VAT) as 'KT_VAT'
	,SUM(T3.KT_NO_VAT) as 'KT_NO_VAT'
	,'VPBILL' as Doc_Type
	from [@BILLVP] T0
	inner join
	(Select U_BPCode,Max(U_Period) as 'Period' from [@BILLVP]
	where U_BPCode is not null
	and U_BType = 2
	and U_Period <= @Period
	and Canceled <> 'Y'
	Group by U_BPCode) T1
	on T0.U_BPCode = T1.U_BPCode and T0.U_Period = T1.Period
	inner join [@BILLVP1] T2 on T0.DocEntry = T2.DocEntry
	left join 
		(Select Convert(decimal,Replace(a.Comments,'Based On Goods Receipt PO ','')) as 'GRPO'
		,b.LineNum
		,b.U_MACP
		,b.U_TENCP
		,b.LineTotal as 'KT_NO_VAT'
		,b.GTotal as 'KT_VAT'
		From OPCH a inner join PCH1 b on a.DocEntry = b.DocEntry 
		where substring(Comments,1,26) = 'Based On Goods Receipt PO'
		and a.Project = 'VPCTY') T3 on  T3.GRPO = T2.U_GRPO_Key  and T3.LineNum = T2.U_GRPO_Line
	where T2.U_DistRule = @VPCODE
	group by T2.U_MaCP,T2.U_TenCP,T0.U_BPCode,T0.U_BPName) Z
	where Z.MA_CP is not null
	and Z.MA_CP != ''
	group by Z.MA_CP, Z.TEN_CP, Z.TEN_NCC
	order by (Select U_NCP from [@CPVP] where Code=Z.MA_CP),Z.MA_CP
END
GO

--Theo doi chi phi Van phong: Data du tru
ALTER PROCEDURE [dbo].[CCM_THEODOI_VP_CE_DATA]
	@VPCODE varchar(200)
   ,@Year as int
AS
BEGIN
	Select --U_IDPB,U_CostCenter,
	U_Code as 'U_MACP',U_Description as 'U_TENCP',U_Amount as 'DuTru'
	from [@KHDTCPHCTB1] 
	where DocEntry = (Select top 1 DocEntry from [@KHDTCPHCTB] where U_Year=@Year order by DocEntry desc)
	and U_IDPB=@VPCODE;
END
GO

--Tong hop chi phi Van phong: Data
ALTER PROCEDURE [dbo].[CCM_TONGHOP_VP_DETAILS_DATA] 
	@ToDate AS DATE
AS
     BEGIN
DECLARE @LUONG_C101 as decimal(19,6)
DECLARE @BHYT_C102 as decimal(19,6)
DECLARE @KHAUHAO as decimal(19,6)

Select @LUONG_C101 = ISNULL(SUM(A1.Debit),0) from OJDT A0 inner join JDT1 A1 on A0.TransID = A1.TransId
where A1.Account like '64211%'
and A0.RefDate <= @ToDate
and A0.TransId not in (SELECT StornoToTr FROM OJDT WHERE StornoToTr IS NOT NULL)
and A0.StornoToTr IS NULL
and A0.RefDate >= datefromparts(YEAR(@ToDate),1,1);

Select @BHYT_C102 = ISNULL(SUM(A1.Debit),0) from OJDT A0 inner join JDT1 A1 on A0.TransID = A1.TransId
where A1.Account like '64212%'
and A0.RefDate <= @ToDate
and A0.TransId not in (SELECT StornoToTr FROM OJDT WHERE StornoToTr IS NOT NULL)
and A0.StornoToTr IS NULL
and A0.RefDate >= datefromparts(YEAR(@ToDate),1,1);

Select @KHAUHAO = ISNULL(SUM(A1.Debit),0) from OJDT A0 inner join JDT1 A1 on A0.TransID = A1.TransId
where A1.Account like '64241%'
and A0.RefDate <= @ToDate
and A0.TransId not in (SELECT StornoToTr FROM OJDT WHERE StornoToTr IS NOT NULL)
and A0.StornoToTr IS NULL
and A0.RefDate >= datefromparts(YEAR(@ToDate),1,1);

SELECT MA_NHOM_CP, TEN_NHOM_CP, MA_CP
--, TEN_CP
,(Select top 1 U_TCP from [@CPVP] where [Code] = MA_CP) as 'TEN_CP'
, DuTru
		, case when MA_CP = 'C101' then @LUONG_C101 
			   when MA_CP = 'C102' then @BHYT_C102
			   when substring(MA_CP,1,2) = 'C8' then @KHAUHAO
				else GT_NO_VAT end as 'GT_NO_VAT'
		, GT_VAT
		, case when MA_CP = 'C101' then @LUONG_C101  
			   when MA_CP = 'C102' then @BHYT_C102 
			   when substring(MA_CP,1,2) = 'C8' then @KHAUHAO
				else KT_NO_VAT end as 'KT_NO_VAT'
		, KT_VAT
FROM ( SELECT ( SELECT U_NCP
                FROM [@CPVP]
                WHERE Code = Y1.MA_CP OR Code = Y2.U_MACP) AS 'MA_NHOM_CP' 
			,(	SELECT U_TNCP
				FROM [@CPVP]
                WHERE Code = Y1.MA_CP OR Code = Y2.U_MACP) AS 'TEN_NHOM_CP' 
			, ISNULL(Y1.MA_CP , Y2.U_MACP) AS 'MA_CP' 
			--, ISNULL(Y1.TEN_CP , Y2.U_TENCP) AS 'TEN_CP' 
			, Y2.DuTru 
			, ISNULL(Y1.GT_NO_VAT , 0) AS 'GT_NO_VAT' 
			, ISNULL(Y1.GT_VAT , 0) AS 'GT_VAT' 
			, ISNULL(Y1.KT_NO_VAT , 0) AS 'KT_NO_VAT' 
			, ISNULL(Y1.KT_VAT , 0) AS 'KT_VAT'
       FROM ( SELECT Z.MA_CP 
					--, Z.TEN_CP 
					, SUM(Z.GT_VAT) AS 'GT_VAT' 
					, SUM(Z.GT_NO_VAT) AS 'GT_NO_VAT' 
					, SUM(Z.KT_VAT) AS 'KT_VAT' 
					, SUM(Z.KT_NO_VAT) AS 'KT_NO_VAT'
              FROM ( SELECT T0.U_MACP AS 'MA_CP' 
							--, T0.U_TENCP AS 'TEN_CP' 
							, T0.TENNCC AS 'TEN_NCC' 
							, SUM(T0.GT_VAT) AS 'GT_VAT' 
							, SUM(T1.GT_NO_VAT) AS 'GT_NO_VAT' 
							, 0 AS 'KT_VAT' 
							, 0 AS 'KT_NO_VAT' 
							, 'JV' AS 'Doc_Type'
                     FROM ( SELECT x.U_MACP 
								   --, x.U_TENCP 
								   , isnull(U_TENNCC , '') AS 'TENNCC' 
								   , SUM(x.Debit) AS GT_VAT
							FROM ( SELECT a.BatchNum , a.TransID , a.Memo , a.U_LCP , b.ProfitCode , b.U_MACP , b.U_TENCP , a.CreateDate , b.U_TENNCC , b.Debit
								   FROM OBTF AS a INNER JOIN BTF1 AS b ON a.BatchNum = b.BatchNum
								   WHERE b.Project = 'VPCTY'
									--and b.ProfitCode = @VPCODE
									--and a.U_KTT <= @Period
								 ) AS x
							GROUP BY x.U_MACP  , U_TENNCC
                     ) AS T0 
					 
					 LEFT JOIN

                     ( SELECT y.U_MACP 
							  , y.U_TENCP 
							  , isnull(U_TENNCC , '') AS 'TENNCC' 
							  , SUM(y.Debit) AS GT_NO_VAT
                       FROM ( SELECT a.BatchNum , a.TransID , a.Memo , a.U_LCP , b.ProfitCode , b.U_MACP , b.U_TENCP , a.CreateDate , b.U_TENNCC , b.Debit
                              FROM OBTF AS a INNER JOIN BTF1 AS b ON a.BatchNum = b.BatchNum
                              WHERE b.Project = 'VPCTY'
                                    AND --and b.ProfitCode = @VPCODE 
                                    SUBSTRING(b.Account , 1 , 4) <> '1331'
									--and a.U_KTT <= @Period
                            ) AS y
                       GROUP BY y.U_MACP , y.U_TENCP , U_TENNCC
                     ) AS T1 ON T0.U_MACP = T1.U_MACP AND T0.TENNCC = T1.TENNCC
                     GROUP BY T0.U_MACP , T0.TENNCC
                     UNION ALL
					--BILLVP
                     SELECT T2.U_MaCP AS 'MA_CP' 
							--, T2.U_TenCP AS 'TEN_CP' 
							,( SELECT CardName FROM OCRD WHERE CardCode = T0.U_BPCode ) AS 'TENNCC' 
							, SUM(T2.U_GrossTotal) AS 'GT_VAT' 
							, SUM(T2.U_Total) AS 'GT_NO_VAT' 
							, SUM(T3.KT_VAT) AS 'KT_VAT' 
							, SUM(T3.KT_NO_VAT) AS 'KT_NO_VAT' 
							, 'VPBILL' AS Doc_Type
                     FROM [@BILLVP] AS T0 
					 
					 INNER JOIN 
					 
					 ( SELECT U_BPCode 
							, MAX(U_Period) AS 'Period'
                       FROM [@BILLVP]
                       WHERE U_BPCode IS NOT NULL
                       AND U_BType = 2
					   and U_FProject ='VPCTY'
					   --and U_Period <= @Period
                       GROUP BY U_BPCode
                      ) AS T1 ON T0.U_BPCode = T1.U_BPCode AND T0.U_Period = T1.Period
                      
					  INNER JOIN [@BILLVP1] AS T2 ON T0.DocEntry = T2.DocEntry
                      LEFT JOIN 
					  ( SELECT CONVERT(DECIMAL , Replace(a.Comments , 'Based On Goods Receipt PO ' , '')) AS 'GRPO' 
								, b.LineNum 
								, b.U_MACP 
								, b.U_TENCP 
								, b.LineTotal AS 'KT_NO_VAT' 
								, b.GTotal AS 'KT_VAT'
                         FROM OPCH AS a INNER JOIN PCH1 AS b ON a.DocEntry = b.DocEntry
                         WHERE SUBSTRING(Comments , 1 , 26) = 'Based On Goods Receipt PO'
                         AND a.Project = 'VPCTY' ) AS T3 ON T3.GRPO = T2.U_GRPO_Key AND T3.LineNum = T2.U_GRPO_Line
							--where T2.U_DistRule = @VPCODE
                     GROUP BY T2.U_MaCP , T2.U_TenCP , T0.U_BPCode , T0.U_BPName
                   ) AS Z
              WHERE Z.MA_CP IS NOT NULL
              GROUP BY Z.MA_CP
            ) AS Y1 FULL OUTER JOIN ( SELECT U_Code AS 'U_MACP' , U_Description AS 'U_TENCP' , SUM(U_Amount) AS 'DuTru'
                                      FROM [@KHDTCPHCTB1]
                                      WHERE DocEntry = ( SELECT TOP 1 DocEntry
                                                         FROM [@KHDTCPHCTB]
                                                         WHERE U_Year = 2018
                                                         ORDER BY DocEntry DESC
                                                       )
                                            AND
                                            U_Code IS NOT NULL
                                      GROUP BY U_Code , U_Description
                                    ) AS Y2 ON Y1.MA_CP = Y2.U_MACP
     ) AS X
WHERE X.MA_NHOM_CP IS NOT NULL
ORDER BY CONVERT(INT , SUBSTRING(X.MA_CP , 2 , LEN(X.MA_CP)-1));
END;
GO

--Thong ke Doanh thu - Loi nhuan: Data
ALTER PROCEDURE [dbo].[CCM_DT_LN_Project_List]
	  @FrDate as date
	, @ToDate as date
AS
BEGIN
Select T0.PrjCode, T0.PrjName,T1.AbsEntry,T1.[NAME],T1.CARDNAME
,(Select [Descr] from UFD1 
	where TableID='OPMG' 
	and FieldID = (Select FieldID from CUFD where TableID='OPMG' 
	and AliasID='PRJGROUP') and FldValue=T1.U_PRJGROUP) as 'PRJGROUP'
,(Select SeriesName from NNM1 where ObjectCode=234000021 and Series= T1.Series) as 'PRJTYPE'
,T2.GTHD
from OPRJ T0 inner join OPMG T1 on T0.PrjCode=T1.FIPROJECT
and T1.[STATUS] <> 'N'
left join 
(Select 
z.U_PRJ
,z.U_GOITHAU
,SUM(z.GTHD) as 'GTHD'
,SUM(z.GGTM) as 'GGTM'
,SUM(z.PA) as 'PA'
,SUM(z.PhiQL) as 'PhiQL'
,SUM(z.PLHD) as 'PLHD'
,SUM(z.KHAC) as 'KHAC'
from (
	Select 
	a.U_PRJ
	,a.U_GOITHAU
	,SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
	,SUM(a.U_GGTM) as 'GGTM'
	,SUM(a.U_PADXTK) as 'PA'
	,SUM(a.U_PQL) as 'PhiQL'
	,'0' as 'PLHD'
	,'0' as 'KHAC'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 47
	and a.BpType = 'C'
	group by a.U_PRJ,a.U_GOITHAU
	union
	Select 
	t1.U_PRJ
	,t1.U_GOITHAU
	,'0' as 'GTHD'
	,'0' as 'GGTM'
	,'0' as 'PA'
	,'0' as 'PhiQL'
	,SUM(t1.PLHD) as PLHD
	,'0' as 'KHAC'
	from (
	Select a.U_PRJ,a.U_GoiTHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 142
	and a.BpType = 'C'
	) t1
	group by t1.U_PRJ,t1.U_GOITHAU
	union
	Select 
	t2.U_PRJ
	,t2.U_GOITHAU
	,'0' as 'GTHD'
	,'0' as 'GGTM'
	,'0' as 'PA'
	,'0' as 'PhiQL'
	,'0' as 'PLHD'
	,SUM(t2.KHAC) as KHAC from (
	Select a.U_PRJ,a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 161
	and a.BpType = 'C') t2
	group by t2.U_PRJ,t2.U_GOITHAU) z
	group by z.U_PRJ,z.U_GOITHAU)T2 on T0.PrjCode=T2.U_PRJ and T1.DocNum = T2.U_GOITHAU
where T0.Active ='Y'
and T0.ValidFrom >= '01-Jan-2017'
and T0.PrjCode <> 'VTTB'
and ISNULL(T0.U_HOANTHANH,'N') <> 'Y'

order by T0.[PrjCode],T1.AbsEntry;
END
GO

--Doanh thu - Loi nhuan: A Index theo Goi thau
ALTER PROCEDURE [dbo].[CCM_BASELINE_A_INDEX]
	@BASELINE_DocEntry as int
	,@FinancialProject as varchar(250)
	,@ProjectId as int
AS
DECLARE @DOANHTHU as decimal(19,6)
DECLARE @Chiphi_DUTRU as decimal(19,6)
DECLARE @Chiphi_BCH as decimal(19,6)
DECLARE @Chiphi_DUPHONG as decimal(19,6)
DECLARE @Chiphi_HOTRO as decimal(19,6)
BEGIN
	if (@BASELINE_DocEntry = -1)
		Select top 1 @BASELINE_DocEntry= t0.DocEntry from [@BASELINE] t0 inner join BASELINE_OPMG t1 
			on t0.DocEntry = t1.DocEntry_BaseLine and t1.AbsEntry=@ProjectId
		where t0.[Status]= 'C'
		and t0.[Canceled] <> 'Y'
		and t0.U_FProject = @FinancialProject
		order by DocEntry asc;
	--Du tru
	Select @Chiphi_DUTRU = SUM(ISNULL(U_CP_NCC,0) + ISNULL(U_CP_NTP,0) + ISNULL(U_CP_DTC,0) + ISNULL(U_CP_VTP,0) 
			 + ISNULL(U_CP_VC,0)  + ISNULL(U_CP_VH,0)  + ISNULL(U_CP_CN,0)  + ISNULL(U_CP_DP,0)
			 + ISNULL(U_CP_DP2,0) + + ISNULL(U_CP_K,0)) 
	from BASELINE_DUTRUB where DocEntry_BaseLine = @BASELINE_DocEntry
	and DocEntry_DUTRU in 
					(Select DocEntry
					from [BASELINE_DUTRU] 
					where DUTRU_TYPE = 1
					and DocEntry_BaseLine = @BASELINE_DocEntry
					and CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [BASELINE_CTG] 
								where DocEntry_BaseLine = @BASELINE_DocEntry
								and U_GoiThauKey = @ProjectID
								group by U_GoiThauKey) a));

	--BCH
	Select @Chiphi_BCH = SUM(ISNULL(U_GTDP,0))
	FROM [BASELINE_CTG4] 
	where DocEntry_BaseLine = @BASELINE_DocEntry
	and DocEntry_CTG in (Select DocEntry from [BASELINE_CTG]  where U_GoiThauKey = @ProjectID)
	and ISNUMERIC(U_TKKT) = 1
	and U_TKKT is not null;
	
	DECLARE @DP1 as decimal(19,6)
	DECLARE @DPBH as decimal(19,6)
	DECLARE @HT1 as decimal(19,6)
	DECLARE @HT2 as decimal(19,6)
	DECLARE @Chiphi_NG as decimal(19,6)
	DECLARE @Table_HD TABLE(
		U_GOITHAU int,
		HT1 decimal(19,6),
		HT2 decimal(19,6),
		CPNG decimal(19,6),
		DPCP decimal(19,6),
		DPBH decimal(19,6),
		CPQLCT decimal(19,6),
		GTHD decimal(19,6),
		GTTM decimal(19,6),
		PA decimal(19,6),
		PhiQL decimal(19,6),
		PLHD decimal(19,6),
		KHAC decimal(19,6),
		Total decimal(19,6)
	);
	--Get HD cho Du an 
	Insert into @Table_HD(U_GOITHAU, HT1, HT2, CPNG, DPCP, DPBH, CPQLCT, GTHD, GTTM, PA, PhiQL, PLHD, KHAC, Total)
	Exec [dbo].[BASELINE_MM_FI_GET_DATA_VII] @BASELINE_DocEntry, @ProjectID;

	Select @DOANHTHU = SUM(ISNULL(Total,0)) --Doanh thu
	,@DP1 = SUM(ISNULL(Total * DPCP/100 ,0)) --Phan tram chi phi du phong
	,@DPBH = SUM(ISNULL(Total * DPBH/100 ,0)) --Phan tram chi phi du phong bao hanh
	,@HT1 = SUM(ISNULL(Total * HT1/100 ,0))  --Ho tro 1
	,@HT2 = SUM(ISNULL(Total * HT2/100 ,0))  --Ho tro 2
	,@Chiphi_NG = SUM(ISNULL(CPNG,0)) -- Chi phi ngoai giao
	from @Table_HD

	--Chi phi DU PHONG
	SET @Chiphi_DUPHONG = @DP1 + @DPBH;
	--Chi phi HOTRO
	SET @Chiphi_HOTRO = @HT1 + @HT2 + @Chiphi_NG

	Select @DOANHTHU as 'Doanhthu'
	, @Chiphi_DUTRU as 'Chiphi_DUTRU'
	, @Chiphi_BCH as 'Chiphi_BCH'
	, @Chiphi_DUPHONG as 'Chiphi_DP'
	, @Chiphi_HOTRO as 'Chiphi_HT'
	,(@DOANHTHU-(@Chiphi_DUTRU+ @Chiphi_BCH+@Chiphi_DUPHONG+@Chiphi_HOTRO))/@DOANHTHU as 'A-INDEX';
END
GO

--Doanh thu - Loi nhuan: A Index theo Goi thau, ngay
ALTER PROCEDURE [dbo].[CCM_BASELINE_DATE_A_INDEX]
	@BASELINE_DocEntry as int
	,@FinancialProject as varchar(250)
	,@ProjectId as int
	,@ToDate as datetime
AS
DECLARE @DOANHTHU as decimal(19,6)
DECLARE @Chiphi_DUTRU as decimal(19,6)
DECLARE @Chiphi_BCH as decimal(19,6)
DECLARE @Chiphi_DUPHONG as decimal(19,6)
DECLARE @Chiphi_HOTRO as decimal(19,6)
BEGIN
	if (@BASELINE_DocEntry = -1)
		Select top 1 @BASELINE_DocEntry= t0.DocEntry from [@BASELINE] t0 inner join BASELINE_OPMG t1 
			on t0.DocEntry = t1.DocEntry_BaseLine and t1.AbsEntry=@ProjectId
		where t0.[Status]= 'C'
		and t0.[Canceled] <> 'Y'
		and t0.U_FProject = @FinancialProject
		and t0.U_BaseDate <= @ToDate
		order by DocEntry desc;
	--Du tru
	Select @Chiphi_DUTRU = SUM(ISNULL(U_CP_NCC,0) + ISNULL(U_CP_NTP,0) + ISNULL(U_CP_DTC,0) + ISNULL(U_CP_VTP,0) 
			 + ISNULL(U_CP_VC,0)  + ISNULL(U_CP_VH,0)  + ISNULL(U_CP_CN,0)  + ISNULL(U_CP_DP,0)
			 + ISNULL(U_CP_DP2,0) + + ISNULL(U_CP_K,0)) 
	from BASELINE_DUTRUB where DocEntry_BaseLine = @BASELINE_DocEntry
	and DocEntry_DUTRU in 
					(Select DocEntry
					from [BASELINE_DUTRU] 
					where DUTRU_TYPE = 1
					and DocEntry_BaseLine = @BASELINE_DocEntry
					and CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [BASELINE_CTG] 
								where DocEntry_BaseLine = @BASELINE_DocEntry
								and U_GoiThauKey = @ProjectID
								group by U_GoiThauKey) a));

	--BCH
	Select @Chiphi_BCH = SUM(ISNULL(U_GTDP,0))
	FROM [BASELINE_CTG4] 
	where DocEntry_BaseLine = @BASELINE_DocEntry
	and DocEntry_CTG in (Select DocEntry from [BASELINE_CTG]  where U_GoiThauKey = @ProjectID)
	and ISNUMERIC(U_TKKT) = 1
	and U_TKKT is not null;
	
	DECLARE @DP1 as decimal(19,6)
	DECLARE @DPBH as decimal(19,6)
	DECLARE @HT1 as decimal(19,6)
	DECLARE @HT2 as decimal(19,6)
	DECLARE @Chiphi_NG as decimal(19,6)
	DECLARE @Table_HD TABLE(
		U_GOITHAU int,
		HT1 decimal(19,6),
		HT2 decimal(19,6),
		CPNG decimal(19,6),
		DPCP decimal(19,6),
		DPBH decimal(19,6),
		CPQLCT decimal(19,6),
		GTHD decimal(19,6),
		GTTM decimal(19,6),
		PA decimal(19,6),
		PhiQL decimal(19,6),
		PLHD decimal(19,6),
		KHAC decimal(19,6),
		Total decimal(19,6)
	);
	--Get HD cho Du an 
	Insert into @Table_HD(U_GOITHAU, HT1, HT2, CPNG, DPCP, DPBH, CPQLCT, GTHD, GTTM, PA, PhiQL, PLHD, KHAC, Total)
	Exec [dbo].[BASELINE_MM_FI_GET_DATA_VII] @BASELINE_DocEntry, @ProjectID;

	Select @DOANHTHU = SUM(ISNULL(Total,0)) --Doanh thu
	,@DP1 = SUM(ISNULL(Total * DPCP/100 ,0)) --Phan tram chi phi du phong
	,@DPBH = SUM(ISNULL(Total * DPBH/100 ,0)) --Phan tram chi phi du phong bao hanh
	,@HT1 = SUM(ISNULL(Total * HT1/100 ,0))  --Ho tro 1
	,@HT2 = SUM(ISNULL(Total * HT2/100 ,0))  --Ho tro 2
	,@Chiphi_NG = SUM(ISNULL(CPNG,0)) -- Chi phi ngoai giao
	from @Table_HD

	--Chi phi DU PHONG
	SET @Chiphi_DUPHONG = @DP1 + @DPBH;
	--Chi phi HOTRO
	SET @Chiphi_HOTRO = @HT1 + @HT2 + @Chiphi_NG

	Select @DOANHTHU as 'Doanhthu'
	, @Chiphi_DUTRU as 'Chiphi_DUTRU'
	, @Chiphi_BCH as 'Chiphi_BCH'
	, @Chiphi_DUPHONG as 'Chiphi_DP'
	, @Chiphi_HOTRO as 'Chiphi_HT'
	,(@DOANHTHU-(@Chiphi_DUTRU+ @Chiphi_BCH+@Chiphi_DUPHONG+@Chiphi_HOTRO))/@DOANHTHU as 'A-INDEX';
END
GO

--Doanh thu - Loi nhuan: A Index theo Du an
ALTER PROCEDURE [dbo].[CCM_BASELINE_FPROJECT_A_INDEX]
	@BASELINE_DocEntry as int
	,@FinancialProject as varchar(250)
AS
DECLARE @DOANHTHU as decimal(19,6)
DECLARE @Chiphi_DUTRU as decimal(19,6)
DECLARE @Chiphi_BCH as decimal(19,6)
DECLARE @Chiphi_DUPHONG as decimal(19,6)
DECLARE @Chiphi_HOTRO as decimal(19,6)
BEGIN
	if (@BASELINE_DocEntry = -1)
		Select top 1 @BASELINE_DocEntry= t0.DocEntry from [@BASELINE] t0 
		where t0.[Status]= 'C'
		and t0.[Canceled] <> 'Y'
		and t0.U_FProject = @FinancialProject
		order by DocEntry asc;
	--Du tru
	Select @Chiphi_DUTRU = SUM(ISNULL(U_CP_NCC,0) + ISNULL(U_CP_NTP,0) + ISNULL(U_CP_DTC,0) + ISNULL(U_CP_VTP,0) 
			 + ISNULL(U_CP_VC,0)  + ISNULL(U_CP_VH,0)  + ISNULL(U_CP_CN,0)  + ISNULL(U_CP_DP,0)
			 + ISNULL(U_CP_DP2,0) + + ISNULL(U_CP_K,0)) 
	from BASELINE_DUTRUB where DocEntry_BaseLine = @BASELINE_DocEntry

	--BCH
	Select @Chiphi_BCH = SUM(ISNULL(U_GTDP,0))
	FROM [BASELINE_CTG4] 
	where DocEntry_BaseLine = @BASELINE_DocEntry
	and U_TKKT not in ('CPQL','CPVTL','MMTB','BCHVP')
	and U_TKKT is not null
	and ISNUMERIC(U_TKKT) = 1;
	
	DECLARE @DP1 as decimal(19,6)
	DECLARE @DPBH as decimal(19,6)
	DECLARE @HT1 as decimal(19,6)
	DECLARE @HT2 as decimal(19,6)
	DECLARE @Chiphi_NG as decimal(19,6)
	DECLARE @Table_HD TABLE(
		U_GOITHAU int,
		HT1 decimal(19,6),
		HT2 decimal(19,6),
		CPNG decimal(19,6),
		DPCP decimal(19,6),
		DPBH decimal(19,6),
		CPQLCT decimal(19,6),
		GTHD decimal(19,6),
		GTTM decimal(19,6),
		PA decimal(19,6),
		PhiQL decimal(19,6),
		PLHD decimal(19,6),
		KHAC decimal(19,6),
		Total decimal(19,6)
	);

	--Get HD cho Du an 
	Insert into @Table_HD(U_GOITHAU, HT1, HT2, CPNG, DPCP, DPBH, CPQLCT, GTHD, GTTM, PA, PhiQL, PLHD, KHAC, Total)
	Exec [dbo].[BASELINE_MM_FI_GET_DATA_VII] @BASELINE_DocEntry, '';

	Select @DOANHTHU = SUM(ISNULL(Total,0)) --Doanh thu
	,@DP1 = SUM(ISNULL(Total * DPCP/100 ,0)) --Phan tram chi phi du phong
	,@DPBH = SUM(ISNULL(Total * DPBH/100 ,0)) --Phan tram chi phi du phong bao hanh
	,@HT1 = SUM(ISNULL(Total * HT1/100 ,0))  --Ho tro 1
	,@HT2 = SUM(ISNULL(Total * HT2/100 ,0))  --Ho tro 2
	,@Chiphi_NG = SUM(ISNULL(CPNG,0)) -- Chi phi ngoai giao
	from @Table_HD

	--Chi phi DU PHONG
	SET @Chiphi_DUPHONG = @DP1 + @DPBH;
	--Chi phi HOTRO
	SET @Chiphi_HOTRO = @HT1 + @HT2 + @Chiphi_NG
	Select @DOANHTHU as 'Doanhthu'
	, @Chiphi_DUTRU as 'Chiphi_DUTRU'
	, @Chiphi_BCH as 'Chiphi_BCH'
	, @Chiphi_DUPHONG as 'Chiphi_DP'
	, @Chiphi_HOTRO as 'Chiphi_HT'
	,(@DOANHTHU-(@Chiphi_DUTRU+ @Chiphi_BCH+@Chiphi_DUPHONG+@Chiphi_HOTRO))/@DOANHTHU as 'A-INDEX';

END
GO

--Doanh thu - Loi nhuan: A Index theo Du an, ngay
ALTER PROCEDURE [dbo].[CCM_BASELINE_FPROJECT_DATE_A_INDEX]
	@BASELINE_DocEntry as int
	,@FinancialProject as varchar(250)
	,@ToDate as datetime
AS
DECLARE @DOANHTHU as decimal(19,6)
DECLARE @Chiphi_DUTRU as decimal(19,6)
DECLARE @Chiphi_BCH as decimal(19,6)
DECLARE @Chiphi_DUPHONG as decimal(19,6)
DECLARE @Chiphi_HOTRO as decimal(19,6)
BEGIN
	if (@BASELINE_DocEntry = -1)
		Select top 1 @BASELINE_DocEntry= t0.DocEntry from [@BASELINE] t0 
		where t0.[Status]= 'C'
		and t0.[Canceled] <> 'Y'
		and t0.U_FProject = @FinancialProject
		and t0.U_BaseDate <= @ToDate
		order by DocEntry desc;
	--Du tru
	Select @Chiphi_DUTRU = SUM(ISNULL(U_CP_NCC,0) + ISNULL(U_CP_NTP,0) + ISNULL(U_CP_DTC,0) + ISNULL(U_CP_VTP,0) 
			 + ISNULL(U_CP_VC,0)  + ISNULL(U_CP_VH,0)  + ISNULL(U_CP_CN,0)  + ISNULL(U_CP_DP,0)
			 + ISNULL(U_CP_DP2,0) + + ISNULL(U_CP_K,0)) 
	from BASELINE_DUTRUB where DocEntry_BaseLine = @BASELINE_DocEntry

	--BCH
	Select @Chiphi_BCH = SUM(ISNULL(U_GTDP,0))
	FROM [BASELINE_CTG4] 
	where DocEntry_BaseLine = @BASELINE_DocEntry
	and U_TKKT not in ('CPQL','CPVTL','MMTB','BCHVP')
	and U_TKKT is not null
	and ISNUMERIC(U_TKKT) = 1;
	
	DECLARE @DP1 as decimal(19,6)
	DECLARE @DPBH as decimal(19,6)
	DECLARE @HT1 as decimal(19,6)
	DECLARE @HT2 as decimal(19,6)
	DECLARE @Chiphi_NG as decimal(19,6)
	DECLARE @Table_HD TABLE(
		U_GOITHAU int,
		HT1 decimal(19,6),
		HT2 decimal(19,6),
		CPNG decimal(19,6),
		DPCP decimal(19,6),
		DPBH decimal(19,6),
		CPQLCT decimal(19,6),
		GTHD decimal(19,6),
		GTTM decimal(19,6),
		PA decimal(19,6),
		PhiQL decimal(19,6),
		PLHD decimal(19,6),
		KHAC decimal(19,6),
		Total decimal(19,6)
	);

	--Get HD cho Du an 
	Insert into @Table_HD(U_GOITHAU, HT1, HT2, CPNG, DPCP, DPBH, CPQLCT, GTHD, GTTM, PA, PhiQL, PLHD, KHAC, Total)
	Exec [dbo].[BASELINE_MM_FI_GET_DATA_VII] @BASELINE_DocEntry, '';

	Select @DOANHTHU = SUM(ISNULL(Total,0)) --Doanh thu
	,@DP1 = SUM(ISNULL(Total * DPCP/100 ,0)) --Phan tram chi phi du phong
	,@DPBH = SUM(ISNULL(Total * DPBH/100 ,0)) --Phan tram chi phi du phong bao hanh
	,@HT1 = SUM(ISNULL(Total * HT1/100 ,0))  --Ho tro 1
	,@HT2 = SUM(ISNULL(Total * HT2/100 ,0))  --Ho tro 2
	,@Chiphi_NG = SUM(ISNULL(CPNG,0)) -- Chi phi ngoai giao
	from @Table_HD

	--Chi phi DU PHONG
	SET @Chiphi_DUPHONG = @DP1 + @DPBH;
	--Chi phi HOTRO
	SET @Chiphi_HOTRO = @HT1 + @HT2 + @Chiphi_NG
	Select @DOANHTHU as 'Doanhthu'
	, @Chiphi_DUTRU as 'Chiphi_DUTRU'
	, @Chiphi_BCH as 'Chiphi_BCH'
	, @Chiphi_DUPHONG as 'Chiphi_DP'
	, @Chiphi_HOTRO as 'Chiphi_HT'
	,(@DOANHTHU-(@Chiphi_DUTRU+ @Chiphi_BCH+@Chiphi_DUPHONG+@Chiphi_HOTRO))/@DOANHTHU as 'A-INDEX';

END
GO

--Tong hop Doanh thu - Loi nhuan: Data
ALTER PROCEDURE [dbo].[CCM_DT_LN_TONGHOP_LST]
	@ToDate as datetime
AS
BEGIN
Select T0.PrjCode, T0.PrjName,T1.AbsEntry,T1.[NAME],T1.CARDNAME
,ISNULL(T1.[OWNER],-1) as 'OWNER'
,(Select ISNULL(lastName+' ','') + ISNULL(middleName + ' ',' ') + ISNULL(firstName,'')  from OHEM where empID = T1.[OWNER]) as 'GDDA'
,(Select [Descr] from UFD1 
	where TableID='OPMG' 
	and FieldID = (Select FieldID from CUFD where TableID='OPMG' 
	and AliasID='PRJGROUP') and FldValue=T1.U_PRJGROUP) as 'PRJGROUP'
,(Select SeriesName from NNM1 where ObjectCode=234000021 and Series= T1.Series) as 'PRJTYPE'
,T2.GTHD
,(Select 
	case MONTH(@ToDate) 
		when 1 then U_DTT1 
		when 2 then U_DTT2 
		when 3 then U_DTT3 
		when 4 then U_DTT4 
		when 5 then U_DTT5 
		when 6 then U_DTT6 
		when 7 then U_DTT7 
		when 8 then U_DTT8 
		when 9 then U_DTT9 
		when 10 then U_DTT10 
		when 11 then U_DTT11 
		when 12 then U_DTT12 
		else 0 end
	from [@DTDA2] 
	where DocEntry = (Select top 1 DocEntry from [@DTDA1] where U_DateFrom <= @ToDate and U_DateTo >= @ToDate order by DocEntry desc)
	and U_PROJECT=T0.PrjCode) as 'DTKehoach'
,(Select SUM(LineTotal) from OINV a inner join INV1 b on a.DocEntry = b.DocEntry
	where a.Project = T0.PrjCode
	and a.CANCELED not in ('Y','C')
	and YEAR(a.DocDate) = YEAR(@ToDate) -1) as 'DTNamtruoc'
,(Select SUM(LineTotal) from OINV a inner join INV1 b on a.DocEntry = b.DocEntry
	where a.Project = T0.PrjCode
	and a.CANCELED not in ('Y','C')
	and YEAR(a.DocDate) = YEAR(@ToDate)
	and a.DocDate <= @ToDate) as 'DTthucte'
from OPRJ T0 inner join OPMG T1 on T0.PrjCode=T1.FIPROJECT and T1.[STATUS] <> 'N'
left join 
(Select 
z.U_PRJ
,z.U_GOITHAU
,SUM(z.GTHD) as 'GTHD'
,SUM(z.GGTM) as 'GGTM'
,SUM(z.PA) as 'PA'
,SUM(z.PhiQL) as 'PhiQL'
,SUM(z.PLHD) as 'PLHD'
,SUM(z.KHAC) as 'KHAC'
from (
	Select 
	a.U_PRJ
	,a.U_GOITHAU
	,SUM(b.PlanQty*b.UnitPrice)+ SUM(b.PlanAmtLC) as 'GTHD'
	,SUM(a.U_GGTM) as 'GGTM'
	,SUM(a.U_PADXTK) as 'PA'
	,SUM(a.U_PQL) as 'PhiQL'
	,'0' as 'PLHD'
	,'0' as 'KHAC'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 47
	and a.BpType = 'C'
	group by a.U_PRJ,a.U_GOITHAU
	union
	Select 
	t1.U_PRJ
	,t1.U_GOITHAU
	,'0' as 'GTHD'
	,'0' as 'GGTM'
	,'0' as 'PA'
	,'0' as 'PhiQL'
	,SUM(t1.PLHD) as PLHD
	,'0' as 'KHAC'
	from (
	Select a.U_PRJ,a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'PLHD'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 142
	and a.BpType = 'C'
	) t1
	group by t1.U_PRJ,t1.U_GOITHAU
	union
	Select 
	t2.U_PRJ
	,t2.U_GOITHAU
	,'0' as 'GTHD'
	,'0' as 'GGTM'
	,'0' as 'PA'
	,'0' as 'PhiQL'
	,'0' as 'PLHD'
	,SUM(t2.KHAC) as KHAC from (
	Select a.U_PRJ,a.U_GOITHAU,case a.Method when 'I' then b.PlanQty*b.UnitPrice else b.PlanAmtLC end as 'KHAC'
	from OOAT a left join OAT1 b on a.AbsID = b.AgrNo
	where a.Series = 161
	and a.BpType = 'C') t2
	group by t2.U_PRJ,t2.U_GOITHAU) z
	group by z.U_PRJ,z.U_GOITHAU)T2 on T0.PrjCode=T2.U_PRJ and (Select AbsEntry from OPMG where DocNum = T2.U_GOITHAU) = T1.AbsEntry
where T0.Active ='Y'
and T0.ValidFrom >= '01-Jan-2017'
and T0.PrjCode <> 'VTTB'
and ISNULL(T0.U_HOANTHANH,'N') <> 'Y'
order by T1.[OWNER],T0.[PrjCode];
END
GO

--San luong: Danh sach Doi tuong 
ALTER PROCEDURE [dbo].[CCM_GET_LST_DT]
AS
BEGIN
Select A.CardCode,A.CardCode+' - ' +B.CardName as 'CardName'from (
Select distinct CARDCODE from OPDN where Project is not null) A inner join OCRD B on A.CardCode = B.CardCode
order by CARDCODE
END
GO

--San luong: Danh sach Hang muc
ALTER PROCEDURE [dbo].[CCM_GET_LST_CT]
AS
BEGIN
	Select distinct A0.U_001,
		(Select top 1 [NAME] from OPHA where U_001=A0.U_001) as 'NAME'
	from OPHA A0 inner join OPMG A1 on A0.ProjectID= A1.AbsEntry
	where A0.Level =2 
	and A0.TYP= 2 
	--and A1.Status <> 'S'
	and A1.U_BPTH = 'XD'
	--and A1.Series not in (135)
	order by U_001;
END
GO

--San luong: Data theo doi tuong
ALTER PROCEDURE [dbo].[CCM_SANLUONG_DATA_DT]
	@FrDate datetime
	,@ToDate datetime
	,@BpCode varchar(50)
AS
BEGIN
	Select A0.Project, A0.CardCode, A0.CardName, A0.U_ParentID3, A0.MA_CT, A0.TEN_CT
	,case A0.NDT when 'PUT01' then 'NCC'
				 when 'PUT02' then 'NTP'
				 when 'PUT09' then 'DTC'
	 else A0.NDT end as 'NDT'
	, A0.NDT as 'NDT_Code'
	,SUM(A0.Total) as 'Total'
	from 
	(
	Select t1.Project,t1.CardCode,t1.CardName,t0.U_ParentID3
	,t2.U_001 as 'MA_CT'
	,t2.[NAME] as 'TEN_CT'
	,t1.U_PUTYPE as 'PUType_Origins'
	,dbo.fnPUType_Convert(t1.U_PUTYPE,t1.CardCode) as 'NDT'
	,t0.LineTotal as 'Total'
	from PDN1 t0 inner join OPDN t1 on t0.DocEntry = t1.DocEntry
	left join OPHA t2 on t0.U_ParentID3 = t2.AbsEntry
	where t1.Project is not null
	and t1.CANCELED not in ('Y','C')
	and t0.U_ParentID3 is not null) A0
	where A0.CardCode = @BpCode
	group by  A0.Project, A0.CardCode, A0.CardName, A0.U_ParentID3, A0.MA_CT, A0.TEN_CT, A0.NDT
	order by A0.Project, A0.CardCode, A0.U_ParentID3;
END
GO

--San luong: Data theo cong tac (Hang muc)
ALTER PROCEDURE [dbo].[CCM_SANLUONG_DATA_CT]
	@FrDate datetime
	,@ToDate datetime
	,@CT varchar(50)
AS
BEGIN
	Select A0.Project, A0.CardCode, A0.CardName, A0.U_ParentID3, A0.MA_CT, A0.TEN_CT, A0.NDT,
	SUM(A0.Total) as 'Total'
	from 
	(
	Select t1.Project,t1.CardCode,t1.CardName,t0.U_ParentID3
	,t2.U_001 as 'MA_CT'
	,t2.[NAME] as 'TEN_CT'
	,dbo.fnPUType_Convert(t1.U_PUTYPE,t1.CardCode) as 'NDT'
	,t0.LineTotal as 'Total'
	from PDN1 t0 inner join OPDN t1 on t0.DocEntry = t1.DocEntry
	left join OPHA t2 on t0.U_ParentID3 = t2.AbsEntry
	where t1.Project is not null
	and t0.U_ParentID3 is not null
	and t1.Canceled not in ('Y','C')) A0
	WHERE A0.MA_CT = @CT
	group by  A0.Project, A0.CardCode, A0.CardName, A0.U_ParentID3, A0.MA_CT, A0.TEN_CT, A0.NDT
	order by A0.Project, A0.U_ParentID3,A0.CardCode
END
GO

--San luong: Gia tri hop dong theo doi tuong
ALTER PROCEDURE [dbo].[CCM_SANLUONG_GTHD]
	@BpCode varchar(50)
	, @FProject as varchar(50)
	, @FrDate datetime
	, @ToDate datetime
AS
BEGIN
	Select (SUM(A1.PlanQty*A1.UnitPrice) + SUM(A1.PlanAmtLC)) as 'GTHD'
	from OOAT A0 inner join OAT1 A1 on A0.AbsId = A1.AgrNo
	where A0.U_PRJ = @FProject
	and A0.Series =48
	and A0.BpCode = @BpCode
	and A0.Status ='A'
	and A0.StartDate <= @ToDate;
END
GO

--San luong: Gia tri hop dong theo cong tac (Hang muc)
ALTER PROCEDURE [dbo].[CCM_SANLUONG_GTHD_CT]
	@BpCode varchar(50)
	, @FProject as varchar(50)
	, @FrDate datetime
	, @ToDate datetime
	, @PuType varchar(50)
AS
BEGIN
	Select (SUM(A1.PlanQty*A1.UnitPrice) + SUM(A1.PlanAmtLC)) as 'GTHD'
	from OOAT A0 inner join OAT1 A1 on A0.AbsId = A1.AgrNo
	where A0.U_PRJ = @FProject
	and A0.Series =48
	and A0.BpCode = @BpCode
	and A0.U_PUTYPE = @PuType
	and A0.Status ='A'
	and A0.StartDate <= @ToDate;
END
GO

--Hieu qua dam phan: Data - Gia tri hop dong
ALTER PROCEDURE [dbo].[CCM_HQDP_GTHD]
	@BpCode varchar(50)
	, @FProject as varchar(50)
	, @FrDate datetime
	, @ToDate datetime
AS
BEGIN
	Select (SUM(A1.PlanQty*A1.UnitPrice) + SUM(A1.PlanAmtLC)) as 'GTHD'
	from OOAT A0 inner join OAT1 A1 on A0.AbsId = A1.AgrNo
	where A0.U_PRJ = @FProject
	and A0.Series =48
	and A0.BpCode = @BpCode
	and A0.Status ='A'
	and A0.StartDate <= @ToDate;
END
GO

--Hieu quan dam phan: Data
ALTER PROCEDURE [dbo].[CCM_HQDP]
	 @FromDate as date
	,@ToDate as date
AS
BEGIN
Select 
Z.Project
, Z.CardCode 
, (Select CardName from OCRD where CardCode = Z.CardCode) as 'CardName'
, Z.U_RECTYPE
--, Z.U_001
, Z.HM_NAME
, Z.NDT
, SUM(Z.Quantity) as 'Quantity'
, SUM(Z.Total_GRPO) as 'Total_GRPO'
, SUM(Z.GT_BOQ) as 'GT_BOQ'
, SUM(Z.GG_DT) as 'GG_DT'
from
(
Select Y.Project, Y.ProjectId, Y.CardCode, Y.U_RECTYPE
, (Select [Name] from OPHA where AbsEntry = (Select ParentID from OPHA where AbsEntry = Y.U_ParentID4)) as 'HM_NAME'
, Y.U_001, Y.NDT, SUM(Y.Quantity) as 'Quantity' 
, SUM(Y.LineTotal) as 'Total_GRPO'
, SUM(Y.GT_BOQ) as 'GT_BOQ'
, SUM(Y.GG_DT) as 'GG_DT'
from
(
Select X0.Project, X0.CardCode, X0.U_RECTYPE, X0.U_ParentID1 as 'ProjectId'
, X0.U_ParentID4
, X0.U_001, X0.NDT, X0.Quantity
, X0.LineTotal 
, (X0.Quantity * X0.U_DG) as 'GT_BOQ'
, (X0.Quantity * X0.U_DGHD) as 'GG_DT'
from 
(
Select A0.Project
, A0.CardCode
, A0.U_RECTYPE
, A1.U_ParentID1
, A1.U_ParentID4
, (Select U_001 from OPHA where AbsEntry = A1.U_ParentID4) as 'U_001'
, A1.ItemCode
, A1.Quantity
, A1.LineTotal
,dbo.fnPUType_Convert(A0.U_PUTYPE,A0.CardCode) as 'NDT'
, ISNULL(A2.U_DG,0) as 'U_DG'
, ISNULL(A2.U_DGHD,0) as 'U_DGHD'
from OPDN A0 inner join PDN1 A1 on A0.DocEntry = A1.DocEntry
left join OPHA A2 on (Select ProjectId from OPHA where AbsEntry = REPLACE(A1.U_ParentID1,',','')) = A2.AbsEntry and A1.U_ParentID4 = A2.AbsEntry
where A1.U_ParentID1 is not null
and A1.U_ParentID3 is not null
and A0.U_RECTYPE is not null
and A0.CANCELED not in ('Y','C')
and A0.DocDate <= @ToDate
and A0.DocDate >= @FromDate) X0
inner join
(
Select A1.U_PrjCode, A0.U_001,A0.U_ITEMNO
from [@CTG1] A0 inner join [@CTG] A1 on A0.DocEntry = A1.DocEntry
where U_001 is not null
and A0.U_ITEMNO is not null
and A1.DocEntry in 
(
	Select DocEntry from [@CTG] B0 inner join
	(Select U_PrjCode,U_GoiThauKey,Max(U_Date) as 'U_DATE' from [@CTG]
	where U_PrjCode is not null 
	and U_GoiThauKey is not null
	group by U_PrjCode,U_GoiThauKey) B1 on B0.U_PrjCode = B1.U_PrjCode and B0.U_GoiThauKey = B1.U_GoiThauKey
	and B0.U_Date = B1.U_DATE
)
) X1
on X0.Project = X1.U_PrjCode and X0.U_001 = X1.U_001 and X0.ItemCode = X1.U_ITEMNO) Y
group by Y.Project,Y.ProjectId,Y.U_ParentID4, Y.CardCode, Y.U_RECTYPE, Y.U_001, Y.NDT) Z
group by Z.Project
, Z.CardCode 
, Z.U_RECTYPE
--, Z.U_001
, Z.HM_NAME
, Z.NDT;
END