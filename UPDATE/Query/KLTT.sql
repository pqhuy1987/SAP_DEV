ALTER PROCEDURE [dbo].[KLTT_GET_ADDITIONALINFO]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BP_Code as varchar(100),
	@Period as int,
	@CGroup as varchar(50),
	@PurchaseType as varchar(50),
	@ToDate as date
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DocEntry as int; -- HD
	DECLARE @DocEntry_HDNT as int; -- HD Nguyen tac
	DECLARE @DocEntry_PLTT as int; -- HD Thay the
	DECLARE @DocEntry_PLT as int; -- PL tang
	DECLARE @HTTU as numeric(19,6);
	DECLARE @GTTU as numeric(19,6);

	--Lay HD Nguyen tac
	Select top 1 @DocEntry_HDNT = isnull(AbsID,-1) 
	from OOAT 
	where U_PRJ is null
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PurchaseType
	and StartDate <= @ToDate
	order by AbsID desc;
	
	--Lay HD
	Select top 1 @DocEntry = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =48
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PurchaseType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL Thay the
	Select top 1 @DocEntry_PLTT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =140
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PurchaseType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL tang
	Select top 1 @DocEntry_PLT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =141
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PurchaseType
	and StartDate <= @ToDate
	order by AbsID desc;
	if (@DocEntry_HDNT > 0)
	begin
		--HD Nguyên t?c n?u có PL thay th? thì l?y PL Thay th?
		if (@DocEntry_PLTT > 0)
			begin
				Select x.*,(x.GTHD * x.PTTU) as 'GTTU'
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLTT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry_PLTT
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
			end
		--else if (@DocEntry_PLT > 0)
		--	begin
			 -- Không x?y ra - PL t?ng gáng trên H? Nguyên t?c
		--	end
		else
			begin
				Select 
					AbsID
					,Number
					,U_SHD
					,StartDate
					,Descript
					,'HDNT' as 'Type'
					,U_PTTU/100 as 'PTTU'
					,U_PTHU/100 as 'PTHU'
					,U_PTBH/100 as 'PTBH'
					,U_PTGL/100 as 'PTGL'
					,U_HTBH as 'HTBH'
					,U_TTTU as 'TTTU'
					,U_CTQLDTC as 'CTQLDTC'
					,0 as 'GTTU'
					,0 as 'GTHD'
				from OOAT 
				where 
					AbsID = @DocEntry_HDNT
					and Status = 'A'
					and Cancelled <> 'Y'
			end
	end
	else if (@DocEntry > 0)
	begin
		--Có H? -- Có PL Thay th? H?
		 Select top 1 @DocEntry_PLTT = isnull(AbsID,-1) from OOAT 
			where U_PRJ = @FinancialProject
			and Series =140
			and BpCode = @BP_Code
			and Status ='A'
			and Cancelled <> 'Y'
			and U_CGroup = @CGroup
			and U_PUTYPE = @PurchaseType
			and StartDate <= @ToDate
			and U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry)
			order by AbsID desc;
		if (@DocEntry_PLTT > 0)
		begin
			--Co PLTT H?p ??ng
			Select x.*,(x.GTHD * x.PTTU) as 'GTTU'
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLTT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,0 as 'GTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry_PLTT
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,0 as 'GTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
		else
		begin
			--Ch? có H? (ho?c có thêm PL t?ng)
			Select x.*,(x.GTHD * x.PTTU) as 'GTTU'
					from (
						--H?p ??ng
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'HD' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,0 as 'GTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,U_CTQLDTC as 'CTQLDTC'
						,0 as 'GTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
	end
END

GO

ALTER PROCEDURE [dbo].[KLTT_GET_FPROJECT]
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

ALTER PROCEDURE [dbo].[KLTT_Get_List_Sub_Level]
	-- Add the parameters for the stored procedure here
	@KLTT_DocEntry as int,
	@Parent_ID as int,
	@Level as int,
	@Type as varchar(10)
AS
BEGIN
	SET NOCOUNT ON;
	if (@Type = 'A')
	begin
		if (@Level =1)
		Select a.U_Sub1,a.U_Sub1Name
		from [@KLTTA] a 
		where a.DocEntry = @KLTT_DocEntry
		group by a.U_Sub1,a.U_Sub1Name;

		if (@Level =2)
		Select a.U_Sub2,a.U_Sub2Name
		from [@KLTTA] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub1 = @Parent_ID
		and a.U_Sub2 != '' 
		and a.U_Sub2 is not null
		group by a.U_Sub2,a.U_Sub2Name;

		if (@Level =3)
		Select a.U_Sub3,a.U_Sub3Name
		from [@KLTTA] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub2 = @Parent_ID
		and a.U_Sub3 != '' 
		and a.U_Sub3 is not null
		group by a.U_Sub3,a.U_Sub3Name;

		if (@Level =4)
		Select a.U_Sub4,a.U_Sub4Name
		from [@KLTTA] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub3 = @Parent_ID
		and a.U_Sub4 != '' 
		and a.U_Sub4 is not null
		group by a.U_Sub4,a.U_Sub4Name;

		if (@Level =5)
		Select a.U_Sub5,a.U_Sub5Name
		from [@KLTTA] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub4 = @Parent_ID
		and a.U_Sub5 != '' 
		and a.U_Sub5 is not null
		group by a.U_Sub5,a.U_Sub5Name;
	end
	if (@Type = 'B')
	begin
		if (@Level =1)
		Select a.U_Sub1,a.U_Sub1Name
		from [@KLTTB] a 
		where a.DocEntry = @KLTT_DocEntry
		group by a.U_Sub1,a.U_Sub1Name;

		if (@Level =2)
		Select a.U_Sub2,a.U_Sub2Name
		from [@KLTTB] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub1 = @Parent_ID
		and a.U_Sub2 != '' 
		and a.U_Sub2 is not null
		group by a.U_Sub2,a.U_Sub2Name;

		if (@Level =3)
		Select a.U_Sub3,a.U_Sub3Name
		from [@KLTTB] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub2 = @Parent_ID
		and a.U_Sub3 != '' 
		and a.U_Sub3 is not null
		group by a.U_Sub3,a.U_Sub3Name;

		if (@Level =4)
		Select a.U_Sub4,a.U_Sub4Name
		from [@KLTTB] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub3 = @Parent_ID
		and a.U_Sub4 != '' 
		and a.U_Sub4 is not null
		group by a.U_Sub4,a.U_Sub4Name;

		if (@Level =5)
		Select a.U_Sub5,a.U_Sub5Name
		from [@KLTTB] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub4 = @Parent_ID
		and a.U_Sub5 != '' 
		and a.U_Sub5 is not null
		group by a.U_Sub5,a.U_Sub5Name;
	end
	if (@Type = 'K')
	begin
		if (@Level =1)
		Select a.U_Sub1,a.U_Sub1Name
		from [@KLTTK] a 
		where a.DocEntry = @KLTT_DocEntry
		group by a.U_Sub1,a.U_Sub1Name;

		if (@Level =2)
		Select a.U_Sub2,a.U_Sub2Name
		from [@KLTTK] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub1 = @Parent_ID
		and a.U_Sub2 != '' 
		and a.U_Sub2 is not null
		group by a.U_Sub2,a.U_Sub2Name;

		if (@Level =3)
		Select a.U_Sub3,a.U_Sub3Name
		from [@KLTTK] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub2 = @Parent_ID
		and a.U_Sub3 != '' 
		and a.U_Sub3 is not null
		group by a.U_Sub3,a.U_Sub3Name;

		if (@Level =4)
		Select a.U_Sub4,a.U_Sub4Name
		from [@KLTTK] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub3 = @Parent_ID
		and a.U_Sub4 != '' 
		and a.U_Sub4 is not null
		group by a.U_Sub4,a.U_Sub4Name;

		if (@Level =5)
		Select a.U_Sub5,a.U_Sub5Name
		from [@KLTTK] a 
		where a.DocEntry = @KLTT_DocEntry
		and a.U_Sub4 = @Parent_ID
		and a.U_Sub5 != '' 
		and a.U_Sub5 is not null
		group by a.U_Sub5,a.U_Sub5Name;
	end
END

GO

ALTER PROCEDURE [dbo].[KLTT_GETDATA]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@To_Date as date,
	@BP_Code as varchar(100),
	@Type as varchar,
	@BGroup as varchar(50),
	@PurchaseType as varchar(50)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @ProjectID as int;
	DECLARE @Last_DocEntry as int
    -- Insert statements for procedure here
	--SELECT top 1 @ProjectID = AbsEntry from OPMG where FIPROJECT = @FinancialProject;
	SELECT @Last_DocEntry = dbo.FN_Get_Last_Period(@BP_Code, @FinancialProject, @BGroup, @PurchaseType);
	IF (@Type = 'A')
		Select 
			--dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			--,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName
			(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as GoiThauKey
			,(Select Name from OPMG where AbsEntry = (Select ProjectID from OPHA where AbsEntry = a.U_ParentID1)) as GoiThauName
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			--,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GPO' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
		left join [@KLTTA] c on c.DocEntry = @Last_DocEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
							and b.U_RECTYPE = @BGroup
		where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = @BGroup
			and b.CardCode = @BP_Code
			and b.DocDate < @To_Date
			and b.U_PUTYPE = @PurchaseType
			and b.CANCELED not in ('Y','C')
			and (Select ISNULL(SUM(TYP),-1) from OPHA where AbsEntry = a.U_ParentID2) not in (11,12,13)
		Union all
		Select
			(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as GoiThauKey
			,(Select Name from OPMG where AbsEntry = (Select ProjectID from OPHA where AbsEntry = a.U_ParentID1)) as GoiThauName
			--dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			--,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName 
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity * -1 as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			--,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GR' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
			from RPD1 a inner join ORPD b on a.DocEntry = b.DocEntry
			left join [@KLTTA] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
			where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = @BGroup
			and b.CardCode = @BP_Code
			and b.U_PUTYPE = @PurchaseType
			and b.DocDate < @To_Date
			and b.CANCELED not in ('Y','C')
			and (Select ISNULL(SUM(TYP),-1) from OPHA where AbsEntry = a.U_ParentID2) not in (11,12,13);
	IF (@Type = 'B')
	Select 
			dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GPO' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
		left join [@KLTTB] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
		where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = 'PS-'+ @BGroup
			and b.CardCode = @BP_Code
			and b.DocDate < @To_Date
			and b.CANCELED not in ('Y','C')
		Union all
		Select
			dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName 
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity*-1 as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GR' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
			from RPD1 a inner join ORPD b on a.DocEntry = b.DocEntry
			left join [@KLTTB] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
			where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = 'PS-'+ @BGroup
			and b.CardCode = @BP_Code
			and b.DocDate < @To_Date
			and b.CANCELED not in ('Y','C');
	IF (@Type = 'C')
		--Goods Issue
		Select
			'' as GoiThauKey
			,'' as GoiThauName
			,'' as SubProjectKey
			,'' as SubProjectName
			,a.DocEntry as GIKey
			,a.LineNum as GIRowKey
			,a.Dscription as DetailsName
			,a.unitMsr as UoM
			,a.Quantity as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,a.U_BPCode as CardCode
			,b.CreateDate
			,'GI' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from IGE1 a inner join OIGE b on a.DocEntry = b.DocEntry
					left join [@KLTTC] c on c.DocEntry = @Last_DocEntry 
										--and c.U_SubProjectKey = c.AbsEntry 
										and c.U_GoodsIssue = a.DocEntry
										and c.U_DetailsKey = a.LineNum
										and ISNULL(c.U_TYPE,'GI') = 'GI'
		where a.Project=@FinancialProject
			and b.U_ISSUETYPE = 1
			and a.U_BPCode = @BP_Code
			and b.DocDate < @To_Date
		UNION ALL
		--Goods Receipt
		Select
			'' as GoiThauKey
			,'' as GoiThauName
			,'' as SubProjectKey
			,'' as SubProjectName
			,a.DocEntry as GIKey
			,a.LineNum as GIRowKey
			,a.Dscription as DetailsName
			,a.unitMsr as UoM
			,a.Quantity as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,a.U_BPCode as CardCode
			,b.CreateDate
			,'GR' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from IGN1 a inner join OIGN b on a.DocEntry = b.DocEntry
					left join [@KLTTC] c on c.DocEntry = @Last_DocEntry 
										--and c.U_SubProjectKey = c.AbsEntry 
										and c.U_GoodsIssue = a.DocEntry
										and c.U_DetailsKey = a.LineNum
										and ISNULL(c.U_TYPE,'') = 'GR'
		where a.Project=@FinancialProject
			and b.U_ISSUETYPE = 1
			and a.U_BPCode = @BP_Code
			and b.DocDate < @To_Date;
	IF (@Type = 'D')
		--Goods Issue
		Select
		'' as GoiThauKey
		,'' as GoiThauName
		,'' as SubProjectKey
		,'' as SubProjectName
		,a.DocEntry as GIKey
		,a.LineNum as GIRowKey
		,a.Dscription as DetailsName
		,a.unitMsr as UoM
		,a.Quantity as Quantity
		,a.Price as UPrice
		,a.LineTotal as Total
		,a.U_BPCode as CardCode
		,b.CreateDate
		,'GI' as 'TYPE'
		,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
		,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from IGE1 a inner join OIGE b on a.DocEntry = b.DocEntry
		left join [@KLTTD] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GoodsIssue = a.DocEntry
							and c.U_DetailsKey = a.LineNum
							and ISNULL(c.U_TYPE,'GI') = 'GI'
		where a.Project=@FinancialProject
		and b.U_ISSUETYPE = 2
		and a.U_BPCode = @BP_Code
		and b.DocDate < @To_Date
		UNION ALL
		--Goods Receipt
		Select
		'' as GoiThauKey
		,'' as GoiThauName
		,'' as SubProjectKey
		,'' as SubProjectName
		,a.DocEntry as GIKey
		,a.LineNum as GIRowKey
		,a.Dscription as DetailsName
		,a.unitMsr as UoM
		,a.Quantity as Quantity
		,a.Price as UPrice
		,a.LineTotal as Total
		,a.U_BPCode as CardCode
		,b.CreateDate
		,'GR' as 'TYPE'
		,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
		,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from IGN1 a inner join OIGN b on a.DocEntry = b.DocEntry
		left join [@KLTTD] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GoodsIssue = a.DocEntry
							and c.U_DetailsKey = a.LineNum
							and ISNULL(c.U_TYPE,'') = 'GR'
		where a.Project=@FinancialProject
		and b.U_ISSUETYPE = 2
		and a.U_BPCode = @BP_Code
		and b.DocDate < @To_Date;
	IF (@Type = 'E')
		Select c.ProjectID as GoiThauKey
			,'' as GoiThauName
			,a.AbsEntry as SubProjectKey
			,a.StageID as StagesKey
			,a.LineID as OpenIssuesKey
			,a.Remarks as Remarks
			,a.U_DVTPS as UoM
			,a.U_KLPS as Quantity
			,a.U_DGPS as UPrice
			,a.EFFORT as Total
			,a.U_NCCPS as CardCode
			,b.START
			,ISNULL(d.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(d.U_CompleteAmount,0) as Last_Complete_Amount
			from PHA2 a inner join PHA1 b on a.AbsEntry = b.AbsEntry and a.StageID = b.LineID
			inner join OPHA c on a.AbsEntry = c.AbsEntry
			left join [@KLTTE] d on d.U_GoithauKey = c.ProjectID
			and d.U_SubProjectKey = a.AbsEntry
			and d.U_StageKey = a.StageID
			and d.U_OpenIssueKey = a.LineID
			and d.DocEntry = @Last_DocEntry 
		where 
			a.U_IssueType= 2
			and c.ProjectId in (Select AbsEntry from OPMG where  FIPROJECT = @FinancialProject)
			and a.U_NCCPS =  @BP_Code
			and b.START <=@To_date;
	IF (@Type = 'F')
		Select c.ProjectID as GoiThauKey
			,'' as GoiThauName
			,a.AbsEntry as SubProjectKey
			,a.StageID as StagesKey
			,a.LineID as OpenIssuesKey
			,a.Remarks as Remarks
			,a.U_DVTPS as UoM
			,a.U_KLPS as Quantity
			,a.U_DGPS as UPrice
			,a.EFFORT as Total
			,a.U_NCCPS as CardCode
			,b.START
			,ISNULL(d.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(d.U_CompleteAmount,0) as Last_Complete_Amount
			from PHA2 a inner join PHA1 b on a.AbsEntry = b.AbsEntry and a.StageID = b.LineID
			inner join OPHA c on a.AbsEntry = c.AbsEntry
			left join [@KLTTF] d on d.U_GoithauKey = c.ProjectID
			and d.U_SubProjectKey = a.AbsEntry
			and d.U_StageKey = a.StageID
			and d.U_OpenIssueKey = a.LineID
			and d.DocEntry = @Last_DocEntry 
		where 
			a.U_IssueType= 3
			and c.ProjectId in (Select AbsEntry from OPMG where  FIPROJECT = @FinancialProject)
			and a.U_NCCPS =  @BP_Code
			and b.START <=@To_date;
	IF (@Type = 'G')
		Select c.ProjectID as GoiThauKey
			,'' as GoiThauName
			,a.AbsEntry as SubProjectKey
			,a.StageID as StagesKey
			,a.LineID as OpenIssuesKey
			,a.Remarks as Remarks
			,a.U_DVTPS as UoM
			,a.U_KLPS as Quantity
			,a.U_DGPS as UPrice
			,a.EFFORT as Total
			,a.U_NCCPS as CardCode
			,b.START
			,ISNULL(d.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(d.U_CompleteAmount,0) as Last_Complete_Amount
			from PHA2 a inner join PHA1 b on a.AbsEntry = b.AbsEntry and a.StageID = b.LineID
			inner join OPHA c on a.AbsEntry = c.AbsEntry
			left join [@KLTTG] d on d.U_GoithauKey = c.ProjectID
			and d.U_SubProjectKey = a.AbsEntry
			and d.U_StageKey = a.StageID
			and d.U_OpenIssueKey = a.LineID
			and d.DocEntry = @Last_DocEntry 
			where 
			a.U_IssueType= 4
			and a.U_NCCPS =  @BP_Code
			and c.ProjectId in (Select AbsEntry from OPMG where  FIPROJECT = @FinancialProject)
			and b.START <=@To_date;
	IF (@Type = 'H')
		Select b.AbsId
			,b.Number
			,b.BpCode
			,b.StartDate
			,b.U_GOITHAU
			,b.Status
			,a.AgrLineNum
			,a.ItemCode
			,a.ItemName
			,a.PlanQty
			,a.InvntryUom
			,a.UnitPrice
			,a.PlanQty*a.UnitPrice as 'Total'
		from OAT1 a left join OOAT b on a.AgrNo = b.AbsID
		where b.U_PRJ = @FinancialProject
			and b.StartDate <= @To_Date
			and b.BpCode = @BP_Code
			and b.Series in (141,140,48)
			and b.U_PUTYPE = @PurchaseType
		order by AbsId;
	IF (@Type = 'K')
	Select 
			dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GPO' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
		from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
		left join [@KLTTK] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
		where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = @BGroup
			and b.CardCode = @BP_Code
			and b.DocDate < @To_Date
			and b.CANCELED not in ('Y','C')
			and (Select ISNULL(TYP,-1) from OPHA where AbsEntry = a.U_ParentID2) in (11,12,13)
		Union all
		Select
			dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
			,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName 
			,a.DocEntry as GRPOKey
			,a.LineNum as GRPORowKey
			,a.Dscription as DetailsName
			,a.U_CTCV as DetailsWork
			,a.unitMsr as UoM
			,a.Quantity*-1 as Quantity
			,a.Price as UPrice
			,a.LineTotal as Total
			,b.CreateDate
			,b.CardCode
			,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
			,a.U_ParentID1
			,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
			,a.U_ParentID2
			,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
			,a.U_ParentID3
			,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
			,a.U_ParentID4
			,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
			,a.U_ParentID5
			,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
			,b.U_RECTYPE
			,'GR' as 'TYPE'
			,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
			from RPD1 a inner join ORPD b on a.DocEntry = b.DocEntry
			left join [@KLTTB] c on c.DocEntry = @Last_DocEntry 
							--and c.U_SubProjectKey = c.AbsEntry 
							and c.U_GPKey = a.DocEntry
							and c.U_GPDetailsKey = a.LineNum
			where a.Project = @FinancialProject
			and a.U_ParentID1 is not null
			and b.U_RECTYPE = @BGroup
			and b.CardCode = @BP_Code
			and b.DocDate < @To_Date
			and b.CANCELED not in ('Y','C')
			and (Select ISNULL(TYP,-1) from OPHA where AbsEntry = a.U_ParentID2) in (11,12,13);
END

GO

ALTER PROCEDURE [dbo].[KLTT_GETLIST]
	-- Add the parameters for the stored procedure here
	@FProject as varchar(200),
	@BP_Code as varchar(100),
	@BGroup as varchar(50),
	@PurchaseType as varchar(50)
AS
BEGIN
	SET NOCOUNT ON;
	Select U_Period as "Period"
	,case U_BType when 1 then N'T?m ?ng'
				  when 2 then N'Thanh toán'
				  when 3 then N'Quy?t toán' end as "Bill Type"
	,U_DATEFROM as "From"
	,U_DATETO as "To"
	,U_CreatedDate as "Created Date"
	,U_FIPROJECT as "Financial Project"
	,DocNum as "Document Number"
	,U_PUType as "Purchase Type"
	,Canceled as 'Rejected'
	from [@KLTT] a
	where a.U_BPCode = @BP_Code
	and a.U_FIPROJECT = @FProject
	and a.U_BGroup = @BGroup
	and a.U_PUType = @PurchaseType
	order by U_Period asc;
END

GO

ALTER PROCEDURE [dbo].[KLTT_GETLIST_APPROVE]
	-- Add the parameters for the stored procedure here
	@BGroup as varchar(50)
	,@Nhantri as varchar(50)
AS
BEGIN
	--Map position OHPS.posID
	SET NOCOUNT ON;
	if (@Nhantri = 'N')
	BEGIN
		IF (@BGroup = 'XD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
			--Select -2 as 'LEVEL', 1 as 'Position';  --pqhuy1987 - 20180619
		IF (@BGroup = 'CD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
			--Select -2 as 'LEVEL', 1 as 'Position';  --pqhuy1987 - 20180619
		IF (@BGroup = 'CDXD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
			--Select -2 as 'LEVEL', 1 as 'Position';  --pqhuy1987 - 20180619
		IF (@BGroup = 'TB')
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
			--Select -2 as 'LEVEL', 1 as 'Position';  --pqhuy1987 - 20180619
		IF (@BGroup = 'TBXD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 2 as 'LEVEL', 2 as 'Position'
			union all
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
			--Select -2 as 'LEVEL', 1 as 'Position';  --pqhuy1987 - 20180619
	END
	-- NTP CHINH lA NHAN TRI
	ELSE IF (@Nhantri = 'YY')	
	BEGIN
		IF (@BGroup = 'XD')
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
		IF (@BGroup = 'CD')
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
		IF (@BGroup = 'CDXD')
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
		IF (@BGroup = 'TB')
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
		IF (@BGroup = 'TBXD')
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select -2 as 'LEVEL', null as 'Position'; --pqhuy1987 - 20180619
	END
	--NHAN TRI
	--ELSE IF (@Nhantri = 'NTP00599')	
	ELSE IF (@Nhantri = 'NTP00611') --update for PRO	
	BEGIN
		IF (@BGroup = 'XD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 20 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 20 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CDXD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 20 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'TB')
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 20 as 'LEVEL', null as 'Position';

		IF (@BGroup = 'TBXD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 2 as 'LEVEL', 2 as 'Position'
			union all
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 20 as 'LEVEL', null as 'Position';
	END
	--NHAN TIEN
	ELSE IF (@Nhantri = 'NTP00601')
	BEGIN
		IF (@BGroup = 'XD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 21 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 21 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CDXD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 21 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'TB')
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 21 as 'LEVEL', null as 'Position';

		IF (@BGroup = 'TBXD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 2 as 'LEVEL', 2 as 'Position'
			union all
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 21 as 'LEVEL', null as 'Position';
	END
	--NHAN TIN
	ELSE IF (@Nhantri = 'NTP00602')
	BEGIN
		IF (@BGroup = 'XD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 22 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 22 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'CDXD')
			Select 3 as 'LEVEL', 6 as 'Position'
			union all
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 5 as 'LEVEL', 2 as 'Position'
			union all
			Select 5 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 22 as 'LEVEL', null as 'Position';
		IF (@BGroup = 'TB')
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 22 as 'LEVEL', null as 'Position';

		IF (@BGroup = 'TBXD')
			Select 3 as 'LEVEL', 5 as 'Position'
			union all
			Select 2 as 'LEVEL', 2 as 'Position'
			union all
			Select 2 as 'LEVEL', 1 as 'Position'
			union all
			Select 1 as 'LEVEL', 2 as 'Position'
			union all
			Select 1 as 'LEVEL', 1 as 'Position'
			union all
			Select 6 as 'LEVEL', 3 as 'Position'
			union all
			Select 22 as 'LEVEL', null as 'Position';
	END
END

GO

ALTER PROCEDURE [dbo].[KLTT_LOADDATA]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BP_Code as varchar(100),
	@Period as int,
	@Type as varchar,
	@DocEntry as int
AS
BEGIN
	SET NOCOUNT ON;
	if (@DocEntry = 0)
		Select @DocEntry = DocEntry from [@KLTT] where U_Period = @Period and U_FIPROJECT = @FinancialProject;
	IF (@Type = 'A')
		Select 
		 [DocEntry]
		  ,[LineId]
		  ,[VisOrder]
		  ,[Object]
		  ,[LogInst]
		  ,[U_SubProjectKey]
		  ,[U_SubProjectName]
		  ,[U_UPrice]
		  ,[U_Sum]
		  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		  ,[U_CompleteAmount]
		  ,[U_Quantity]
		  ,[U_GoiThauKey]
		  ,[U_GoiThau]
		  ,[U_GPKey]
		  ,[U_GPDetailsKey]
		  ,[U_GPDetailsName]
		  ,[U_UoM]
		  ,[U_CTCV]
		  ,[U_Sub1]
		  ,[U_Sub2]
		  ,[U_Sub3]
		  ,[U_Sub4]
		  ,[U_Sub5]
		  ,[U_Sub1Name]
		  ,[U_Sub2Name]
		  ,[U_Sub3Name]
		  ,[U_Sub4Name]
		  ,[U_Sub5Name]
      ,[U_Type]--DocEntry,LineId,U_GoiThauKey,U_GoiThau, U_SubProjectKey,U_SubProjectName,U_GPKey,U_GPDetailsKey,U_GPDetailsName,U_CTCV,U_UoM,U_Quantity,U_UPrice,U_Sum,U_CompleteAmount,U_CompleteRate/100 as 'U_CompleteRate'
		from [@KLTTA] 
		where DocEntry = @DocEntry
		order by U_Sub1,U_Sub2,U_Sub3,U_Sub4,U_Sub5;
	IF (@Type = 'B')
		Select 
		 [DocEntry]
		  ,[LineId]
		  ,[VisOrder]
		  ,[Object]
		  ,[LogInst]
		  ,[U_SubProjectKey]
		  ,[U_SubProjectName]
		  ,[U_UPrice]
		  ,[U_Sum]
		  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		  ,[U_CompleteAmount]
		  ,[U_Quantity]
		  ,[U_GoiThauKey]
		  ,[U_GoiThau]
		  ,[U_GPKey]
		  ,[U_GPDetailsKey]
		  ,[U_GPDetailsName]
		  ,[U_UoM]
		  ,[U_CTCV]
		  ,[U_Sub1]
		  ,[U_Sub2]
		  ,[U_Sub3]
		  ,[U_Sub4]
		  ,[U_Sub5]
		  ,[U_Sub1Name]
		  ,[U_Sub2Name]
		  ,[U_Sub3Name]
		  ,[U_Sub4Name]
		  ,[U_Sub5Name]
      ,[U_Type]--DocEntry,LineId,U_GoiThauKey,U_GoiThau, U_SubProjectKey,U_SubProjectName,U_GPKey,U_GPDetailsKey,U_GPDetailsName,U_CTCV,U_UoM,U_Quantity,U_UPrice,U_Sum,U_CompleteAmount,U_CompleteRate/100 as 'U_CompleteRate'
		from [@KLTTB] 
		where DocEntry = @DocEntry
		order by U_Sub1,U_Sub2,U_Sub3,U_Sub4,U_Sub5;
	IF (@Type = 'C')
		Select [DocEntry]
			  ,[LineId]
			  ,[VisOrder]
			  ,[Object]
			  ,[LogInst]
			  ,[U_GoodsIssue]
			  ,[U_DetailsKey]
			  ,[U_DetailsName]
			  ,[U_UoM]
			  ,[U_UPrice]
			  ,[U_Quantity]
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'U_SUM'
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'U_CompleteAmount'
			  ,[U_GoiThauKey]
			  ,[U_GoiThau]
			  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		from [@KLTTC] 
		where DocEntry = @DocEntry;
	IF (@Type = 'D')
		Select [DocEntry]
			  ,[LineId]
			  ,[VisOrder]
			  ,[Object]
			  ,[LogInst]
			  ,[U_GoodsIssue]
			  ,[U_DetailsKey]
			  ,[U_DetailsName]
			  ,[U_UoM]
			  ,[U_UPrice]
			  ,[U_Quantity]
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'U_SUM'
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'U_CompleteAmount'
			  ,[U_GoiThauKey]
			  ,[U_GoiThau]
			  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		from [@KLTTD] 
		where DocEntry = @DocEntry;
	IF (@Type = 'E')
		Select [DocEntry]
			  ,[LineId]
			  ,[VisOrder]
			  ,[Object]
			  ,[LogInst]
			  ,[U_SubprojectKey]
			  ,[U_StageKey]
			  ,[U_OpenIssueKey]
			  ,[U_OpenIssueRemark]
			  ,[U_UoM]
			  ,[U_UPrice]
			  ,[U_Quantity]
			  ,[U_Sum]
			  ,[U_CompleteAmount]
			  ,[U_GoiThauKey]
			  ,[U_GoiThau]
			  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		from [@KLTTE] 
		where DocEntry = @DocEntry;
	IF (@Type = 'F')
		Select [DocEntry]
			  ,[LineId]
			  ,[VisOrder]
			  ,[Object]
			  ,[LogInst]
			  ,[U_SubProjectKey]
			  ,[U_StageKey]
			  ,[U_OpenIssueKey]
			  ,[U_OpenIssueRemark]
			  ,[U_UoM]
			  ,[U_UPrice]
			  ,[U_Quantity]
			  ,[U_Sum]
			  ,[U_CompleteAmount]
			  ,[U_GoiThauKey]
			  ,[U_GoiThau]
			  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		from [@KLTTF] 
		where DocEntry = @DocEntry;
	IF (@Type = 'G')
		Select [DocEntry]
			  ,[LineId]
			  ,[VisOrder]
			  ,[Object]
			  ,[LogInst]
			  ,[U_SubProjectKey]
			  ,[U_StageKey]
			  ,[U_OpenIssueKey]
			  ,[U_OpenIssueRemark]
			  ,[U_UoM]
			  ,[U_UPrice]
			  ,[U_Quantity]
			  ,[U_Sum]
			  ,[U_CompleteAmount]
			  ,[U_GoiThauKey]
			  ,[U_GoiThau]
			  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		from [@KLTTG] 
		where DocEntry = @DocEntry;
	IF (@Type = 'H')
		Select *
		from [@KLTTH] 
		where DocEntry = @DocEntry;
	IF (@Type = 'K')
		Select 
		 [DocEntry]
		  ,[LineId]
		  ,[VisOrder]
		  ,[Object]
		  ,[LogInst]
		  --,[U_SubProjectKey]
		  --,[U_SubProjectName]
		  ,[U_UPrice]
		  ,[U_Sum]
		  ,[U_CompleteRate]/100 as 'U_CompleteRate'
		  ,[U_CompleteAmount]
		  ,[U_Quantity]
		  ,[U_GoiThauKey]
		  ,[U_GoiThau]
		  ,[U_GPKey]
		  ,[U_GPDetailsKey]
		  ,[U_GPDetailsName]
		  ,[U_UoM]
		  ,[U_CTCV]
		  ,[U_Sub1]
		  ,[U_Sub2]
		  ,[U_Sub3]
		  ,[U_Sub4]
		  ,[U_Sub5]
		  ,[U_Sub1Name]
		  ,[U_Sub2Name]
		  ,[U_Sub3Name]
		  ,[U_Sub4Name]
		  ,[U_Sub5Name]
      ,[U_Type]--DocEntry,LineId,U_GoiThauKey,U_GoiThau, U_SubProjectKey,U_SubProjectName,U_GPKey,U_GPDetailsKey,U_GPDetailsName,U_CTCV,U_UoM,U_Quantity,U_UPrice,U_Sum,U_CompleteAmount,U_CompleteRate/100 as 'U_CompleteRate'
		from [@KLTTK] 
		where DocEntry = @DocEntry
		order by U_Sub1,U_Sub2,U_Sub3,U_Sub4,U_Sub5;
END

GO

ALTER PROCEDURE [dbo].[KLTT_TOTAL]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BP_Code as varchar(100),
	@Period as int,
	@BGroup as varchar(50),
	@PUType as varchar(10)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @ProjectID as int;
	DECLARE @DocEntry as int;
	DECLARE @VAT as float;
	DECLARE @Last_Bill_Period as int;
	DECLARE @ToDate as date;

	DECLARE @Table_HD TABLE(
		AbsID int,
		[Number] int,
		U_SHD decimal(19,6),
		StartDate date,
		Descript nvarchar(254),
		[Type] varchar(50),
		PTTU decimal(19,6),
		PTHU decimal(19,6),
		PTBH decimal(19,6),
		PTGL decimal(19,6),
		HTBH varchar(50),
		TTTU decimal(19,6),
		CTQLDTC varchar(50),
		GTTU decimal(19,6),
		GTHD decimal(19,6)
	);

	Select top 1 @DocEntry = DocEntry, @VAT = ISNULL(U_VAT,10)/100 , @Last_Bill_Period = U_Period, @ToDate = U_DATETO
	from [@KLTT] 
	where U_Period <= @Period 
	and U_FIPROJECT = @FinancialProject
	and U_BGroup = @BGroup
	and U_BPCode = @BP_Code
	and U_PUType = @PUType
	and U_BType in (2,1)
	and Canceled ='N'
	order by U_Period desc;
	--Get HD cho bill 
	--Insert into @Table_HD(AbsID, [Number], U_SHD, StartDate, Descript, [Type], PTTU, PTHU, PTBH, PTGL, HTBH, TTTU, CTQLDTC, GTTU, GTHD)
	--Exec [dbo].[KLTT_GET_ADDITIONALINFO] @FinancialProject, @BP_Code, @Last_Bill_Period, @BGroup, @PUType, @ToDate;

	Select b.* 
	,(Select ISNULL(SUM(U_GTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period <= @Period
		and U_BType = 1
		) as 'TOTAL_TU'
	,(Select ISNULL(SUM(U_HTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period <= @Period + 1
		and U_BType = 2
		) as 'TOTAL_HU'
	,(Select ISNULL(SUM(U_GTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period <= @Last_Bill_Period
		and U_BType = 1
		) as 'TOTAL_TU_LASTBILL'
	,(Select ISNULL(SUM(U_HTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period <= @Last_Bill_Period
		and U_BType = 2
		) as 'TOTAL_HU_LASTBILL'
	,(Select U_PTQuanLy from [@KLTT] where DocEntry =@DocEntry) as 'PhiQL'
	from 
	(
	Select SUM(a.Sum_PL) as 'SUM_PL', SUM(a.SUM_CA) * (1+@VAT) as 'SUM_CA',  SUM(a.SUM_CA) as 'SUM_CA_NOVAT'
	from
	(
	Select 'A' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTA] 
	where DocEntry =@DocEntry
	union all
	Select 'B' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTB] 
	where DocEntry =@DocEntry
	union all
	Select 'K' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTK] 
	where DocEntry =@DocEntry
	union all
	Select 'C' as Type--,-SUM(U_SUM) as 'Sum_PL',-SUM(U_CompleteAmount) as 'SUM_CA'
					, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'Sum_PL'
					, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'SUM_CA'
	from [@KLTTC] 
	where DocEntry =@DocEntry
	union all
	Select 'D' as Type--,-SUM(U_SUM) as 'Sum_PL',-SUM(U_CompleteAmount) as 'SUM_CA'
					, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'Sum_PL'
					, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'SUM_CA'
	from [@KLTTD] 
	where DocEntry =@DocEntry
	union all
	Select 'E' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTE] 
	where DocEntry =@DocEntry
	union all
	Select 'F' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTF] 
	where DocEntry =@DocEntry
	union all
	Select 'G' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTG] 
	where DocEntry =@DocEntry
	--union
	--Select 'H' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_SUM) as 'SUM_CA'
	--from [@KLTTH] 
	--where DocEntry =@DocEntry
	) a)b;
END

GO

ALTER PROCEDURE [dbo].[KLTT_GetList_Bill_Approve]
	-- Add the parameters for the stored procedure here
	@UserName as varchar(50)
AS
BEGIN

DECLARE @Dept_Code as int
DECLARE @Pos_Code as int
DECLARE @Dept_Name as nvarchar(100)
DECLARE @Pos_Name as nvarchar(100)
--Get User Info - Dept - Position
Select @Dept_Code=a.dept
	,@Dept_Name = a.deptName
	,@Pos_Code = a.position
	,@Pos_Name = a.posName
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

--Get Lasting Bill Post
Select z.*
,(Select [Name] from OUDP where Code= z.POST_LVL) as POST_LVL_Des
from
(
Select a.DocEntry
,a.U_FIPROJECT as Project
,a.U_Period as 'Period'
,a.U_BGroup as 'BGroup'
,a.U_PUType as 'Purchase Type'
,case a.U_BType when 1 then N'T?m ?ng'
			when 2 then N'Thanh toán'
			when 3 then N'Quy?t toán'
			end as 'Bill Type'
,a.U_BPCode as 'BPCode'
,a.U_BPName as 'BPName'
,a.U_DATEFROM as 'From Date'
,a.U_DATETO as 'To Date'
,ISNULL(a.U_Link,'') as 'Link'
,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
	where userId = (Select UserId from OUSR where User_Code = 
(Select top 1 U_Usr from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by  c.LineId desc))) as 'Last Approved'
--,case when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and a.U_BPCode in ('NTP00599','NTP00601','NTP00602') and U_BGroup ='XD'
,case when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and a.U_BPCode in ('NTP00611','NTP00601','NTP00602') and U_BGroup ='XD' -- Update For PRO
	then 6
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='XD'
	then 3 
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CDXD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TB'
	then 2
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TBXD'
	then 3
	else (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is null order by c.LineId asc)
	end as 'POST_LVL'
,a.U_BPCode2
from [@KLTT] a where 
a.Canceled = 'N'
and a.Status = 'O') z
where z.POST_LVL = @Dept_Code
and z.Project in 
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) 
order by z.Project,z.Period,z.BPCode asc;
END

GO

ALTER PROCEDURE [dbo].[KLTT_GetList_Bill_Approved]
	-- Add the parameters for the stored procedure here
	@UserName as varchar(50)
AS
BEGIN

DECLARE @Dept_Code as int
DECLARE @Pos_Code as int
DECLARE @Dept_Name as nvarchar(100)
DECLARE @Pos_Name as nvarchar(100)
--Get User Info - Dept - Position
Select @Dept_Code=a.dept
	,@Dept_Name = a.deptName
	,@Pos_Code = a.position
	,@Pos_Name = a.posName
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

--Get Lasting Bill Post
Select z.*
,(Select [Name] from OUDP where Code= z.POST_LVL) as POST_LVL_Des
from
(
Select a.DocEntry
,a.U_FIPROJECT as Project
,a.U_Period as 'Period'
,a.U_BGroup as 'BGroup'
,a.U_PUType as 'Purchase Type'
,case a.U_BType when 1 then N'T?m ?ng'
			when 2 then N'Thanh toán'
			when 3 then N'Quy?t toán'
			end as 'Bill Type'
,a.U_BPCode as 'BPCode'
,a.U_BPName as 'BPName'
,a.U_DATEFROM as 'From Date'
,a.U_DATETO as 'To Date'
,a.UserSign as 'Creator'
,ISNULL(a.U_Link,'') as 'Link'
,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
where userId = (Select UserId from OUSR where User_Code = a.Creator)) as 'Creator Name'
,case when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='XD'
	then 3 
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CDXD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TB'
	then 2
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TBXD'
	then 3
	else (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is null order by c.LineId asc)
	end as 'POST_LVL'
,a.U_BPCode2
from [@KLTT] a where
a.Canceled = 'N'
and a.Status = 'C'
) z
where 
--z.Creator = (Select USERID from OUSR where USER_CODE=@Username)
--and 
z.Project in
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
	order by z.Project,z.Period,z.BPCode asc;
END
GO

ALTER PROCEDURE [dbo].[KLTT_GetList_Bill_Rejected]
	-- Add the parameters for the stored procedure here
	@UserName as varchar(50)
AS
BEGIN

DECLARE @Dept_Code as int
DECLARE @Pos_Code as int
DECLARE @Dept_Name as nvarchar(100)
DECLARE @Pos_Name as nvarchar(100)
--Get User Info - Dept - Position
Select @Dept_Code=a.dept
	,@Dept_Name = a.deptName
	,@Pos_Code = a.position
	,@Pos_Name = a.posName
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

--Get Lasting Bill Post
Select z.*
,(Select [Name] from OUDP where Code= z.POST_LVL) as POST_LVL_Des
from
(
Select a.DocEntry
,a.U_FIPROJECT as Project
,a.U_Period as 'Period'
,a.U_BGroup as 'BGroup'
,a.U_PUType as 'Purchase Type'
,case a.U_BType when 1 then N'T?m ?ng'
			when 2 then N'Thanh toán'
			when 3 then N'Quy?t toán'
			end as 'Bill Type'
,a.U_BPCode as 'BPCode'
,a.U_BPName as 'BPName'
,a.U_DATEFROM as 'From Date'
,a.U_DATETO as 'To Date'
,a.UserSign as 'Creator'
,ISNULL(a.U_Link,'') as 'Link'
,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
where userId = (Select UserId from OUSR where User_Code = a.Creator)) as 'Creator Name'
,case when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='XD'
	then 3 
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CDXD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TB'
	then 2
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TBXD'
	then 3
	else (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is null order by c.LineId asc)
	end as 'POST_LVL'
,a.U_BPCode2
from [@KLTT] a where 
a.Canceled = 'Y'
--and a.Status = 'C'
) z
where 
--z.Creator = (Select USERID from OUSR where USER_CODE=@Username)
--and 
z.Project in 
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) 
	order by z.Project,z.Period,z.BPCode asc;
END
GO

ALTER PROCEDURE [dbo].[KLTT_GetList_Bill_All]
	-- Add the parameters for the stored procedure here
	@UserName as varchar(50)
AS
BEGIN

DECLARE @Dept_Code as int
DECLARE @Pos_Code as int
DECLARE @Dept_Name as nvarchar(100)
DECLARE @Pos_Name as nvarchar(100)
--Get User Info - Dept - Position
Select @Dept_Code=a.dept
	,@Dept_Name = a.deptName
	,@Pos_Code = a.position
	,@Pos_Name = a.posName
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

--Get Lasting Bill Post
Select z.*
,(Select [Name] from OUDP where Code= z.POST_LVL) as POST_LVL_Des
from
(
Select a.DocEntry
,a.U_FIPROJECT as Project
,a.U_Period as 'Period'
,a.U_BGroup as 'BGroup'
,a.U_PUType as 'Purchase Type'
,case a.U_BType when 1 then N'T?m ?ng'
			when 2 then N'Thanh toán'
			when 3 then N'Quy?t toán'
			end as 'Bill Type'
,a.U_BPCode as 'BPCode'
,a.U_BPName as 'BPName'
,a.U_DATEFROM as 'From Date'
,a.U_DATETO as 'To Date'
,ISNULL(a.U_Link,'') as 'Link'
,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
	where userId = (Select UserId from OUSR where User_Code = 
(Select top 1 U_Usr from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by  c.LineId desc))) as 'Last Approved'
,case when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='XD'
	then 3 
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='CDXD'
	then 3
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TB'
	then 2
	when (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is not null order by c.LineId desc) is null and U_BGroup ='TBXD'
	then 3
	else (Select top 1 U_Level from [@KLTT_APPROVE] c where c.DocEntry=a.DocEntry and c.U_Status is null order by c.LineId asc)
	end as 'POST_LVL'
,a.U_BPCode2
from [@KLTT] a-- where 
--a.Canceled = 'N'
--and a.Status = 'O'
) z
where 
--z.POST_LVL = @Dept_Code
--and 
z.Project in 
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) 
order by z.Project,z.Period,z.BPCode asc;
END

GO

ALTER PROCEDURE [dbo].[KLTT_Approve_LV]
	-- Add the parameters for the stored procedure here
	@UserName as varchar(50),
	@DocEntry as int,
	@Status as nvarchar(10),
	@Comment as nvarchar(200)
AS
BEGIN

DECLARE @Dept_Code as int
DECLARE @Pos_Code as int
DECLARE @Dept_Name as nvarchar(100)
DECLARE @Pos_Name as nvarchar(100)
DECLARE @Update_Row as int
--Get User Info - Dept - Position
Select @Dept_Code=a.dept
	,@Dept_Name = a.deptName
	,@Pos_Code = a.position
	,@Pos_Name = a.posName
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

--Update Level Posting
Update [@KLTT_APPROVE] 
set U_Usr = @UserName
	,U_Time=CONVERT(varchar(30), GETDATE(), 113)
	,U_Status = @Status
	,U_Comment = @Comment
where DocEntry = @DocEntry
and (ISNULL(U_Position,@Pos_Code) = @Pos_Code or ((@UserName='Thuy.nguyen') and U_Level = 6))
and (Select Canceled from [@KLTT] where DocEntry = @DocEntry) <> 'Y'
and U_Level = @Dept_Code
and U_Status is null;
--Update them truong hop khi truong phong duyet ko qua nhan vien
if @Pos_Code = 1
		Update [@KLTT_APPROVE] 
		set U_Usr = @UserName
			,U_Time=CONVERT(varchar(30), GETDATE(), 113)
			,U_Status = '3'
		where DocEntry = @DocEntry
		and U_Position = 2
		and U_Level = @Dept_Code
		and U_Status is null;
SELECT @Update_Row = @@ROWCOUNT;
RETURN @Update_Row;
END

GO

ALTER PROCEDURE [dbo].[KLTT_Approve_Process]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
	Select 
	U_Level
	,(Select [Name] from OUDP where Code=U_Level) as 'DeptName'
	,U_Position
	,(Select [Name] from OHPS where posID=U_Position) as 'Position'
	,case U_Status when 1 then 'Approved' when 2 then 'Rejected' when 3 then 'By Pass' when 4 then 'Approved with Comment' end as 'Status'
	,U_Usr as 'Usr Approved by' 
	,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') as 'NAME' from OHEM
		where userId = (Select UserId from OUSR where User_Code = U_Usr)) as 'Approved by'
	,U_Time as 'Approved on'
	,ISNULL(U_Comment,'') as 'Comment' 
	From [@KLTT_APPROVE]
	where DocEntry = @DocEntry;
END

GO

ALTER PROCEDURE [dbo].[KLTT_APPROVE_GET_ADDITIONALINFO]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BP_Code as varchar(100),
	@Period as int,
	@CGroup as varchar(50),
	@PUType as varchar(50),
	@ToDate as date
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @DocEntry as int; -- HD
	DECLARE @DocEntry_HDNT as int; -- HD Nguyen tac
	DECLARE @DocEntry_PLTT as int; -- HD Thay the
	DECLARE @DocEntry_PLT as int; -- PL tang
	DECLARE @HTTU as numeric(19,6);
	DECLARE @GTTU as numeric(19,6);

	--Lay HD Nguyen tac
	Select top 1 @DocEntry_HDNT = isnull(AbsID,-1) 
	from OOAT 
	where U_PRJ is null
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;
	
	--Lay HD
	Select top 1 @DocEntry = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =48
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL Thay the
	Select top 1 @DocEntry_PLTT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =140
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL tang
	Select top 1 @DocEntry_PLT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =141
	and BpCode = @BP_Code
	and Status ='A'
	and Cancelled <> 'Y'
	and U_CGroup = @CGroup
	and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;
	if (@DocEntry_HDNT > 0)
	begin 
		--HD Nguyên t?c n?u có PL thay th? thì l?y PL Thay th?
		if (@DocEntry_PLTT > 0)
			begin
				Select x.*
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLTT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry_PLTT
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
			end
		--else if (@DocEntry_PLT > 0)
		--	begin
			 -- Không x?y ra - PL t?ng gáng trên H? Nguyên t?c
		--	end
		else
			begin
				Select 
					AbsID
					,Number
					,U_SHD
					,StartDate
					,Descript
					,'HDNT' as 'Type'
					,U_PTTU/100 as 'PTTU'
					,U_PTHU/100 as 'PTHU'
					,U_PTBH/100 as 'PTBH'
					,U_PTGL/100 as 'PTGL'
					,U_HTBH as 'HTBH'
					,U_TTTU as 'TTTU'
					,0 as 'GTHD'
				from OOAT 
				where 
					AbsID = @DocEntry_HDNT
					and Status = 'A'
					and Cancelled <> 'Y'
			end
	end
	else if (@DocEntry > 0)
	begin
		--Có H? -- Có PL Thay th? H?
		 Select top 1 @DocEntry_PLTT = isnull(AbsID,-1) from OOAT 
			where U_PRJ = @FinancialProject
			and Series =140
			and BpCode = @BP_Code
			and Status ='A'
			and Cancelled <> 'Y'
			and U_CGroup = @CGroup
			and U_PUTYPE = @PUType
			and StartDate <= @ToDate
			and U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry)
			order by AbsID desc;
		if (@DocEntry_PLTT > 0)
		begin
			--Co PLTT H?p ??ng
			Select x.*
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLTT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry_PLTT
						and Series in (140,141)
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
						and Series in (140,141)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
		else
		begin
			--Ch? có H? (ho?c có thêm PL t?ng)
			Select x.*
					from (
						--H?p ??ng
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'HD' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						AbsID = @DocEntry
						and Status = 'A'
						and Cancelled <> 'Y'
					union all
						--Phu luc Tang
						Select 
						AbsID
						,Number
						,U_SHD
						,StartDate
						,Descript
						,'PLT' as 'Type'
						,U_PTTU/100 as 'PTTU'
						,U_PTHU/100 as 'PTHU'
						,U_PTBH/100 as 'PTBH'
						,U_PTGL/100 as 'PTGL'
						,U_HTBH as 'HTBH'
						,U_TTTU as 'TTTU'
						,(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
							from OAT1 b
							where  b.AgrNo = AbsID) as 'GTHD'
						from OOAT 
						where 
						U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry)
						and StartDate <= @ToDate
						and Series in (140,141)
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
	end
END

GO

ALTER PROCEDURE [dbo].[KLTT_APPROVE_TOTAL]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BP_Code as varchar(100),
	@Period as int,
	@BGroup as varchar(50),
	@PUType as varchar(10)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @ProjectID as int;
	DECLARE @DocEntry as int;
	DECLARE @VAT as float;
	DECLARE @Last_Bill_Period as int;

	Select top 1 @DocEntry = DocEntry, @VAT = ISNULL(U_VAT,10)/100 , @Last_Bill_Period = U_Period
	from [@KLTT] 
	where U_Period <= @Period 
	and U_FIPROJECT = @FinancialProject
	and U_BGroup = @BGroup
	and U_BPCode = @BP_Code
	and U_PUType = @PUType
	and U_BType in (1,2,3)
	and Canceled = 'N'
	order by U_Period desc;

	Select b.* 
	,(Select ISNULL(SUM(U_GTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period < @Period
		and U_PUType = @PUType
		and U_BType = 1
		) as 'TOTAL_TU'
	,(Select ISNULL(SUM(U_HTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period <= @Period
		and U_PUType = @PUType
		and U_BType = 2
		) as 'TOTAL_HU'
	,(Select ISNULL(SUM(U_GTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_PUType = @PUType
		and U_Period <= @Last_Bill_Period
		and U_BType = 1
		) as 'TOTAL_TU_LASTBILL'
	,(Select ISNULL(SUM(U_HTTU),0) from [@KLTT] 
		where U_FIPROJECT = @FinancialProject
		and U_BGroup = @BGroup
		and U_BPCode = @BP_Code
		and U_Period < @Last_Bill_Period
		and U_PUType = @PUType
		and U_BType = 2
		) as 'TOTAL_HU_LASTBILL'
	,(Select ISNULL(U_PTQuanLy,0) from [@KLTT] where DocEntry =@DocEntry) as 'PhiQL'
	from 
	(
	Select SUM(a.Sum_PL) as 'SUM_PL',SUM(a.Sum_PL) * (1+@VAT) as 'SUM_PL_VAT', SUM(a.SUM_CA) * (1+@VAT) as 'SUM_CA',  SUM(a.SUM_CA) as 'SUM_CA_NOVAT'
	from
	(
	Select 'A' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTA] 
	where DocEntry =@DocEntry
	union all
	Select 'B' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTB] 
	where DocEntry =@DocEntry
	union all
	Select 'K' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTK] 
	where DocEntry =@DocEntry
	union all
	Select 'C' as Type --, -SUM(U_SUM) as 'Sum_PL',-SUM(U_CompleteAmount) as 'SUM_CA'
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'Sum_PL'
			  , case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'SUM_CA'
	from [@KLTTC] 
	where DocEntry =@DocEntry
	union all
	Select 'D' as Type--, -SUM(U_SUM) as 'Sum_PL',-SUM(U_CompleteAmount) as 'SUM_CA'
				, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_Sum] else [U_Sum] end as 'Sum_PL'
				, case ISNULL([U_TYPE],'GI') when 'GI' then -[U_CompleteAmount] else [U_CompleteAmount] end as 'SUM_CA'
	from [@KLTTD] 
	where DocEntry =@DocEntry
	union all
	Select 'E' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTE] 
	where DocEntry =@DocEntry
	union all
	Select 'F' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTF] 
	where DocEntry =@DocEntry
	union all
	Select 'G' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTG] 
	where DocEntry =@DocEntry
	--union
	--Select 'H' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_SUM) as 'SUM_CA'
	--from [@KLTTH] 
	--where DocEntry =@DocEntry
	) a)b;
END

GO

ALTER PROCEDURE [dbo].[KLTT_APPROVE_DUTRU]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
	,@BpCode as varchar(100)
	,@CGroup as varchar(50)
	,@PUType as varchar(50)
AS
BEGIN
	SET NOCOUNT ON;
	begin
		Select
			case	when @PUType = 'PUT01' then z.U_CP_NCC
					when @PUType = 'PUT02' then z.U_CP_NTP
					when @PUType = 'PUT03' then z.U_CP_NCC
					when @PUType = 'PUT04' then z.U_CP_NCC
					when @PUType = 'PUT09' then z.U_CP_DTC
					when @PUType = 'PUT05' and z.Series in (70,71) then z.U_CP_NCC
					when @PUType = 'PUT05' and z.Series in (72,73) then z.U_CP_NTP
					when @PUType = 'PUT05' and z.Series in (778) then z.U_CP_DTC
					when @PUType = 'PUT06' and z.Series in (70,71) then z.U_CP_NCC
					when @PUType = 'PUT06' and z.Series in (72,73) then z.U_CP_NTP
					when @PUType = 'PUT06' and z.Series in (778) then z.U_CP_DTC
					else z.U_CP_NCC
			end as 'DUTRU'
		from (
		Select [U_BPCode]
			  ,[U_BPName]
		      ,a.[U_TYPE]
			  ,b.Series
			  ,case when b.Series in (70,71) then SUM([U_CP_NCC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])  
					else SUM([U_CP_NCC]) end 
				  as 'U_CP_NCC'
			  ,SUM([U_CP_DP2]) as 'U_CP_DP2'
			  ,case when b.Series in (72,73) then SUM([U_CP_NTP]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				   else SUM([U_CP_NTP]) end
				   as 'U_CP_NTP'
			  ,case when b.Series in (78) then SUM([U_CP_DTC]) + SUM([U_CP_CN]) + SUM([U_CP_DP]) + SUM([U_CP_Prelims]) + SUM([U_CP_TB]) +  SUM([U_CP_K]) + SUM([U_CP_VTP]) + SUM([U_CP_VC]) + SUM([U_CP_MB]) + SUM([U_CP_T]) + SUM([U_CP_VH])
				  else SUM([U_CP_DTC]) end
				   as 'U_CP_DTC'
			FROM [@DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where
			a.U_BPCode = @BpCode
			and a.[U_TYPE] = @CGroup
			and a.DocEntry in 
						(Select DocEntry
						from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in (
							Select a.CTG_KEY 
							from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FinancialProject group by U_GoiThauKey) a))
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]) z
	end
END
GO

ALTER FUNCTION [dbo].[FN_Get_Last_Period]
(@BpCode  AS nvarchar(100),@FProject as nvarchar(200), @BGroup as varchar(50), @PurchaseType as varchar(50))
RETURNS int 
    BEGIN   
        DECLARE @DocEntry int;
		Select Top 1 @DocEntry = DocEntry from [@KLTT]
		where U_BPCode= @BpCode
		and U_FIPROJECT = @FProject
		and U_BGroup = @BGroup
		and U_PUType = @PurchaseType
		and U_Period = (Select Max(U_period)  from [@KLTT]
						where U_BPCode= @BpCode
						and U_FIPROJECT = @FProject
						and U_BGroup = @BGroup
						and U_PUType = @PurchaseType
						and U_BType = 2
						and Canceled = 'N');
		RETURN @DocEntry;
    END 
GO

ALTER PROCEDURE [dbo].[KLTT_GT_KYNAY]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
	SET NOCOUNT ON;
	Select SUM(a.SUM_CA) as 'SUM_CA_NOVAT'
	from
	(
	Select 'A' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTA] 
	where DocEntry =@DocEntry
	union all
	Select 'B' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTB] 
	where DocEntry =@DocEntry
	union all
	Select 'K' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTK] 
	where DocEntry =@DocEntry
	union all
	Select 'C' as Type,-SUM(case ISNULL(U_TYPE,'GI') when 'GR' then -U_SUM else U_SUM end) as 'Sum_PL'
					  ,-SUM(case ISNULL(U_TYPE,'GI') when 'GR' then -U_CompleteAmount else U_CompleteAmount end) as 'SUM_CA'
	from [@KLTTC] 
	where DocEntry =@DocEntry
	union all
	Select 'D' as Type,-SUM(case ISNULL(U_TYPE,'GI') when 'GR' then -U_SUM else U_SUM end) as 'Sum_PL'
					  ,-SUM(case ISNULL(U_TYPE,'GI') when 'GR' then -U_CompleteAmount else U_CompleteAmount end) as 'SUM_CA'
	from [@KLTTD] 
	where DocEntry =@DocEntry
	union all
	Select 'E' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTE] 
	where DocEntry =@DocEntry
	union all
	Select 'F' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTF] 
	where DocEntry =@DocEntry
	union all
	Select 'G' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_CompleteAmount) as 'SUM_CA'
	from [@KLTTG] 
	where DocEntry =@DocEntry
	--union
	--Select 'H' as Type,SUM(U_SUM) as 'Sum_PL',SUM(U_SUM) as 'SUM_CA'
	--from [@KLTTH] 
	--where DocEntry =@DocEntry
	) a;
END
GO

ALTER PROCEDURE [dbo].[KLTT_GETDATA_NHANTRI]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@Period as int,
	@BP_Code as varchar(100),
	@Type as varchar,
	@BGroup as varchar(50),
	@PurchaseType as varchar(50)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @TableTmp TABLE(
		DocEntry int NOT NULL,
		BPCode varchar(250) NOT NULL,
		BGroup nvarchar(254) ,
		BType nvarchar(254) ,
		BPuType nvarchar(250),
		BPeriod int NOT NULL
	);
	INSERT INTO @TableTmp (DocEntry, BPCode,BGroup,BType,BPuType,BPeriod)
	Select B.DocEntry,A.*  from (
	Select U_BPCode,U_BGroup,U_BType,U_PUType,Max(U_period) as 'Period' from [@KLTT] 
	where U_BPCode2= @BP_Code
	and Canceled = 'N'
	and U_BGroup = @BGroup
	--and U_PUType = @PurchaseType
	and U_Period <= @Period
	and U_FIPROJECT = @FinancialProject
	group by U_BPCode,U_BGroup,U_BType,U_PUType) A
	Inner join [@KLTT] B on A.U_BPCode=B.U_BPCode and A.U_BGroup=B.U_BGroup and A.U_BType = B.U_BType and A.Period = B.U_Period;

	IF (@Type = 'A')
		Select U_GoiThauKey as 'GoiThauKey'
		,U_GoiThau as 'GoiThauName'
		,U_GPKey as 'GRPOKey'
		,U_GPDetailsKey as 'GRPORowKey'
		,U_GPDetailsName as 'DetailsName'
		,U_CTCV as 'DetailsWork'
		,U_UoM as 'UoM'
		,U_Quantity as 'Quantity'
		,U_UPrice as 'UPrice'
		,U_Sum as 'Total'
		,U_Sub1 as 'U_ParentID1'
		,U_Sub1Name as 'Name1'
		,U_Sub2 as 'U_ParentID2'
		,U_Sub2Name as 'Name2'
		,U_Sub3 as 'U_ParentID3'
		,U_Sub3Name as 'Name3'
		,U_Sub4 as 'U_ParentID4'
		,U_Sub4Name as 'Name4'
		,U_Sub5 as 'U_ParentID5'
		,U_Sub5Name as 'Name5'
		,U_Type as 'TYPE'
		,U_CompleteRate as 'Last_Complete_Rate'
		,U_CompleteAmount as 'Last_Complete_Amount'
		from [@KLTTA] where DocEntry in (Select DocEntry From @TableTmp);
	IF (@Type = 'B')
		Select U_GoiThauKey as 'GoiThauKey'
		,U_GoiThau as 'GoiThauName'
		,U_GPKey as 'GRPOKey'
		,U_GPDetailsKey as 'GRPORowKey'
		,U_GPDetailsName as 'DetailsName'
		,U_Details as 'DetailsWork'
		,U_UoM as 'UoM'
		,U_Quantity as 'Quantity'
		,U_UPrice as 'UPrice'
		,U_Sum as 'Total'
		,U_Sub1 as'U_ParentID1'
		,U_Sub1Name as 'Name1'
		,U_Sub2 as'U_ParentID2'
		,U_Sub2Name as 'Name2'
		,U_Sub3 as'U_ParentID3'
		,U_Sub3Name as 'Name3'
		,U_Sub4 as'U_ParentID4'
		,U_Sub4Name as 'Name4'
		,U_Sub5 as'U_ParentID5'
		,U_Sub5Name as 'Name5'
		,U_Type as 'TYPE'
		,U_CompleteRate as 'Last_Complete_Rate'
		,U_CompleteAmount as 'Last_Complete_Amount'
		from [@KLTTB] 
		where DocEntry in (Select DocEntry From @TableTmp) 
		and U_Type is not null;
	IF (@Type = 'C')
		Select
		U_GoiThauKey as GoiThauKey
		,U_GoiThau as GoiThauName
		,'' as SubProjectKey
		,'' as SubProjectName
		,U_GoodsIssue as GIKey
		,U_DetailsKey as GIRowKey
		,U_DetailsName as DetailsName
		,U_UoM as UoM
		,U_Quantity as Quantity
		,U_UPrice as UPrice
		,U_Sum as Total
		--,a.U_BPCode as CardCode
		--,b.CreateDate
		,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
		,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
		from [@KLTTC] 
		where DocEntry in (Select DocEntry From @TableTmp)
		and U_GoodsIssue is not null;
	IF (@Type = 'D')
		Select
		U_GoiThauKey as GoiThauKey
		,U_GoiThau as GoiThauName
		,'' as SubProjectKey
		,'' as SubProjectName
		,U_GoodsIssue as GIKey
		,U_DetailsKey as GIRowKey
		,U_DetailsName as DetailsName
		,U_UoM as UoM
		,U_Quantity as Quantity
		,U_UPrice as UPrice
		,U_Sum as Total
		--,a.U_BPCode as CardCode
		--,b.CreateDate
		,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
		,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
		from [@KLTTD] 
		where DocEntry in (Select DocEntry From @TableTmp)
		and U_GoodsIssue is not null;
	IF (@Type = 'E')
		Select U_GoiThauKey as GoiThauKey
			,U_GoiThau as GoiThauName
			,U_SubprojectKey as SubProjectKey
			,U_StageKey as StagesKey
			,U_OpenIssueKey as OpenIssuesKey
			,U_OpenIssueRemark as Remarks
			,U_UoM as UoM
			,U_Quantity as Quantity
			,U_UPrice as UPrice
			,U_Sum as Total
			--,a.U_NCCPS as CardCode
			--,b.START
			,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
			from [@KLTTE] 
			where DocEntry in (Select DocEntry From @TableTmp)
			and U_StageKey is not null;
	IF (@Type = 'F')
		Select U_GoiThauKey as GoiThauKey
			,U_GoiThau as GoiThauName
			,U_SubProjectKey as SubProjectKey
			,U_StageKey as StagesKey
			,U_OpenIssueKey as OpenIssuesKey
			,U_OpenIssueRemark as Remarks
			,U_UoM as UoM
			,U_Quantity as Quantity
			,U_UPrice as UPrice
			,U_Sum as Total
			--,a.U_NCCPS as CardCode
			--,b.START
			,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
			from [@KLTTF] where DocEntry in (Select DocEntry From @TableTmp)
			and U_StageKey is not null;
	IF (@Type = 'G')
		Select U_GoiThauKey as GoiThauKey
			,U_GoiThau as GoiThauName
			,U_SubProjectKey as SubProjectKey
			,U_StageKey as StagesKey
			,U_OpenIssueKey as OpenIssuesKey
			,U_OpenIssueRemark as Remarks
			,U_UoM as UoM
			,U_Quantity as Quantity
			,U_UPrice as UPrice
			,U_Sum as Total
			--,a.U_NCCPS as CardCode
			--,b.START
			,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
			from [@KLTTG] where DocEntry in (Select DocEntry From @TableTmp)
			and U_StageKey is not null;
	IF (@Type = 'K')
	Select 
			U_GoiThauKey as GoiThauKey
			,U_GoiThau as GoiThauName
			,U_GPKey as GRPOKey
			,U_GPDetailsKey as GRPORowKey
			,U_GPDetailsName as DetailsName
			,U_CTCV as DetailsWork
			,U_UoM as UoM
			,U_Quantity as Quantity
			,U_UPrice as UPrice
			,U_Sum as Total
			--,b.CreateDate
			--,b.CardCode
			,(Select ProjectID from OPHA where AbsEntry = U_Sub1) as ProjectNo
			,U_Sub1 as U_ParentID1
			,U_Sub1Name as Name1
			,U_Sub2 as U_ParentID2
			,U_Sub2Name as Name2
			,U_Sub3 as U_ParentID3
			,U_Sub3Name as Name3
			,U_Sub4 as U_ParentID4
			,U_Sub4Name as Name4
			,U_Sub5 as U_ParentID5
			,U_Sub5Name as Name5
			--,b.U_RECTYPE
			,U_Type as 'TYPE'
			,ISNULL(U_CompleteRate,0) as Last_Complete_Rate
			,ISNULL(U_CompleteAmount,0) as Last_Complete_Amount
		from [@KLTTK] where DocEntry in (Select DocEntry From @TableTmp)
		and U_GPKey is not null;
END

GO

ALTER PROCEDURE [dbo].[KLTT_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
--Get User Info - Dept - Position
Declare @DeptCode as int
Declare @PosCode as int
Declare @FProject as varchar(250)
Select top 1 @DeptCode = ISNULL(U_Level,''),@PosCode=ISNULL(U_Position,'') from [@KLTT_APPROVE] 
where DocEntry = @DocEntry 
and U_Status is null 
order by LineID;

Select @FProject=ISNULL(U_FIPROJECT,'') from [@KLTT] where DocEntry=@DocEntry;

Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name
from OHEM a inner join OUSR b on a.USERID = b.UserID
left join HTM1 c on c.empID=a.empID
inner join OHTM d on c.teamID = d.teamID
where a.dept = @DeptCode
and a.position = @PosCode
and d.name = @FProject;
END

GO

CREATE PROCEDURE [dbo].[GET_EMAIL_CONF]
	-- Add the parameters for the stored procedure here
AS
BEGIN
	SET NOCOUNT ON;
	Select
	 (Select [Name] from [@Email_Conf] where Code='Host') as 'Host_Address'
	 ,(Select [Name] from [@Email_Conf] where Code='Host_Port') as 'Host_Port'
	 ,(Select [Name] from [@Email_Conf] where Code='EnableSSL') as 'EnableSSL'
	 ,(Select [Name] from [@Email_Conf] where Code='User') as 'User'
	 ,(Select [Name] from [@Email_Conf] where Code='Pwd') as 'Pwd'
	 ,(Select [Name] from [@Email_Conf] where Code='Email_From') as 'Email_From'
	 ,(Select [Name] from [@Email_Conf] where Code='Email_From_Name') as 'Email_From_Name'
END
GO

ALTER PROCEDURE [dbo].[KLTT_GETDATA_COVER]
	@DocEntry as int
AS
BEGIN
DECLARE @FProject as varchar(250)
DECLARE @Period as int
DECLARE @BpCode as varchar(250)
DECLARE @BpName as nvarchar(254)
DECLARE @BGroup as varchar(250)
DECLARE @PUType as varchar(250)
DECLARE @Todate as date
DECLARE @BpCode2 as varchar(250)
DECLARE @PhiQL as numeric(19,6)
DECLARE @VAT as numeric(19,6)
DECLARE @HTTU as numeric(19,6)
DECLARE @GTTU as numeric(19,6)
DECLARE @BType as int
DECLARE @PTGL as numeric(19,6)

Select @FProject = U_FIPROJECT
, @Period = U_Period
, @BpCode = U_BPCode
, @BpName = U_BPName
, @BGroup = U_BGroup
, @PUType = U_PUType
, @Todate = U_DATETO
, @BPCode2 = U_BPCode2
, @PhiQL = U_PTQuanLy
, @VAT = U_VAT
, @HTTU = U_HTTU
, @GTTU = U_GTTU
, @BType = U_BType
from [@KLTT] 
where DocEntry = @DocEntry;

--Lay H?
DECLARE @DocEntry_HDNT as int
DECLARE @DocEntry_PLTT as int
DECLARE @DocEntry_PLT as int
DECLARE @DocEntry_HD as int

DECLARE @GTHD as decimal -- Gia tri hop dong
DECLARE @PLT as decimal --Gia tri phu luc tang
DECLARE @NDHD as nvarchar(254) -- noi dung HD
DECLARE @SHD as int -- So hop dong
DECLARE @NGAYHD as date -- Ngay HD
DECLARE @HTBH as varchar(50) --Hinh thuc bao hanh
DECLARE @PTBH as decimal --Phan tram bao hanh
--Lay HD Nguyen tac
Select top 1 @DocEntry_HDNT = isnull(AbsID,-1) 
from OOAT
where U_PRJ is null
and BpCode = @BpCode
and Status ='A'
and Cancelled <> 'Y'
and U_CGroup = @BGroup
and U_PUTYPE = @PUType
and StartDate <= @ToDate
order by AbsID desc;
	
--Lay HD
Select top 1 @DocEntry_HD = isnull(AbsID,-1)
from OOAT 
where U_PRJ = @FProject
and Series =48
and BpCode = @BpCode
and U_PUTYPE = @PUType
and Status ='A'
and Cancelled <> 'Y'
and U_CGroup = @BGroup
and StartDate <= @ToDate
order by AbsID desc;

--Lay PL Thay the
Select top 1 @DocEntry_PLTT = isnull(AbsID,-1)
from OOAT 
where U_PRJ = @FProject
and Series =141
and BpCode = @BpCode
and Status ='A'
and Cancelled <> 'Y'
and U_PUTYPE = @PUType
and U_CGroup = @BGroup
and StartDate <= @ToDate
order by AbsID desc;

--Lay PL tang
Select top 1 @DocEntry_PLT = isnull(AbsID,-1)
from OOAT 
where U_PRJ = @FProject
and Series =140
and BpCode = @BpCode
and Status ='A'
and Cancelled <> 'Y'
and U_CGroup = @BGroup
and U_PUTYPE = @PUType
and StartDate <= @ToDate
order by AbsID desc;

if (ISNULL(@DocEntry_HDNT,0) > 0)
begin
	--HD Nguyên t?c n?u có PL thay th? thì l?y PL Thay th?
	if (ISNULL(@DocEntry_PLTT,0) > 0)
		begin
			--Phu luc thay the
			Select 
			  @SHD = Number
			, @NDHD = Descript
			, @NGAYHD = StartDate
			, @HTBH = U_HTBH
			, @PTGL = ISNULL(U_PTGL,0)
			, @PTBH = ISNULL(U_PTBH,0)
			, @GTHD = (Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
				from OAT1 b
				where  b.AgrNo = AbsID)
			from OOAT 
			where 
			AbsID = @DocEntry_PLTT
			and Status = 'A'
			and Cancelled <> 'Y';
				
			--Phu luc Tang
			Select 
			@PLT = (Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
				from OAT1 b
				where  b.AgrNo = AbsID)
			from OOAT 
			where 
			U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
			and StartDate <= @ToDate
			and Status ='A'
			and Cancelled <> 'Y'
		end
	else
		begin
			Select 
			@SHD = U_SHD
			, @NDHD = Descript
			, @GTHD = 0
			, @HTBH = U_HTBH
			, @PTGL = case @BType when '3' then ISNULL(U_PTBH,0) else ISNULL(U_PTGL,0) end
			from OOAT 
			where 
				AbsID = @DocEntry_HDNT
				and Status = 'A'
				and Cancelled <> 'Y'
		end
end
else if (ISNULL(@DocEntry_HD,0) > 0)
begin
	--Có H? -- Có PL Thay th? H?
		Select top 1 @DocEntry_PLTT = isnull(AbsID,-1) from OOAT 
		where U_PRJ = @FProject
		and Series =141
		and BpCode = @BpCode
		and Status ='A'
		and Cancelled <> 'Y'
		and U_CGroup = 'VP'
		and StartDate <= @ToDate
		and U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_HD)
		order by AbsID desc;
	if (ISNULL(@DocEntry_PLTT,0) > 0)
	begin
		--Co PLTT H?p ??ng
		Select 
			@SHD = Number
			,@NDHD = Descript
			, @NGAYHD = StartDate
			, @PTGL = ISNULL(U_PTGL,0)
			, @PTBH = ISNULL(U_PTBH,0)
			, @HTBH = U_HTBH
			,@GTHD = (Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
				from OAT1 b
				where  b.AgrNo = AbsID)
		from OOAT 
		where 
			AbsID = @DocEntry_PLTT
			and Series in (140,141)
			and Status = 'A'
			and Cancelled <> 'Y';
						
		--Phu luc Tang
		Select 
			@PLT =(Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
					from OAT1 b
					where  b.AgrNo = AbsID)
		from OOAT 
		where 
			U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
			and Series in (140,141)
			and StartDate <= @ToDate
			and Status ='A'
			and Cancelled <> 'Y';
	end
	else
	begin
		--Ch? có H? (ho?c có thêm PL t?ng)
		--H?p ??ng
		Select 
			@SHD =  Number
			, @NDHD = Descript
			, @NGAYHD = StartDate
			, @PTGL = ISNULL(U_PTGL,0)
			, @PTBH = ISNULL(U_PTBH,0)
			, @HTBH = U_HTBH
			, @GTHD = ( Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
						from OAT1 b
						where  b.AgrNo = AbsID) 
		from OOAT 
		where 
			AbsID = @DocEntry_HD
			and Status = 'A'
			and Cancelled <> 'Y';

		--Phu luc Tang
		Select 
				@PLT = ( Select (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) 
						from OAT1 b
						where  b.AgrNo = AbsID)
		from OOAT 
		where 
			U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_HD)
			and 
			StartDate <= @ToDate
			and Series in (140)
			and Status ='A'
			and Cancelled <> 'Y';
	end
end

--Lay KLTT
if (@BType > 1)
	begin
	DECLARE @KLTT_Cur TABLE(
			SUM_PL numeric(19,6),
			SUM_PL_VAT numeric(19,6),
			SUM_CA numeric(19,6),
			SUM_CA_NOVAT numeric(19,6),
			TOTAL_TU numeric(19,6),
			TOTAL_HU numeric(19,6),
			TOTAL_TU_LASTBILL numeric(19,6),
			TOTAL_HU_LASTBILL numeric(19,6),
			PhiQL numeric(19,6)
		);
	--Ky nay
	DECLARE @TongGT_Thicong as numeric(19,6)
	DECLARE @GT_thuchien_denKN as numeric(19,6)
	DECLARE @GT_thanhtoan_denKN as numeric(19,6)
	DECLARE @Tamung as numeric(19,6)
	DECLARE @HoanTamung as numeric(19,6)
	DECLARE @GT_giulai as numeric(19,6)
	DECLARE @TongGT_thanhtoan_denKN as numeric(19,6)
	DECLARE @TongGT_thanhtoan_kytruoc as numeric(19,6)
	DECLARE @GT_Denghi_thanhtoan_KN as numeric(19,6)
	--Ky truoc
	DECLARE @Pre_Period as int
	DECLARE @GT_thanhtoan_Kytruoc as numeric(19,6)

	Insert into @KLTT_Cur(SUM_PL, SUM_PL_VAT, SUM_CA, SUM_CA_NOVAT, TOTAL_TU, TOTAL_HU, TOTAL_TU_LASTBILL, TOTAL_HU_LASTBILL, PhiQL)
	Exec [dbo].[KLTT_APPROVE_TOTAL] @FProject, @BpCode, @Period, @BGroup, @PUType;

	Select @TongGT_Thicong = SUM_PL_VAT * (1 + ISNULL(PhiQL,0)/100) 
		  ,@GT_thuchien_denKN = SUM_CA * (1 + ISNULL(PhiQL,0)/100)
		  ,@Tamung = ISNULL(TOTAL_TU, 0) 
		  ,@HoanTamung = ISNULL(TOTAL_HU, 0)
	from @KLTT_Cur;

	SET @Pre_Period = @Period -1 ;
	Delete from @KLTT_Cur;
	Insert into @KLTT_Cur(SUM_PL, SUM_PL_VAT, SUM_CA, SUM_CA_NOVAT, TOTAL_TU, TOTAL_HU, TOTAL_TU_LASTBILL, TOTAL_HU_LASTBILL, PhiQL)
	Exec [dbo].[KLTT_APPROVE_TOTAL] @FProject, @BpCode, @Pre_Period, @BGroup, @PUType;

	Select @GT_thanhtoan_Kytruoc = (SUM_CA * (1 - ISNULL(@PTGL,0)/100)) + (SUM_CA * ISNULL(PhiQL,0)/100) + ISNULL(TOTAL_TU_LASTBILL,0) - ISNULL(TOTAL_HU,0)
	from @KLTT_Cur;

	SET @GT_thanhtoan_denKN = (1- ISNULL(@PTGL,0)/100) * @GT_thuchien_denKN;
	SET @GT_giulai = (@PTGL/100) * @GT_thuchien_denKN;
	SET @TongGT_thanhtoan_denKN = @GT_thanhtoan_denKN + ISNULL(@Tamung,0) - ISNULL(@HoanTamung,0);
	SET @GT_Denghi_thanhtoan_KN = ISNULL(@TongGT_thanhtoan_denKN, 0) - ISNULL(@GT_thanhtoan_Kytruoc,0);
	
	if(@BType = 2)
		--Get DATA for Cover
		Select @Period as 'So'
			, @BpCode as 'BpCode'
			, @BpCode2 as 'BpCode2'
			, @BpName as 'BpName'
			, (Select PrjName from OPRJ where PrjCode = @FProject) as 'ProjectName'
			, @SHD as 'SoHD'
			, @NGAYHD as 'NgayHD'
			, @NDHD as 'Noidung'
			, @BType as 'BillType'
			, @BGroup as 'BillGroup'
			, ISNULL(@GTHD,0) as 'GTHD'
			, ISNULL(@PLT,0) as 'PLT'
			, ISNULL(@TongGT_Thicong,0) as 'TongGT_Thicong'
			, ISNULL(@GT_thuchien_denKN,0) as 'GT_thuchien_denKN'
			, ISNULL(@GT_thanhtoan_denKN,0) as 'GT_thanhtoan_denKN'
			, ISNULL(@Tamung,0) as 'Tamung'
			, ISNULL(@HoanTamung,0) as 'HoanTamung'
			, ISNULL(@GT_giulai,0) as 'GT_giulai'
			, ISNULL(@TongGT_thanhtoan_denKN,0) as 'TongGT_thanhtoan_denKN'
			, ISNULL(@GT_thanhtoan_Kytruoc,0) as 'TongGT_thanhtoan_kytruoc'
			, ISNULL(@GT_Denghi_thanhtoan_KN,0) as 'GT_Denghi_thanhtoan_KN';
	else if (@BType = 3)
		begin
			if (ISNULL(@HTBH,'') <> 'TM') SET @GT_giulai = 0;
			else SET @GT_giulai = (@PTBH/100) * @GT_thuchien_denKN;
			SET @GT_thanhtoan_denKN =  @GT_thuchien_denKN;
			SET @TongGT_thanhtoan_denKN = @GT_thanhtoan_denKN;
			SET @GT_Denghi_thanhtoan_KN = ISNULL(@TongGT_thanhtoan_denKN, 0) - ISNULL(@GT_thanhtoan_Kytruoc,0);

			Select @Period as 'So'
				, @BpCode as 'BpCode'
				, @BpCode2 as 'BpCode2'
				, @BpName as 'BpName'
				, (Select PrjName from OPRJ where PrjCode = @FProject) as 'ProjectName'
				, @SHD as 'SoHD'
				, @NGAYHD as 'NgayHD'
				, @NDHD as 'Noidung'
				, @BType as 'BillType'
				, @BGroup as 'BillGroup'
				, ISNULL(@GTHD,0) as 'GTHD'
				, ISNULL(@PLT,0) as 'PLT'
				, ISNULL(@TongGT_Thicong,0) as 'TongGT_Thicong'
				, ISNULL(@GT_thuchien_denKN,0) as 'GT_thuchien_denKN'
				, ISNULL(@GT_thanhtoan_denKN,0) as 'GT_thanhtoan_denKN'
				, ISNULL(@Tamung,0) as 'Tamung'
				, ISNULL(@Tamung,0) as 'HoanTamung'
				, ISNULL(@GT_giulai,0) as 'GT_giulai'
				, ISNULL(@TongGT_thanhtoan_denKN,0) as 'TongGT_thanhtoan_denKN'
				, ISNULL(@GT_thanhtoan_Kytruoc,0) as 'TongGT_thanhtoan_kytruoc'
				, ISNULL(@GT_Denghi_thanhtoan_KN,0) as 'GT_Denghi_thanhtoan_KN';
		end
	end
else
	begin
		--Get DATA for Cover
		Select @Period as 'So'
			, @BpCode as 'BpCode'
			, @BpCode2 as 'BpCode2'
			, @BpName as 'BpName'
			, (Select PrjName from OPRJ where PrjCode = @FProject) as 'ProjectName'
			, @SHD as 'SoHD'
			, @NGAYHD as 'NgayHD'
			, @NDHD as 'Noidung'
			, @BType as 'BillType'
			, @BGroup as 'BillGroup'
			, ISNULL(@GTHD,0) as 'GTHD'
			, ISNULL(@PLT,0) as 'PLT'
			, 0 as 'TongGT_Thicong'
			, 0 as 'GT_thuchien_denKN'
			, 0 as 'GT_thanhtoan_denKN'
			, @GTTU as 'Tamung'
			, 0 as 'HoanTamung'
			, 0 as 'GT_giulai'
			, @GTTU as 'TongGT_thanhtoan_denKN'
			, 0 as 'TongGT_thanhtoan_kytruoc'
			, @GTTU as 'GT_Denghi_thanhtoan_KN';
	end
END;
GO

ALTER PROCEDURE [dbo].[KLTT_Get_Approve_Process_Cover]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
	DECLARE @NT as varchar(10)
	DECLARE @BGroup as varchar(50)
	DECLARE @BPCode as varchar(50)
	DECLARE @CreateDate as datetime

	DECLARE @1_Name as nvarchar(254)
	DECLARE @1_Arrpove as nvarchar(254)
	DECLARE @1_Time as datetime
	DECLARE @1_Comm as nvarchar(254)

	DECLARE @2_Name as nvarchar(254)
	DECLARE @2_Arrpove as nvarchar(254)
	DECLARE @2_Time as datetime
	DECLARE @2_Comm as nvarchar(254)

	DECLARE @3_Name as nvarchar(254)
	DECLARE @3_Arrpove as nvarchar(254)
	DECLARE @3_Time as datetime
	DECLARE @3_Comm as nvarchar(254)

	DECLARE @CCM_Name as nvarchar(254)
	DECLARE @CCM_Arrpove as nvarchar(254)
	DECLARE @CCM_Time as datetime
	DECLARE @CCM_Comm as nvarchar(254)

	DECLARE @BGD_Name as nvarchar(254)
	DECLARE @BGD_Arrpove as nvarchar(254)
	DECLARE @BGD_Time as datetime
	DECLARE @BGD_Comm as nvarchar(254)

	DECLARE @KT_Name as nvarchar(254)
	DECLARE @KT_Arrpove as nvarchar(254)
	DECLARE @KT_Time as datetime
	DECLARE @KT_Comm as nvarchar(254)

	DECLARE @TableTmp TABLE(
		LineId int NOT NULL,
		Approve_Level nvarchar(254) NOT NULL,
		Usr nvarchar(254) ,
		UserName nvarchar(254),
		Position nvarchar(254),
		Approve_Time datetime,
		Comment nvarchar(254),
		Approve_Status nvarchar(254)
	);

	Select @CreateDate = CreateDate, @BPCode=U_BPCode, @NT=U_BPCode2, @BGroup = U_BGroup from [@KLTT] where DocEntry = @DocEntry;

	INSERT INTO @TableTmp (LineId, Approve_Level,Usr,UserName,Position,Approve_Time,Comment,Approve_Status)
	Select a.LineId,a.U_Level,a.U_Usr, ISNULL(c.LastName,'') +' '+ ISNULL(c.MiddleName,'')+ ' '+ ISNULL(c.FirstName,''), a.U_Position,CONVERT(VARCHAR(50), a.U_Time, 113) as 'U_Time',a.U_Comment, case a.U_Status when 2 then 'Rejected' when 1 then 'Approved' when 3 then 'Bypass' end as 'U_Status'
	from [@KLTT_APPROVE] a left join OUSR b on a.U_Usr = b.USER_CODE
	left join OHEM c on b.USERID = c.userId
	where DocEntry = @DocEntry;

	--Select @TBP_Name = UserName
	--	, @TBP_Arrpove = Approve_Status
	--	, @TBP_Time = Approve_Time
	--	, @TBP_Comm = Comment
	--from @TableTmp 
	--where LineId =2;
	if (@BGroup = 'CD')
	begin
		Select @1_Name = UserName
		, @1_Arrpove = Approve_Status
		, @1_Time = Approve_Time
		, @1_Comm = Comment
		from @TableTmp 
		where LineId =1;

		Select @2_Name = ''
		, @2_Arrpove = ''
		, @2_Time = ''
		, @2_Comm = '';

		Select @3_Name = UserName
		, @3_Arrpove = Approve_Status
		, @3_Time = Approve_Time
		, @3_Comm = Comment
		from @TableTmp 
		where LineId =3;

		Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
		from @TableTmp 
		where LineId =5;

		Select @BGD_Name = UserName
			, @BGD_Arrpove = Approve_Status
			, @BGD_Time = Approve_Time
			, @BGD_Comm = Comment
		from @TableTmp 
		where LineId =6;

		Select @KT_Name = UserName
			, @KT_Arrpove = Approve_Status
			, @KT_Time = Approve_Time
			, @KT_Comm = Comment
		from @TableTmp 
		where LineId = case ISNULL(@NT,'') when '' then 7 else 7 end; --pqhuy1987 20180619 gi?m 1 b?c k? toán
	end
	--else if (@BGroup = 'XD' and @BPCode <>'NTP00599')
	else if (@BGroup = 'XD' and @BPCode <>'NTP00611') -- update for PRO
	begin
		Select @1_Name = ''
		, @1_Arrpove = ''
		, @1_Time = ''
		, @1_Comm = '';

		Select @2_Name = UserName
		, @2_Arrpove = Approve_Status
		, @2_Time = Approve_Time
		, @2_Comm = Comment
		from @TableTmp 
		where LineId =1;

		Select @3_Name = ''
		, @3_Arrpove = ''
		, @3_Time = ''
		, @3_Comm = '';

		Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
		from @TableTmp 
		where LineId =3;

		Select @BGD_Name = UserName
			, @BGD_Arrpove = Approve_Status
			, @BGD_Time = Approve_Time
			, @BGD_Comm = Comment
		from @TableTmp 
		where LineId =4;

		Select @KT_Name = UserName
			, @KT_Arrpove = Approve_Status
			, @KT_Time = Approve_Time
			, @KT_Comm = Comment
		from @TableTmp 
		where LineId = case ISNULL(@NT,'') when '' then 5 else 5 end; --pqhuy1987 20180619 gi?m 1 b?c k? toán
	end
	--else if (@BGroup = 'XD' and @BPCode ='NTP00599')
	else if (@BGroup = 'XD' and @BPCode ='NTP00611') -- Update for PRO
	begin
		Select @1_Name = ''
		, @1_Arrpove = ''
		, @1_Time = ''
		, @1_Comm = '';

		Select @2_Name = ''
		, @2_Arrpove = ''
		, @2_Time = ''
		, @2_Comm = ''
		from @TableTmp 
		where LineId =1;

		Select @3_Name = ''
		, @3_Arrpove = ''
		, @3_Time = ''
		, @3_Comm = '';

		Select @CCM_Name = ''
		, @CCM_Arrpove = ''
		, @CCM_Time = ''
		, @CCM_Comm = ''
		from @TableTmp 
		where LineId =3;
		if (@CreateDate <= '19-june-2018' )
		begin
			Select @BGD_Name = UserName
				, @BGD_Arrpove = Approve_Status
				, @BGD_Time = Approve_Time
				, @BGD_Comm = Comment
			from @TableTmp 
			where LineId = 4;
		end
		else
		begin
			Select @BGD_Name = UserName
				, @BGD_Arrpove = Approve_Status
				, @BGD_Time = Approve_Time
				, @BGD_Comm = Comment
			from @TableTmp 
			where LineId = 1;
		end
		if (@CreateDate <= '19-june-2018' ) begin
				Select @KT_Name = UserName
					, @KT_Arrpove = Approve_Status
					, @KT_Time = Approve_Time
					, @KT_Comm = ''
				from @TableTmp 
				where LineId = case ISNULL(@NT,'') when '' then 5 else 5 end;
		end
		else 
		begin
				Select @KT_Name = UserName
					, @KT_Arrpove = Approve_Status
					, @KT_Time = Approve_Time
					, @KT_Comm = ''
				from @TableTmp 
				where LineId = case ISNULL(@NT,'') when '' then 2 else 2 end;			
		end
	end
	else if (@BGroup = 'CDXD')
	begin
		Select @1_Name = UserName
		, @1_Arrpove = Approve_Status
		, @1_Time = Approve_Time
		, @1_Comm = Comment
		from @TableTmp 
		where LineId =1;

		Select @2_Name = UserName
		, @2_Arrpove = Approve_Status
		, @2_Time = Approve_Time
		, @2_Comm = Comment
		from @TableTmp 
		where LineId =2;

		Select @3_Name = UserName
		, @3_Arrpove = Approve_Status
		, @3_Time = Approve_Time
		, @3_Comm = Comment
		from @TableTmp 
		where LineId =4;

		Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
		from @TableTmp 
		where LineId =6;

		Select @BGD_Name = UserName
			, @BGD_Arrpove = Approve_Status
			, @BGD_Time = Approve_Time
			, @BGD_Comm = Comment
		from @TableTmp 
		where LineId =7;

		Select @KT_Name = UserName
			, @KT_Arrpove = Approve_Status
			, @KT_Time = Approve_Time
			, @KT_Comm = Comment
		from @TableTmp 
		where LineId = case ISNULL(@NT,'') when '' then 8 else 8 end; --pqhuy1987 20180619 gi?m 1 b?c k? toán
	end
	else if (@BGroup = 'TB')
	begin
		Select @1_Name = UserName
		, @1_Arrpove = Approve_Status
		, @1_Time = Approve_Time
		, @1_Comm = Comment
		from @TableTmp 
		where LineId =1;

		Select @2_Name = ''
		, @2_Arrpove = ''
		, @2_Time = ''
		, @2_Comm = '';

		Select @3_Name = ''
		, @3_Arrpove = ''
		, @3_Time = ''
		, @3_Comm = '';

		Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
		from @TableTmp 
		where LineId =3;

		Select @BGD_Name = UserName
			, @BGD_Arrpove = Approve_Status
			, @BGD_Time = Approve_Time
			, @BGD_Comm = Comment
		from @TableTmp 
		where LineId =4;

		Select @KT_Name = UserName
			, @KT_Arrpove = Approve_Status
			, @KT_Time = Approve_Time
			, @KT_Comm = Comment
		from @TableTmp 
		where LineId = case ISNULL(@NT,'') when '' then 5 else 5 end; --pqhuy1987 20180619 gi?m 1 b?c k? toán
	end
	else if (@BGroup = 'TBXD')
	begin
		Select @1_Name = ''
		, @1_Arrpove = ''
		, @1_Time = ''
		, @1_Comm = '';

		Select @2_Name = UserName
		, @2_Arrpove = Approve_Status
		, @2_Time = Approve_Time
		, @2_Comm = Comment
		from @TableTmp 
		where LineId =1;

		Select @3_Name = UserName
		, @3_Arrpove = Approve_Status
		, @3_Time = Approve_Time
		, @3_Comm = Comment
		from @TableTmp 
		where LineId =3;

		Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
		from @TableTmp 
		where LineId =5;

		Select @BGD_Name = UserName
			, @BGD_Arrpove = Approve_Status
			, @BGD_Time = Approve_Time
			, @BGD_Comm = Comment
		from @TableTmp 
		where LineId =6;

		Select @KT_Name = UserName
			, @KT_Arrpove = Approve_Status
			, @KT_Time = Approve_Time
			, @KT_Comm = Comment
		from @TableTmp 
		where LineId = case ISNULL(@NT,'') when '' then 7 else 7 end;--pqhuy1987 20180619 gi?m 1 b?c k? toán
	end
	
	if (ISNULL(@NT,'') <> '')
	begin
		SET @BGD_Name = ''
		SET @KT_Name = ''
	end
	Select --@TBP_Name as 'TBP_Name', @TBP_Arrpove as 'TBP_Approve', @TBP_Time as 'TBP_Time', ISNULL(@TBP_Comm,'') as 'TBP_Comm'
		    @1_Name as '1_Name', @1_Arrpove as '1_Arrpove', @1_Time as '1_Time', ISNULL(@1_Comm,'') as '1_Comm'
		  , @2_Name as '2_Name', @2_Arrpove as '2_Arrpove', @2_Time as '2_Time', ISNULL(@2_Comm,'') as '2_Comm'
		  , @3_Name as '3_Name', @3_Arrpove as '3_Arrpove', @3_Time as '3_Time', ISNULL(@3_Comm,'') as '3_Comm'
		  , @CCM_Name as 'CCM_Name', @CCM_Arrpove as 'CCM_Arrpove', @CCM_Time as 'CCM_Time', ISNULL(@CCM_Comm,'') as 'CCM_Comm'
		  , @BGD_Name as 'BGD_Name', @BGD_Arrpove as 'BGD_Arrpove', @BGD_Time as 'BGD_Time', ISNULL(@BGD_Comm,'') as 'BGD_Comm'
		  , @KT_Name as 'KT_Name', @KT_Arrpove as 'KT_Arrpove', @KT_Time as 'KT_Time', ISNULL(@KT_Comm,'') as 'KT_Comm';
END
GO

CREATE PROCEDURE [dbo].[KLTT_Check_Approve]
	@DocEntry int
as
begin
Select COUNT(*) from [@KLTT] where DocEntry = @DocEntry and [Status]='C' and [Canceled] <> 'Y';
--select *
--from
--(
--select DocNum,'So' as SCT from [@KLTT] where status = 'C' and DocNum =@DocEntry)z1
--FULL OUTER JOIN
--(
--	select 'So' as SCT1
--)z2

--ON(z1.SCT = z2.SCT1)
end