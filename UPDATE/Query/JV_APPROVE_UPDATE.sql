ALTER PROCEDURE [dbo].[JV_Approve_LV]
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
Update [@JV_APROVE_D] 
set U_Usr = @UserName
	,U_Time=CONVERT(varchar(30), GETDATE(), 113)
	,U_Status = @Status
	,U_Comment = @Comment
where DocEntry = @DocEntry
and (U_Position = @Pos_Code or (@UserName='Thuy.nguyen' and U_Level = 6))
and U_Level = @Dept_Code
and (U_Status is null or U_Status ='4')
and (Select ISNULL(Canceled,'N') from [@JV_APPROVE] where DocEntry = @DocEntry) <> 'Y'
and LineID = (Select Min(LineID) from [@JV_APROVE_D]
				where  DocEntry = @DocEntry
				and U_Level = @Dept_Code
				and (U_Position = @Pos_Code or (@UserName='Thuy.nguyen' and U_Level = 6))
				and (U_Status is null or U_Status ='4'));
--Update them truong hop khi truong phong duyet ko qua nhan vien
if @Pos_Code = 1
		Update [@JV_APROVE_D] 
		set U_Usr = @UserName
			,U_Time=CONVERT(varchar(30), GETDATE(), 113)
			,U_Status = '3'
		where DocEntry = @DocEntry
		and U_Position = 2
		and U_Level = @Dept_Code
		and U_Status is null
		and LineID = (Select Min(LineID) from [@JV_APROVE_D]
						where  DocEntry = @DocEntry
						and U_Level = @Dept_Code
						and U_Position = 2
						and (U_Status is null));
SELECT @Update_Row = @@ROWCOUNT;
RETURN @Update_Row;
END

GO

ALTER PROCEDURE [dbo].[JV_Check_BCH_VP]
	-- Add the parameters for the stored procedure here
	@BatchNum as int
AS
BEGIN
Select top 1 c.U_LCP,ISNULL(b.U_Status,-99) as 'APPROVED' 
from [@JV_APPROVE] a inner join [@JV_APROVE_D] b
on a.DocEntry = b.DocEntry 
inner join [OBTF] c on a.U_JVBatchNum = c.BatchNum
where BatchNum = @BatchNum
order by b.LineID desc;
END

GO

ALTER PROCEDURE [dbo].[JV_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
Declare @DeptCode as int
Declare @PosCode as int
Declare @FProject as varchar(50)

	Select top 1 @DeptCode = ISNULL(U_Level,''),@PosCode=ISNULL(U_Position,'') from [@JV_APROVE_D] 
	where DocEntry = @DocEntry 
	and U_Status is null 
	order by LineID;

	Select Top 1 @FProject = Project from OBTF 
	where BatchNum = (Select U_JVBatchNum from [@JV_APPROVE] where DocEntry = @DocEntry);

	Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name
	from OHEM a inner join OUSR b on a.USERID = b.UserID
	left join HTM1 c on c.empID=a.empID
	inner join OHTM d on c.teamID = d.teamID
	where a.dept = @DeptCode
	and a.position = @PosCode
	and d.name = @FProject;
END

GO

ALTER PROCEDURE [dbo].[JV_Get_Total_Approve]
	-- Add the parameters for the stored procedure here
	@BatchNum as int
AS
BEGIN
Select 
ProfitCode as 'Ma Phong Ban'
,(Select OcrName from OOCR where OcrCode = ProfitCode) as 'Ten Phong Ban'
,U_MANCC as 'Ma NCC'
,U_TENNCC as 'Ten NCC'
,U_MACP as 'Ma CP'
,U_TENCP as 'Ten CP'
,U_NOIDUNG as 'Noi dung'
,Project as 'Du an'
,Format( Debit ,'N0','en-US' ) as 'Ghi no'
,Format( Credit ,'N0','en-US' ) as 'Ghi co'
from BTF1 where BatchNum = @BatchNum;
END

GO

ALTER PROCEDURE [dbo].[JV_GetList_Approve]
	-- Add the parameters for the stored procedure here
	@Type as varchar(50),
	@Usr as varchar(100),
	@FinancialPrj as varchar(100)
AS
BEGIN

	DECLARE @Usr_Position as int
	DECLARE @Usr_Dept as int
	DECLARE @CHT as int
	SET NOCOUNT ON;
	Select @Usr_Dept = dept
	--, (Select [Name] from OUDP where Code=dept) as deptName
	,@Usr_Position = position 
	--, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@Usr);

	Select @CHT = SUM(b.Temp) from 
		(Select  distinct case right(U_BPTH,2) when 'ME' then -100 
		else 100 end as 'Temp' from OPMG where FIPROJECT = @FinancialPrj and STATUS != 'T') b;
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS


	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME
	IF (@Type = 'VP')
		Select @Usr_Dept as 'LEVEL', 2 as 'Position'
		union all
		Select @Usr_Dept as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 4 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position';
	IF (@Type = 'BCH')
		Select 3 as 'LEVEL', case  when @CHT <= -100 then 6 else 5 end as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position';
END

GO

--Show Current
ALTER PROCEDURE [dbo].[JV_GetList_Approve_Current]
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
--JV
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode/BPCode]
,(Select OcrName from OOCR where OcrCode=z.[ProfitCode/BPCode]) as 'Department/BPName'
,z.Remark
from 
(Select a.BatchNum
	,b.DocEntry
	,a.NumOfTrans
	,Format( a.LocTotal ,'N0','en-US' ) as 'Total'
	,CONVERT(varchar(11), a.DateID, 113) as 'Create Date'
	,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
	,(Select Top 1 Project from OBTF where BatchNum = a.BatchNum) as 'Project'
	,b.U_Type as 'Type'
	,(Select top 1 U_KTT from [OBTF] where BatchNum = a.BatchNum) as 'Period'
	,(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = b.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
	,(Select top 1 ProfitCode from [BTF1] where Batchnum = a.BatchNum and ProfitCode is not null) as 'ProfitCode/BPCode'
	,a.Remarks as 'Remark'
	from OBTD a inner join [@JV_APPROVE] b on a.BatchNum = b.U_JVBatchNum where a.Status = 'O' and b.Status not in ('C') and b.Canceled <> 'Y') z
	where z.POST_LVL =@Dept_Code
	and z.Project in 
	(Select y.name as 'FProject' from (
		Select * from HTM1 where empID =
		(Select empID from OHEM
		where UserID = (
		Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
Union all
--BILl VP
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode/BPCode]
,z.[Department/BPName]
,z.Remark
from
	(
		Select -1 as 'BatchNum',a.DocEntry,1 as 'NumofTrans'
		, case when a.U_BType = 1 then Format(a.U_Tamung,'N0','en-US') 
			else (
			--Select Format(SUM(isnull(U_GrossTotal,0)),'N0','en-US' ) from [@BILLVP1] where DocEntry = a.DocEntry
			Select Format(SUM(isnull(a1.U_GrossTotal,0)) - SUM(isnull(a2.U_KhautruTU,0)),'N0','en-US' )
			from [@BILLVP1] a1 inner join [@BILLVP] a2 on a1.DocEntry= a2.DocEntry 
			where a1.DocEntry = a.DocEntry
					and a1.U_GRPO_Key not in 
					(Select a3.U_GRPO_Key from [@BILLVP1] a3 inner join [@BILLVP] a4 on a3.DocEntry=a4.DocEntry 
					where a4.U_BPCode = a2.U_BPCode
					and a4.U_Period < a2.U_Period
					and a4.U_FProject = a2.U_FProject
					and a4.Canceled not in ('Y','C'))
			) end as 'Total'
		,CONVERT(varchar(11), a.CreateDate, 113) as 'Create Date'
		,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
		,a.U_FProject as 'Project'
		,'BILLVP' as 'Type'
		,a.U_Period as 'Period'
		,(Select top 1 U_Level from [@BILLVP2] where DocEntry = a.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
		,a.U_BPCode as 'ProfitCode/BPCode'
		,(Select CardName from OCRD where CardCode = a.U_BPCode) as 'Department/BPName'
		,a.Remark as 'Remark'
		from [@BILLVP] a
		where a.Canceled <> 'Y'
		and a.[Status] <> 'C'
	) z
where z.Project in 
	(
		Select y.name as 'FProject' from (
				Select * from HTM1 where empID =
				(Select empID from OHEM
				where UserID = (
				Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID
	)
and z.POST_LVL = @Dept_Code;
END

GO

--Show Approved
ALTER PROCEDURE [dbo].[JV_GetList_Approved_Current]
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

--JV
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode]
,(Select OcrName from OOCR where OcrCode=z.[ProfitCode]) as 'Department/BPName'
,z.Remark
from 
(Select a.BatchNum
,b.DocEntry
,a.NumOfTrans
,Format( a.LocTotal ,'N0','en-US' ) as 'Total'
,CONVERT(varchar(11), a.DateID, 113) as 'Create Date'
,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
,(Select Top 1 Project from OBTF where BatchNum = a.BatchNum) as 'Project'
,b.U_Type as 'Type'
,(Select top 1 U_KTT from [OBTF] where BatchNum = a.BatchNum) as 'Period'
,(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = b.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
,(Select top 1 ProfitCode from [BTF1] where Batchnum = a.BatchNum and ProfitCode is not null) as 'ProfitCode'
,a.Remarks as 'Remark'
from OBTD a inner join [@JV_APPROVE] b 
on a.BatchNum = b.U_JVBatchNum 
where a.Status = 'O' 
and b.Status = 'C'
) z
where 
z.Project in 
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
--and ((Select COUNT(U_Status) from [@JV_APROVE_D] where U_Usr = @UserName and DocEntry=z.DocEntry) >= 1
--	or
--and  z.Username = @UserName

Union all

--BILl VP
Select  z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode]
,z.[Department/BPName]
,z.Remark
from
	(
		Select -1 as 'BatchNum'
		, a.DocEntry
		, 1 as 'NumofTrans'
		, case when a.U_BType = 1 then Format(a.U_Tamung,'N0','en-US') 
			   else (
			   --Select Format(SUM(isnull(U_GrossTotal,0)),'N0','en-US' ) from [@BILLVP1] where DocEntry = a.DocEntry
			   Select Format(SUM(isnull(a1.U_GrossTotal,0)),'N0','en-US' ) 
			   from [@BILLVP1] a1 inner join [@BILLVP] a2 on a1.DocEntry= a2.DocEntry 
			   where a1.DocEntry = a.DocEntry
					and a1.U_GRPO_Key not in 
					(Select a3.U_GRPO_Key from [@BILLVP1] a3 inner join [@BILLVP] a4 on a3.DocEntry=a4.DocEntry 
					where a4.U_BPCode = a2.U_BPCode
					and a4.U_Period < a2.U_Period
					and a4.U_FProject = a2.U_FProject
					and a4.Canceled not in ('Y','C'))
			   ) end as 'Total'
		, CONVERT(varchar(11), a.CreateDate, 113) as 'Create Date'
		, (Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
		, a.U_FProject as 'Project'
		, 'BILLVP' as 'Type'
		, a.U_Period as 'Period'
		, (Select top 1 U_Level from [@BILLVP2] where DocEntry = a.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
		, a.U_BPCode as 'ProfitCode'
		,(Select CardName from OCRD where CardCode = a.U_BPCode) as 'Department/BPName'
		,a.Remark as 'Remark'
		from [@BILLVP] a
		where a.Canceled <> 'Y'
		and a.Status = 'C'
	) z
where z.Project in 
	(
		Select y.name as 'FProject' from (
				Select * from HTM1 where empID =
				(Select empID from OHEM
				where UserID = (
				Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID
	);
--and ((Select COUNT(U_Status) from [@BILLVP2] where U_Usr = @UserName and DocEntry=z.DocEntry) >= 1
	--or 
--and z.Username = @UserName;

END;

GO

--Show Rejected
ALTER PROCEDURE [dbo].[JV_GetList_Rejected_Current]
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

--JV
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode]
,(Select OcrName from OOCR where OcrCode=z.[ProfitCode]) as 'Department/BPName'
,z.Remark
from 
(Select a.BatchNum
,b.DocEntry
,a.NumOfTrans
,Format( a.LocTotal ,'N0','en-US' ) as 'Total'
,CONVERT(varchar(11), a.DateID, 113) as 'Create Date'
,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
,(Select Top 1 Project from OBTF where BatchNum = a.BatchNum) as 'Project'
,b.U_Type as 'Type'
,(Select top 1 U_KTT from [OBTF] where BatchNum = a.BatchNum) as 'Period'
,(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = b.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
,(Select top 1 ProfitCode from [BTF1] where Batchnum = a.BatchNum and ProfitCode is not null) as 'ProfitCode'
,a.Remarks as 'Remark'
from OBTD a inner join [@JV_APPROVE] b 
on a.BatchNum = b.U_JVBatchNum 
where --a.Status = 'O' 
--and 
b.Canceled = 'Y'
) z
where 
z.Project in 
(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
--and (Select COUNT(U_Status) from [@JV_APROVE_D] where U_Usr = @UserName and DocEntry=z.DocEntry) >= 1

Union all
--BILl VP
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode]
,z.[Department/BPName]
,z.Remark
from
	(
		Select -1 as 'BatchNum'
		, a.DocEntry
		, 1 as 'NumofTrans'
		, case when a.U_BType = 1 then Format(a.U_Tamung,'N0','en-US') 
			   else (
			   --Select Format(SUM(isnull(U_GrossTotal,0)),'N0','en-US' ) from [@BILLVP1] where DocEntry = a.DocEntry
				Select Format(SUM(isnull(a1.U_GrossTotal,0)),'N0','en-US' ) 
				from [@BILLVP1] a1 inner join [@BILLVP] a2 on a1.DocEntry= a2.DocEntry 
				where a1.DocEntry = a.DocEntry
					and a1.U_GRPO_Key not in 
					(Select a3.U_GRPO_Key from [@BILLVP1] a3 inner join [@BILLVP] a4 on a3.DocEntry=a4.DocEntry 
					where a4.U_BPCode = a2.U_BPCode
					and a4.U_Period < a2.U_Period
					and a4.U_FProject = a2.U_FProject
					and a4.Canceled not in ('Y','C'))
			   ) end as 'Total'
		, CONVERT(varchar(11), a.CreateDate, 113) as 'Create Date'
		, (Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
		, a.U_FProject as 'Project'
		, 'BILLVP' as 'Type'
		, a.U_Period as 'Period'
		, (Select top 1 U_Level from [@BILLVP2] where DocEntry = a.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
		, a.U_BPCode as 'ProfitCode'
		,(Select CardName from OCRD where CardCode = a.U_BPCode) as 'Department/BPName'
		,a.Remark as 'Remark'
		from [@BILLVP] a
		where a.Canceled = 'Y'
	) z
where z.Project in 
	(
		Select y.name as 'FProject' from (
				Select * from HTM1 where empID =
				(Select empID from OHEM
				where UserID = (
				Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID
	);
--and (Select COUNT(U_Status) from [@BILLVP2] where U_Usr = @UserName and DocEntry=z.DocEntry) >= 1;

END;

GO

--Show All
ALTER PROCEDURE [dbo].[JV_GetList_Current]
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
--JV
Select z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode/BPCode]
,(Select OcrName from OOCR where OcrCode=z.[ProfitCode/BPCode]) as 'Department/BPName'
,z.Remark
from 
(Select a.BatchNum
	,b.DocEntry
	,a.NumOfTrans
	,Format( a.LocTotal ,'N0','en-US' ) as 'Total'
	,CONVERT(varchar(11), a.DateID, 113) as 'Create Date'
	,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
	,(Select Top 1 Project from OBTF where BatchNum = a.BatchNum) as 'Project'
	,b.U_Type as 'Type'
	,(Select top 1 U_KTT from [OBTF] where BatchNum = a.BatchNum) as 'Period'
	,(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = b.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
	,(Select top 1 ProfitCode from [BTF1] where Batchnum = a.BatchNum and ProfitCode is not null) as 'ProfitCode/BPCode'
	,a.Remarks as 'Remark'
	from OBTD a inner join [@JV_APPROVE] b on a.BatchNum = b.U_JVBatchNum 
	--where a.Status = 'O' and b.Status not in ('C')
	) z
	where 
	--z.POST_LVL =@Dept_Code
	--and 
	z.Project in 
	(Select y.name as 'FProject' from (
		Select * from HTM1 where empID =
		(Select empID from OHEM
		where UserID = (
		Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
Union all
--BILl VP
Select  z.BatchNum, z.DocEntry, z.NumOfTrans, z.Total, z.[Create Date], z.Username, z.Project
, z.[Type], z.[Period], z.POST_LVL, z.[ProfitCode/BPCode]
,z.[Department/BPName]
,z.Remark
from
	(
		Select -1 as 'BatchNum',a.DocEntry,1 as 'NumofTrans'
		, case when a.U_BType = 1 then Format(a.U_Tamung,'N0','en-US') 
			else (
			--Select Format(SUM(isnull(U_GrossTotal,0)),'N0','en-US' ) from [@BILLVP1] where DocEntry = a.DocEntry
			Select Format(SUM(isnull(a1.U_GrossTotal,0)),'N0','en-US' ) 
			from [@BILLVP1] a1 inner join [@BILLVP] a2 on a1.DocEntry= a2.DocEntry 
			where a1.DocEntry = a.DocEntry
					and a1.U_GRPO_Key not in 
					(Select a3.U_GRPO_Key from [@BILLVP1] a3 inner join [@BILLVP] a4 on a3.DocEntry=a4.DocEntry 
					where a4.U_BPCode = a2.U_BPCode
					and a4.U_Period < a2.U_Period
					and a4.U_FProject = a2.U_FProject
					and a4.Canceled not in ('Y','C'))
			) end as 'Total'
		,CONVERT(varchar(11), a.CreateDate, 113) as 'Create Date'
		,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
		,a.U_FProject as 'Project'
		,'BILLVP' as 'Type'
		,a.U_Period as 'Period'
		,(Select top 1 U_Level from [@BILLVP2] where DocEntry = a.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
		,a.U_BPCode as 'ProfitCode/BPCode'
		,(Select CardName from OCRD where CardCode = a.U_BPCode) as 'Department/BPName'
		,a.Remark as 'Remark'
		from [@BILLVP] a
	) z
where z.Project in 
	(
		Select y.name as 'FProject' from (
				Select * from HTM1 where empID =
				(Select empID from OHEM
				where UserID = (
				Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID
	);
--and z.POST_LVL = @Dept_Code;
END

GO

ALTER PROCEDURE [dbo].[JV_Reset_Approve_LV1]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
DECLARE @First_LV as nvarchar(100)
DECLARE @Update_Row as int
--Get User Info - Dept - Position
Select top 1 @First_LV = U_Level from [@JV_APROVE_D] where DocEntry = @DocEntry order by LineId asc;
Update [@JV_APROVE_D] SET U_Usr = null, U_Time = null, U_Status = null, U_Comment = null
where DocEntry = @DocEntry
and U_Level = @First_LV;
SELECT @Update_Row = @@ROWCOUNT;
RETURN @Update_Row;
END

GO

ALTER PROCEDURE [dbo].[JV_Get_Data_BCH_Cover]
	-- Add the parameters for the stored procedure here
	@BatchNum as int
AS
BEGIN
DECLARE @LCP as varchar(10)
DECLARE @Period as int
DECLARE @Year_F as int
DECLARE @Project as varchar(100)
DECLARE @GT_KYNAY as decimal
DECLARE @TONG_GT as decimal
DECLARE @GT_KYTRUOC as decimal
DECLARE @DTTT as varchar(100)
DECLARE @PBTT as nvarchar(100)

Select top 1 @LCP = ISNULL(U_LCP,'')
, @Period = ISNULL(U_KTT,-1)
, @Year_F = ISNULL(YEAR(CreateDate),0)
, @Project = ISNULL(Project,'-1')
, @PBTT = ISNULL(U_PBTT,'-1')
from OBTF where BatchNum = @BatchNum order by TransId desc;

Select @GT_KYNAY = ISNULL(SUM(LocTotal),0) from OBTF
where BatchNum = @BatchNum;

Select @GT_KYTRUOC = ISNULL(SUM(LocTotal),0) from OBTF
where BatchNum in 
(
	Select BatchNum from OBTF
	where U_LCP = @LCP
	and ISNULL(Project,'-1') = @Project
	and U_KTT <= @Period - 1
	and ISNULL(U_PBTT,'-1') = @PBTT
);

Select @TONG_GT = ISNULL(SUM(LocTotal),0) from OBTF
where BatchNum in 
(
	Select BatchNum from OBTF
	where U_LCP = @LCP
	and ISNULL(Project,'-1') = @Project
	and YEAR(CreateDate) = @Year_F
	and BatchNum <= @BatchNum
);

if (@LCP = 'VP')
	Select top 1 @DTTT = isnull(U_DTTT,'VP_TYPE') from BTF1 where BatchNum = @BatchNum and U_DTTT is not null;
Select @GT_KYNAY as 'KYNAY'
,@GT_KYTRUOC as 'KYTRUOC'
,@GT_KYNAY + @GT_KYTRUOC as 'TONGGT'
,ISNULL(@DTTT,'BCH') as 'VP_TYPE'
,ISNULL((Select OcrName from OOCR where OcrCode = @PBTT),'') as 'PBTT'
,dbo.Num2Text(ISNULL(@GT_KYNAY,0)) as 'Sotienbangchu'
,@Project as 'FProject'
,@Period as 'Period';
END

GO

ALTER PROCEDURE [dbo].[JV_GET_HDINFO]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@BatchNum as int,
	--@Period as int,
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
	DECLARE @BP_Code as varchar(100);
	DECLARE @BP_Name as varchar(200);
	--Lay BP Code

	Select top 1 @BP_Code = isnull(U_MANCC,''), @BP_Name = isnull(U_TENNCC,'') from BTF1 where BatchNum = @BatchNum

	--Lay HD Nguyen tac
	Select top 1 @DocEntry_HDNT = isnull(AbsID,-1) 
	from OOAT
	where U_PRJ is null
	and BpCode = @BP_Code
	and Status ='A'
	and U_CGroup = 'VP'
	--and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;
	
	--Lay HD
	Select top 1 @DocEntry = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =48
	and BpCode = @BP_Code
	and Status ='A'
	and U_CGroup = 'VP'
	--and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL Thay the
	Select top 1 @DocEntry_PLTT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =140
	and BpCode = @BP_Code
	and Status ='A'
	and U_CGroup = @CGroup
	--and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL tang
	Select top 1 @DocEntry_PLT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =141
	and BpCode = @BP_Code
	and Status ='A'
	and U_CGroup ='VP'
	--and U_PUTYPE = @PUType
	and StartDate <= @ToDate
	order by AbsID desc;
	if (@DocEntry_HDNT > 0)
	begin
		--HD Nguyên tắc nếu có PL thay thế thì lấy PL Thay thế
		if (@DocEntry_PLTT > 0)
			begin
				Select x.*,@BP_Code as 'BPCode',@BP_Name as 'BPName'
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						--,U_SHD
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
						--,U_SHD
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
						--U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry_PLTT)
						--and 
						StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
			end
		--else if (@DocEntry_PLT > 0)
		--	begin
			 -- Không xảy ra - PL tăng gáng trên HĐ Nguyên tắc
		--	end
		else
			begin
				Select 
					AbsID
					,Number
					--,U_SHD
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
					,@BP_Code as 'BPCode'
					,@BP_Name as 'BPName'
				from OOAT 
				where 
					AbsID = @DocEntry_HDNT
					and Status = 'A'
					and Cancelled <> 'Y'
			end
	end
	else if (@DocEntry > 0)
	begin
		--Có HĐ -- Có PL Thay thế HĐ
		 Select top 1 @DocEntry_PLTT = isnull(AbsID,-1) from OOAT 
			where U_PRJ = @FinancialProject
			and Series =140
			and BpCode = @BP_Code
			and Status ='A'
			and Cancelled <> 'Y'
			and U_CGroup = 'VP'
			--and U_PUTYPE = @PUType
			and StartDate <= @ToDate
			and U_SHD in (Select NUMBER from OOAT where AbsID= @DocEntry)
			order by AbsID desc;
		if (@DocEntry_PLTT > 0)
		begin
			--Co PLTT Hợp đồng
			Select x.*,@BP_Code as 'BPCode',@BP_Name as 'BPName'
					from (
						--Phu luc thay the
						Select 
						AbsID
						,Number
						--,U_SHD
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
						--,U_SHD
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
						and 
						Series in (140,141)
						and StartDate <= @ToDate
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
		else
		begin
			--Chỉ có HĐ (hoặc có thêm PL tăng)
			Select x.*,@BP_Code as 'BPCode',@BP_Name as 'BPName'
					from (
						--Hợp đồng
						Select 
						AbsID
						,Number
						--,U_SHD
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
						--,U_SHD
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
						and 
						StartDate <= @ToDate
						and Series in (140,141)
						and Status ='A'
						and Cancelled <> 'Y') x
					order by x.AbsID desc
		end
	end
END
GO

ALTER PROCEDURE [dbo].[VPBILL_GET_FPROJECT]
	-- Add the parameters for the stored procedure here
	@Username as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 
	WHERE 
	--T0.[ValidFrom] >= '01-01-2017' 
	T0.[Active] = 'Y'
	and (T0.[PrjCode] like 'VPCTY%' or T0.[PrjCode] like 'VTTB%')
	and T0.[PrjCode] in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
END
GO

ALTER PROCEDURE [dbo].[VPBILL_GETLIST]
	-- Add the parameters for the stored procedure here
	@FProject as varchar(200),
	@BP_Code as varchar(100)
AS
BEGIN
	SET NOCOUNT ON;
	Select U_Period as "Period"
	,case U_BType when 1 then N'Tạm ứng'
				  when 2 then N'Thanh toán'
				  when 3 then N'Quyết toán' end as "Bill Type"
	,U_DateFr as "From"
	,U_DateTo as "To"
	,CreateDate as "Created Date"
	,U_FProject as "Financial Project"
	,DocNum as "Document Number"
	,Canceled as 'Rejected'
	from [@BILLVP] a
	where a.U_BPCode = @BP_Code
	and a.U_FProject = @FProject
	order by U_Period asc;
END
GO

ALTER PROCEDURE [dbo].[VPBILL_GETDATA]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100),
	@To_Date as date,
	@BP_Code as varchar(100)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @ProjectID as int;

	Select 
		(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as GoiThauKey
		,(Select Name from OPMG where AbsEntry = (Select ProjectID from OPHA where AbsEntry = a.U_ParentID1)) as GoiThauName
		,a.DocEntry as GRPOKey
		,a.LineNum as GRPORowKey
		,a.Dscription as DetailsName
		,a.U_CTCV as DetailsWork
		,a.unitMsr as UoM
		,a.Quantity as Quantity
		,a.Price as UPrice
		,ISNULL(a.LineTotal,0) + ISNULL(a.LineVat,0) as Gross_Total
		,a.LineTotal as Total
		,b.CreateDate
		,b.CardCode
		,a.Project as 'DA'
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
		,a.OcrCode as 'MaPB'
		,(Select OcrName from OOCR where OcrCode = a.OcrCode) as 'TenPB'
		,a.U_MACP as 'MACP'
		,a.U_TENCP as 'TenCP'
		--,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
		--,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
	from PDN1 a inner join OPDN b on a.DocEntry = b.DocEntry
	--left join [@KLTTA] c on c.DocEntry = @Last_DocEntry 
	--					and c.U_GPKey = a.DocEntry
	--					and c.U_GPDetailsKey = a.LineNum
	--					and b.U_RECTYPE = @BGroup
	where b.Project = @FinancialProject
		--and a.U_ParentID1 is not null
		and b.CardCode = @BP_Code
		and b.DocDate <= @To_Date
		and b.U_RECTYPE = 'vp'
		and b.CANCELED not in ('Y','C')
		--and (Select ISNULL(SUM(TYP),-1) from OPHA where AbsEntry = a.U_ParentID2) not in (11,12,13);
	--Union all
	--Select
	--	(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as GoiThauKey
	--	,(Select Name from OPMG where AbsEntry = (Select ProjectID from OPHA where AbsEntry = a.U_ParentID1)) as GoiThauName
	--	--dbo.FN_Get_Goi_Thau(a.U_ParentID1) as GoiThauKey
	--	--,(Select Name from OPHA where AbsEntry = dbo.FN_Get_Goi_Thau(a.U_ParentID1)) as GoiThauName 
	--	,a.DocEntry as GRPOKey
	--	,a.LineNum as GRPORowKey
	--	,a.Dscription as DetailsName
	--	,a.U_CTCV as DetailsWork
	--	,a.unitMsr as UoM
	--	,a.Quantity * -1 as Quantity
	--	,a.Price as UPrice
	--	,a.LineTotal as Total
	--	,b.CreateDate
	--	,b.CardCode
	--	--,(Select ProjectID from OPHA where AbsEntry = a.U_ParentID1) as ProjectNo
	--	,a.U_ParentID1
	--	,(Select Name from OPHA where AbsEntry = a.U_ParentID1) as Name1
	--	,a.U_ParentID2
	--	,(Select Name from OPHA where AbsEntry = a.U_ParentID2) as Name2
	--	,a.U_ParentID3
	--	,(Select Name from OPHA where AbsEntry = a.U_ParentID3) as Name3
	--	,a.U_ParentID4
	--	,(Select Name from OPHA where AbsEntry = a.U_ParentID4) as Name4
	--	,a.U_ParentID5
	--	,(Select Name from OPHA where AbsEntry = a.U_ParentID5) as Name5
	--	,b.U_RECTYPE
	--	,'GR' as 'TYPE'
	--	,ISNULL(c.U_CompleteRate,0) as Last_Complete_Rate
	--	,ISNULL(c.U_CompleteAmount,0) as Last_Complete_Amount
	--	from RPD1 a inner join ORPD b on a.DocEntry = b.DocEntry
	--	left join [@KLTTA] c on c.DocEntry = @Last_DocEntry 
	--					--and c.U_SubProjectKey = c.AbsEntry 
	--					and c.U_GPKey = a.DocEntry
	--					and c.U_GPDetailsKey = a.LineNum
	--	where a.Project = @FinancialProject
	--	and a.U_ParentID1 is not null
	--	and b.U_RECTYPE = @BGroup
	--	and b.CardCode = @BP_Code
	--	and b.U_PUTYPE = @PurchaseType
	--	and b.DocDate < @To_Date
	--	and b.CANCELED not in ('Y','C')
	--	and (Select ISNULL(SUM(TYP),-1) from OPHA where AbsEntry = a.U_ParentID2) not in (11,12,13);
END

GO

ALTER PROCEDURE [dbo].[VPBILL_GetList_Approve]
	-- Add the parameters for the stored procedure here
	@Usr as varchar(100)
AS
BEGIN

	DECLARE @Usr_Position as int
	DECLARE @Usr_Dept as int
	DECLARE @CHT as int
	SET NOCOUNT ON;
	Select @Usr_Dept = dept
	,@Usr_Position = position 
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@Usr);
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS

	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME
	Select X.*,Y.Name as 'DeptName',Z.name as 'PosName'
	from (
	Select @Usr_Dept as 'LEVEL', 2 as 'Position'
	union all
	Select @Usr_Dept as 'LEVEL', 1 as 'Position'
	union all
	Select 1 as 'LEVEL', 2 as 'Position'
	union all
	Select 1 as 'LEVEL', 1 as 'Position'
	union all
	Select 6 as 'LEVEL', 4 as 'Position'
	union all
	Select -2 as 'LEVEL', 2 as 'Position'
	union all
	Select -2 as 'LEVEL', 1 as 'Position') X
	inner join OUDP Y on X.LEVEL = Y.Code
	inner join OHPS Z on X.Position = Z.posID;
END

GO

ALTER PROCEDURE [dbo].[VPBILL_Get_Total_Approve]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
DECLARE @DocEntry_Pre as int
Select top 1 @DocEntry_Pre = ISNULL(DocEntry,-1) from [@BILLVP] 
where U_BType = 2
and U_Period < (Select U_Period from [@BILLVP] where DocEntry = @DocEntry)
and U_FProject = (Select U_FProject from [@BILLVP] where DocEntry = @DocEntry)
and U_BPCode = (Select U_BPCode from [@BILLVP] where DocEntry = @DocEntry)
and Canceled not in ('Y','C')
order by U_Period desc;

Select 
U_GRPO_Key as 'So chung tu'
,U_DistRule as 'Ma Phong Ban'
,U_DisRule_Name as 'Ten Phong Ban'
,U_MACP as 'Ma CP'
,U_TENCP as 'Ten CP'
,U_NOIDUNG as 'Noi dung'
,U_Project as 'Du an'
,Format( U_GrossTotal ,'N0','en-US' ) as 'Gia tri (bao gom VAT)'
from [@BILLVP1] where DocEntry = @DocEntry
and U_GRPO_Key not in 
(Select b.U_GRPO_Key from [@BILLVP1] b where b.DocEntry= @DocEntry_Pre);
END
GO

ALTER PROCEDURE [dbo].[VPBILL_Approve_LV]
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
Update [@BILLVP2] 
set U_Usr = @UserName
	,U_Time=CONVERT(varchar(30), GETDATE(), 113)
	,U_Status = @Status
	,U_Comment = @Comment
where DocEntry = @DocEntry
and (ISNULL(U_Position,@Pos_Code) = @Pos_Code or ((@UserName='Lan.nguyen') and U_Level = 6))
and (Select Canceled from [@BILLVP] where DocEntry = @DocEntry) <> 'Y'
and U_Level = @Dept_Code
and U_Status is null;
--Update them truong hop khi truong phong duyet ko qua nhan vien
if @Pos_Code = 1
		Update [@BILLVP2] 
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

ALTER PROCEDURE [dbo].[VPBILL_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
--Get User Info - Dept - Position
Declare @DeptCode as int
Declare @PosCode as int
Declare @FProject as varchar(250)
	Select top 1 @DeptCode = ISNULL(U_Level,''),@PosCode=ISNULL(U_Position,'') from [@BILLVP2] 
	where DocEntry = @DocEntry 
	and U_Status is null 
	order by LineID;

	Select @FProject=ISNULL(U_FPROJECT,'') from [@BILLVP] where DocEntry=@DocEntry;

	Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name
	from OHEM a inner join OUSR b on a.USERID = b.UserID
	left join HTM1 c on c.empID=a.empID
	inner join OHTM d on c.teamID = d.teamID
	where a.dept = @DeptCode
	and a.position = @PosCode
	and d.name = @FProject;
END
GO

ALTER PROCEDURE [dbo].[VPBILL_Get_Data_Cover]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
	DECLARE @GTDenKyNay as decimal
	DECLARE @GTKN as decimal
	DECLARE @GTKT as decimal
	DECLARE @TUKT as decimal
	DECLARE @TUKN as decimal
	DECLARE @HU as decimal
	DECLARE @TNCN as decimal
	DECLARE @BpCode as varchar(200)
	DECLARE @BpName as nvarchar(254)
	DECLARE @ToDate as date
	DECLARE @FinancialProject as varchar(200)
	DECLARE @Type as int
	DECLARE @Note as nvarchar(254)
	DECLARE @DocEntry_HDNT as int
	DECLARE @DocEntry_PLTT as int
	DECLARE @DocEntry_PLT as int
	DECLARE @DocEntry_HD as int

	DECLARE @GTHD as decimal -- Gia tri hop dong
	DECLARE @PLT as decimal --Gia tri phu luc tang
	DECLARE @NDHD as nvarchar(254) -- noi dung HD
	DECLARE @SHD as int -- So hop dong
	DECLARE @NGAYHD as date -- Ngay HD

	Select 
	 @GTDenKyNay = (Select  SUM(ISNULL(U_GrossTotal,0)) from [@BILLVP1] where DocEntry = a.DocEntry)
	,@TUKN = a.U_Tamung
	,@GTKT = (Select SUM(ISNULL(t1.U_GrossTotal,0))
				from [@BILLVP1] t1 inner join [@BILLVP] t2 on t1.DocEntry= t2.DocEntry
				where t2.U_BPCode=a.U_BPCode 
				and t2.U_Period = 			
					(Select Max(U_Period)
					from [@BILLVP] t 
					where t.U_BPCode=a.U_BPCode 
					and t.U_Period < a.U_Period
					and t.U_FProject = a.U_FProject
					and t.Canceled <> 'Y')
				and t2.Canceled <> 'Y'
				and t2.U_FProject = a.U_FProject)

				
	,@TUKT = (Select SUM(ISNULL(t3.U_Tamung,0))
			from [@BILLVP] t3 
			where t3.U_BPCode=a.U_BPCode 
			and t3.U_Period < a.U_Period 
			and t3.U_FProject = a.U_FProject
			and t3.U_BType = 1
			and t3.Canceled <> 'Y')
	,@HU = (Select SUM(ISNULL(t3.U_KhautruTU,0))
			from [@BILLVP] t3 
			where t3.U_BPCode=a.U_BPCode 
			and t3.U_Period <= a.U_Period 
			and t3.U_FProject = a.U_FProject
			and t3.U_BType = 2
			and t3.Canceled <> 'Y')
	--,@TNCN = case when SUBSTRING(a.U_BPCode,1,3) = 'DTC' then @GTKN * 0.1 else 0 end
	,@BpCode = a.U_BPCode
	,@BpName = (Select top 1 CardName from OCRD where CardCode = a.U_BPCode)
	,@ToDate = a.U_DateTo
	,@FinancialProject = a.U_FProject
	,@Type = a.U_BType
	,@Note = a.Remark
	from [@BILLVP] a
	where a.DocEntry= @DocEntry;

	--Lay HD Nguyen tac
	Select top 1 @DocEntry_HDNT = isnull(AbsID,-1) 
	from OOAT
	where U_PRJ is null
	and BpCode = @BpCode
	and Status ='A'
	and U_CGroup = 'VP'
	and StartDate <= @ToDate
	order by AbsID desc;
	
	--Lay HD
	Select top 1 @DocEntry_HD = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =48
	and BpCode = @BpCode
	and Status ='A'
	and U_CGroup = 'VP'
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL Thay the
	Select top 1 @DocEntry_PLTT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =141
	and BpCode = @BpCode
	and Status ='A'
	and U_CGroup = 'VP'
	and StartDate <= @ToDate
	order by AbsID desc;

	--Lay PL tang
	Select top 1 @DocEntry_PLT = isnull(AbsID,-1)
	from OOAT 
	where U_PRJ = @FinancialProject
	and Series =140
	and BpCode = @BpCode
	and Status ='A'
	and U_CGroup ='VP'
	and StartDate <= @ToDate
	order by AbsID desc;

	if (ISNULL(@DocEntry_HDNT,0) > 0)
	begin
		--HD Nguyên tắc nếu có PL thay thế thì lấy PL Thay thế
		if (ISNULL(@DocEntry_PLTT,0) > 0)
			begin
				--Phu luc thay the
				Select 
				  @SHD = U_SHD
				, @NDHD = Descript
				, @NGAYHD = StartDate
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
				from OOAT 
				where 
					AbsID = @DocEntry_HDNT
					and Status = 'A'
					and Cancelled <> 'Y'
			end
	end
	else if (ISNULL(@DocEntry_HD,0) > 0)
	begin
		--Có HĐ -- Có PL Thay thế HĐ
		 Select top 1 @DocEntry_PLTT = isnull(AbsID,-1) from OOAT 
			where U_PRJ = @FinancialProject
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
			--Co PLTT Hợp đồng
			Select 
				 @SHD = U_SHD
				,@NDHD = Descript
				, @NGAYHD = StartDate
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
			--Chỉ có HĐ (hoặc có thêm PL tăng)
			--Hợp đồng
			Select 
				  @SHD =  U_SHD
				, @NDHD = Descript
				, @NGAYHD = StartDate
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

	if (@Type = 1)
		begin
		Set @GTKN = ISNULL(@TUKN,0);
		Set @GTDenKyNay = ISNULL(@GTKT,0);
		end
	else
		Set @GTKN = ISNULL(@GTDenKyNay,0) - ISNULL(@GTKT,0) - ISNULL(@HU,0);

	if (SUBSTRING(@BpCode,1,3) = 'DTC')
		--Set @TNCN = (ISNULL(@GTKN,0) / 1.1) * 0.1;
		Set @TNCN = ISNULL(@GTKN,0) * 0.1;
	else 
		Set @TNCN = 0;

	Select @DocEntry as 'So'
		, @BpCode as 'BpCode'
		, @BpName as 'BpName'
		, @Note as 'Note'
		, @SHD as 'SoHD'
		, @NGAYHD as 'NgayHD'
		, @NDHD as 'Noidung'
		, ISNULL(@HU,0) as 'HU'
		, ISNULL(@GTHD,0) as 'GTHD'
		, ISNULL(@PLT,0) as 'PLT'
		, ISNULL(@GTKT,0) + ISNULL(@TUKT,0) as 'GTKT'
		, ISNULL(@GTKN,0) as 'GTKN'
		, ISNULL(@GTDenKyNay,0) + ISNULL(@TUKT,0) + ISNULL(@TUKN,0)  as 'GTDenKynay'
		, ISNULL(@TNCN,0) as 'TNCN'
		, ISNULL(@GTKN,0) - ISNULL(@TNCN,0) as 'GTthucnhan'
		, dbo.Num2Text(ISNULL(@GTKN,0) - ISNULL(@TNCN,0)) as 'Sotienbangchu';
END
GO

CREATE PROCEDURE [dbo].[VPBILL_Get_Approve_Process_Cover]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
	DECLARE @TBP_Name as nvarchar(254)
	DECLARE @TBP_Arrpove as nvarchar(254)
	DECLARE @TBP_Time as datetime
	DECLARE @TBP_Comm as nvarchar(254)

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

	INSERT INTO @TableTmp (LineId, Approve_Level,Usr,UserName,Position,Approve_Time,Comment,Approve_Status)
	Select a.LineId,a.U_Level,a.U_Usr, ISNULL(c.LastName,'') +' '+ ISNULL(c.MiddleName,'')+ ' '+ ISNULL(c.FirstName,''), a.U_Position,CONVERT(VARCHAR(50), a.U_Time, 113) as 'U_Time',a.U_Comment, case a.U_Status when 2 then 'Rejected' when 1 then 'Approved' when 3 then 'Bypass' end as 'U_Status'
	from [@BILLVP2] a left join OUSR b on a.U_Usr = b.USER_CODE
	left join OHEM c on b.USERID = c.userId
	where DocEntry = @DocEntry
	and LineId in (2,4,5,7);

	Select @TBP_Name = UserName
		, @TBP_Arrpove = Approve_Status
		, @TBP_Time = Approve_Time
		, @TBP_Comm = Comment
	from @TableTmp 
	where LineId =2;

	Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
	from @TableTmp 
	where LineId =4;

	Select @BGD_Name = UserName
		, @BGD_Arrpove = Approve_Status
		, @BGD_Time = Approve_Time
		, @BGD_Comm = Comment
	from @TableTmp 
	where LineId =5;

	Select @KT_Name = UserName
		, @KT_Arrpove = Approve_Status
		, @KT_Time = Approve_Time
		, @KT_Comm = Comment
	from @TableTmp 
	where LineId =7;
	
	Select @TBP_Name as 'TBP_Name', @TBP_Arrpove as 'TBP_Approve', @TBP_Time as 'TBP_Time', ISNULL(@TBP_Comm,'') as 'TBP_Comm'
		  , @CCM_Name as 'CCM_Name', @CCM_Arrpove as 'CCM_Arrpove', @CCM_Time as 'CCM_Time', ISNULL(@CCM_Comm,'') as 'CCM_Comm'
		  , @BGD_Name as 'BGD_Name', @BGD_Arrpove as 'BGD_Arrpove', @BGD_Time as 'BGD_Time', ISNULL(@BGD_Comm,'') as 'BGD_Comm'
		  , @KT_Name as 'KT_Name', @KT_Arrpove as 'KT_Arrpove', @KT_Time as 'KT_Time', ISNULL(@KT_Comm,'') as 'KT_Comm';

END
GO

ALTER PROCEDURE [dbo].[GET_MENUUID]
	@ReportName as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	SELECT top 1 MenuUID FROM OCMN where [Name]=@ReportName;
END;
GO

CREATE PROCEDURE [dbo].[JV_Get_Approve_Process_Cover]
	-- Add the parameters for the stored procedure here
	@BatchNum as int
AS
BEGIN
	DECLARE @JV_TYPE as nvarchar(50)

	DECLARE @1_Name as nvarchar(254)
	DECLARE @1_Arrpove as nvarchar(254)
	DECLARE @1_Time as datetime
	DECLARE @1_Comm as nvarchar(254)

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

	INSERT INTO @TableTmp (LineId, Approve_Level,Usr,UserName,Position,Approve_Time,Comment,Approve_Status)
	Select a.LineId,a.U_Level,a.U_Usr, ISNULL(c.LastName,'') +' '+ ISNULL(c.MiddleName,'')+ ' '+ ISNULL(c.FirstName,''), a.U_Position,CONVERT(VARCHAR(50), a.U_Time, 113) as 'U_Time',a.U_Comment, case a.U_Status when 2 then 'Rejected' when 1 then 'Approved' when 3 then 'Bypass' end as 'U_Status'
	from [@JV_APROVE_D] a left join OUSR b on a.U_Usr = b.USER_CODE
	left join OHEM c on b.USERID = c.userId
	where DocEntry = (Select top 1 DocEntry from [@JV_APPROVE] where U_JVBatchNum = @BatchNum);

	Select top 1 @JV_TYPE = U_Type from [@JV_APPROVE] 
	where U_JVBatchNum = @BatchNum;

	Select @1_Name = UserName
		, @1_Arrpove = Approve_Status
		, @1_Time = Approve_Time
		, @1_Comm = Comment
	from @TableTmp 
	where LineId = case ISNULL(@JV_TYPE,'') 
					when 'VP' then 2
					when 'BCH' then 1
					else 1 end;

	Select @CCM_Name = UserName
		, @CCM_Arrpove = Approve_Status
		, @CCM_Time = Approve_Time
		, @CCM_Comm = Comment
	from @TableTmp 
	where LineId = case ISNULL(@JV_TYPE,'') 
					when 'VP' then 4
					when 'BCH' then 3
					else 1 end;

	Select @BGD_Name = UserName
		, @BGD_Arrpove = Approve_Status
		, @BGD_Time = Approve_Time
		, @BGD_Comm = Comment
	from @TableTmp 
	where LineId = case ISNULL(@JV_TYPE,'') 
					when 'VP' then 5
					when 'BCH' then 4
					else 1 end;

	Select @KT_Name = UserName
		, @KT_Arrpove = Approve_Status
		, @KT_Time = Approve_Time
		, @KT_Comm = Comment
	from @TableTmp 
	where LineId =  case ISNULL(@JV_TYPE,'') 
					when 'VP' then 7
					when 'BCH' then 6
					else 1 end;
	
	Select @1_Name as '1_Name', @1_Arrpove as '1_Approve', @1_Time as '1_Time', ISNULL(@1_Comm,'') as '1_Comm'
		  , @CCM_Name as 'CCM_Name', @CCM_Arrpove as 'CCM_Arrpove', @CCM_Time as 'CCM_Time', ISNULL(@CCM_Comm,'') as 'CCM_Comm'
		  , @BGD_Name as 'BGD_Name', @BGD_Arrpove as 'BGD_Arrpove', @BGD_Time as 'BGD_Time', ISNULL(@BGD_Comm,'') as 'BGD_Comm'
		  , @KT_Name as 'KT_Name', @KT_Arrpove as 'KT_Arrpove', @KT_Time as 'KT_Time', ISNULL(@KT_Comm,'') as 'KT_Comm';

END
GO