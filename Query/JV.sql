CREATE PROCEDURE [dbo].[JV_Approve_LV]
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
and U_Position = @Pos_Code
and U_Level = @Dept_Code
and (U_Status is null or U_Status ='4')
and LineID = (Select Min(LineID) from [@JV_APROVE_D]
				where  DocEntry = @DocEntry
				and U_Level = @Dept_Code
				and U_Position = @Pos_Code
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

CREATE PROCEDURE [dbo].[JV_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
--Get User Info - Dept - Position
Select USER_CODE, a.FirstName +' '+ a.LastName as 'NAME' from OHEM a inner join OUSR b on a.USERID = b.UserID
where a.dept =
(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = @DocEntry and U_Status is null order by LineID)
and a.position =
(Select top 1 U_Position from [@JV_APROVE_D] where DocEntry = @DocEntry and U_Status is null order by LineID);
END

GO

CREATE PROCEDURE [dbo].[JV_Reset_Approve_LV1]
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

CREATE PROCEDURE [dbo].[JV_GetList_Approve_Current]
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
Select * from 
(Select a.BatchNum
,b.DocEntry
,a.NumOfTrans
,Format( a.LocTotal ,'N0','en-US' ) as 'Total'
--, a.Project
,CONVERT(varchar(11), a.DateID, 113) as 'Create Date'
,(Select t.USER_CODE from OUSR t where t.Userid = a.UserSign) as 'Username'
,(Select Top 1 Project from OBTF where BatchNum = a.BatchNum) as 'Project'
,b.U_Type as 'Type'
,(Select top 1 U_Level from [@JV_APROVE_D] where DocEntry = b.DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) as 'POST_LVL'
from OBTD a inner join [@JV_APPROVE] b on a.BatchNum = b.U_JVBatchNum where a.Status = 'O' and b.Status not in ('C')) z
where z.POST_LVL =@Dept_Code;
END

GO

CREATE PROCEDURE [dbo].[JV_Get_Total_Approve]
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

CREATE PROCEDURE [dbo].[JV_GetList_Approve]
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

CREATE PROCEDURE [dbo].[JV_Check_BCH_VP]
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