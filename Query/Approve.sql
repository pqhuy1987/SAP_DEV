Select DocEntry
,U_FIPROJECT as Project
,U_Period as 'Period'
,U_BGroup as 'BGroup'
,case U_BType when 1 then N'Tạm ứng'
			when 2 then N'Thanh toán'
			when 3 then N'Quyết toán'
			end as 'Bill Type'
,U_BPCode as 'BPCode'
,U_BPName as 'BPName'
,U_DATEFROM as 'From Date'
,U_DATETO as 'To Date'
from [@KLTT]
where 
Canceled = 'N'
and Status = 'O';

--Get Last Posting
Select top 1 * from [@KLTT_APPROVE] c where c.DocEntry=62 and c.U_Status is not null order by c.LineId desc;

--Get User Info
Select dept
, (Select [Name] from OUDP where Code=dept) as deptName
,position 
,(Select [Name] from OHPS where posID=position) as posName
from OHEM 
where userID = (Select t.USERID from OUSR t where t.User_Code='manager');

------------------------------------------------
ALTER PROCEDURE [dbo].[GetList_Bill_Approve]
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
,case a.U_BType when 1 then N'Tạm ứng'
			when 2 then N'Thanh toán'
			when 3 then N'Quyết toán'
			end as 'Bill Type'
,a.U_BPCode as 'BPCode'
,a.U_BPName as 'BPName'
,a.U_DATEFROM as 'From Date'
,a.U_DATETO as 'To Date'
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
from [@KLTT] a where 
a.Canceled = 'N'
and a.Status = 'O') z
where z.POST_LVL = @Dept_Code;
END

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

-----------------------------------------------------------------------------------------------
ALTER PROCEDURE [dbo].[Approve_Bill_LV]
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
and U_Position = @Pos_Code
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

----------------------------------------------------------------------
ALTER PROCEDURE [dbo].[GetList_Approve]
	-- Add the parameters for the stored procedure here
	@BGroup as varchar(50)
AS
BEGIN
	--Map position OHPS.posID
	SET NOCOUNT ON;
	IF (@BGroup = 'XD')
		Select 3 as 'LEVEL', 5 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position';
	IF (@BGroup = 'CD')
		Select 3 as 'LEVEL', 6 as 'Position'
		union all
		Select 5 as 'LEVEL', 2 as 'Position'
		union all
		Select 5 as 'LEVEL', 1 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position'
	IF (@BGroup = 'CDXD')
		Select 3 as 'LEVEL', 6 as 'Position'
		union all
		Select 3 as 'LEVEL', 5 as 'Position'
		union all
		Select 5 as 'LEVEL', 2 as 'Position'
		union all
		Select 5 as 'LEVEL', 1 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position'
	IF (@BGroup = 'TB')
		Select 2 as 'LEVEL', 1 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position'
	IF (@BGroup = 'TBXD')
		Select 3 as 'LEVEL', 5 as 'Position'
		union all
		Select 2 as 'LEVEL', 2 as 'Position'
		union all
		Select 2 as 'LEVEL', 1 as 'Position'
		union all
		Select -2 as 'LEVEL', 2 as 'Position'
		union all
		Select -2 as 'LEVEL', 1 as 'Position'
		union all
		Select 1 as 'LEVEL', 2 as 'Position'
		union all
		Select 1 as 'LEVEL', 1 as 'Position'
		union all
		Select 6 as 'LEVEL', 3 as 'Position'
END