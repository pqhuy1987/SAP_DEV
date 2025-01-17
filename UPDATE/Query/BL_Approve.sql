ALTER PROCEDURE [dbo].[BL_Update_Post_Level_HD] 
	@AbsID as int,
	@Blanket_Type as varchar,
	@Blanket_Level as int,
	@Usr as varchar(50),
	@Approve as int,
	@Usr_Comment as nvarchar(254)
AS
SET NOCOUNT OFF
DECLARE @Update_Row as int
BEGIN
	IF @Blanket_Level = 1
		BEGIN
		UPDATE OOAT SET U_Apprv1 = @Approve, U_UsrApprv1 = @Usr, U_DTApprv1 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv1 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
		END
	IF @Blanket_Level = 2
	BEGIN
		UPDATE OOAT SET U_Apprv2 = @Approve, U_UsrApprv2 = @Usr, U_DTApprv2 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv2 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 3
	BEGIN
		UPDATE OOAT SET U_Apprv3 = @Approve, U_UsrApprv3 = @Usr, U_DTApprv3 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv3 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 4
	BEGIN
		UPDATE OOAT SET U_Apprv4 = @Approve, U_UsrApprv4 = @Usr, U_DTApprv4 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv4 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 5
	BEGIN
		UPDATE OOAT SET U_Apprv5 = @Approve, U_UsrApprv5 = @Usr, U_DTApprv5 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv5 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 6
	BEGIN
		UPDATE OOAT SET U_Apprv6 = @Approve, U_UsrApprv6 = @Usr, U_DTApprv6 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv6 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 7
	BEGIN
		UPDATE OOAT SET U_Apprv7 = @Approve, U_UsrApprv7 = @Usr, U_DTApprv7 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv7 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 8
	BEGIN
		UPDATE OOAT SET U_Apprv8 = @Approve, U_UsrApprv8 = @Usr, U_DTApprv8 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv8 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 9
	BEGIN
		UPDATE OOAT SET U_Apprv9 = @Approve, U_UsrApprv9 = @Usr, U_DTApprv9 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv9 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 10
	BEGIN
		UPDATE OOAT SET U_Apprv10 = @Approve, U_UsrApprv10 = @Usr, U_DTApprv10 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv10 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 11
	BEGIN
		UPDATE OOAT SET U_Apprv11 = @Approve, U_UsrApprv11 = @Usr, U_DTApprv11 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv11 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 12
	BEGIN
		UPDATE OOAT SET U_Apprv12 = @Approve, U_UsrApprv12 = @Usr, U_DTApprv12 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv12 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	RETURN @Update_Row;
END
RETURN

GO

ALTER PROCEDURE [dbo].[BL_Update_Post_Level_HD_WithNote] 
	@AbsID as int,
	@Blanket_Type as varchar,
	@Blanket_Level as int,
	@Usr as varchar(50),
	@Lvl1 as int,
	@Approve as int,
	@Usr_Comment as nvarchar(254)
AS
SET NOCOUNT OFF
DECLARE @Update_Row as int
DECLARE @Dept_Create as int
DECLARE @CGroup as varchar(50)
BEGIN
Select @Dept_Create = b.dept, @CGroup = a.U_CGroup
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.AbsID = @AbsID;

	IF @Blanket_Level = 1
		BEGIN
		UPDATE OOAT SET U_Apprv1 = @Approve, U_UsrApprv1 = @Usr, U_DTApprv1 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv1 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
		END
	IF @Blanket_Level = 2
	BEGIN
		UPDATE OOAT SET U_Apprv2 = @Approve, U_UsrApprv2 = @Usr, U_DTApprv2 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv2 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 3
	BEGIN
		UPDATE OOAT SET U_Apprv3 = @Approve, U_UsrApprv3 = @Usr, U_DTApprv3 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv3 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 4
	BEGIN
		UPDATE OOAT SET U_Apprv4 = @Approve, U_UsrApprv4 = @Usr, U_DTApprv4 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv4 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 5
	BEGIN
		UPDATE OOAT SET U_Apprv5 = @Approve, U_UsrApprv5 = @Usr, U_DTApprv5 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv5 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 6
	BEGIN
		UPDATE OOAT SET U_Apprv6 = @Approve, U_UsrApprv6 = @Usr, U_DTApprv6 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv6 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 7
	BEGIN
		UPDATE OOAT SET U_Apprv7 = @Approve, U_UsrApprv7 = @Usr, U_DTApprv7 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv7 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 8
	BEGIN
		UPDATE OOAT SET U_Apprv8 = @Approve, U_UsrApprv8 = @Usr, U_DTApprv8 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv8 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 9
	BEGIN
		UPDATE OOAT SET U_Apprv9 = @Approve, U_UsrApprv9 = @Usr, U_DTApprv9 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv9 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 10
	BEGIN
		UPDATE OOAT SET U_Apprv10 = @Approve, U_UsrApprv10 = @Usr, U_DTApprv10 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv10 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 11
	BEGIN
		UPDATE OOAT SET U_Apprv11 = @Approve, U_UsrApprv11 = @Usr, U_DTApprv11 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv11 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Blanket_Level = 12
	BEGIN
		UPDATE OOAT SET U_Apprv12 = @Approve, U_UsrApprv12 = @Usr, U_DTApprv12 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv12 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		SELECT @Update_Row = @@ROWCOUNT;
	END
	IF @Lvl1 = 1
	BEGIN
		IF  (@CGroup ='CDXD')
			UPDATE OOAT SET U_Apprv2 = '2', U_UsrApprv2 = @Usr, U_DTApprv2 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv2 = @Usr_Comment
			where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
		else
			UPDATE OOAT SET U_Apprv1 = '2', U_UsrApprv1 = @Usr, U_DTApprv1 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv1 = @Usr_Comment
			where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
	END
	IF @Lvl1 = 2
	BEGIN
		UPDATE OOAT SET U_Apprv2 = '2', U_UsrApprv2 = @Usr, U_DTApprv2 = CONVERT(VARCHAR, GETDATE(), 103) +' ' + CONVERT(VARCHAR, GETDATE(), 24), U_CommApprv2 = @Usr_Comment
		where AbsID = @AbsID and BpType = @Blanket_Type and Status = 'D' and Cancelled = 'N';
	END
	RETURN @Update_Row;
END
RETURN

GO

-- lấy danh sách duyệt hợp đồng
ALTER PROCEDURE [dbo].[BL_Get_List] 
	@UserName as nvarchar(100)
AS
SET NOCOUNT OFF
DECLARE @position as int
DECLARE @dept as int
BEGIN
	Select @dept=a.dept
	,@position = a.position
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

	Select [Agreement No],[Project],[BpCode],[BpName],[Descript],[GTHĐ],[Status],[Contract Group],[Purchase Type],[Creator],[Last Approved]
	
	from
	(Select AbsId , Number as 'Agreement No',U_PRJ as 'Project',BpCode,BpName,Descript,a.Status,U_CGroup as 'Contract Group',U_PUTYPE as 'Purchase Type' 
	,(Select Format( (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) ,'N0','en-US' ) from OAT1 b where b.AgrNo = AbsId) as N'GTHĐ'
	,case when U_CGroup = 'XD' and ISNULL(b.dept,-9) not in (1,2) and @position = 5 and (U_Apprv2 is null or U_Apprv10='2')  then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv2 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv2 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null  and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv2 is not null)) then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) not in (1,2) and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and (U_Apprv10 is not null and U_Apprv10 <> '2') then 1
		  
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and @position = 6 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv1 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 5 and U_Apprv1 is not null  and U_Apprv6 is null then 1
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv1 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null)) then 1
		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) not in (1,2) and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and (U_Apprv10 is not null and U_Apprv10 <> '2') then 1

		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @position = 6 and U_Apprv1 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @position = 5 and U_Apprv1 is not null and (U_Apprv2 is null or U_Apprv10='2')  then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv2 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 5 and U_Apprv2 is not null  and U_Apprv6 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv2 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv2 is not null))  then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) not in (1,2) and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) not in (1,2) and @position = 1 and @dept = 2 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv1 is not null and U_Apprv4 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv1 is not null and U_Apprv8 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null))  then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) not in (1,2) and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and @position = 5 and (U_Apprv2 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv2 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 2 and U_Apprv2 is not null  and U_Apprv6 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv2 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv2 is not null)) then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) not in (1,2) and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @position = 1 and @dept = ISNULL(b.dept,-9) and U_Apprv1 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv1 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv1 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null)) then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and (@position = 4 or @UserName in ('Lan.nguyen','Thuy.nguyen'))and U_Apprv10 is not null then 1

		  -- Bộ phận IT
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @position = 1 and (@dept = 18 or @dept = 7) and U_Apprv1 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @position = 1 and @dept = 23 and U_Apprv1 is not null and U_Apprv12 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = 4 and U_Apprv1 is not null and U_Apprv12 is not null and U_Apprv4 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = -2 and U_Apprv1 is not null and U_Apprv12 is not null and U_Apprv8 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv12 is not null)) then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and (@position = 4 or @UserName in ('Lan.nguyen','Thuy.nguyen'))and U_Apprv10 is not null then 1

		  --CCM Creates
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and @position = 1 and (U_Apprv1 is null or U_Apprv10 = '2') then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) = 1 and @dept = 4  and U_Apprv4 is null then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) = 1 and @dept = -2 and U_Apprv8 is null then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and ((U_Apprv4 is not null  and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv1 <> '2')) then 1
		  when U_CGroup = 'XD' and ISNULL(b.dept,-9) = 1 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'CD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and @position = 1 and (U_Apprv1 is null or U_Apprv10 = '2') then 1
		  when U_CGroup = 'CD' and @dept = 4 and ISNULL(b.dept,-9) = 1 and U_Apprv4 is null then 1
		  when U_CGroup = 'CD' and @dept = 5 and ISNULL(b.dept,-9) = 1 and U_Apprv6 is null then 1
		  when U_CGroup = 'CD' and @dept = -2 and ISNULL(b.dept,-9) = 1 and U_Apprv8 is null then 1
		  when U_CGroup = 'CD' and @dept = 1 and ISNULL(b.dept,-9) = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv1 <> '2')) then 1
		  when U_CGroup = 'CD' and @position = 3 and ISNULL(b.dept,-9) = 1 and U_Apprv10 is not null then 1

		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and @position = 1 and (U_Apprv1 is null or U_Apprv10 = '2') then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and @dept = 4 and U_Apprv4 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and @dept = 5 and U_Apprv6 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and @dept = -2 and U_Apprv8 is null then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv1 <> '2')) then 1
		  when U_CGroup = 'CDXD' and ISNULL(b.dept,-9) = 1 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = 2 and @position = 1 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = 4 and U_Apprv4 is null then 1
		  --when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = 2 and U_Apprv6 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = -2 and U_Apprv8 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv1 <> '2')) then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 1 and @position = 1 and @dept = 2 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 1 and @dept = 4 and U_Apprv1 is not null and U_Apprv4 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 1 and @dept = -2 and U_Apprv1 is not null and U_Apprv8 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 1 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null))  then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 1 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) = 1 and @position = 1 and @dept = ISNULL(b.dept,-9) and U_Apprv1 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) = 1 and @dept = 4 and U_Apprv1 is not null  and U_Apprv4 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) = 1 and @dept = -2 and U_Apprv1 is not null  and U_Apprv8 is null then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) = 1 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null)) then 1
		  when U_CGroup = 'VP' and ISNULL(b.dept,-9) not in (1,2) and (@position = 4 or @UserName in ('Lan.nguyen','Thuy.nguyen'))and U_Apprv10 is not null then 1

		  --TB Creates

		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 2 and @dept = 2 and @position = 1 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 2 and @dept = 4 and U_Apprv4 is null then 1
		  --when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 1 and @dept = 2 and U_Apprv6 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 2 and @dept = -2 and U_Apprv8 is null then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 2 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv6 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null and U_Apprv1 <> '2')) then 1
		  when U_CGroup = 'TBXD' and ISNULL(b.dept,-9) = 2 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 2 and @position = 1 and @dept = 2 and (U_Apprv1 is null or U_Apprv10='2' )then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 2 and @dept = 4 and U_Apprv1 is not null and U_Apprv4 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 2 and @dept = -2 and U_Apprv1 is not null and U_Apprv8 is null then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 2 and @dept = 1 and ((U_Apprv4 is not null and U_Apprv8 is not null and U_Apprv10 is null) or (U_Apprv10 = '2' and U_Apprv1 is not null))  then 1
		  when U_CGroup = 'TB' and ISNULL(b.dept,-9) = 2 and (@position = 3 or @UserName in ('Lan.nguyen','Thuy.nguyen')) and U_Apprv10 is not null then 1

		  else 0
	 end as 'Show'
	 ,ISNULL(b.lastName,'') +' ' + ISNULL(b.middleName,'') +' ' +ISNULL(b.firstName,'') as 'Creator'
	 --Last Approved
	 ,
	 (
	 Select top 1 (Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
		where userId = (Select UserId from OUSR where User_Code = z.Usr_Appr))  
	 from (
	 Select U_Apprv1 as 'Status',U_usrApprv1 as 'Usr_Appr',Convert(Datetime,U_DTApprv1,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv2 as 'Status',U_usrApprv2 as 'Usr_Appr',Convert(Datetime,U_DTApprv2,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv3 as 'Status',U_usrApprv3 as 'Usr_Appr',Convert(Datetime,U_DTApprv3,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv4 as 'Status',U_usrApprv4 as 'Usr_Appr',Convert(Datetime,U_DTApprv4,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv5 as 'Status',U_usrApprv5 as 'Usr_Appr',Convert(Datetime,U_DTApprv5,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv6 as 'Status',U_usrApprv6 as 'Usr_Appr',Convert(Datetime,U_DTApprv6,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv7 as 'Status',U_usrApprv7 as 'Usr_Appr',Convert(Datetime,U_DTApprv7,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv8 as 'Status',U_usrApprv8 as 'Usr_Appr',Convert(Datetime,U_DTApprv8,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv9 as 'Status',U_usrApprv9 as 'Usr_Appr',Convert(Datetime,U_DTApprv9,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv10 as 'Status',U_usrApprv10 as 'Usr_Appr',Convert(Datetime,U_DTApprv10,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv11 as 'Status',U_usrApprv11 as 'Usr_Appr',Convert(Datetime,U_DTApprv11,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv12 as 'Status',U_usrApprv12 as 'Usr_Appr',Convert(Datetime,U_DTApprv12,103) as 'Date_Appr' from OOAT where Number= a.Number
	 )z
	 where a.Status is not null
	 order by Date_Appr desc
	 ) as 'Last Approved'
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Status ='D' 
	and a.BpType = 'S'
	and a.Cancelled ='N') x
	where x.Show = 1
	and ( x.Project in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) or x.Project is null) ;
END
RETURN
GO

ALTER PROCEDURE [dbo].[BL_Get_List_Approved] 
	@UserName as nvarchar(100)
AS
SET NOCOUNT OFF
DECLARE @position as int
DECLARE @dept as int
BEGIN
	Select @dept=a.dept
	,@position = a.position
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

	Select [Agreement No],[Project],[BpCode],[BpName],[Descript],[GTHĐ],[Status],[Contract Group],[Purchase Type],[Creator],[Last Approved]
	
	from
	(Select AbsId , Number as 'Agreement No',U_PRJ as 'Project',BpCode,BpName,Descript,a.Status,U_CGroup as 'Contract Group',U_PUTYPE as 'Purchase Type' 
	,(Select Format( (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) ,'N0','en-US' ) from OAT1 b where b.AgrNo = AbsId) as N'GTHĐ'
	,1 as 'Show'
	 ,ISNULL(b.lastName,'') +' ' + ISNULL(b.middleName,'') +' ' +ISNULL(b.firstName,'') as 'Creator'
	 --Last Approved
	 ,
	 (
	 Select top 1 (Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
		where userId = (Select UserId from OUSR where User_Code = z.Usr_Appr))  
	 from (
	 Select U_Apprv1 as 'Status',U_usrApprv1 as 'Usr_Appr',Convert(Datetime,U_DTApprv1,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv2 as 'Status',U_usrApprv2 as 'Usr_Appr',Convert(Datetime,U_DTApprv2,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv3 as 'Status',U_usrApprv3 as 'Usr_Appr',Convert(Datetime,U_DTApprv3,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv4 as 'Status',U_usrApprv4 as 'Usr_Appr',Convert(Datetime,U_DTApprv4,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv5 as 'Status',U_usrApprv5 as 'Usr_Appr',Convert(Datetime,U_DTApprv5,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv6 as 'Status',U_usrApprv6 as 'Usr_Appr',Convert(Datetime,U_DTApprv6,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv7 as 'Status',U_usrApprv7 as 'Usr_Appr',Convert(Datetime,U_DTApprv7,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv8 as 'Status',U_usrApprv8 as 'Usr_Appr',Convert(Datetime,U_DTApprv8,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv9 as 'Status',U_usrApprv9 as 'Usr_Appr',Convert(Datetime,U_DTApprv9,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv10 as 'Status',U_usrApprv10 as 'Usr_Appr',Convert(Datetime,U_DTApprv10,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv11 as 'Status',U_usrApprv11 as 'Usr_Appr',Convert(Datetime,U_DTApprv11,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv12 as 'Status',U_usrApprv12 as 'Usr_Appr',Convert(Datetime,U_DTApprv12,103) as 'Date_Appr' from OOAT where Number= a.Number
	 )z
	 where a.Status is not null
	 order by Date_Appr desc
	 ) as 'Last Approved'
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Status ='A' 
	and a.Cancelled ='N') x
	where x.Show = 1
	and ( x.Project in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) 
	or x.Project is null);
END
RETURN

GO

ALTER PROCEDURE [dbo].[BL_Get_List_Rejected] 
	@UserName as nvarchar(100)
AS
SET NOCOUNT OFF
DECLARE @position as int
DECLARE @dept as int
BEGIN
	Select @dept=a.dept
	,@position = a.position
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

	Select [Agreement No],[Project],[BpCode],[BpName],[Descript],[GTHĐ],[Status],[Contract Group],[Purchase Type],[Creator],[Last Approved]
	
	from
	(Select AbsId , Number as 'Agreement No',U_PRJ as 'Project',BpCode,BpName,Descript,a.Status,U_CGroup as 'Contract Group',U_PUTYPE as 'Purchase Type' 
	,(Select Format( (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) ,'N0','en-US' ) from OAT1 b where b.AgrNo = AbsId) as N'GTHĐ'
	,1 as 'Show'
	 ,ISNULL(b.lastName,'') +' ' + ISNULL(b.middleName,'') +' ' +ISNULL(b.firstName,'') as 'Creator'
	 
	 --Last Approved
	 ,
	 (
	 Select top 1 (Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
		where userId = (Select UserId from OUSR where User_Code = z.Usr_Appr))  
	 from (
	 Select U_Apprv1 as 'Status',U_usrApprv1 as 'Usr_Appr',Convert(Datetime,U_DTApprv1,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv2 as 'Status',U_usrApprv2 as 'Usr_Appr',Convert(Datetime,U_DTApprv2,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv3 as 'Status',U_usrApprv3 as 'Usr_Appr',Convert(Datetime,U_DTApprv3,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv4 as 'Status',U_usrApprv4 as 'Usr_Appr',Convert(Datetime,U_DTApprv4,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv5 as 'Status',U_usrApprv5 as 'Usr_Appr',Convert(Datetime,U_DTApprv5,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv6 as 'Status',U_usrApprv6 as 'Usr_Appr',Convert(Datetime,U_DTApprv6,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv7 as 'Status',U_usrApprv7 as 'Usr_Appr',Convert(Datetime,U_DTApprv7,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv8 as 'Status',U_usrApprv8 as 'Usr_Appr',Convert(Datetime,U_DTApprv8,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv9 as 'Status',U_usrApprv9 as 'Usr_Appr',Convert(Datetime,U_DTApprv9,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv10 as 'Status',U_usrApprv10 as 'Usr_Appr',Convert(Datetime,U_DTApprv10,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv11 as 'Status',U_usrApprv11 as 'Usr_Appr',Convert(Datetime,U_DTApprv11,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv12 as 'Status',U_usrApprv12 as 'Usr_Appr',Convert(Datetime,U_DTApprv12,103) as 'Date_Appr' from OOAT where Number= a.Number
	 )z
	 where a.Status is not null
	 order by Date_Appr desc
	 ) as 'Last Approved'
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Status not in('A','D')) x
	where x.Show = 1
	and (x.Project in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID) 
	or x.Project is null);
END
RETURN

GO

ALTER PROCEDURE [dbo].[BL_Get_List_All] 
	@UserName as nvarchar(100)
AS
SET NOCOUNT OFF
DECLARE @position as int
DECLARE @dept as int
BEGIN
	Select @dept=a.dept
	,@position = a.position
from 
(
	Select dept
	, (Select [Name] from OUDP where Code=dept) as deptName
	,position 
	, (Select [Name] from OHPS where posID=position) as posName
	from OHEM 
	where userID = (Select t.USERID from OUSR t where t.User_Code=@UserName)) a;

	Select [Agreement No],[Project],[BpCode],[BpName],[Descript],[GTHĐ],[Status],[Contract Group],[Purchase Type],[Creator],[Last Approved]
	
	from
	(Select AbsId , Number as 'Agreement No',U_PRJ as 'Project',BpCode,BpName,Descript,a.Status,U_CGroup as 'Contract Group',U_PUTYPE as 'Purchase Type' 
	,(Select Format( (SUM(b.PlanQty*b.UnitPrice) + SUM(b.PlanAmtLC)) ,'N0','en-US' ) from OAT1 b where b.AgrNo = AbsId) as N'GTHĐ'
	,1 as 'Show'
	 ,ISNULL(b.lastName,'') +' ' + ISNULL(b.middleName,'') +' ' +ISNULL(b.firstName,'') as 'Creator'
	 ,b.dept
	 --Last Approved
	 ,
	 (
	 Select top 1 (Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') from OHEM
		where userId = (Select UserId from OUSR where User_Code = z.Usr_Appr))  
	 from (
	 Select U_Apprv1 as 'Status',U_usrApprv1 as 'Usr_Appr',Convert(Datetime,U_DTApprv1,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv2 as 'Status',U_usrApprv2 as 'Usr_Appr',Convert(Datetime,U_DTApprv2,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv3 as 'Status',U_usrApprv3 as 'Usr_Appr',Convert(Datetime,U_DTApprv3,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv4 as 'Status',U_usrApprv4 as 'Usr_Appr',Convert(Datetime,U_DTApprv4,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv5 as 'Status',U_usrApprv5 as 'Usr_Appr',Convert(Datetime,U_DTApprv5,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv6 as 'Status',U_usrApprv6 as 'Usr_Appr',Convert(Datetime,U_DTApprv6,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv7 as 'Status',U_usrApprv7 as 'Usr_Appr',Convert(Datetime,U_DTApprv7,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv8 as 'Status',U_usrApprv8 as 'Usr_Appr',Convert(Datetime,U_DTApprv8,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv9 as 'Status',U_usrApprv9 as 'Usr_Appr',Convert(Datetime,U_DTApprv9,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv10 as 'Status',U_usrApprv10 as 'Usr_Appr',Convert(Datetime,U_DTApprv10,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv11 as 'Status',U_usrApprv11 as 'Usr_Appr',Convert(Datetime,U_DTApprv11,103) as 'Date_Appr' from OOAT where Number= a.Number
	 Union all
	 Select U_Apprv12 as 'Status',U_usrApprv12 as 'Usr_Appr',Convert(Datetime,U_DTApprv12,103) as 'Date_Appr' from OOAT where Number= a.Number
	 )z
	 where a.Status is not null
	 order by Date_Appr desc
	 ) as 'Last Approved'
	from OOAT a left join OHEM b on a.UserSign = b.userId
	--where a.Status not in('A','D')
	) x
	where x.Show = 1
	and (x.Project in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Username))) x inner join OHTM y on x.teamID = y.teamID)
	or x.Project is null ) ;
END
RETURN

GO

ALTER PROCEDURE [dbo].[BL_Get_Aprrove_Process] 
	@BlanketNo as int
AS
SET NOCOUNT OFF
DECLARE @BlanketType as varchar(100)
DECLARE @Dept_Create as int
BEGIN
	Select  @BlanketType = a.U_CGroup, @Dept_Create = ISNULL(b.dept,-9)
	from OOAT a left join OHEM b on a.UserSign = b.userId 
	where a.Number = @BlanketNo;
	if (@Dept_Create not in (1,2)) --Not CCM Created
	begin
		if (@BlanketType ='XD')
			Select N'Dư án XD - Chỉ huy trưởng XD' as 'Dept',dbo.BL_Status_Name(U_Apprv2) as 'Status',U_usrApprv2 as 'Approved by',U_DTApprv2 as 'Approved on' ,U_CommApprv2 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3 as 'Approved by',U_DTApprv3 as 'Approved on',U_CommApprv3 as 'Comment'  from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4 as 'Approved by',U_DTApprv4 as 'Approved on',U_CommApprv4 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7 as 'Approved by',U_DTApprv7 as 'Approved on',U_CommApprv7 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8 as 'Approved by',U_DTApprv8 as 'Approved on',U_CommApprv8 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9 as 'Approved by',U_DTApprv9 as 'Approved on',U_CommApprv9 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10 as 'Approved by',U_DTApprv10 as 'Approved on',U_CommApprv10 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept',dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11 as 'Approved by',U_DTApprv11 as 'Approved on',U_CommApprv11 as 'Comment' from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='CD')
			Select N'Dư án XD - Chỉ huy trưởng ME' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='CDXD')
			Select N'Dư án XD - Chỉ huy trưởng ME' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Dư án XD - Chỉ huy trưởng XD' as 'Dept',dbo.BL_Status_Name(U_Apprv2) as 'Status',U_usrApprv2 as 'Approved by',U_DTApprv2 as 'Approved on' ,U_CommApprv2 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='TB')
			Select N'Thiết bị - Trưởng phòng TB' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='TBXD')
			Select N'Dư án XD - Chỉ huy trưởng XD' as 'Dept',dbo.BL_Status_Name(U_Apprv2) as 'Status',U_usrApprv2 as 'Approved by',U_DTApprv2 as 'Approved on' ,U_CommApprv2 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Thiết bị - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			Union all
			Select N'Thiết bị - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='VP')
		begin
			if (@Dept_Create = 18 or @Dept_Create = 15 or @Dept_Create = 11 or @Dept_Create = 7 )
			begin
				Select N'Phòng Ban - Trưởng Bộ Phận' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
				Union all
				Select N'Trưởng phòng Tổng Hợp' as 'Dept',dbo.BL_Status_Name(U_Apprv12) as 'Status',U_usrApprv12 as 'Approved by',U_DTApprv12 as 'Approved on' ,U_CommApprv12 as 'Comment' from OOAT where Number= @BlanketNo
				Union all
				Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
				Union all
				Select N'Hợp đồng - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
				Union all
				Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
				Union all
				Select N'Kế toán - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
				Union all
				Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
				Union all
				Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
				Union all
				Select N'BGĐ - Phó tổng GĐ' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
			end
			else
			begin
				Select N'Phòng Ban - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
				Union all
				Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
				Union all
				Select N'Hợp đồng - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
				Union all
				Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
				Union all
				Select N'Kế toán - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
				Union all
				Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
				Union all
				Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
				Union all
				Select N'BGĐ - Phó tổng GĐ' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;				
			end 
		end
	end
	else if (@Dept_Create = 1)  --CCM Created
	begin
		if (@BlanketType ='XD')
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1,U_DTApprv1,U_CommApprv1 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status' ,U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='CD')
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1,U_DTApprv1,U_CommApprv1 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='CDXD')
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1,U_DTApprv1,U_CommApprv1 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			Union all
			Select N'ME - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;

		if (@BlanketType ='TB')
			Select N'Thiết bị - Trưởng phòng TB' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;

		if (@BlanketType ='TBXD')
			Select N'Thiết bị - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1,U_DTApprv1,U_CommApprv1 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			--Select N'Thiết bị - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			--Union all
			--Select N'Thiết bị - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			--Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept',dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
		
		if (@BlanketType ='VP')
			Select N'Phòng Ban - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept',dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Phó tổng GĐ' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
	end
	else if (@Dept_Create =2) --TB Created
	begin
		if (@BlanketType ='TB')
			Select N'Thiết bị - Trưởng phòng TB' as 'Dept',dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1 as 'Approved by',U_DTApprv1 as 'Approved on' ,U_CommApprv1 as 'Comment' from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept', dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;

		if (@BlanketType ='TBXD')
			Select N'Thiết bị - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv1) as 'Status',U_usrApprv1,U_DTApprv1,U_CommApprv1 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv3) as 'Status',U_usrApprv3,U_DTApprv3,U_CommApprv3 from OOAT where Number= @BlanketNo
			Union all
			Select N'Hợp đồng - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv4) as 'Status',U_usrApprv4,U_DTApprv4,U_CommApprv4 from OOAT where Number= @BlanketNo
			Union all
			--Select N'Thiết bị - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv5) as 'Status',U_usrApprv5,U_DTApprv5,U_CommApprv5 from OOAT where Number= @BlanketNo
			--Union all
			--Select N'Thiết bị - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv6) as 'Status',U_usrApprv6,U_DTApprv6,U_CommApprv6 from OOAT where Number= @BlanketNo
			--Union all
			Select N'Kế toán - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv7) as 'Status',U_usrApprv7,U_DTApprv7,U_CommApprv7 from OOAT where Number= @BlanketNo
			Union all
			Select N'Kế toán - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv8) as 'Status',U_usrApprv8,U_DTApprv8,U_CommApprv8 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Nhân viên' as 'Dept',dbo.BL_Status_Name(U_Apprv9) as 'Status',U_usrApprv9,U_DTApprv9,U_CommApprv9 from OOAT where Number= @BlanketNo
			Union all
			Select N'CCM - Trưởng phòng' as 'Dept', dbo.BL_Status_Name(U_Apprv10) as 'Status',U_usrApprv10,U_DTApprv10,U_CommApprv10 from OOAT where Number= @BlanketNo
			Union all
			Select N'BGĐ - Giám đốc dự án' as 'Dept',dbo.BL_Status_Name(U_Apprv11) as 'Status',U_usrApprv11,U_DTApprv11,U_CommApprv11 from OOAT where Number= @BlanketNo;
	end
END
RETURN

GO

ALTER PROCEDURE [dbo].[BL_Check_Current_Level]
	@BlanketNo as int
AS
SET NOCOUNT OFF
DECLARE @BlanketType as varchar(100)
DECLARE @Level1 as varchar(10)
DECLARE @Level2 as varchar(10)
DECLARE @Level3 as varchar(10)
DECLARE @Level4 as varchar(10)
DECLARE @Level5 as varchar(10)
DECLARE @Level6 as varchar(10)
DECLARE @Level7 as varchar(10)
DECLARE @Level8 as varchar(10)
DECLARE @Level9 as varchar(10)
DECLARE @Level10 as varchar(10)
DECLARE @Level12 as varchar(10)
DECLARE @Dept_Create as int
DECLARE @Result as int
BEGIN
	Select  @BlanketType = a.U_CGroup
	,@Level1 = ISNULL(a.U_Apprv1,'')
	,@Level2 = ISNULL(a.U_Apprv2,'')
	,@Level3 = ISNULL(a.U_Apprv3,'')
	,@Level4 = ISNULL(a.U_Apprv4,'')
	,@Level5 = ISNULL(a.U_Apprv5,'')
	,@Level6 = ISNULL(a.U_Apprv6,'')
	,@Level7 = ISNULL(a.U_Apprv7,'')
	,@Level8 = ISNULL(a.U_Apprv8,'')
	,@Level9 = ISNULL(a.U_Apprv9,'')
	,@Level10 = ISNULL(a.U_Apprv10,'')
	,@Level12 = ISNULL(a.U_Apprv12,'')
	,@Dept_Create = b.dept
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Number = @BlanketNo;
	if (@BlanketType = 'XD')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			if (@Dept_Create = 1) -- CCM Create
			begin
				Set @Result =1;
				if (@Level1 <> '' and @Result =1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 2;
				if (@Level2 <> '' and @Result = 2)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
		end
		else --CMM Comment with Note
		begin
			if (@Dept_Create = 1) --CCM Create
			begin
				Set @Result =1;
				if (@Level1 <> '' and @Level1 <> '2' and @Result =1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
						Set @Result = 5;
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 2;
				if (@Level2 <> '2' and @Result = 2)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
	if (@BlanketType = 'CD')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			if (@Dept_Create = 1) -- CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result =1)
					begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
						begin
							Set @Result = 4;
							if (@Level10 = '1' and @Result = 4)
								Set @Result = 5;
						end
					end
			end
			else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result = 1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
		end
		else --CMM Comment with Note
		begin
			if (@Dept_Create = 1) --CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Level1 <> '2' and @Result = 1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
						Set @Result = 5;
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '2' and @Result = 1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
	if (@BlanketType = 'CDXD')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			if (@Dept_Create = 1) -- CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result =1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Level1 <> '2' and @Result = 1)
				begin
					Set @Result = 2;
					if (@Level2 <> '' and @Result = 2)
						begin
							Set @Result = 3;
							if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
							begin
								Set @Result = 4;
								if (@Level10 = '1' and @Result = 4)
									Set @Result = 5;
							end
						end
					end
			end
		end
		else --CMM Comment with Note
		begin
			if (@Dept_Create = 1) --CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Level1 <> '2' and @Result =1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
						Set @Result = 5;
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 2;
				if (@Level2 <> '2' and @Result = 2)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
	if (@BlanketType = 'TB')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			if (@Dept_Create in (1,2)) -- CCM Create or TB Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result = 1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result = 1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
		end
		else --CMM Comment with Note
		begin
			--if (@Dept_Create = 1) --CCM Create
			--begin
			--	Set @Result = 4;
            --    if (@Level10 = '1' and @Result = 4)
			--		Set @Result = 5;
			--end
			--else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '2' and @Result = 1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
	if (@BlanketType = 'TBXD')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			if (@Dept_Create in (1,2)) -- CCM Create or TB Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Result =1)
				begin
					Set @Result = 3;
					if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
					begin
						Set @Result = 4;
						if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
					end
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 2;
				if (@Level2 <> '' and @Level2 <> '2' and @Result = 2)
					begin
					Set @Result = 3;
						if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
						begin
							Set @Result = 4;
							if (@Level10 = '1' and @Result = 4)
								Set @Result = 5;
						end
					end
			end
		end
		else --CMM Comment with Note
		begin
			if (@Dept_Create = 1) --CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Level1 <> '2' and @Result =1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
						Set @Result = 5;
				end
			end
			else --Not CCM Create
			begin
				Set @Result = 2;
				if (@Level2 <> '2' and @Result = 2)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
	if (@BlanketType = 'VP')
	begin
		if (@Level10 <> '2') --Not CCM Comment with Note
		begin
			--if (@Dept_Create = 1) -- CCM Create
			--begin
			--	Set @Result = 3;
			--	if (@Level4 <> '' and @Level6 <>'' and @Level8 <>'' and @Result = 3)
			--	begin
			--		Set @Result = 4;
			--		if (@Level10 = '1' and @Result = 4)
			--			Set @Result = 5;
			--	end
			--end
			--else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '' and @Level12 <> '' and @Result = 1)
					begin
					Set @Result = 3;
						if (@Level4 <> '' and @Level8 <>'' and @Result = 3)
						begin
							Set @Result = 4;
							if (@Level10 = '1' and @Result = 4)
								Set @Result = 5;
						end
					end
			end
		end
		else --CMM Comment with Note
		begin
			--if (@Dept_Create = 1) --CCM Create
			--begin
			--	Set @Result = 4;
            --    if (@Level10 = '1' and @Result = 4)
			--		Set @Result = 5;
			--end
			--else --Not CCM Create
			begin
				Set @Result = 1;
				if (@Level1 <> '2' and @Result = 1)
				begin
					Set @Result = 4;
					if (@Level10 = '1' and @Result = 4)
							Set @Result = 5;
				end
			end
		end
	end
END
RETURN @Result;
GO

ALTER PROCEDURE [dbo].[BL_Check_User_Level]
	@UserName as varchar(100)
	,@BlanketNo as int
AS
SET NOCOUNT OFF
DECLARE @Dept as int
DECLARE @BlanketType as varchar(50)
DECLARE @Dept_Create as int
DECLARE @Position as int
DECLARE @Result as int
DECLARE @Apprlv1 as varchar(10)
BEGIN
	Set @Result = -9;
	--Dept
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS
	
	--Position
	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME
	Select  @BlanketType = a.U_CGroup
	,@Dept_Create = b.dept
	,@Apprlv1 = a.U_Apprv1
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Number = @BlanketNo;

	SELECT @Position = b.position
	,@Dept = b.dept 
	FROM OUSR a LEFT JOIN OHEM b ON a.USERID = b.userId 
	WHERE a.USER_CODE =@UserName;

	if (@BlanketType = 'CD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 6) Set @Result = 1;
			if (@Dept = 4) Set @Result = 3;
			if (@Dept = 5) Set @Result = 3;
			if (@Dept = -2) Set @Result = 3;
			if (@Dept = 1) Set @Result = 4;
			if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = 2)) Set @Result = 1;
			else if (@Dept = 4) Set @Result = 3;
			else if (@Dept = 5) Set @Result = 3;
			else if (@Dept = -2) Set @Result = 3;
			else if (@Dept = 1) Set @Result = 4;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
	end
	if (@BlanketType = 'XD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 5) Set @Result = 2;
			if (@Dept = 4) Set @Result = 3;
			if (@Dept = -2) Set @Result = 3;
			if (@Dept = 1) Set @Result = 4;
			if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = 2)) Set @Result = 1;
			else if (@Dept = 4) Set @Result = 3;
			else if (@Dept = -2) Set @Result = 3;
			else if (@Dept = 1) Set @Result = 4;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
	end
	if (@BlanketType = 'CDXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 6) Set @Result = 1;
			if (@Dept = 3 and @Position = 5) Set @Result = 2;
			if (@Dept = 4) Set @Result = 3;
			if (@Dept = 5) Set @Result = 3;
			if (@Dept = -2) Set @Result = 3;
			if (@Dept = 1) Set @Result = 4;
			if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = 2)) Set @Result = 1;
			else if (@Dept = 4) Set @Result = 3;
			else if (@Dept = 5) Set @Result = 3;
			else if (@Dept = -2) Set @Result = 3;
			else if (@Dept = 1) Set @Result = 4;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
	end
	if (@BlanketType = 'TBXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 5) Set @Result = 2;
			if (@Dept = 4) Set @Result = 3;
			if (@Dept = 2) Set @Result = 3;
			if (@Dept = -2) Set @Result = 3;
			if (@Dept = 1) Set @Result = 4;
			if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
		else --CCM Create
		begin
			if (@Dept = 2 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = 2)) Set @Result = 1;
			else if (@Dept = 4) Set @Result = 3;
			--else if (@Dept = 2) Set @Result = 3;
			else if (@Dept = -2) Set @Result = 3;
			else if (@Dept = 1) Set @Result = 4;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
		end
	end
	if (@BlanketType = 'TB')
	begin
		--if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 2 and @Position = 1) Set @Result = 1;
			if (@Dept = 4) Set @Result = 3;
			if (@Dept = -2) Set @Result = 3;
			if (@Dept = 1) Set @Result = 4;
			if (@Dept = 6 and @Position = 3) Set @Result = 5;
		end
		--else --CCM Create
		--begin
		--	if (@Dept = 4) Set @Result = 3;
		--	if (@Dept = 5) Set @Result = 3;
		--	if (@Dept = -2) Set @Result = 3;
		--	if (@Dept = 1) Set @Result = 4;
		--	if (@Dept = 6 and @Position = 3) Set @Result = 5;
		--end
	end
	if (@BlanketType = 'VP')
	begin
		if(@Dept_Create = @Dept and @Position = 1 and @Apprlv1 is null) Set @Result = 1;
		else if (@Dept = 23) Set @Result = 1;
		else if (@Dept = 4) Set @Result = 3;
		else if (@Dept = -2) Set @Result = 3;
		else if (@Dept = 1) Set @Result = 4;
		else if (@Dept = 6 and (@Position = 4 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 5;
	end

END
RETURN @Result;

Go

ALTER PROCEDURE [dbo].[BL_Get_Level_1]
	@BlanketNo as int
AS
SET NOCOUNT OFF
DECLARE @BlanketType as varchar(50)
DECLARE @Dept_Create as int
DECLARE @Result as int
BEGIN
	Set @Result = -9;
	--Dept
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS
	
	--Position
	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME
	Select  @BlanketType = a.U_CGroup
	,@Dept_Create = b.dept
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Number = @BlanketNo;

	if (@BlanketType = 'CD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
			Set @Result = 1;
		else --CCM Create
			Set @Result = 1;
	end
	if (@BlanketType = 'XD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
			Set @Result = 2;
		else --CCM Create
			Set @Result = 1;
	end
	if (@BlanketType = 'CDXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
			Set @Result = 1;
		else --CCM Create
			Set @Result = 1;
	end
	if (@BlanketType = 'TBXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
			Set @Result = 2;
		else --CCM Create
			Set @Result = 1;
	end
	if (@BlanketType = 'TB')
		Set @Result = 1;
	if (@BlanketType = 'VP')
		Set @Result = 1;
END
RETURN @Result;

GO

ALTER FUNCTION [dbo].[BL_Status_Name]
(@Status  AS int)
RETURNS varchar(100) 
    BEGIN   
	DECLARE @Result as nvarchar(100);
		if (@Status = 0) set @Result = 'Rejected';
		else if (@Status = 1) set @Result = 'Approved' ;
		else if (@Status = 2) set @Result = 'Approved with Note' ;
		else set @Result = @Status;
		RETURN @Result
	END  
GO

ALTER PROCEDURE [dbo].[BL_Get_User_Posting_Level]
	@UserName as varchar(100)
	,@BlanketNo as int
AS
SET NOCOUNT OFF
DECLARE @Dept as int
DECLARE @BlanketType as varchar(50)
DECLARE @Dept_Create as int
DECLARE @Position as int
DECLARE @Result as int
DECLARE @Apprlv1 as varchar(10)
BEGIN
	Set @Result = -9;
	--Dept
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS
	
	--Position
	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME
	Select  @BlanketType = a.U_CGroup
	,@Dept_Create = b.dept
	,@Apprlv1 = a.U_Apprv1
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Number = @BlanketNo;

	SELECT @Position = b.position
	,@Dept = b.dept 
	FROM OUSR a LEFT JOIN OHEM b ON a.USERID = b.userId 
	WHERE a.USER_CODE =@UserName;

	if (@BlanketType = 'CD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 6) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 5 and @Position = 2) Set @Result = 5;
			else if (@Dept = 5 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen' )) Set @Result = 11;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = '2')) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 5 and @Position = 2) Set @Result = 5;
			else if (@Dept = 5 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
	end
	if (@BlanketType = 'XD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 5) Set @Result = 2;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = '2')) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
	end
	if (@BlanketType = 'CDXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 6) Set @Result = 1;
			else if (@Dept = 3 and @Position = 5) Set @Result = 2;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 5 and @Position = 2) Set @Result = 5;
			else if (@Dept = 5 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
		else --CCM Create
		begin
			if (@Dept = 1 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = '2')) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 5 and @Position = 2) Set @Result = 5;
			else if (@Dept = 5 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
	end
	if (@BlanketType = 'TBXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 3 and @Position = 5) Set @Result = 2;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 2 and @Position = 2) Set @Result = 5;
			else if (@Dept = 2 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
		else --CCM Create
		begin
			if (@Dept = 2 and @Position = 1 and (@Apprlv1 is null or @Apprlv1 = '2')) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = 2 and @Position = 2) Set @Result = 5;
			else if (@Dept = 2 and @Position = 1) Set @Result = 6;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and (@Position = 3 or @UserName='lan.nguyen' or @UserName='thuy.nguyen')) Set @Result = 11;
		end
	end
	if (@BlanketType = 'TB')
	begin
		--if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@Dept = 2 and @Position = 1) Set @Result = 1;
			else if (@Dept = 4 and @Position = 2) Set @Result = 3;
			else if (@Dept = 4 and @Position = 1) Set @Result = 4;
			else if (@Dept = -2 and @Position = 2) Set @Result = 7;
			else if (@Dept = -2 and @Position = 1) Set @Result = 8;
			else if (@Dept = 1 and @Position = 2) Set @Result = 9;
			else if (@Dept = 1 and @Position = 1) Set @Result = 10;
			else if (@Dept = 6 and @Position = 3) Set @Result = 11;
		end
		--else --CCM Create
		--begin
		--	if (@Dept = 4) Set @Result = 3;
		--	if (@Dept = 5) Set @Result = 3;
		--	if (@Dept = -2) Set @Result = 3;
		--	if (@Dept = 1) Set @Result = 4;
		--	if (@Dept = 6 and @Position = 3) Set @Result = 5;
		--end
	end
	if (@BlanketType = 'VP')
	begin
		if(@Dept_Create = @Dept and @Position = 1 and @Apprlv1 is null) Set @Result = 1;
		else if (@Dept = 23 and @Position = 1) Set @Result = 12;
		else if (@Dept = 4 and @Position = 2) Set @Result = 3;
		else if (@Dept = 4 and @Position = 1) Set @Result = 4;
		else if (@Dept = -2 and @Position = 2) Set @Result = 7;
		else if (@Dept = -2 and @Position = 1) Set @Result = 8;
		else if (@Dept = 1 and @Position = 2) Set @Result = 9;
		else if (@Dept = 1 and @Position = 1) Set @Result = 10;
		else if (@Dept = 6 and (@Position = 4 or @UserName='lan.nguyen')) Set @Result = 11;
	end

END
RETURN @Result;
GO

CREATE PROCEDURE [dbo].[BL_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@BlanketNo as int,
	@LVL_Posting as int
AS
BEGIN
--Get User Info - Dept - Position
Declare @FProject as varchar(250)
Declare @BlanketType as varchar(10)
Declare @Dept_Create as int
DECLARE @Table_Dept TABLE(
		DeptId int NOT NULL,
		PosId int
	);

--Lay thong tin Hop dong
Select  @BlanketType = a.U_CGroup
	,@Dept_Create = b.dept
	,@FProject = U_PRJ
	from OOAT a left join OHEM b on a.UserSign = b.userId
	where a.Number = @BlanketNo;

	--Dept
	-- -2	Kế Toán	
	--  1	CCM
	--  2	Thiết Bị
	--  3	Dự Án XD
	--  4	Pháp chế
	--  5	Cơ điện
	--  6	BGĐ
	--  7	HCNS
	
	--Position
	--1	Trưởng phòng
	--2	Nhân viên
	--3	Giám đốc dự án
	--4	Phó tổng giám đốc
	--5	Chỉ huy trưởng DA
	--6	Chỉ huy trưởng ME

if (@BlanketType = 'CD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 3 as DeptID,6 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 5 as DeptID,1 as PosID
				union
				Select 5 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
		else --CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 5 as DeptID,1 as PosID
				union
				Select 5 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
	end
if (@BlanketType = 'XD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@LVL_Posting = 2)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 3 as DeptID, 5 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
		else --CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID, 1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
	end
if (@BlanketType = 'CDXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 3 as DeptID, 6 as PosID;
			else if (@LVL_Posting = 2)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 3 as DeptID, 5 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 5 as DeptID,1 as PosID
				union
				Select 5 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
		else --CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID, 1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 5 as DeptID,1 as PosID
				union
				Select 5 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
	end
if (@BlanketType = 'TBXD')
	begin
		if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@LVL_Posting = 2)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 3 as DeptID, 5 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 2 as DeptID,1 as PosID
				union
				Select 2 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
		else --CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 2 as DeptID, 1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 2 as DeptID,1 as PosID
				union
				Select 2 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
	end
if (@BlanketType = 'TB')
	begin
		--if (@Dept_Create <> 1) --NOT CCM Create
		begin
			if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 2 as DeptID, 1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select 2 as DeptID,1 as PosID
				union
				Select 2 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,3 as PosID;
		end
		--else --CCM Create
		--begin
		--	if (@Dept = 4) Set @Result = 3;
		--	if (@Dept = 5) Set @Result = 3;
		--	if (@Dept = -2) Set @Result = 3;
		--	if (@Dept = 1) Set @Result = 4;
		--	if (@Dept = 6 and @Position = 3) Set @Result = 5;
		--end
	end
if (@BlanketType = 'VP')
	begin
		if (@LVL_Posting = 1)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select @Dept_Create as DeptID, 1 as PosID;
			else if (@LVL_Posting = 3)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 4 as DeptID,1 as PosID
				union
				Select 4 as DeptID,2 as PosID
				union
				Select -2 as DeptID,1 as PosID
				union
				Select -2 as DeptID,2 as PosID;
			else if (@LVL_Posting = 4)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 1 as DeptID,1 as PosID
				union
				Select 1 as DeptID,2 as PosID;
			else if (@LVL_Posting = 5)
				INSERT INTO @Table_Dept (DeptID,PosId)
				Select 6 as DeptID,4 as PosID;
	end

--Select @FProject=ISNULL(U_FIPROJECT,'') from [@KLTT] where DocEntry=@DocEntry;

Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name
from OHEM a inner join OUSR b on a.USERID = b.UserID
left join HTM1 c on c.empID=a.empID
inner join OHTM d on c.teamID = d.teamID
where a.dept in (Select distinct DeptId from @Table_Dept)
and a.position in (Select distinct PosId from @Table_Dept)
and d.name = @FProject;
END
GO
