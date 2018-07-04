--Purchase Request
Select * from OPRQ;
--Purchase Request Details
Select * from PRQ1;
--Approve Template
Select * from OWTM;
--User Approve
Select * from WTM1;
--Stage Approve
Select * from WTM2;
--Approve Document
Select * from WTM3;
Select * from WTM4;

Select * from OWST;
Select * from WST1;
--Draft Document
Select * From ODRF where ObjType=1470000113;

Select * from OWDD where Status = 'W';
Select * from WDD1;

Select T0.DocEntry
	,T0.DocNum
	,T0.CANCELED
	,T0.DocDate
	,T0.DocTotal
	,T0.Project
	,T0.U_GOITHAU
	,T0.U_RECTYPE
	,T0.U_PUTYPE
	--,(Select 
from ODRF T0 left join OWDD T1 on T0.DocEntry = T1.DocEntry
where T0.ObjType=1470000113
and T0.DocEntry = 22

ALTER PROCEDURE [dbo].[PR_Get_List_Approve]
	@Username as varchar(50)
AS
BEGIN
	DECLARE @UserID as int;
	Select @UserID=USERID from OUSR where USER_CODE=@Username;

	Select T0.DocEntry
		, T0.DocNum
		, T1.WddCode
		--,T0.CANCELED
		, CONVERT(varchar(11), T0.DocDate, 113) as 'Create Date'
		, Format(T0.DocTotal ,'N0','en-US' ) as 'DocTotal'
		, ISNULL(T0.Project,'') as 'Project'
		, ISNULL(T0.U_GOITHAU,'') as 'GoiThau'
		, ISNULL(T0.U_RECTYPE,'') as 'RECTYPE'
		, ISNULL(T0.U_PUTYPE,'') as 'PUType'
		--,(Select 
	from ODRF T0 left join OWDD T1 on T0.DocEntry = T1.DocEntry
	where T0.ObjType=1470000113
	and T0.WddStatus = 'W'
	and @UserID in (Select UserID 
					from WDD1 
					where WddCode= T1.WddCode and Status='W'
					and StepCode in (Select StepCode
									 from WDD1 
									 where WddCode=T1.WddCode
									 group by StepCode 
									 having SUM(case [Status] when 'W' then 0 else 1 end) = 0))
	
END
GO

ALTER PROCEDURE [dbo].[PR_Get_List_Rejected]
	@Username as varchar(50)
AS
BEGIN
	DECLARE @UserID as int;
	Select @UserID=USERID from OUSR where USER_CODE=@Username;

	Select T0.DocEntry
		, T0.DocNum
		, T1.WddCode
		--,T0.CANCELED
		, CONVERT(varchar(11), T0.DocDate, 113) as 'Create Date'
		, Format(T0.DocTotal ,'N0','en-US' ) as 'DocTotal'
		, ISNULL(T0.Project,'') as 'Project'
		, ISNULL(T0.U_GOITHAU,'') as 'GoiThau'
		, ISNULL(T0.U_RECTYPE,'') as 'RECTYPE'
		, ISNULL(T0.U_PUTYPE,'') as 'PUType'
		--,(Select 
	from ODRF T0 left join OWDD T1 on T0.DocEntry = T1.DocEntry
	where T0.ObjType=1470000113
	and WddStatus ='N'
	--and @UserID in (Select UserID 
	--				from WDD1 
	--				where WddCode= T1.WddCode and Status='W'
	--				and StepCode in (Select StepCode
	--								 from WDD1 
	--								 where WddCode=T1.WddCode
	--								 group by StepCode 
	--								 having SUM(case [Status] when 'W' then 0 else 1 end) = 0))
	
END
GO

CREATE PROCEDURE [dbo].[PR_Get_List_Approved]
	@Username as varchar(50)
AS
BEGIN
	DECLARE @UserID as int;
	Select @UserID=USERID from OUSR where USER_CODE=@Username;

	Select T0.DocEntry
		, T0.DocNum
		, T1.WddCode
		--,T0.CANCELED
		, CONVERT(varchar(11), T0.DocDate, 113) as 'Create Date'
		, Format(T0.DocTotal ,'N0','en-US' ) as 'DocTotal'
		, ISNULL(T0.Project,'') as 'Project'
		, ISNULL(T0.U_GOITHAU,'') as 'GoiThau'
		, ISNULL(T0.U_RECTYPE,'') as 'RECTYPE'
		, ISNULL(T0.U_PUTYPE,'') as 'PUType'
		--,(Select 
	from ODRF T0 left join OWDD T1 on T0.DocEntry = T1.DocEntry
	where T0.ObjType=1470000113
	and WddStatus ='Y'
	--and @UserID in (Select UserID 
	--				from WDD1 
	--				where WddCode= T1.WddCode and Status='W'
	--				and StepCode in (Select StepCode
	--								 from WDD1 
	--								 where WddCode=T1.WddCode
	--								 group by StepCode 
	--								 having SUM(case [Status] when 'W' then 0 else 1 end) = 0))
	
END
GO

CREATE PROCEDURE [dbo].[PR_Get_List_All]
	@Username as varchar(50)
AS
BEGIN
	DECLARE @UserID as int;
	Select @UserID=USERID from OUSR where USER_CODE=@Username;

	Select T0.DocEntry
		, T0.DocNum
		, T1.WddCode
		--,T0.CANCELED
		, CONVERT(varchar(11), T0.DocDate, 113) as 'Create Date'
		, Format(T0.DocTotal ,'N0','en-US' ) as 'DocTotal'
		, ISNULL(T0.Project,'') as 'Project'
		, ISNULL(T0.U_GOITHAU,'') as 'GoiThau'
		, ISNULL(T0.U_RECTYPE,'') as 'RECTYPE'
		, ISNULL(T0.U_PUTYPE,'') as 'PUType'
		--,(Select 
	from ODRF T0 left join OWDD T1 on T0.DocEntry = T1.DocEntry
	where T0.ObjType=1470000113
	--and @UserID in (Select UserID 
	--				from WDD1 
	--				where WddCode= T1.WddCode and Status='W'
	--				and StepCode in (Select StepCode
	--								 from WDD1 
	--								 where WddCode=T1.WddCode
	--								 group by StepCode 
	--								 having SUM(case [Status] when 'W' then 0 else 1 end) = 0))
	
END
GO

ALTER PROCEDURE [dbo].[PR_Get_Approve_Process]
	@DocEntry_ODRF as int
AS
BEGIN
	DECLARE @WtmCode as int
	DECLARE @WddCode as int

	Select @WtmCode = ISNULL(T1.WtmCode,-1) , @WddCode = ISNULL(T1.WddCode,-1)
	from ODRF T0 inner join OWDD T1 on T0.DocEntry=T1.DocEntry
	where T0.DocEntry = @DocEntry_ODRF;

	Select T0.Remarks as 'Process'
	,(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') as 'NAME' from OHEM
		where userId = T1.UserID) as 'Approved by'
	,CONVERT(varchar(11), T1.UpdateDate, 113) as 'Approved on'
	,case T1.[Status] when 'Y' then 'Approved' when 'N' then 'Rejected' else '' end as 'Status'
	,T1.Remarks as 'Comment'
	from WTM2 T0 left join (Select * from WDD1 where WddCode=@WddCode and [Status] <> 'W')  T1 on  T0.WstCode=T1.StepCode
	where T0.WtmCode= @WtmCode
	order by T0.SortId;
END
GO

ALTER PROCEDURE [dbo].[PR_Get_Document_Details]
	@DocEntry_ODRF as int
AS
BEGIN
	Select LineNum, ItemCode, Dscription
	, Format( Quantity ,'N0','en-US' ) as 'Quantity'
	, Format( Price ,'N0','en-US' ) as 'Price'
	, Format( LineTotal ,'N0','en-US' ) as 'LineTotal' 
	, Format( VatSum ,'N0','en-US' ) as 'LineVAT' 
	from DRF1 where DocEntry = @DocEntry_ODRF;
END
GO

ALTER PROCEDURE [dbo].[PR_GETDATA_COVER]
	@DocNum as int
AS
BEGIN
	Select DocTotal, OPOR.CreateDate, DocEntry, Project, OUSR.Department, OUDP.Remarks from  OPOR
	inner join 
	OUSR on OUSR.USERID = OPOR.UserSign
	inner join
	OUDP on OUDP.Code = OUSR.Department
	where OPOR.DocEntry = @DocNum
END
GO

CREATE PROCEDURE [dbo].[PR_GETDATA_COVER_DETAIL]
	@DocEntry as int
AS
BEGIN
	Select ItemCode, Dscription, UomCode, Quantity, Price, LineTotal from  POR1
	--inner join 
	--OUSR on OUSR.USERID = OPOR.UserSign
	--inner join
	--OUDP on OUDP.Code = OUSR.Department
	where POR1.DocEntry = @DocEntry
END
GO