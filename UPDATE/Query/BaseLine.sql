--Table Project
CREATE TABLE BASELINE_OPMG (
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	[AbsEntry] [int] NOT NULL,
	[OWNER] [int] NULL,
	[NAME] [nvarchar](254) NULL,
	[START] [datetime] NULL,
	[FINISHED] [numeric](19, 6) NULL,
	[DocNum] [int] NULL,
	[Series] [smallint] NULL,
	[TYP] [char](1) NULL,
	[CARDCODE] [nvarchar](15) NULL,
	[CARDNAME] [nvarchar](100) NULL,
	[CONTACT] [int] NULL,
	[TERRITORY] [int] NULL,
	[EMPLOYEE] [int] NULL,
	[WithPhases] [char](1) NULL,
	[STATUS] [char](1) NULL,
	[DUEDATE] [datetime] NULL,
	[CLOSING] [datetime] NULL,
	[FIPROJECT] [nvarchar](20) NULL,
	[RISK] [char](1) NULL,
	[INDUSTRY] [int] NULL,
	[REASON] [ntext] NULL,
	[Free_Text] [ntext] NULL,
	[BPLid] [int] NULL,
	[AtcEntry] [int] NULL,
	[Attachment] [ntext] NULL,
	[LogInstanc] [int] NULL,
	[UpdateDate] [datetime] NULL,
	[UserSign] [smallint] NULL,
	[UserSign2] [smallint] NULL,
	[CreateDate] [datetime] NULL,
	[UpdateTS] [int] NULL,
	[U_BPTH] [nvarchar](10) NULL,
	[U_PRJGROUP] [nvarchar](10) NULL,
	[U_PRJTYPE] [nvarchar](10) NULL,
	[U_CPHT1] [numeric](19, 6) NULL,
	[U_CPHT2] [numeric](19, 6) NULL,
	[U_DPBH] [numeric](19, 6) NULL,
	[U_DPCP] [numeric](19, 6) NULL,
	[U_CPNG] [numeric](19, 6) NULL,
	[U_CPQLCT] [numeric](19, 6) NULL,
	[U_VT] [nvarchar](40) NULL,
	[U_SUMTT] [numeric](19, 6) NULL,
	[U_SUMTTDT] [numeric](19, 6) NULL,
	[U_SUMTTHD] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_OPHA (
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	[AbsEntry] [int] NOT NULL,
	[OWNER] [int] NULL,
	[NAME] [nvarchar](254) NULL,
	[START] [datetime] NULL,
	[FINISHED] [numeric](19, 6) NULL,
	[ParentID] [int] NULL,
	[ProjectID] [int] NULL,
	[Code] [int] NULL,
	[TYP] [int] NULL,
	[CONTRIB] [numeric](19, 6) NULL,
	[STATUS] [char](1) NULL,
	[END] [datetime] NULL,
	[COST] [numeric](19, 6) NULL,
	[PLANNED] [numeric](19, 6) NULL,
	[Level] [int] NULL,
	[DUEDATE] [datetime] NULL,
	[LogInstanc] [int] NULL,
	[UpdateDate] [datetime] NULL,
	[UserSign] [smallint] NULL,
	[UserSign2] [smallint] NULL,
	[CreateDate] [datetime] NULL,
	[UpdateTS] [int] NULL,
	[U_001] [nvarchar](50) NULL,
	[U_002] [nvarchar](10) NULL,
	[U_KLDT] [numeric](19, 6) NULL,
	[U_DG] [numeric](19, 6) NULL,
	[U_TTBV] [numeric](19, 6) NULL,
	[U_TTDT] [numeric](19, 6) NULL,
	[U_003] [numeric](19, 6) NULL,
	[U_REMARK] [ntext] NULL,
	[U_TTHD] [numeric](19, 6) NULL,
	[U_DGHD] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_PHA1(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	[AbsEntry] [int] NOT NULL,
	[LineID] [int] NOT NULL,
	[StageID] [int] NULL,
	[POS] [int] NULL,
	[START] [datetime] NULL,
	[CLOSE] [datetime] NULL,
	[Task] [int] NULL,
	[DSCRIPTION] [ntext] NULL,
	[EXPCOSTS] [numeric](19, 6) NULL,
	[InvAmtAR] [numeric](19, 6) NULL,
	[OpenAmtAR] [numeric](19, 6) NULL,
	[InvAmtAP] [numeric](19, 6) NULL,
	[OpenAmtAP] [numeric](19, 6) NULL,
	[PERCENT] [numeric](19, 6) NULL,
	[FINISH] [char](1) NULL,
	[OWNER] [int] NULL,
	[StageDep1] [int] NULL,
	[StageDep2] [int] NULL,
	[StageDep3] [int] NULL,
	[StageDep4] [int] NULL,
	[StDp1Type] [char](1) NULL,
	[StDp2Type] [char](1) NULL,
	[StDp3Type] [char](1) NULL,
	[StDp4Type] [char](1) NULL,
	[StDp1Abs] [int] NULL,
	[StDp2Abs] [int] NULL,
	[StDp3Abs] [int] NULL,
	[StDp4Abs] [int] NULL,
	[LogInstanc] [int] NULL,
	[AtcEntry] [int] NULL);
CREATE TABLE BASELINE_PHA2(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	[AbsEntry] [int] NOT NULL,
	[LineID] [int] NOT NULL,
	[StageID] [int] NULL,
	[AREA] [int] NULL,
	[PRIORITY] [int] NULL,
	[REMARKS] [ntext] NULL,
	[CLOSED] [char](1) NULL,
	[SOLUTIONID] [int] NULL,
	[SOLUTION] [nvarchar](254) NULL,
	[RESPNSIBLE] [int] NULL,
	[ENTERED] [int] NULL,
	[DATE] [datetime] NULL,
	[EFFORT] [numeric](19, 6) NULL,
	[LogInstanc] [int] NULL,
	[U_NCCPS] [nvarchar](15) NULL,
	[U_TENNCCPS] [nvarchar](100) NULL,
	[U_DVTPS] [nvarchar](10) NULL,
	[U_KLPS] [numeric](19, 6) NULL,
	[U_DGPS] [numeric](19, 6) NULL,
	[U_TTPS] [numeric](19, 6) NULL,
	[U_Issuetype] [nvarchar](10) NULL);

--Table CTG
Create table BASELINE_CTG
(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry int NOT NULL,
	U_PrjCode  nvarchar(20),
	U_PrjName nvarchar(250),
	U_Date datetime,
	U_GoiThauKey nvarchar(50),
	U_GoiThauName nvarchar(200)
);
Create table BASELINE_CTG1
(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_CTG int NOT NULL,
	LineID int NOT NULL,
	[U_001] nvarchar(30),
    [U_ITEMNO] nvarchar(20),
    [U_ITEMNAME] nvarchar(200),
    [U_DVT] nvarchar(20),
    [U_DGDAUTHAU] [numeric](19, 6),
    [U_DGDUPHONG] [numeric](19, 6),
    [U_DinhMuc] [numeric](19, 6),
    [U_HAOHUT] [numeric](19, 6),
    [U_TTDAUTHAU] [numeric](19, 6)
);
CREATE TABLE BASELINE_CTG2(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_CTG int NOT NULL,
	[LineId] [int] NOT NULL,
	[U_001] [nvarchar](30) NULL,
	[U_MATHIETBI] [nvarchar](20) NULL,
	[U_SLDUTRU] [numeric](19, 6) NULL,
	[U_DVTTB] [nvarchar](10) NULL,
	[U_DGMUABAN] [numeric](19, 6) NULL,
	[U_DGVCTB] [numeric](19, 6) NULL,
	[U_DGVH] [numeric](19, 6) NULL,
	[U_GTMB] [numeric](19, 6) NULL,
	[U_GTTHUE] [numeric](19, 6) NULL,
	[U_GTVANCHUYEN] [numeric](19, 6) NULL,
	[U_GTVANHANH] [numeric](19, 6) NULL,
	[U_NGAYCAP] [datetime] NULL,
	[U_NGAYTRA] [datetime] NULL,
	[U_SLTHUE] [numeric](19, 6) NULL,
	[U_SLVANCHUYEN] [numeric](19, 6) NULL,
	[U_SLVANHANH] [numeric](19, 6) NULL,
	[U_TENTHIETBI] [nvarchar](250) NULL,
	[U_TENHM] [nvarchar](250) NULL,
	[U_DGTHUE] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_CTG3(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_CTG int NOT NULL,
	[LineId] [int] NOT NULL,
	[U_001] [nvarchar](30) NULL,
	[U_LOAICHIPHI] [nvarchar](50) NULL,
	[U_DGNCC] [numeric](19, 6) NULL,
	[U_DGNTP] [numeric](19, 6) NULL,
	[U_DGVTP] [numeric](19, 6) NULL,
	[U_DGVC] [numeric](19, 6) NULL,
	[U_DGCN] [numeric](19, 6) NULL,
	[U_DGDTC] [numeric](19, 6) NULL,
	[U_DGDP] [numeric](19, 6) NULL,
	[U_DGDP2] [numeric](19, 6) NULL,
	[U_DGPRELIM] [numeric](19, 6) NULL,
	[U_DGTB] [numeric](19, 6) NULL,
	[U_DGK] [numeric](19, 6) NULL,
	[U_TENHM] [nvarchar](250) NULL);
CREATE TABLE BASELINE_CTG4(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_CTG int NOT NULL,
	[LogInst] [int] NULL,
	[U_001] [nvarchar](30) NULL,
	[U_TKKT] [nvarchar](15) NULL,
	[U_TTKKT] [nvarchar](50) NULL,
	[U_GTDP] [numeric](19, 6) NULL);

--Table DUTRU
Create table BASELINE_DUTRU
(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry int NOT NULL,
	CTG_Key  int NOT NULL,
	DUTRU_TYPE nvarchar(10),
	FProject nvarchar(250),
	ProjectID int
);
Create table BASELINE_DUTRUA
(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_DUTRU  int NOT NULL,
	LineID int NOT NULL,
	[U_SubProjectCode] [nvarchar](100) NOT NULL,
	[U_SubProjectDesc] [nvarchar](250) NULL,
	[U_CP_NCC] [numeric](19, 6) NULL,
	[U_CP_NTP] [numeric](19, 6) NULL,
	[U_CP_DTC] [numeric](19, 6) NULL,
	[U_CP_VTP] [numeric](19, 6) NULL,
	[U_CP_MB] [numeric](19, 6) NULL,
	[U_CP_T] [numeric](19, 6) NULL,
	[U_CP_VH] [numeric](19, 6) NULL,
	[U_CP_VC] [numeric](19, 6) NULL,
	[U_CP_CN] [numeric](19, 6) NULL,
	[U_CP_DP] [numeric](19, 6) NULL,
	[U_CP_DP2] [numeric](19, 6) NULL,
	[U_CP_Prelims] [numeric](19, 6) NULL,
	[U_CP_TB] [numeric](19, 6) NULL,
	[U_CP_K] [numeric](19, 6) NULL,
	[U_SplitTo] [int] NOT NULL
);
Create table BASELINE_DUTRUB
(
	ID int IDENTITY(1,1) PRIMARY KEY,
	DocEntry_BaseLine int NOT NULL,
	DocEntry_DUTRU  int NOT NULL,
	LineID int NOT NULL,
	[U_DTT_LineID] [int] NOT NULL,
	[U_CP_NCC] [numeric](19, 6) NULL,
	[U_BPCode] [nvarchar](50) NOT NULL,
	[U_BPName] [nvarchar](250) NULL,
	[U_CP_NTP] [numeric](19, 6) NULL,
	[U_CP_DTC] [numeric](19, 6) NULL,
	[U_CP_VTP] [numeric](19, 6) NULL,
	[U_CP_VC] [numeric](19, 6) NULL,
	[U_CP_MB] [numeric](19, 6) NULL,
	[U_CP_T] [numeric](19, 6) NULL,
	[U_CP_VH] [numeric](19, 6) NULL,
	[U_CP_CN] [numeric](19, 6) NULL,
	[U_CP_DP] [numeric](19, 6) NULL,
	[U_CP_DP2] [numeric](19, 6) NULL,
	[U_CP_Prelims] [numeric](19, 6) NULL,
	[U_SubProjectCode] [nvarchar](100) NOT NULL,
	[U_SubProjectDesc] [nvarchar](250) NULL,
	[U_CP_TB] [numeric](19, 6) NULL,
	[U_CP_K] [numeric](19, 6) NULL,
	[U_TYPE] [nvarchar](2) NULL,
	[U_TGDK] [datetime] NULL,
	[U_NCTN] [nvarchar](10) NULL,
	[U_PVCV] [nvarchar](200) NULL
);

--Table KLTT
CREATE TABLE BASELINE_KLTT(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[Canceled] [char](1) NULL,
	[UserSign] [int] NULL,
	[Status] [char](1) NULL,
	[CreateDate] [datetime] NULL,
	[CreateTime] [smallint] NULL,
	[UpdateDate] [datetime] NULL,
	[UpdateTime] [smallint] NULL,
	[Creator] [nvarchar](8) NULL,
	[U_FIPROJECT] [nvarchar](50) NOT NULL,
	[U_DATEFROM] [datetime] NULL,
	[U_DATETO] [datetime] NULL,
	[U_BPName] [nvarchar](100) NULL,
	[U_BPCode] [nvarchar](15) NOT NULL,
	[U_Period] [int] NOT NULL,
	[U_CreatedDate] [datetime] NULL,
	[U_VAT] [numeric](19, 6) NOT NULL,
	[U_GTTU] [numeric](19, 6) NULL,
	[U_BGroup] [nvarchar](10) NOT NULL,
	[U_BType] [int] NOT NULL,
	[U_HTTU] [numeric](19, 6) NULL,
	[U_PUType] [nvarchar](10) NULL,
	[U_BPCode2] [nvarchar](100) NULL,
	[U_PTQuanLy] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_KLTTA(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_SubProjectKey] [int] NULL,
	[U_SubProjectName] [nvarchar](254) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_GoiThauKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_GPKey] [int] NULL,
	[U_GPDetailsKey] [int] NULL,
	[U_GPDetailsName] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CTCV] [nvarchar](50) NULL,
	[U_Sub1] [nvarchar](250) NULL,
	[U_Sub2] [nvarchar](250) NULL,
	[U_Sub3] [nvarchar](250) NULL,
	[U_Sub4] [nvarchar](250) NULL,
	[U_Sub5] [nvarchar](250) NULL,
	[U_Sub1Name] [nvarchar](250) NULL,
	[U_Sub2Name] [nvarchar](250) NULL,
	[U_Sub3Name] [nvarchar](250) NULL,
	[U_Sub4Name] [nvarchar](250) NULL,
	[U_Sub5Name] [nvarchar](250) NULL,
	[U_Type] [nvarchar](50) NULL);
CREATE TABLE BASELINE_KLTTB(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_SubProjectKey] [int] NULL,
	[U_SubProjectName] [nvarchar](254) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL,
	[U_GoiThauKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_StageKey] [int] NULL,
	[U_OpenIssueKey] [int] NULL,
	[U_OpenIssueRemark] [nvarchar](254) NULL,
	[U_Details] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_Sub1] [nvarchar](250) NULL,
	[U_Sub2] [nvarchar](250) NULL,
	[U_Sub3] [nvarchar](250) NULL,
	[U_Sub4] [nvarchar](250) NULL,
	[U_Sub5] [nvarchar](250) NULL,
	[U_Sub1Name] [nvarchar](250) NULL,
	[U_Sub2Name] [nvarchar](250) NULL,
	[U_Sub3Name] [nvarchar](250) NULL,
	[U_Sub4Name] [nvarchar](250) NULL,
	[U_Sub5Name] [nvarchar](250) NULL,
	[U_GPKey] [int] NULL,
	[U_GPDetailsKey] [int] NULL,
	[U_GPDetailsName] [nvarchar](250) NULL,
	[U_Type] [nvarchar](50) NULL,
	[U_CTCV] [nvarchar](250) NULL);
CREATE TABLE BASELINE_KLTTC(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_GoodsIssue] [int] NULL,
	[U_DetailsKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_DetailsName] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL,
	[U_GoiThauKey] [int] NULL);
CREATE TABLE BASELINE_KLTTD(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_GoodsIssue] [int] NULL,
	[U_DetailsKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_DetailsName] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL,
	[U_GoiThauKey] [int] NULL);
CREATE TABLE BASELINE_KLTTE(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_SubprojectKey] [int] NULL,
	[U_StageKey] [int] NULL,
	[U_GoiThauKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_OpenIssueKey] [int] NULL,
	[U_OpenIssueRemark] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_KLTTF(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_SubProjectKey] [int] NULL,
	[U_StageKey] [int] NULL,
	[U_GoiThauKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_OpenIssueKey] [int] NULL,
	[U_OpenIssueRemark] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_KLTTG(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_SubProjectKey] [int] NULL,
	[U_StageKey] [int] NULL,
	[U_GoiThauKey] [int] NULL,
	[U_GoiThau] [nvarchar](254) NULL,
	[U_OpenIssueKey] [int] NULL,
	[U_OpenIssueRemark] [nvarchar](254) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_KLTTH(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_PBAKey] [int] NULL,
	[U_PBANumber] [int] NULL,
	[U_PBADate] [datetime] NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_PBADetailsKey] [int] NULL,
	[U_Type] [nvarchar](10) NULL,
	[U_ItemNo] [nvarchar](50) NULL,
	[U_ItemName] [nvarchar](100) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL);
CREATE TABLE BASELINE_KLTTK(
	[ID] [int] IDENTITY(1,1) PRIMARY KEY,
	[DocEntry_BaseLine] [int] NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineId] [int] NOT NULL,
	[U_GoiThauKey] [nvarchar](200) NULL,
	[U_GoiThau] [nvarchar](200) NULL,
	[U_GPKey] [int] NULL,
	[U_GPDetailsKey] [int] NULL,
	[U_GPDetailsName] [nvarchar](250) NULL,
	[U_Type] [nvarchar](50) NULL,
	[U_CTCV] [nvarchar](200) NULL,
	[U_Sub1] [nvarchar](200) NULL,
	[U_Sub2] [nvarchar](200) NULL,
	[U_Sub3] [nvarchar](200) NULL,
	[U_Sub4] [nvarchar](200) NULL,
	[U_Sub5] [nvarchar](200) NULL,
	[U_Sub1Name] [nvarchar](250) NULL,
	[U_Sub2Name] [nvarchar](250) NULL,
	[U_Sub3Name] [nvarchar](250) NULL,
	[U_Sub4Name] [nvarchar](250) NULL,
	[U_Sub5Name] [nvarchar](250) NULL,
	[U_UoM] [nvarchar](50) NULL,
	[U_Quantity] [numeric](19, 6) NULL,
	[U_UPrice] [numeric](19, 6) NULL,
	[U_Sum] [numeric](19, 6) NULL,
	[U_CompleteRate] [numeric](19, 6) NULL,
	[U_CompleteAmount] [numeric](19, 6) NULL);
GO

ALTER PROCEDURE [dbo].[BASELINE_GetList_Approve]
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
	Select 3 as 'LEVEL', 5 as 'Position'
	union all
	Select 6 as 'LEVEL', 3 as 'Position'
	union all
	Select 1 as 'LEVEL', 2 as 'Position'
	union all
	Select 1 as 'LEVEL', 1 as 'Position') X
	inner join OUDP Y on X.LEVEL = Y.Code
	inner join OHPS Z on X.Position = Z.posID;
END
GO

ALTER PROCEDURE [dbo].[BASELINE_GetList_Current]
	@Usr as varchar(200)
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
	where userID = (Select t.USERID from OUSR t where t.User_Code=@Usr)) a;

	Select DocEntry,U_FProject as 'Financial Project'
	,(Select PrjName from OPRJ where PrjCode=U_FProject) as 'Project Name'
	,U_BaseDate as 'BaseLine Date',U_Note as 'Note'
	,[Status],[Canceled] from [@BASELINE]
	where U_FProject in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Usr))) x inner join OHTM y on x.teamID = y.teamID)
	and (Select top 1 U_Level from [@BASELINE_APPR] where DocEntry = DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) = @Dept_Code;
END
GO

ALTER PROCEDURE [dbo].[BASELINE_GetList_Approved_Current]
	@Usr as varchar(200)
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
	where userID = (Select t.USERID from OUSR t where t.User_Code=@Usr)) a;

	Select DocEntry,U_FProject as 'Financial Project'
	,(Select PrjName from OPRJ where PrjCode=U_FProject) as 'Project Name'
	,U_BaseDate as 'BaseLine Date',U_Note as 'Note'
	,[Status],[Canceled]
	from [@BASELINE]
	where [Status]= 'C' 
	and [Canceled] not in ('Y','C')
	and U_FProject in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Usr))) x inner join OHTM y on x.teamID = y.teamID)
	and (Select top 1 U_Level from [@BASELINE_APPR] where DocEntry = DocEntry and (U_Status is null or U_Status = '4') order by LineId asc ) = @Dept_Code
END
GO

ALTER PROCEDURE [dbo].[BASELINE_GetList_Rejected_Current]
	@Usr as varchar(200)
AS
BEGIN
	Select DocEntry,U_FProject as 'Financial Project'
	,(Select PrjName from OPRJ where PrjCode=U_FProject) as 'Project Name'
	,U_BaseDate as 'BaseLine Date',U_Note as 'Note'
	,[Status],[Canceled] from [@BASELINE]
	where [Canceled] in ('Y','C')
	and U_FProject in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Usr))) x inner join OHTM y on x.teamID = y.teamID);
END
GO

ALTER PROCEDURE [dbo].[BASELINE_GetList]
	@Usr as varchar(200)
AS
BEGIN
	Select DocEntry,U_FProject as 'Financial Project',
	(Select PrjName from OPRJ where PrjCode=U_FProject) as 'Project Name'
	,U_BaseDate as 'BaseLine Date',U_Note as 'Note'
	,[Status],[Canceled] from [@BASELINE]
	where U_FProject in 
	(Select y.name as 'FProject' from (
	Select * from HTM1 where empID =
	(Select empID from OHEM
	where UserID = (
	Select USERID from OUSR where USER_CODE=@Usr))) x inner join OHTM y on x.teamID = y.teamID);
END
GO

ALTER PROCEDURE [dbo].[BASELINE_Add_Data]
	@BaseLine_DocEntry as int
AS
BEGIN
DECLARE @FProject as nvarchar(250)
SET XACT_ABORT ON;
BEGIN TRANSACTION;
BEGIN TRY
	Select @FProject = U_FProject from [@BASELINE] where DocEntry=@BaseLine_DocEntry;
	--Insert BASELINE_OPMG
	INSERT INTO BASELINE_OPMG([DocEntry_BaseLine], [AbsEntry], [OWNER], [NAME], [START]
							,[FINISHED] ,[DocNum] ,[Series] ,[TYP] ,[CARDCODE] ,[CARDNAME]
							,[CONTACT] ,[TERRITORY] ,[EMPLOYEE] ,[WithPhases] ,[STATUS]
							,[DUEDATE] ,[CLOSING] ,[FIPROJECT] ,[RISK] ,[INDUSTRY] ,[REASON]
							,[Free_Text] ,[BPLid] ,[AtcEntry] ,[Attachment] ,[LogInstanc]
							,[UpdateDate] ,[UserSign] ,[UserSign2] ,[CreateDate] ,[UpdateTS]
							,[U_BPTH] ,[U_PRJGROUP] ,[U_PRJTYPE] ,[U_CPHT1] ,[U_CPHT2]
							,[U_DPBH] ,[U_DPCP] ,[U_CPNG] ,[U_CPQLCT] ,[U_VT] ,[U_SUMTT]
							,[U_SUMTTDT] ,[U_SUMTTHD]) 
	Select @BaseLine_DocEntry, [AbsEntry], [OWNER], [NAME], [START]
							,[FINISHED] ,[DocNum] ,[Series] ,[TYP] ,[CARDCODE] ,[CARDNAME]
							,[CONTACT] ,[TERRITORY] ,[EMPLOYEE] ,[WithPhases] ,[STATUS]
							,[DUEDATE] ,[CLOSING] ,[FIPROJECT] ,[RISK] ,[INDUSTRY] ,[REASON]
							,[Free_Text] ,[BPLid] ,[AtcEntry] ,[Attachment] ,[LogInstanc]
							,[UpdateDate] ,[UserSign] ,[UserSign2] ,[CreateDate] ,[UpdateTS]
							,[U_BPTH] ,[U_PRJGROUP] ,[U_PRJTYPE] ,[U_CPHT1] ,[U_CPHT2]
							,[U_DPBH] ,[U_DPCP] ,[U_CPNG] ,[U_CPQLCT] ,[U_VT] ,[U_SUMTT]
							,[U_SUMTTDT] ,[U_SUMTTHD] from OPMG where FIPROJECT = @FProject and Status not in ('N','T');

	--Insert BASELINE_OPHA
	INSERT INTO BASELINE_OPHA ([DocEntry_BaseLine], [AbsEntry], [OWNER], [NAME], [START]
								,[FINISHED], [ParentID], [ProjectID] , [Code] , [TYP]
								,[CONTRIB], [STATUS], [END], [COST], [PLANNED], [Level]
								,[DUEDATE], [LogInstanc], [UpdateDate], [UserSign], [UserSign2]
								,[CreateDate], [UpdateTS], [U_001], [U_002], [U_KLDT], [U_DG]
								,[U_TTBV], [U_TTDT], [U_003], [U_REMARK], [U_TTHD], [U_DGHD])
	Select @BaseLine_DocEntry, [AbsEntry], [OWNER], [NAME], [START]
								,[FINISHED], [ParentID], [ProjectID] , [Code] , [TYP]
								,[CONTRIB], [STATUS], [END], [COST], [PLANNED], [Level]
								,[DUEDATE], [LogInstanc], [UpdateDate], [UserSign], [UserSign2]
								,[CreateDate], [UpdateTS], [U_001], [U_002], [U_KLDT], [U_DG]
								,[U_TTBV], [U_TTDT], [U_003], [U_REMARK], [U_TTHD], [U_DGHD] 
	from OPHA where ProjectID in (Select AbsEntry from BASELINE_OPMG where DocEntry_BaseLine = @BaseLine_DocEntry);

	--Insert BASELINE_PHA1
	INSERT INTO BASELINE_PHA1([DocEntry_BaseLine], [AbsEntry], [LineID], [StageID], [POS], [START]
							, [CLOSE], [Task], [DSCRIPTION], [EXPCOSTS], [InvAmtAR], [OpenAmtAR], [InvAmtAP]
							, [OpenAmtAP], [PERCENT], [FINISH], [OWNER], [StageDep1], [StageDep2], [StageDep3]
							, [StageDep4], [StDp1Type], [StDp2Type], [StDp3Type], [StDp4Type], [StDp1Abs]
							, [StDp2Abs], [StDp3Abs], [StDp4Abs], [LogInstanc], [AtcEntry])
	Select @BaseLine_DocEntry, [AbsEntry], [LineID], [StageID], [POS], [START]
							, [CLOSE], [Task], [DSCRIPTION], [EXPCOSTS], [InvAmtAR], [OpenAmtAR], [InvAmtAP]
							, [OpenAmtAP], [PERCENT], [FINISH], [OWNER], [StageDep1], [StageDep2], [StageDep3]
							, [StageDep4], [StDp1Type], [StDp2Type], [StDp3Type], [StDp4Type], [StDp1Abs]
							, [StDp2Abs], [StDp3Abs], [StDp4Abs], [LogInstanc], [AtcEntry] 
	from PHA1 
	where  AbsEntry in (Select AbsEntry from BASELINE_OPHA where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_PHA2
	INSERT INTO BASELINE_PHA2([DocEntry_BaseLine], [AbsEntry], [LineID], [StageID], [AREA], [PRIORITY]
							, [REMARKS], [CLOSED], [SOLUTIONID], [SOLUTION], [RESPNSIBLE], [ENTERED]
							, [DATE], [EFFORT], [LogInstanc], [U_NCCPS], [U_TENNCCPS], [U_DVTPS]
							, [U_KLPS], [U_DGPS], [U_TTPS], [U_Issuetype])
	Select @BaseLine_DocEntry, [AbsEntry], [LineID], [StageID], [AREA], [PRIORITY]
							, [REMARKS], [CLOSED], [SOLUTIONID], [SOLUTION], [RESPNSIBLE], [ENTERED]
							, [DATE], [EFFORT], [LogInstanc], [U_NCCPS], [U_TENNCCPS], [U_DVTPS]
							, [U_KLPS], [U_DGPS], [U_TTPS], [U_Issuetype]
	from PHA2 
	where AbsEntry in (Select distinct AbsEntry from BASELINE_PHA1 where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_CTG
	INSERT INTO BASELINE_CTG(DocEntry_BaseLine, DocEntry, U_PrjCode, U_PrjName, U_Date, U_GoiThauKey, U_GoiThauName)
	Select @BaseLine_DocEntry, DocEntry, U_PrjCode, U_PrjName, U_Date, U_GoiThauKey, U_GoiThauName from [@CTG]
	where DocEntry in (Select a.CTG_KEY 
						from (Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
								from [@CTG] 
								where U_PrjCode = @FProject
								group by U_GoiThauKey) a
						);

	--Insert BASELINE_CTG1
	INSERT INTO BASELINE_CTG1(DocEntry_BaseLine, DocEntry_CTG, LineID, U_001, U_ITEMNO, U_ITEMNAME, U_DVT, U_DGDAUTHAU, U_DGDUPHONG, U_DinhMuc, U_HAOHUT, U_TTDAUTHAU)
	Select @BaseLine_DocEntry, DocEntry, LineId, U_001, U_ITEMNO, U_ITEMNAME, U_DVT, U_DGDAUTHAU, U_DGDUPHONG, U_DinhMuc, U_HAOHUT, U_TTDAUTHAU 
	from [@CTG1] where DocEntry in (Select DocEntry from BASELINE_CTG where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_CTG2
	INSERT INTO BASELINE_CTG2(DocEntry_BaseLine, DocEntry_CTG, LineId, U_001, U_MATHIETBI, U_SLDUTRU, U_DVTTB, U_DGMUABAN, U_DGVCTB, U_DGVH, U_GTMB, U_GTTHUE, U_GTVANCHUYEN, U_GTVANHANH, U_NGAYCAP, U_NGAYTRA, U_SLTHUE, U_SLVANCHUYEN, U_SLVANHANH, U_TENTHIETBI, U_TENHM, U_DGTHUE)
	Select @BaseLine_DocEntry, DocEntry, LineId, U_001, U_MATHIETBI, U_SLDUTRU, U_DVTTB, U_DGMUABAN, U_DGVCTB, U_DGVH, U_GTMB, U_GTTHUE, U_GTVANCHUYEN, U_GTVANHANH, U_NGAYCAP, U_NGAYTRA, U_SLTHUE, U_SLVANCHUYEN, U_SLVANHANH, U_TENTHIETBI, U_TENHM, U_DGTHUE
	from [@CTG2] where DocEntry in (Select DocEntry from BASELINE_CTG where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_CTG3
	INSERT INTO BASELINE_CTG3(DocEntry_BaseLine, DocEntry_CTG, LineId, U_001, U_LOAICHIPHI, U_DGNCC, U_DGNTP, U_DGVTP, U_DGVC, U_DGCN, U_DGDTC, U_DGDP, U_DGDP2, U_DGPRELIM, U_DGTB, U_DGK, U_TENHM )
	Select @BaseLine_DocEntry, DocEntry, LineId, U_001, U_LOAICHIPHI, U_DGNCC, U_DGNTP, U_DGVTP, U_DGVC, U_DGCN, U_DGDTC, U_DGDP, U_DGDP2, U_DGPRELIM, U_DGTB, U_DGK, U_TENHM
	from [@CTG3] where DocEntry in (Select DocEntry from BASELINE_CTG where DocEntry_BaseLine=@BaseLine_DocEntry);
	
	--Insert BASELINE_CTG4
	INSERT INTO BASELINE_CTG4(DocEntry_BaseLine, DocEntry_CTG, LineId, U_001, U_TKKT, U_TTKKT, U_GTDP )
	Select @BaseLine_DocEntry, DocEntry, LineId, U_001, U_TKKT, U_TTKKT, U_GTDP
	from [@CTG4] where DocEntry in (Select DocEntry from BASELINE_CTG where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_DUTRU
	INSERT INTO BASELINE_DUTRU(DocEntry_BaseLine,DocEntry,CTG_Key,DUTRU_TYPE,FProject,ProjectID)
	Select @BaseLine_DocEntry
		, DocEntry
		, U_CTG_Key
		, U_DUTRU_TYPE
		, U_FinancialPrj
		, U_Goithau
	From [@DUTRU]
	where DocEntry in (
						Select DocEntry from [@DUTRU] 
						where U_DUTRU_TYPE = 1
						and U_CTG_Key in 
							(Select a.CTG_KEY from 
								(Select U_GoiThauKey,max(DocEntry) as CTG_KEY from [@CTG] where U_PrjCode = @FProject group by U_GoiThauKey) a
							)
					   );

	--Insert BASELINE_DUTRUA
	INSERT INTO BASELINE_DUTRUA(DocEntry_BaseLine, DocEntry_DUTRU, LineID, U_SubProjectCode
								, U_SubProjectDesc, U_CP_NCC, U_CP_NTP, U_CP_DTC
								, U_CP_VTP, U_CP_MB, U_CP_T, U_CP_VH
								, U_CP_VC, U_CP_CN, U_CP_DP, U_CP_DP2
								, U_CP_Prelims, U_CP_TB, U_CP_K, U_SplitTo)
	Select @BaseLine_DocEntry, DocEntry, LineId, U_SubProjectCode 
		, U_SubProjectDesc, U_CP_NCC, U_CP_NTP, U_CP_DTC
		, U_CP_VTP, U_CP_MB, U_CP_T, U_CP_VH
		, U_CP_VC, U_CP_CN, U_CP_DP, U_CP_DP2
		, U_CP_Prelims, U_CP_TB, U_CP_K, U_SplitTo
	from [@DUTRUA]
	where DocEntry in (Select DocEntry from BASELINE_DUTRU where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_DUTRUB
	INSERT INTO BASELINE_DUTRUB(DocEntry_BaseLine, DocEntry_DUTRU, LineID, U_DTT_LineID, U_SubProjectCode
								, U_SubProjectDesc, U_CP_NCC, U_BPCode, U_BPName
								, U_CP_NTP, U_CP_DTC , U_CP_VTP, U_CP_VC
								, U_CP_MB, U_CP_T , U_CP_VH, U_CP_CN, U_CP_DP, U_CP_DP2
								, U_CP_Prelims, U_CP_TB, U_CP_K, U_TYPE
								, U_TGDK, U_NCTN, U_PVCV)
	Select @BaseLine_DocEntry, 
		 DocEntry, LineID, U_DTT_LineID, U_SubProjectCode
		, U_SubProjectDesc, U_CP_NCC, U_BPCode, U_BPName
		, U_CP_NTP, U_CP_DTC , U_CP_VTP, U_CP_VC
		, U_CP_MB, U_CP_T , U_CP_VH, U_CP_CN, U_CP_DP, U_CP_DP2
		, U_CP_Prelims, U_CP_TB, U_CP_K, U_TYPE
		, U_TGDK, U_NCTN, U_PVCV
	from [@DUTRUB]
	where DocEntry in (Select DocEntry from BASELINE_DUTRU where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTT
	INSERT INTO BASELINE_KLTT([DocEntry_BaseLine], [DocEntry], [Canceled], [UserSign], [Status], [CreateDate]
							, [CreateTime], [UpdateDate], [UpdateTime], [Creator], [U_FIPROJECT], [U_DATEFROM]
							, [U_DATETO], [U_BPName], [U_BPCode], [U_Period], [U_CreatedDate], [U_VAT]
							, [U_GTTU], [U_BGroup], [U_BType], [U_HTTU], [U_PUType], [U_BPCode2], [U_PTQuanLy])
	Select @BaseLine_DocEntry, [DocEntry], [Canceled], [UserSign], [Status], [CreateDate]
							, [CreateTime], [UpdateDate], [UpdateTime], [Creator], [U_FIPROJECT], [U_DATEFROM]
							, [U_DATETO], [U_BPName], [U_BPCode], [U_Period], [U_CreatedDate], [U_VAT]
							, [U_GTTU], [U_BGroup], [U_BType], [U_HTTU], [U_PUType], [U_BPCode2], [U_PTQuanLy]
	from [@KLTT] 
	where DocEntry in (Select DocEntry 
						from [@KLTT] x inner join (Select U_BPCode,U_BGroup,U_PUType, MAx(U_Dateto) as Dateto 
													from [@KLTT] 
													where U_FIPROJECT = @FProject 
													and U_BType in (2,3) and Canceled not in ('Y','C') 
													group by U_BPCode,U_BGroup,U_PUType) y
						on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto 
						and x.U_BGroup = y.U_BGroup and x.U_PUType = y.U_PUType
						)
	Union all

	Select @BaseLine_DocEntry, [DocEntry], [Canceled], [UserSign], [Status], [CreateDate]
							, [CreateTime], [UpdateDate], [UpdateTime], [Creator], [U_FIPROJECT], [U_DATEFROM]
							, [U_DATETO], [U_BPName], [U_BPCode], [U_Period], [U_CreatedDate], [U_VAT]
							, [U_GTTU], [U_BGroup], [U_BType], [U_HTTU], [U_PUType], [U_BPCode2], [U_PTQuanLy]
	from [@KLTT] 
	where U_FIPROJECT = @FProject
	and U_BType = 1
	and Canceled not in ('Y','C');

	--Insert BASELINE_KLTTA
	INSERT INTO BASELINE_KLTTA ([DocEntry_BaseLine], [DocEntry], [LineId], [U_SubProjectKey], [U_SubProjectName]
							, [U_CompleteAmount], [U_Quantity], [U_GoiThauKey], [U_GoiThau], [U_GPKey]
							, [U_GPDetailsKey], [U_GPDetailsName], [U_UoM], [U_UPrice], [U_Sum], [U_CompleteRate]
							, [U_CTCV], [U_Sub1], [U_Sub2], [U_Sub3], [U_Sub4], [U_Sub5], [U_Sub1Name]
							, [U_Sub2Name], [U_Sub3Name], [U_Sub4Name], [U_Sub5Name], [U_Type])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_SubProjectKey], [U_SubProjectName]
							, [U_CompleteAmount], [U_Quantity], [U_GoiThauKey], [U_GoiThau], [U_GPKey]
							, [U_GPDetailsKey], [U_GPDetailsName], [U_UoM], [U_UPrice], [U_Sum], [U_CompleteRate]
							, [U_CTCV], [U_Sub1], [U_Sub2], [U_Sub3], [U_Sub4], [U_Sub5], [U_Sub1Name]
							, [U_Sub2Name], [U_Sub3Name], [U_Sub4Name], [U_Sub5Name], [U_Type] 
	from [@KLTTA] 
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTB
	INSERT INTO BASELINE_KLTTB([DocEntry_BaseLine], [DocEntry], [LineId], [U_SubProjectKey], [U_SubProjectName]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey], [U_GoiThau], [U_StageKey]
							 , [U_OpenIssueKey], [U_OpenIssueRemark], [U_Details], [U_UoM], [U_UPrice]
							 , [U_Quantity], [U_Sum], [U_Sub1], [U_Sub2], [U_Sub3], [U_Sub4], [U_Sub5]
							 , [U_Sub1Name], [U_Sub2Name], [U_Sub3Name], [U_Sub4Name], [U_Sub5Name]
							 , [U_GPKey], [U_GPDetailsKey], [U_GPDetailsName], [U_Type], [U_CTCV])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_SubProjectKey], [U_SubProjectName]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey], [U_GoiThau], [U_StageKey]
							 , [U_OpenIssueKey], [U_OpenIssueRemark], [U_Details], [U_UoM], [U_UPrice]
							 , [U_Quantity], [U_Sum], [U_Sub1], [U_Sub2], [U_Sub3], [U_Sub4], [U_Sub5]
							 , [U_Sub1Name], [U_Sub2Name], [U_Sub3Name], [U_Sub4Name], [U_Sub5Name]
							 , [U_GPKey], [U_GPDetailsKey], [U_GPDetailsName], [U_Type], [U_CTCV] 
	from [@KLTTB]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTC
	INSERT INTO BASELINE_KLTTC([DocEntry_BaseLine], [DocEntry], [LineId], [U_GoodsIssue], [U_DetailsKey]
							 , [U_GoiThau], [U_DetailsName], [U_UoM], [U_UPrice], [U_Quantity], [U_Sum]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_GoodsIssue], [U_DetailsKey]
							 , [U_GoiThau], [U_DetailsName], [U_UoM], [U_UPrice], [U_Quantity], [U_Sum]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey]
	from [@KLTTC]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTD
	INSERT INTO BASELINE_KLTTD([DocEntry_BaseLine], [DocEntry], [LineId], [U_GoodsIssue], [U_DetailsKey]
							 , [U_GoiThau], [U_DetailsName], [U_UoM], [U_UPrice], [U_Quantity], [U_Sum]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_GoodsIssue], [U_DetailsKey]
							 , [U_GoiThau], [U_DetailsName], [U_UoM], [U_UPrice], [U_Quantity], [U_Sum]
							 , [U_CompleteRate], [U_CompleteAmount], [U_GoiThauKey]
	from [@KLTTD]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTE
	INSERT INTO BASELINE_KLTTE([DocEntry_BaseLine], [DocEntry], [LineId], [U_SubprojectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_SubprojectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount]
	from [@KLTTE]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTF
	INSERT INTO BASELINE_KLTTF([DocEntry_BaseLine], [DocEntry], [LineId], [U_SubProjectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_SubProjectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount]
	from [@KLTTF]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTG
	INSERT INTO BASELINE_KLTTG([DocEntry_BaseLine], [DocEntry], [LineId], [U_SubProjectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_SubProjectKey], [U_StageKey]
							 , [U_GoiThauKey], [U_GoiThau], [U_OpenIssueKey], [U_OpenIssueRemark], [U_UoM]
							 , [U_UPrice], [U_Quantity], [U_Sum], [U_CompleteRate], [U_CompleteAmount]
	from [@KLTTG]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	--Insert BASELINE_KLTTH
	INSERT INTO BASELINE_KLTTH([DocEntry_BaseLine], [DocEntry], [LineId], [U_PBAKey], [U_PBANumber], [U_PBADate]
							 , [U_UoM], [U_PBADetailsKey], [U_Type], [U_ItemNo], [U_ItemName], [U_Quantity]
							 , [U_UPrice], [U_Sum])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_PBAKey], [U_PBANumber], [U_PBADate]
							 , [U_UoM], [U_PBADetailsKey], [U_Type], [U_ItemNo], [U_ItemName], [U_Quantity]
							 , [U_UPrice], [U_Sum]
	from [@KLTTH]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);
	
	--Insert BASELINE_KLTTK
	INSERT INTO BASELINE_KLTTK([DocEntry_BaseLine], [DocEntry], [LineId], [U_GoiThauKey], [U_GoiThau], [U_GPKey]
							 , [U_GPDetailsKey], [U_GPDetailsName], [U_Type], [U_CTCV], [U_Sub1], [U_Sub2]
							 , [U_Sub3], [U_Sub4], [U_Sub5], [U_Sub1Name], [U_Sub2Name], [U_Sub3Name], [U_Sub4Name]
							 , [U_Sub5Name], [U_UoM], [U_Quantity], [U_UPrice], [U_Sum], [U_CompleteRate], [U_CompleteAmount])
	Select @BaseLine_DocEntry, [DocEntry], [LineId], [U_GoiThauKey], [U_GoiThau], [U_GPKey]
							 , [U_GPDetailsKey], [U_GPDetailsName], [U_Type], [U_CTCV], [U_Sub1], [U_Sub2]
							 , [U_Sub3], [U_Sub4], [U_Sub5], [U_Sub1Name], [U_Sub2Name], [U_Sub3Name], [U_Sub4Name]
							 , [U_Sub5Name], [U_UoM], [U_Quantity], [U_UPrice], [U_Sum], [U_CompleteRate], [U_CompleteAmount]
	from [@KLTTK]
	where DocEntry in (Select DocEntry from BASELINE_KLTT where DocEntry_BaseLine=@BaseLine_DocEntry);

	COMMIT; 
END TRY

BEGIN CATCH
	ROLLBACK
	DECLARE @ErrorMessage VARCHAR(2000)
	SELECT @ErrorMessage = 'Error: ' + ERROR_MESSAGE()
	RAISERROR(@ErrorMessage, 16, 1)
END CATCH

END
GO

ALTER PROCEDURE [dbo].[BASELINE_GET_DATA_BCDT_A]
	-- Add the parameters for the stored procedure here
	@FinancialProject as varchar(100)
AS
BEGIN
	SET NOCOUNT ON;
	--DECLARE @ProjectID as int;
	--DECLARE @DocEntry as int;
	--SELECT top 1 @ProjectID = AbsEntry from OPMG where FIPROJECT = @FinancialProject;

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
	union
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
	and a.BpType = 'C') t1
	union
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
	and a.Series = 161
	and a.BpType = 'C') t2) z
END
GO

CREATE PROCEDURE [dbo].[BASELINE_MM_CE_GETDATA_SUM]
	-- Add the parameters for the stored procedure here
	@DocEntry_BaseLine as int
	,@GoiThauKey as varchar(250)
AS
BEGIN
	SET NOCOUNT ON;
    -- Insert statements for procedure here
	if (@GoiThauKey = '')
			Select * 
			FROM [BASELINE_DUTRUA] 
			where DocEntry_BaseLine = @DocEntry_BaseLine; 
	else
			Select * 
			FROM [BASELINE_DUTRUA] 
			where DocEntry_BaseLine= @DocEntry_BaseLine
			and DocEntry_DUTRU in  
			(Select DocEntry 
			 From BASELINE_DUTRU 
			 where ProjectID in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
			 and DocEntry_BaseLine = @DocEntry_BaseLine
			);
END
GO

CREATE PROCEDURE [dbo].[BASELINE_MM_CE_GETDATA_DETAILS]
	-- Add the parameters for the stored procedure here
	@DocEntry_BaseLine as int
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
	from [BASELINE_KLTT] a inner join
		(
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTA] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTB] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTK] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTC] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTD] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTE] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTF] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTG]
		where DocEntry_BaseLine = @DocEntry_BaseLine )b on a.DocEntry = b.DocEntry
	where 
	(Select GroupCode from OCRD where CardCode=a.U_BPCode) <> 112
	and a.U_BType <> 1
	and a.DocEntry_BaseLine = @DocEntry_BaseLine
	group by a.U_BPCOde,a.U_BPName,a.U_BPCode2,b.U_Sub3Name,a.U_PUTYPE;

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
				FROM [BASELINE_DUTRUB] 
				where DocEntry_BaseLine = @DocEntry_BaseLine
				group by [U_BPCode],[U_BPName],[U_SubProjectDesc],[U_DTT_LineID]) T0
			FULL JOIN
			@TableTmp_KLTT T1 on T0.U_BPCode = T1.U_BPCode and T0.U_SubProjectDesc = T1.U_Sub3Name ;
	end
		
	else
	begin
	--Get Data KLTT All Project
	INSERT INTO @TableTmp_KLTT(U_BPCode,U_BPName,U_BPCode2,U_Sub3Name,U_GoiThauKey,U_PUTYPE,KL_HD,KL_TT,KL_TT_DD)
	Select a.U_BPCOde
		,a.U_BPName
		,a.U_BPCode2
		,b.U_Sub3Name
		,b.U_GoiThauKey
		,a.U_PUTYPE
		,SUM(SUM_PL) as 'KL_HD'
		,SUM(SUM_CA) as 'KL_TT' 
		,SUM(case a.Status when 'C' then b.SUM_CA else 0 end) as 'KL_TT_DD'
	from [BASELINE_KLTT] a inner join
		(
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTA] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTB] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTK] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTC] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTD] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTE] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTF] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		union
		Select DocEntry,U_GoiThauKey,'' as U_Sub3Name,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
		from [BASELINE_KLTTG]
		where DocEntry_BaseLine = @DocEntry_BaseLine)b on a.DocEntry = b.DocEntry
	
	where 
	(Select GroupCode from OCRD where CardCode=a.U_BPCode) <> 112
	and a.U_BType <> 1
	and a.DocEntry_BaseLine = @DocEntry_BaseLine
	and b.U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
	group by a.U_BPCOde,a.U_BPName,a.U_BPCode2,b.U_GoiThauKey,a.U_PUTYPE,b.U_Sub3Name;

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
				FROM [BASELINE_DUTRUB] 
				where DocEntry_DUTRU in  
					(Select DocEntry 
					 From BASELINE_DUTRU 
					 where ProjectID in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
					 and DocEntry_BaseLine = @DocEntry_BaseLine
					)
				and DocEntry_BaseLine = @DocEntry_BaseLine
				group by [U_BPCode],[U_BPName],[U_SubProjectDesc],[U_DTT_LineID]) T0
			FULL JOIN
			@TableTmp_KLTT T1 on T0.U_BPCode = T1.U_BPCode and T0.U_SubProjectDesc = T1.U_Sub3Name ;
		end
END

GO

CREATE PROCEDURE [dbo].[BASELINE_MM_CE_GET_DATA_BCH]
	-- Add the parameters for the stored procedure here
	@DocEntry_BaseLine as int
	,@GoiThauKey as varchar(200)
	--,@CTG_Entry as int
AS
BEGIN
	SET NOCOUNT ON;
	--DECLARE @ProjectID as int;
	DECLARE @FinancialProject as nvarchar(200);
	Select @FinancialProject = U_FProject from [@BASELINE] where DocEntry = @DocEntry_BaseLine;
    -- Insert statements for procedure here
	--SELECT top 1 @ProjectID = AbsEntry from OPMG where FIPROJECT = @FinancialProject;
	if (@GoiThauKey = '')
		Select * from
		(
		Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' 
		FROM [BASELINE_CTG4] 
		where DocEntry_BaseLine = @DocEntry_BaseLine
		group by U_TKKT,U_TTKKT) a
		left join 
		(Select b.Account,SUM(b.Debit) as TOTAL_BCH
		From OJDT a inner join JDT1 b on a.TransID=b.TransId
		where a.Project = @FinancialProject
		and a.U_LCP = 'BCH'
		group by b.Account) b on a.U_TKKT=b.Account;
	else
		Select * from
		(Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,U_GTDP FROM [BASELINE_CTG4]
			where DocEntry_CTG in 
			(Select DocEntry from [BASELINE_CTG] where U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,',')))
		) a
		left join 
			(Select b.Account,SUM(b.Debit) as TOTAL_BCH
			From OJDT a inner join JDT1 b on a.TransID=b.TransId
			where a.Project = @FinancialProject
			and a.U_LCP = 'BCH'
			group by b.Account) b 
		on a.U_TKKT=b.Account ;
END;

GO

ALTER PROCEDURE [dbo].[BASELINE_MM_FI_GET_DATA_B]
	-- Add the parameters for the stored procedure here
	@DocEntry_BaseLine as int
	,@Goithau_Key as varchar(200)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @FinancialProject as varchar(200)
	Select @FinancialProject=U_FProject from [@BASELINE] where DocEntry = @DocEntry_BaseLine;
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_BaseLine = @DocEntry_BaseLine 
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_BaseLine = @DocEntry_BaseLine
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_BaseLine = @DocEntry_BaseLine
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]) T0
			LEFT JOIN
			(Select a.U_BPCode
				,a.U_BPName
				,a.U_BGroup
				,case when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT05','PUT06','PUT07','PUT08') and c.Series in (70,71) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (70,71) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (70,71) then 'PUT09'
					  --NTP
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (72,73) then 'PUT01'
					  when a.U_PUType  in ('PUT02','PUT05','PUT06') and c.Series in (72,73) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (72,73) then 'PUT09'
					  --DTC
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (78) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (78) then 'PUT02'
					  when a.U_PUType  in ('PUT05','PUT06','PUT09') and c.Series in (78) then 'PUT09'
					  else ISNULL(a.U_PUType,'') end as 'U_PUType'
				,a.U_BPCode2
				,a.U_PTQuanly
				,SUM(b.U_SUM) as 'KL_HD'
				,SUM(b.U_CompleteAmount) as 'KL_TT'
				,SUM(case a.Status when 'C' then (b.U_CompleteAmount) else 0 end) as 'KL_TT_DD'
				from [BASELINE_KLTT] a inner join (
				Select z1.DocEntry,SUM(z1.Sum_PL) as 'U_SUM', SUM(z1.SUM_CA) as 'U_CompleteAmount'
				from (
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTA]
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTB] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTK] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTC] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTD] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTE] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTF] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTG]
					where DocEntry_BaseLine= @DocEntry_BaseLine) z1
					group by z1.DocEntry
				) b on a.DocEntry = b.DocEntry and a.DocEntry_BaseLine= @DocEntry_BaseLine
				inner join  OCRD c on a.U_BPCode = c.CardCode
				--where a.DocEntry in (
				--Select --y.U_BPCode,
				--DocEntry from [@KLTT] x inner join (
				--Select U_BPCode,U_BGroup,U_PUType,MAx(U_Dateto) as Dateto from [@KLTT] where U_FIPROJECT = @FinancialProject group by U_BPCode,U_BGroup,U_PUType) y
				--on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto and x.U_BGroup = y.U_BGroup and x.U_PUType = y.U_PUType)
				group by  a.U_BPCode,a.U_BPName,a.U_BGroup
				,case when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT05','PUT06','PUT07','PUT08') and c.Series in (70,71) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (70,71) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (70,71) then 'PUT09'
					  --NTP
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (72,73) then 'PUT01'
					  when a.U_PUType  in ('PUT02','PUT05','PUT06') and c.Series in (72,73) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (72,73) then 'PUT09'
					  --DTC
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (78) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (78) then 'PUT02'
					  when a.U_PUType  in ('PUT05','PUT06','PUT09') and c.Series in (78) then 'PUT09'
					  else ISNULL(a.U_PUType,'') end
				,a.U_BPCode2,a.U_PTQuanly
				--,a.U_PUType
				) T1
				on T0.U_BPCode = T1.U_BPCode and T0.U_TYPE = T1.U_BGroup and T0.PUType = T1.U_PUType
				and (T0.U_CP_NCC <> 0 or T0.U_CP_DTC <> 0 or T0.U_CP_NTP <>0 )
			LEFT JOIN
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
				group by BpCode	,U_CGroup,U_PUType) y on x.BpCode = y.BpCode and x.U_CGroup = y.U_CGroup and x.U_PUTYPE = y.U_PUTYPE and x.StartDate = y.Last_Date
				where Series =48
				and Status = 'A') T2
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
		,T0.U_TYPE
		,T0.U_CP_NCC
		,T0.U_CP_NTP
		,T0.U_CP_DTC
		,T0.U_CP_DP2
		,T1.U_BGroup
		,T1.U_PUType
		,T1.KL_HD
		,T1.KL_TT
		,T1.KL_TT_DD
		,T2.GTHD
		,T3.TOTAL as 'TOTAL_AP_INVOICE' from 
			(
			--DU TRU
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_DUTRU in 
						(Select DocEntry
						from [BASELINE_DUTRU] 
						where DUTRU_TYPE = 1
						and DocEntry_BaseLine = @DocEntry_BaseLine
						and ProjectID in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
						)
			and DocEntry_BaseLine = @DocEntry_BaseLine
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_DUTRU in 
						(Select DocEntry
						from [BASELINE_DUTRU] 
						where DUTRU_TYPE = 1
						and DocEntry_BaseLine = @DocEntry_BaseLine
						and ProjectID in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
						)
			and DocEntry_BaseLine = @DocEntry_BaseLine
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
			FROM [BASELINE_DUTRUB] a inner join  OCRD b on a.U_BPCode = b.CardCode
			where a.DocEntry_DUTRU in 
						(Select DocEntry
						from [BASELINE_DUTRU] 
						where DUTRU_TYPE = 1
						and DocEntry_BaseLine = @DocEntry_BaseLine
						and ProjectID in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
						)
			and DocEntry_BaseLine = @DocEntry_BaseLine
			group by [U_BPCode],[U_BPName],b.Series,a.[U_TYPE]) T0
			LEFT JOIN
			--Khoi luong thanh toan
			(Select a.U_BPCode
				,a.U_BPName
				,a.U_BGroup
				--,a.U_PUType
				--NCC
				,case when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT05','PUT06','PUT07','PUT08') and c.Series in (70,71) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (70,71) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (70,71) then 'PUT09'
					  --NTP
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (72,73) then 'PUT01'
					  when a.U_PUType  in ('PUT02','PUT05','PUT06') and c.Series in (72,73) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (72,73) then 'PUT09'
					  --DTC
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (78) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (78) then 'PUT02'
					  when a.U_PUType  in ('PUT05','PUT06','PUT09') and c.Series in (78) then 'PUT09'
					  else ISNULL(a.U_PUType,'') end as 'U_PUType'
				,SUM(b.U_SUM) as 'KL_HD'
				,SUM(b.U_CompleteAmount) as 'KL_TT'
				,SUM(case a.Status when 'C' then b.U_CompleteAmount else 0 end) as 'KL_TT_DD'
				from [BASELINE_KLTT] a inner join (
				Select z1.DocEntry,SUM(z1.Sum_PL) as 'U_SUM', SUM(z1.SUM_CA) as 'U_CompleteAmount'
				from (
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTA]
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTB] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTK] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTC] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,-U_SUM as 'Sum_PL',-U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTD] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTE] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTF] 
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))
					union
					Select DocEntry,U_SUM as 'Sum_PL',U_CompleteAmount as 'SUM_CA'
					from [BASELINE_KLTTG]
					where DocEntry_BaseLine= @DocEntry_BaseLine
					and U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@Goithau_Key,','))) z1
					group by z1.DocEntry
				) b on a.DocEntry = b.DocEntry and a.DocEntry_BaseLine = @DocEntry_BaseLine
				inner join  OCRD c on a.U_BPCode = c.CardCode
				--where a.DocEntry in (
				--Select --y.U_BPCode,
				--DocEntry from [@KLTT] x inner join (
				--Select U_BPCode,U_BGroup,U_PUType,MAx(U_Dateto) as Dateto from [@KLTT] where U_FIPROJECT = @FinancialProject and U_GoiThauKey=@Goithau_Key group by U_BPCode,U_BGroup,U_PUType) y
				--on x.U_BPCode = y.U_BPCode and x.U_DATETO = y.Dateto and x.U_BGroup = y.U_BGroup and x.U_PUType = y.U_PUType
				--where b.U_Sub1 in (Select AbsEntry from OPHA where ProjectID = @Goithau_Key))
				group by  a.U_BPCode,a.U_BPName,a.U_BGroup
				,case when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT05','PUT06','PUT07','PUT08') and c.Series in (70,71) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (70,71) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (70,71) then 'PUT09'
					  --NTP
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (72,73) then 'PUT01'
					  when a.U_PUType  in ('PUT02','PUT05','PUT06') and c.Series in (72,73) then 'PUT02'
					  when a.U_PUType  in ('PUT09') and c.Series in (72,73) then 'PUT09'
					  --DTC
					  when a.U_PUType  in ('PUT01','PUT03','PUT04','PUT07','PUT08') and c.Series in (78) then 'PUT01'
					  when a.U_PUType  in ('PUT02') and c.Series in (78) then 'PUT02'
					  when a.U_PUType  in ('PUT05','PUT06','PUT09') and c.Series in (78) then 'PUT09'
					  else ISNULL(a.U_PUType,'') end
				--,--a.U_PUType
				) T1
				on T0.U_BPCode = T1.U_BPCode and T0.U_TYPE = T1.U_BGroup 
			LEFT JOIN
			--Hop dong
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
				and U_GOITHAU = @Goithau_Key
				group by BpCode	,U_CGroup,U_PUType) y on x.BpCode = y.BpCode and x.U_CGroup = y.U_CGroup and x.U_PUTYPE = y.U_PUTYPE and x.StartDate = y.Last_Date
				where Series = 48
				and Status = 'A') T2
			on T0.U_BPCode = T2.BpCode 
			and T1.U_BGroup = T2.U_CGroup
			and T1.U_PUType = T2.U_PUType
			LEFT JOIN 
			--Hoa don - AP Invoice
			(Select a.CardCode
				,a.U_RECTYPE
				,a.U_PUTYPE
				,(Select SUM(LineTotal) from PCH1 where DocEntry=a.DocEntry and U_ParentID1 in (Select AbsEntry from OPHA where ProjectID = @Goithau_Key)) as 'TOTAL'
			from OPCH a
			where a.Project = @FinancialProject
			and a.CANCELED not in ('Y','C'))T3
			on T2.BpCode = T3.CardCode
			and T2.U_CGroup = T3.U_RECTYPE
			and T2.U_PUTYPE = T3.U_PUTYPE
	end
END

GO

ALTER PROCEDURE [dbo].[BASELINE_MM_FI_GET_DATA_BCH]
	-- Add the parameters for the stored procedure here
	@DocEntry_BaseLine as int
	,@GoiThauKey as varchar(200)
	--,@CTG_Entry as int
AS
BEGIN
	SET NOCOUNT ON;
	--DECLARE @ProjectID as int;
	DECLARE @FinancialProject as nvarchar(200);
	Select @FinancialProject = U_FProject from [@BASELINE] where DocEntry = @DocEntry_BaseLine;
    -- Insert statements for procedure here
	--SELECT top 1 @ProjectID = AbsEntry from OPMG where FIPROJECT = @FinancialProject;
	if (@GoiThauKey = '')
		Select * from
		(
		Select left(U_TKKT + '00000000',8) as 'U_TKKT',U_TTKKT,SUM(U_GTDP) as 'U_GTDP' 
		FROM [BASE_LINECTG4] 
		where DocEntry_BaseLine = @DocEntry_BaseLine 
		--in 
		--(Select DocEntry 
		--from [@DUTRU] 
		--where 
		--U_DUTRU_TYPE=1
		--and U_CTG_Key in 
		--		(Select x.CTG_KEY 
		--		from 
		--			(Select U_GoiThauKey,max(DocEntry) as CTG_KEY 
		--			from [@CTG] 
		--			where U_PrjCode = @FinancialProject group by U_GoiThauKey) x))
		group by U_TKKT,U_TTKKT) a
		left join 
		(Select b.Account,SUM(b.Debit) as TOTAL_BCH
		From OJDT a inner join JDT1 b on a.TransID=b.TransId
		where a.Project = @FinancialProject
		and a.U_LCP = 'BCH'
		group by b.Account) b on a.U_TKKT=b.Account;
	else
		Select * from
		(Select left(x.U_TKKT + '00000000',8) as 'U_TKKT',x.U_TTKKT,x.U_GTDP FROM [BASELINE_CTG4] x inner join [BASELINE_CTG] y on x.DocEntry_CTG=y.DocEntry and y.DocEntry_BaseLine=@DocEntry_BaseLine
			where x.DocEntry_BaseLine = @DocEntry_BaseLine 
			and y.U_GoiThauKey in (Select splitdata from dbo.fnSplitString(@GoiThauKey,','))
			--in 
			--(Select DocEntry from [@DUTRU] where U_CTG_Key in 
			--	(Select DocEntry From [@CTG] where U_PrjCode = @FinancialProject and U_GoiThauKey = @GoiThauKey)
			--)
			) a
		left join 
			(Select b.Account,SUM(b.Debit) as TOTAL_BCH
			From OJDT a inner join JDT1 b on a.TransID=b.TransId
			where a.Project = @FinancialProject
			and a.U_LCP = 'BCH'
			group by b.Account) b 
		on a.U_TKKT=b.Account ;
END;
GO

CREATE PROCEDURE [dbo].[BASELINE_Approve_LV]
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
Update [@BASELINE_APPR] 
set U_Usr = @UserName
	,U_Time=CONVERT(varchar(30), GETDATE(), 113)
	,U_Status = @Status
	,U_Comment = @Comment
where DocEntry = @DocEntry
and U_Posistion = @Pos_Code
and U_Level = @Dept_Code
and (U_Status is null or U_Status ='4')
and LineID = (Select Min(LineID) from [@BASELINE_APPR]
				where  DocEntry = @DocEntry
				and U_Level = @Dept_Code
				and U_Posistion = @Pos_Code
				and (U_Status is null or U_Status ='4'));
--Update them truong hop khi truong phong duyet ko qua nhan vien
if @Pos_Code = 1
		Update [@BASELINE_APPR]
		set U_Usr = @UserName
			,U_Time=CONVERT(varchar(30), GETDATE(), 113)
			,U_Status = '3'
		where DocEntry = @DocEntry
		and U_Posistion = 2
		and U_Level = @Dept_Code
		and U_Status is null
		and LineID = (Select Min(LineID) from [@BASELINE_APPR]
						where  DocEntry = @DocEntry
						and U_Level = @Dept_Code
						and U_Posistion = 2
						and (U_Status is null));
SELECT @Update_Row = @@ROWCOUNT;
RETURN @Update_Row;
END
GO

CREATE PROCEDURE [dbo].[BASELINE_Get_Lst_Usr_LV]
	-- Add the parameters for the stored procedure here
	@DocEntry as int
AS
BEGIN
--Get User Info - Dept - Position
Declare @DeptCode as int
Declare @PosCode as int
Declare @FProject as varchar(250)
	Select top 1 @DeptCode = ISNULL(U_Level,''),@PosCode=ISNULL(U_Posistion,'') from [@BASELINE_APPR] 
	where DocEntry = @DocEntry 
	and U_Status is null 
	order by LineID;

	Select @FProject=ISNULL(U_FPROJECT,'') from [@BASELINE] where DocEntry=@DocEntry;

	Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name
	from OHEM a inner join OUSR b on a.USERID = b.UserID
	left join HTM1 c on c.empID=a.empID
	inner join OHTM d on c.teamID = d.teamID
	where a.dept = @DeptCode
	and a.position = @PosCode
	and d.name = @FProject;
END
GO
