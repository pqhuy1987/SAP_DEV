CREATE PROCEDURE [dbo].[DeleteJournalEntry_Delivery] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM ODLN WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE ODLN SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_GoodIssue] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OIGE WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OIGE SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_GoodReceipt] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OIGN WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OIGN SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_GoodReceiptPO] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OPDN WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OPDN SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_GoodReturn] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM ORPD WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE ORPD SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30'AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_InventoryTransfer] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OWTR WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OWTR SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_IssueforProduction] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OIGE WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OIGE SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN

GO

CREATE PROCEDURE [dbo].[DeleteJournalEntry_ReceiptfromProduction] @DocNum as int
AS
SET NOCOUNT OFF
DECLARE @SoDongXoa AS Int
DECLARE @TransID as Int
SELECT @TransID = ISNULL(TransId,0) FROM OIGN WHERE DocNum = @DocNum
IF @TransID > 0
BEGIN
	DELETE FROM JDT1 WHERE TransId = @TransID
	DELETE FROM OJDT WHERE TransId = @TransID
	SELECT @SoDongXoa = @@ROWCOUNT
	IF @SoDongXoa > 0
	BEGIN
		UPDATE OIGN SET TransId = 0 WHERE DocNum = @DocNum
		DECLARE @NextNumber as int 
		DECLARE @AutoKey as int
		SELECT @NextNumber = NextNumber FROm NNM1 WHERE ObjectCode = '30' AND Series = '17'
		SELECT @AutoKey = AutoKey FROM ONNM WHERE ObjectCode = '30' AND DfltSeries = '17'
		UPDATE NNM1 SET NextNumber = @NextNumber -1 WHERE ObjectCode = '30' AND Series = '17'
		UPDATE ONNM SET AutoKey = @AutoKey - 1 WHERE ObjectCode = '30' AND DfltSeries = '17'
	END
END
RETURN