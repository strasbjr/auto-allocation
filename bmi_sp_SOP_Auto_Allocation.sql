USE [BLU]
GO

/****** Object:  StoredProcedure [dbo].[bmi_sp_SOP_Auto_Allocation]    Script Date: 5/22/2018 2:46:26 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE PROC [dbo].[bmi_sp_SOP_Auto_Allocation]
	(@I_vUserID VARCHAR(50) = 'sa')

-- VERSION 5/31/2017
/*---------------------------------------------------------------------------------------
Created by InterDyn BMI.  All Rights Reserved.

TITLE:				bmi_sp_SOP_Auto_Allocation
USAGE:				SQL Procedure used in both SQL job and called manually
					from Dex window to auto allocate SOP documents.
			
DEPENDENCIES:		eConnect Stored Procs:
						taSOPLineIvcInsert
						taSopLineIvcInsertComponent
					SQL
						DB Mail must be installed, configured, and working.

CREATE IN:			Any company database that uses bmiSOPAutoAllocation modifications.

28-AUG-2016			ZMAN		Created.
12-SEP-2016			ZMAN		Added updates for kits so that quantities would be 
								allocated immediately.
13-SEP-2016			ZMAN		Fixed issue with partially cancelled items.
14-SEP-2016			ZMAN		Fixed issue with qty to invoice not getting set correctly.
16-SEP-2016			ZMAN		Change Order - Added fields to tables for reporting purposes.
23-SEP-2016			ZMAN		Fixed issue with email sends when running manually from Dex.
26-SEP-2016						PROMOTED TO PROD
30-SEP-206			ZMAN		Added code to clear locks prior to next run in eConnectOutTemp
								and DEX_LOCK
13-OCT-2016			ZMAN		Fixed kit component partial allocations.
17-OCT-2016			ZMAN		Added another order by clause to cursor.
18-OCT-2016			ZMAN		Added ways to skip records that are locked and report them to user.
20-OCT-2016			ZMAN		Added pricing fix.
25-OCT-2016			ZMAN		Worked on locks issues, made sure all selectes were nolock
29-MAR-2017			ZMAN		Added report subscription
28-APR-2017			ZMAN		Added start and end email notifications.
31-MAY-2017			ZMAN		Added markdown amounts.
02-JAN-2018			JSTRASBURG	Altered HTML email to only send rows that did not show successful.
09-JAN-2019			ZMAN		SOP Auto Allocation Revisions Project (See Dev Doc)
27-MAR-2018			JSTRASBURG	New version with revisions (from ZMAN's 1/9/18 entry) released to PROD
26-APR-2018			JSTRASBURG	Added Deadlock Priority
01-MAY-2018			JSTRASBURG	Disabled the SmartConnect Item Update Trigger during the execution
11-MAY-2018			JSTRASBURG	Removed Deadlock Priority
15-MAY-2018			JSTRASBURG	Disabled the trigger that updates the timestamp on SOP10200
								and the SmartConnect SOP Header trigger (not necessary for this).
								This is an attempt to alleviate deadlocks.
								Added an update to Sales Orders to catch up the missed time 
								while the SC trigger was disabled.
17-MAY-2018			JSTRASBURG	Added CDPPARTS (site) and TGTPTS BO (bachnumb) as an
								additional scenario for AA.
---------------------------------------------------------------------------------------*/
AS
SET NOCOUNT ON

if (SUSER_NAME() = 'NT SERVICE\SQLSERVERAGENT')
begin
	-- Disable certain triggers that commonly cause deadlocking during this time.  This won't
	-- execute as a regular user.
	alter table IV00101 DISABLE TRIGGER sc_trgen_IV00101_update_ONGOING_GP_ITEMS_TO_WOW_BLU_DOT
	alter table SOP10200 DISABLE TRIGGER zDT_SOP10200U
	alter table SOP10100 DISABLE TRIGGER sc_trgen_SOP10100_update_ONGOING_GO_SOP_TO_WOW_UPDATE_BLU_DOT
	alter table SOP10200 DISABLE TRIGGER sc_trgen_SOP10200_update_ONGOING_GO_SOP_TO_WOW_UPDATE_BLU_DOT
end

DECLARE @ProfileName VARCHAR(MAX) = 'Administrator'
DECLARE @Receipients VARCHAR(MAX) = 'bdb-carl@bludot.com;ssweerin@bludot.com;jstrasburg@bludot.com'
--DECLARE @Receipients VARCHAR(MAX) = 'jstrasburg@bludot.com'

DECLARE
	@I_vSOPTYPE integer, 
	@I_vSOPNUMBE varchar(25), 
	@I_vCUSTNMBR varchar(20),
	@I_vDOCDATE datetime,
	@I_vDOCID char(15),
	@I_vLNITMSEQ integer, --FOR KIT COMPONENTS, PASS IN KIT PARENT ITEM LNITMSEQ
	@I_vCMPNTSEQ integer,
	@I_vITEMNMBR varchar(50),
	@I_vLOCNCODE varchar(20),
	@I_vUNITPRCE numeric(19,5),
	@I_vQUANTITY numeric(19,5),
	@I_vReqShipDate datetime,
	@I_vACTLSHIP datetime,
	@I_vSHIPMTHD char(15),
	@I_vPRSTADCD char(15),
	@I_vShipToName char(64),
	@I_vCNTCPRSN char(60),
	@I_vADDRESS1 char(60),
	@I_vADDRESS2 char(60),
	@I_vADDRESS3 char(60),
	@I_vCITY char(35),
	@I_vSTATE char(29),
	@I_vZIPCODE char(10),
	@I_vCOUNTRY char(60),
	@I_vPHONE1 char(21),
	@I_vPHONE2 char(21),
	@I_vPHONE3 char(21),
	@I_vFAXNUMBR char(21),
	@I_vPrint_Phone_NumberGB integer

/*OTHER REQUIRED VARIABLES*/
DECLARE @I_vUSERDATE datetime = GetDate()
DECLARE @I_vAutoAssignBin integer = 1
DECLARE @I_vXTNDPRCE numeric(19,5) = 0
DECLARE @I_vMRKDNAMT numeric(19,5)
DECLARE @I_vMRKDNPCT numeric(19,5)
DECLARE @I_vMRKDNAMT_HDR numeric(19,5)
DECLARE @I_vCOMMNTID char(15) = ''
DECLARE @I_vCOMMENT_1 char(50) = ''
DECLARE @I_vCOMMENT_2 char(50) = ''
DECLARE @I_vCOMMENT_3 char(50) = ''
DECLARE @I_vCOMMENT_4 char(50) = ''
DECLARE @I_vUNITCOST numeric(19,5)
DECLARE @I_vPRCLEVEL char(10) = ''
DECLARE @I_vITEMDESC char(100) = ''
DECLARE @I_vTAXAMNT numeric(19,5) = 0
DECLARE @I_vQTYONHND numeric(19,5) = 0
DECLARE @I_vQTYRTRND numeric(19,5) = 0
DECLARE @I_vQTYINUSE numeric(19,5) = 0
DECLARE @I_vQTYINSVC numeric(19,5) = 0
DECLARE @I_vQTYDMGED numeric(19,5) = 0
DECLARE @I_vNONINVEN integer = 0
DECLARE @I_vDROPSHIP integer = 0
DECLARE @I_vQTYTBAOR numeric(19,5) = 0
DECLARE @I_vSALSTERR char(15) = ''
DECLARE @I_vSLPRSNID char(15) = ''
DECLARE @I_vITMTSHID char(15) = ''
DECLARE @I_vIVITMTXB integer = 0
DECLARE @I_vTAXSCHID char(15) = ''
DECLARE @I_vEXCEPTIONALDEMAND integer = 0
DECLARE @I_vFUFILDAT datetime = '1/1/1900'
DECLARE @I_vINVINDX varchar(75) = ''
DECLARE @I_vCSLSINDX varchar(75) = ''
DECLARE @I_vSLSINDX varchar(75) = ''
DECLARE @I_vMKDNINDX varchar(75) = ''
DECLARE @I_vRTNSINDX varchar(75) = ''
DECLARE @I_vINUSINDX varchar(75) = ''
DECLARE @I_vINSRINDX varchar(75) = ''
DECLARE @I_vDMGDINDX varchar(75) = ''
DECLARE @I_vAUTOALLOCATESERIAL integer = 0
DECLARE @I_vAUTOALLOCATELOT integer = 0
DECLARE @I_vGPSFOINTEGRATIONID char(30) = ''
DECLARE @I_vINTEGRATIONSOURCE integer = 0
DECLARE @I_vINTEGRATIONID char(30) =  ''
DECLARE @I_vRequesterTrx integer = 0
DECLARE @I_vQTYCANCE numeric(19,5) = 0
DECLARE @I_vQTYFULFI numeric(19,5)    --Do not pass anything in, including a zero, as it will use the value to popuate the quantity fulfilled
DECLARE @I_vALLOCATE integer = 0
DECLARE @I_vUpdateIfExists integer = 1  --document will exist
DECLARE @I_vRecreateDist integer = 1
DECLARE @I_vQUOTEQTYTOINV numeric(19,5) = 0
DECLARE @I_vTOTALQTY numeric(19,5) = 0
DECLARE @I_vCMMTTEXT varchar(500) = ''
DECLARE @I_vDEFPRICING integer = 0
DECLARE @I_vDEFEXTPRICE integer = 0
DECLARE @I_vCURNCYID char(15) = ''
DECLARE @I_vUOFM char(8) = ''
DECLARE @I_vIncludePromo integer = 0
DECLARE @I_vCKCreditLimit integer = 0
DECLARE @I_vRECREATETAXES integer = 1
DECLARE @I_vRECREATECOMM integer = 1
DECLARE @I_vUSRDEFND1 char(50) = ''
DECLARE @I_vUSRDEFND2 char(50) = ''
DECLARE @I_vUSRDEFND3 char(50) = ''
DECLARE @I_vUSRDEFND4 varchar(8000) = ''
DECLARE @I_vUSRDEFND5 varchar(8000) = ''
DECLARE @I_vATYALLOC numeric(19,5) = 0
DECLARE @I_vCMPITUOM char(9) = ''
DECLARE @I_vKitCompMan integer = 0
DECLARE @I_vQtyShrtOpt integer = 0
DECLARE @I_vQTYTOINV numeric(19,5) = 0

DECLARE @QTYTOALLOCATE numeric(19,5) = 0
DECLARE @QTYALLOCATE NUMERIC(19,5) = 0
DECLARE @DEX_ROW_ID INT
DECLARE @LOCKTBL VARCHAR(MAX) = RTRIM(db_name()) + '.dbo.SOP10100'
DECLARE @LOCKINT TINYINT = 0
DECLARE @iStatus INT = 0

--ERROR TRAPPING - ACTIVE LOCKING
DECLARE @oExists INT
DECLARE @OInsStatus INT
DECLARE @DexLockErrorState INT
DECLARE @iError INT
DECLARE @O_oErrorState INT
DECLARE @RECORDLOCKED INT
DECLARE @ECTEMPLOCK INT
DECLARE @DLLOCK INT

--ERROR TRAPPING VARIABLES
DECLARE 
	@ErrStringLines varchar(255),
	@RtnString varchar(2000),
	@O_iErrorState integer = 0,
	@oErrString varchar(255) = '',
	@str varchar(255),
	@strlen int,
	@codestrt int,
	@codeend int,
	@code char(10), 
	@codedesc varchar(2000),
	@codedescsingle varchar(2000),
	@seqno int

DECLARE @SEQNUM INT = 1
DECLARE @EXECDATE DATE

--VARIABLES FOR RECORD LOCKING
DECLARE @cSOPTYPE INT = 0
DECLARE @cSOPNUMBE VARCHAR(25) = ''

DECLARE @QTYAVLIV NUMERIC(19,5)
DECLARE @QTYAVLORD NUMERIC(19,5)
DECLARE @NEWQTYTOINV NUMERIC(19,5)
DECLARE @MSTRNUMB INT = 0
DECLARE @TIME TIME(0) = SYSDATETIMEOFFSET()
DECLARE @DATETIME DATETIME = @TIME
DECLARE @DTVC VARCHAR(10) = CONVERT(VARCHAR(10), @DATETIME, 108)

--DBMAIL VARIABLES FOR START AND END EMAILS
DECLARE @NOTIFICATIONSSubject NVARCHAR(MAX)
DECLARE @NOTIFICATIONSBody NVARCHAR(MAX)
DECLARE @NOTIFICATIONSProcessTime DATETIME = GETDATE()

--NEW VARIABLES FOR REVISIONS
DECLARE @Cv_SOPTYPE INT  = 0
DECLARE @Cv_SOPNUMBE VARCHAR(25) = ''
DECLARE @Cv_DEX_ROW_ID INT = 0
DECLARE @Cv_MSTRNUMB INT = 0
DECLARE @Cv_ID INT = 0
DECLARE @ORDFROM INT = 0
DECLARE @ORDTO INT = 0
DECLARE @I_vID INT = 0
DECLARE @LINEFROM INT = 0
DECLARE @LINETO INT = 0

SELECT @I_vUserID = RTRIM(@I_vUserID)

--SET THE SEQ NUMBER FOR THE PROCESS
SELECT @SEQNUM = COALESCE(MAX(SEQNUMBR),0) FROM bmiSOPAutoAllocationLog with (NOLOCK)
IF @SEQNUM = 0 
BEGIN
	SELECT @SEQNUM = 1
END
ELSE
BEGIN
	SET @SEQNUM = @SEQNUM + 1
END

DECLARE @SOP10200TEMP AS TABLE
(
	I_vID int not null primary key identity(1,1),
	I_vSOPTYPE smallint, 
	I_vSOPNUMBE varchar(25), 
	I_vCUSTNMBR varchar(20),
	I_vDOCDATE datetime,
	I_vDOCID varchar(20),
	I_vLNITMSEQ int, 
	I_vCMPNTSEQ int, 
	I_vITEMNMBR varchar(35), 
	I_vLOCNCODE varchar(15), 
	I_vUNITPRCE numeric(19,5), 
	I_vXTNDPRCE numeric(19,5),
	I_vQUANTITY numeric(19,5),
	I_vQTYCANCE numeric(19,5),
	I_vATYALLOC numeric(19,5),
	I_vTOTALQTY numeric(19,5),
	I_vQTYTOINV numeric(19,5),
	I_vReqShipDate datetime,
	I_vACTLSHIP datetime, 
	I_vSHIPMTHD varchar(15), 
	I_vPRSTADCD varchar(15),
	I_vShipToName varchar(65), 
	I_vCNTCPRSN varchar(61), 
	I_vADDRESS1 varchar(61), 
	I_vADDRESS2 varchar(61), 
	I_vADDRESS3 varchar(61), 
	I_vCITY varchar(35), 
	I_vSTATE varchar(29), 
	I_vZIPCODE varchar(11), 
	I_vCOUNTRY varchar(61), 
	I_vPHONE1 varchar(21), 
	I_vPHONE2 varchar(21), 
	I_vPHONE3 varchar(21), 
	I_vFAXNUMBR varchar(21), 
	I_vPrint_Phone_NumberGB smallint,
	I_vCMPITUOM CHAR(9),
	I_vITMTSHID CHAR(15),
	I_vMRKDNAMT_HDR NUMERIC(19,5),
	I_vMRKDNAMT_LINE NUMERIC(19,5),
	I_vMRKDNPCT int,
	MSTRNUMB INT,
	RECORDLOCKED INT,
	DEX_ROW_ID INT
)

DECLARE @ORDSTEMP AS TABLE
(
	Cv_ID int not null primary key identity(1,1),
	Cv_SOPTYPE smallint, 
	Cv_SOPNUMBE varchar(25),
	Cv_MSTRNUMB INT,
	Cv_DEX_ROW_ID INT
)

INSERT INTO @ORDSTEMP
(
	Cv_SOPTYPE, Cv_SOPNUMBE, Cv_MSTRNUMB, Cv_DEX_ROW_ID
)
SELECT 
	DISTINCT L.SOPTYPE, L.SOPNUMBE, H.MSTRNUMB, H.DEX_ROW_ID
FROM 
	SOP10200 L WITH(NOLOCK)
	INNER JOIN SOP10100 H WITH(NOLOCK) ON H.SOPTYPE = L.SOPTYPE AND H.SOPNUMBE = L.SOPNUMBE
	INNER JOIN RM00101 C WITH(NOLOCK) ON C.CUSTNMBR = H.CUSTNMBR
	INNER JOIN IV00101 I WITH(NOLOCK) ON I.ITEMNMBR = L.ITEMNMBR
WHERE 
	L.SOPTYPE=2 
	AND L.QTYCANCE < L.QUANTITY
	AND 
	(
	(L.LOCNCODE = 'BDMN' AND H.BACHNUMB IN ('VERIFY','HOLD TO SHIP', 'BACKORDER', 'VERIFY-SC'))
	OR
	(L.LOCNCODE = 'CDPPARTS' AND H.BACHNUMB = 'TGTPTS BO')
	)					
	AND L.PURCHSTAT IN (1,2)
	AND I.ITEMTYPE IN (1,3)
	AND H.VOIDSTTS = 0
ORDER BY
	H.MSTRNUMB ASC, L.SOPTYPE, L.SOPNUMBE

--IF THERE ARE NO ORDERS, THEN STOP HERE
DECLARE @RECRET INT
SELECT @RECRET = COUNT(Cv_SOPNUMBE) FROM @ORDSTEMP
IF (@RECRET) = 0 
BEGIN
	--NO RECORDS RETURNED SEND EMAIL AND STOP HERE
	
	--SEND NOTIFICATION EMAIL
	SELECT @NOTIFICATIONSSubject = N'SOP AUTO ALLOCATION ' + CONVERT(NVARCHAR(20),@SEQNUM) + N' FOUND NO RECORDS TO PROCESS ' + CONVERT(NVARCHAR(20),@NOTIFICATIONSProcessTime)
	SELECT @NOTIFICATIONSBody = N'SOP Auto Allocation has found no records to process.'
	EXEC msdb.dbo.sp_send_dbmail  
		@profile_name = @ProfileName,  
		@recipients = @Receipients,
		@body = @NOTIFICATIONSBody,
		@subject = @NOTIFICATIONSSubject
	
	--END PROCESSING
	RETURN
END

/*---------------------------------------------------------------------------------------------------------------------
THERE ARE ORDERS TO PROCESS, BEGIN PROCESS
---------------------------------------------------------------------------------------------------------------------*/
--SEND NOTIFICATION EMAIL START
SELECT @NOTIFICATIONSSubject = N'SOP AUTO ALLOCATION ' + CONVERT(NVARCHAR(20),@SEQNUM) + N' HAS STARTED ' + CONVERT(NVARCHAR(20),@NOTIFICATIONSProcessTime)
SELECT @NOTIFICATIONSBody = N'SOP Auto Allocation has started on orders in VERIFY, VERIFY-SC, HOLD TO SHIP, and BACKORDER batches in site BDMN.  It is also running for batch TGTPTS BO in site CDPPARTS.'
EXEC msdb.dbo.sp_send_dbmail  
	@profile_name = @ProfileName,  
	@recipients = @Receipients,
	@body = @NOTIFICATIONSBody,
	@subject = @NOTIFICATIONSSubject

SELECT @ORDFROM = MIN(Cv_ID) FROM @ORDSTEMP
SELECT @ORDTO = MAX(Cv_ID) FROM @ORDSTEMP

--GET THE FIRST ORDER FROM THE TEMP TABLE
SELECT @Cv_ID = Cv_ID, @Cv_SOPTYPE = Cv_SOPTYPE, @Cv_SOPNUMBE = Cv_SOPNUMBE, @Cv_MSTRNUMB = Cv_MSTRNUMB, @Cv_DEX_ROW_ID = Cv_DEX_ROW_ID 
FROM @ORDSTEMP
WHERE Cv_ID = @ORDFROM

WHILE @ORDTO >= @ORDFROM
BEGIN
	SELECT @RECORDLOCKED = 0

	--IS RECORD IN ECONNECT TEMP OUT
	IF EXISTS(SELECT * FROM eConnectOutTemp WITH (NOLOCK) WHERE INDEX1 = @Cv_SOPNUMBE)
	BEGIN 
		SELECT @RECORDLOCKED = 1
	END
		
	--IS RECORD IN DEX_LOCK
	IF EXISTS(SELECT * FROM TEMPDB.DBO.DEX_LOCK WITH(NOLOCK) WHERE ROW_ID = @Cv_DEX_ROW_ID)
	BEGIN
		SELECT @RECORDLOCKED = 1
	END

	IF @RECORDLOCKED = 1
	BEGIN --RECORD LOCKED
		--CREATE LOG RECORD
		SELECT @DATETIME = GETDATE()
		SELECT @DTVC = CONVERT(VARCHAR(10), @DATETIME, 108)
		SELECT @EXECDATE = CONVERT(DATE,GETDATE())
		
		INSERT INTO bmiSOPAutoAllocationLog (USERID, SEQNUMBR, SOPTYPE, SOPNUMBE, LNITMSEQ, CMPNTSEQ, ITEMNMBR, bmiErrorMessage, DATE1, CUSTCLAS, SLPRSNID, BACHNUMB, QTYTBAOR, ATYALLOC, QUANTITY)
		VALUES (@I_vUserID, @SEQNUM, @Cv_SOPTYPE, @Cv_SOPNUMBE, 0, 0, '', 'Document locked, allocation skipped - ' + @DTVC, @EXECDATE, '', '', '', 0, 0, 0)
	END

	--IF RECORD NOT LOCKED, CONTINUE
	IF @RECORDLOCKED = 0
	BEGIN
		--CLEAR RECORDS FROM THE TEMP TABLE
		DELETE FROM @SOP10200TEMP
		--INSERT NEW RECORDS INTO TEMP TABLE FOR CURRENT ORDER
		INSERT INTO @SOP10200TEMP
		(
			I_vSOPTYPE, I_vSOPNUMBE, I_vCUSTNMBR, I_vDOCDATE, I_vDOCID, I_vLNITMSEQ, I_vCMPNTSEQ, I_vITEMNMBR, I_vLOCNCODE, 
			I_vUNITPRCE, I_vXTNDPRCE, I_vQUANTITY, I_vQTYCANCE, I_vATYALLOC, I_vTOTALQTY, I_vQTYTOINV, I_vReqShipDate, I_vACTLSHIP, I_vSHIPMTHD, I_vPRSTADCD, I_vShipToName, I_vCNTCPRSN, 
			I_vADDRESS1, I_vADDRESS2, I_vADDRESS3, I_vCITY, I_vSTATE, I_vZIPCODE, I_vCOUNTRY, I_vPHONE1, I_vPHONE2, I_vPHONE3,
			I_vFAXNUMBR, I_vPrint_Phone_NumberGB, I_vCMPITUOM, I_vITMTSHID, I_vMRKDNAMT_HDR, I_vMRKDNAMT_LINE, I_vMRKDNPCT, MSTRNUMB, RECORDLOCKED, DEX_ROW_ID
		)
		SELECT 
			L.SOPTYPE, L.SOPNUMBE, H.CUSTNMBR, H.DOCDATE, H.DOCID, L.LNITMSEQ, L.CMPNTSEQ, L.ITEMNMBR, L.LOCNCODE,
			L.UNITPRCE, l.XTNDPRCE, L.QTYTOINV, L.QTYCANCE, L.ATYALLOC, L.QUANTITY, L.QTYTOINV, L.ReqShipDate, L.ACTLSHIP, L.SHIPMTHD, L.PRSTADCD, L.ShipToName, L.CNTCPRSN,
			L.ADDRESS1, L.ADDRESS2, L.ADDRESS3, L.CITY, L.[STATE], L.ZIPCODE, L.COUNTRY, L.PHONE1, L.PHONE2, L.PHONE3,
			L.FAXNUMBR, L.Print_Phone_NumberGB, L.UOFM, L.ITMTSHID, H.MRKDNAMT, L.MRKDNAMT, L.MRKDNPCT, H.MSTRNUMB, 0, H.DEX_ROW_ID
		FROM 
			SOP10200 L WITH(NOLOCK)
			INNER JOIN SOP10100 H WITH(NOLOCK) ON H.SOPTYPE = L.SOPTYPE AND H.SOPNUMBE = L.SOPNUMBE
			INNER JOIN RM00101 C WITH(NOLOCK) ON C.CUSTNMBR = H.CUSTNMBR
			INNER JOIN IV00101 I WITH(NOLOCK) ON I.ITEMNMBR = L.ITEMNMBR
		WHERE 
			L.SOPTYPE = @Cv_SOPTYPE
			AND L.SOPNUMBE = @Cv_SOPNUMBE
			AND L.QTYCANCE < L.QUANTITY
			AND
			(
			(L.LOCNCODE = 'BDMN' AND H.BACHNUMB IN ('VERIFY','HOLD TO SHIP', 'BACKORDER', 'VERIFY-SC'))
			OR
			(L.LOCNCODE = 'CDPPARTS' AND H.BACHNUMB = 'TGTPTS BO')
			)
			AND L.PURCHSTAT IN (1,2)
			AND I.ITEMTYPE IN (1,3)
			AND H.VOIDSTTS = 0
		ORDER BY
			L.LNITMSEQ, L.CMPNTSEQ

		--SELECT * FROM @SOP10200TEMP

		SELECT @LINEFROM = MIN(I_vID) FROM @SOP10200TEMP WHERE I_vSOPTYPE = @Cv_SOPTYPE AND I_vSOPNUMBE = @Cv_SOPNUMBE
		SELECT @LINETO = MAX(I_vID) FROM @SOP10200TEMP WHERE I_vSOPTYPE = @Cv_SOPTYPE AND I_vSOPNUMBE = @Cv_SOPNUMBE
	
		--SELECT RECORDS FROM LINE TEMP TABLE
		SELECT
			@I_vID = I_vID, @I_vSOPTYPE = I_vSOPTYPE, @I_vSOPNUMBE = I_vSOPNUMBE, @I_vCUSTNMBR = I_vCUSTNMBR, 
			@I_vDOCDATE = I_vDOCDATE, @I_vDOCID = I_vDOCID, @I_vLNITMSEQ = I_vLNITMSEQ, 
			@I_vCMPNTSEQ = I_vCMPNTSEQ, @I_vITEMNMBR = I_vITEMNMBR, @I_vLOCNCODE = I_vLOCNCODE, 
			@I_vUNITPRCE = I_vUNITPRCE, @I_vXTNDPRCE = I_vXTNDPRCE, @I_vQUANTITY = I_vQUANTITY, 
			@I_vQTYCANCE = I_vQTYCANCE, @I_vATYALLOC = I_vATYALLOC, @I_vTOTALQTY = I_vTOTALQTY, 
			@I_vQTYTOINV = I_vQTYTOINV, @I_vReqShipDate = I_vReqShipDate, @I_vACTLSHIP = I_vACTLSHIP, 
			@I_vSHIPMTHD = I_vSHIPMTHD, @I_vPRSTADCD = I_vPRSTADCD, @I_vShipToName = I_vShipToName, 
			@I_vCNTCPRSN = I_vCNTCPRSN, @I_vADDRESS1 = I_vADDRESS1, @I_vADDRESS2 = I_vADDRESS2, 
			@I_vADDRESS3 = I_vADDRESS3, @I_vCITY = I_vCITY, @I_vSTATE = I_vSTATE, 
			@I_vZIPCODE = I_vZIPCODE, @I_vCOUNTRY = I_vCOUNTRY, @I_vPHONE1 = I_vPHONE1, 
			@I_vPHONE2 = I_vPHONE2, @I_vPHONE3 = I_vPHONE3, @I_vFAXNUMBR = I_vFAXNUMBR, 
			@I_vPrint_Phone_NumberGB = I_vPrint_Phone_NumberGB, @I_vCMPITUOM = I_vCMPITUOM, @I_vITMTSHID = I_vITMTSHID, 
			@I_vMRKDNAMT_HDR = I_vMRKDNAMT_HDR, @I_vMRKDNAMT = I_vMRKDNAMT_LINE, @I_vMRKDNPCT = I_vMRKDNPCT, 
			@RECORDLOCKED = RECORDLOCKED, @DEX_ROW_ID = DEX_ROW_ID
		FROM
			@SOP10200TEMP
		WHERE
			I_vID = @LINEFROM
			
		WHILE @LINETO >= @LINEFROM
		BEGIN
			--INITIALIZE ALL ERROR VARIABLES
			SELECT 
				@str = ''
				,@strlen = 0
				,@codestrt = 0
				,@codeend = 0
				,@code = '' 
				,@codedesc = ''
				,@codedescsingle = ''
				,@seqno = 0
				,@LOCKINT = 0
				,@iStatus = 0

			--CREATE LOG RECORD
			SELECT @DATETIME = GETDATE()
			SELECT @DTVC = CONVERT(VARCHAR(10), @DATETIME, 108)
			SELECT @EXECDATE = CONVERT(DATE,GETDATE())

			INSERT INTO bmiSOPAutoAllocationLog (USERID, SEQNUMBR, SOPTYPE, SOPNUMBE, LNITMSEQ, CMPNTSEQ, ITEMNMBR, bmiErrorMessage, DATE1, CUSTCLAS, SLPRSNID, BACHNUMB, QTYTBAOR, ATYALLOC, QUANTITY)
			VALUES (@I_vUserID, @SEQNUM, @I_vSOPTYPE, @I_vSOPNUMBE, @I_vLNITMSEQ, @I_vCMPNTSEQ, @I_vITEMNMBR, 'Record processed successfully  - ' + @DTVC, @EXECDATE, '', '', '', 0, 0, 0)

			--CHANGE ORDER - UPDATE ALLOCATION LOG FOR ADDITIONAL FIELDS FOR REPORTING
			UPDATE bmiSOPAutoAllocationLog SET CUSTCLAS = C.CUSTCLAS, BACHNUMB = S.BACHNUMB, SLPRSNID = S.SLPRSNID
			FROM bmiSOPAutoAllocationLog A INNER JOIN SOP10100 S WITH(NOLOCK) ON S.SOPTYPE = A.SOPTYPE AND S.SOPNUMBE = A.SOPNUMBE INNER JOIN RM00101 C ON C.CUSTNMBR = S.CUSTNMBR
			WHERE USERID = @I_vUserID AND SEQNUMBR = @SEQNUM

			SET @O_iErrorState = 0
			--IS ITEM A KIT COMPONENT
			IF @I_vCMPNTSEQ > 0
			BEGIN
				--KIT COMPONENT
				SET @I_vKitCompMan = 1 -- SET TO 1 A KIT COMPONENT, OTHERWISE 0
				SET @I_vQtyShrtOpt = 4 -- SET TO 3 FOR A KIT PARENT, SET TO 4 FOR REGULAR ITEMS AND KIT COMPONENTS
			
				EXEC taSopLineIvcInsertComponent
					@I_vSOPTYPE, @I_vSOPNUMBE, @I_vUSERDATE, @I_vLOCNCODE, @I_vLNITMSEQ, @I_vITEMNMBR, @I_vAutoAssignBin, @I_vITEMDESC, @I_vTOTALQTY, @I_vQTYTBAOR
					,@I_vQTYCANCE, @I_vQTYFULFI, @I_vQUOTEQTYTOINV, @I_vQTYONHND, @I_vQTYRTRND, @I_vQTYINUSE, @I_vQTYINSVC, @I_vQTYDMGED, @I_vCUSTNMBR, @I_vDOCID
					,@I_vUNITCOST, @I_vNONINVEN, @I_vAUTOALLOCATESERIAL, @I_vAUTOALLOCATELOT, @I_vCMPNTSEQ, @I_vCMPITUOM, @I_vCURNCYID, @I_vUpdateIfExists, @I_vRecreateDist
					,@I_vRequesterTrx, @I_vQtyShrtOpt, @I_vRECREATECOMM, @I_vUSRDEFND1, @I_vUSRDEFND2, @I_vUSRDEFND3, @I_vUSRDEFND4, @I_vUSRDEFND5, @O_iErrorState OUTPUT, @oErrString OUTPUT

				--SEE IF THIS DOCUMENT WAS PARTIALLY ALLOCATED, IF SO, HAVE TO DEAL WITH IT
				SET @NEWQTYTOINV = 0
				SET @QTYALLOCATE = 0
			
				SELECT @NEWQTYTOINV = QTYTOINV FROM SOP10200 WITH(NOLOCK) WHERE SOPTYPE = @I_vSOPTYPE AND SOPNUMBE = @I_vSOPNUMBE AND LNITMSEQ = @I_vLNITMSEQ AND CMPNTSEQ = @I_vCMPNTSEQ
				SELECT @QTYALLOCATE = ATYALLOC FROM SOP10200 WITH(NOLOCK) WHERE SOPTYPE = @I_vSOPTYPE AND SOPNUMBE = @I_vSOPNUMBE AND LNITMSEQ = @I_vLNITMSEQ AND CMPNTSEQ = @I_vCMPNTSEQ
			
				IF @NEWQTYTOINV <> @QTYALLOCATE
				--PARTIAL ALLOCATION
				BEGIN
					SET @QTYAVLIV = 0
					-- GO TO IV00102 AND SEE WHAT IS AVAILABLE
					SELECT @QTYAVLIV = 
						CASE	WHEN QTYONHND - ATYALLOC > 0 THEN QTYONHND - ATYALLOC
								WHEN @I_vQUANTITY < 0 THEN 0
								ELSE 0 END
								FROM IV00102 WITH(NOLOCK) WHERE ITEMNMBR = @I_vITEMNMBR AND LOCNCODE = @I_vLOCNCODE

					IF @QTYAVLIV > 0
					-- HAVE QUANTITY AVAILABLE TO ALLOCATE  IN IV00102
					BEGIN
						-- UPDATE QTY ALLOCATED ON KIT COMPONENTS IF THERE WAS A CHANGE
						UPDATE SOP10200 
						SET ATYALLOC = QTYTOINV --, EXTQTYAL = @QTYTOALLOCATE
						WHERE SOPTYPE = @I_vSOPTYPE AND SOPNUMBE = @I_vSOPNUMBE AND LNITMSEQ = @I_vLNITMSEQ AND CMPNTSEQ = @I_vCMPNTSEQ
					
						--MUST UPDATE IV00102 NOW
						UPDATE IV00102 SET ATYALLOC = ATYALLOC + @NEWQTYTOINV WHERE ITEMNMBR = @I_vITEMNMBR AND LOCNCODE = @I_vLOCNCODE
						UPDATE IV00102 SET ATYALLOC = ATYALLOC + @NEWQTYTOINV WHERE ITEMNMBR = @I_vITEMNMBR AND RCRDTYPE = 1 AND LOCNCODE = ''
					END	
				END

			END
			ELSE
			BEGIN

				--PARENT ITEM OR REGULAR ITEM
				SELECT @I_vKitCompMan = 0 -- SET TO 1 A KIT COMPONENT, OTHERWISE 0
			
				--COMPONENT SEQ IS 0, IS THIS A KIT TYPE ITEM (PARENT)
				IF (SELECT ITEMTYPE FROM IV00101 WITH(NOLOCK) WHERE ITEMNMBR = @I_vITEMNMBR) = 3
				BEGIN
					SET @I_vQtyShrtOpt = 3 -- SET TO 3 FOR A KIT PARENT
					SET @I_vQUANTITY = @I_vTOTALQTY
				END
				ELSE
				BEGIN
					SET @I_vQtyShrtOpt = 4 -- SET TO 4 FOR REGULAR ITEMS AND KIT COMPONENTS
					SET @I_vQUANTITY = @I_vTOTALQTY - @I_vQTYCANCE
				END

				/*05312017 - ZMAN - EVALUATE THE MARK DOWN PERCENT TO SET THE DEF PRICING FLAG*/
				IF @I_vMRKDNPCT <> 0
				BEGIN
					SELECT @I_vDEFPRICING = 1
					SELECT @I_vMRKDNAMT = null
				END
				ELSE
				BEGIN
					SELECT @I_vMRKDNPCT = null
				END

				/*MAKE ECONNECT CALL*/
				exec taSOPLineIvcInsert
					@I_vSOPTYPE,@I_vSOPNUMBE,@I_vCUSTNMBR,@I_vDOCDATE,@I_vUSERDATE,@I_vLOCNCODE,@I_vITEMNMBR,@I_vAutoAssignBin,@I_vUNITPRCE,@I_vXTNDPRCE,@I_vQUANTITY,
					@I_vMRKDNAMT,@I_vMRKDNPCT,@I_vCOMMNTID,@I_vCOMMENT_1,@I_vCOMMENT_2,@I_vCOMMENT_3,@I_vCOMMENT_4,@I_vUNITCOST,@I_vPRCLEVEL,@I_vITEMDESC,@I_vTAXAMNT,
					@I_vQTYONHND,@I_vQTYRTRND,@I_vQTYINUSE,@I_vQTYINSVC,@I_vQTYDMGED,@I_vNONINVEN,@I_vLNITMSEQ,@I_vDROPSHIP,@I_vQTYTBAOR,@I_vDOCID,@I_vSALSTERR,@I_vSLPRSNID,
					@I_vITMTSHID,@I_vIVITMTXB,@I_vTAXSCHID,@I_vPRSTADCD,@I_vShipToName,@I_vCNTCPRSN,@I_vADDRESS1,@I_vADDRESS2,@I_vADDRESS3,@I_vCITY,@I_vSTATE,@I_vZIPCODE,
					@I_vCOUNTRY,@I_vPHONE1,@I_vPHONE2,@I_vPHONE3,@I_vFAXNUMBR,@I_vPrint_Phone_NumberGB,@I_vEXCEPTIONALDEMAND,@I_vReqShipDate,@I_vFUFILDAT,@I_vACTLSHIP,
					@I_vSHIPMTHD,@I_vINVINDX,@I_vCSLSINDX,@I_vSLSINDX,@I_vMKDNINDX,@I_vRTNSINDX,@I_vINUSINDX,@I_vINSRINDX,@I_vDMGDINDX,@I_vAUTOALLOCATESERIAL,@I_vAUTOALLOCATELOT,
					@I_vGPSFOINTEGRATIONID,@I_vINTEGRATIONSOURCE,@I_vINTEGRATIONID,@I_vRequesterTrx,@I_vQTYCANCE,@I_vQTYFULFI,@I_vALLOCATE,@I_vUpdateIfExists,@I_vRecreateDist,
					@I_vQUOTEQTYTOINV,@I_vTOTALQTY,@I_vCMMTTEXT,@I_vKitCompMan,@I_vDEFPRICING,@I_vDEFEXTPRICE,@I_vCURNCYID,@I_vUOFM,@I_vIncludePromo,@I_vCKCreditLimit,
					@I_vQtyShrtOpt,@I_vRECREATETAXES,@I_vRECREATECOMM,@I_vUSRDEFND1,@I_vUSRDEFND2,@I_vUSRDEFND3,@I_vUSRDEFND4,@I_vUSRDEFND5,
					@O_iErrorState OUTPUT,@oErrString OUTPUT

			END

			-- ERROR TRAPPING FOR RECORD UPDATES
			IF @O_iErrorState <> 0
			BEGIN

				SET @ErrStringLines = LTRIM(RTRIM(@oErrString))		--Line error message	
				SET @str = (CASE WHEN @ErrStringLines = '' THEN '' ELSE @ErrStringLines + ' ' END)
				SET @strlen = LEN(@str)
				SET @codestrt = 1
				SET @codedesc = ''
			
				WHILE @strlen > 0 
				BEGIN
					SET @codeend = CASE WHEN PATINDEX('% %', @str) = 0 THEN @strlen ELSE PATINDEX('% %', @str) END
					SET @code = SUBSTRING(@str, @codestrt, @codeend)
					SET @codedesc = CASE @code WHEN '0' THEN @codedesc + '' ELSE @codedesc + 
						ISNULL((SELECT RTRIM(ErrorDesc) FROM DYNAMICS.dbo.taErrorCode WITH(NOLOCK)
						WHERE ErrorCode = CAST(@code AS INT)),'') END
					SET @str = ISNULL(SUBSTRING(@str, @codeend + 1, @strlen),'')
					SET @strlen = LEN(@str)
					SET @codedesc = CASE @strlen WHEN 0 THEN @codedesc + '.' 
						ELSE CASE WHEN @code = 0 THEN @codedesc ELSE @codedesc + '; 'END END
				END

				--UPDATE bmiSOPAutoAllocationLog with Errors
				SELECT @DATETIME = GETDATE()
				SELECT @DTVC = CONVERT(VARCHAR(10), @DATETIME, 108)
				SELECT @EXECDATE = CONVERT(DATE,GETDATE())
			
				UPDATE bmiSOPAutoAllocationLog
					SET bmiErrorMessage = (@codedesc + ' - ' +  @DTVC), DATE1 = @EXECDATE
					WHERE USERID = @I_vUserID 
					AND SEQNUMBR = @SEQNUM
					AND SOPTYPE = @I_vSOPTYPE 
					AND SOPNUMBE = @I_vSOPNUMBE
					AND LNITMSEQ = @I_vLNITMSEQ
					AND CMPNTSEQ = @I_vCMPNTSEQ 

			END -- ERROR TRAPPING ON RECORD UPDATES
			ELSE
			BEGIN
				--NO ERRORS ON RECORD UPDATES, UPDATE LOG FOR QUANTITES
				SELECT @DATETIME = GETDATE()
				SELECT @DTVC = CONVERT(VARCHAR(10), @DATETIME, 108)
				SELECT @EXECDATE = CONVERT(DATE,GETDATE())

				UPDATE bmiSOPAutoAllocationLog 
					SET ATYALLOC = L.ATYALLOC, QTYTBAOR = L.QTYTBAOR, QUANTITY = L.QUANTITY, bmiErrorMessage = 'Record processed successfully  - ' + @DTVC, DATE1 = @EXECDATE
					FROM bmiSOPAutoAllocationLog A INNER JOIN SOP10200 L WITH(NOLOCK) ON L.SOPTYPE = A.SOPTYPE AND L.SOPNUMBE = A.SOPNUMBE AND L.LNITMSEQ = A.LNITMSEQ AND L.CMPNTSEQ = A.CMPNTSEQ
					WHERE USERID = @I_vUserID AND SEQNUMBR = @SEQNUM
					AND a.SOPNUMBE = @I_vSOPNUMBE
					AND a.LNITMSEQ = @I_vLNITMSEQ
					AND a.CMPNTSEQ = @I_vCMPNTSEQ 
			END

			SELECT @LOCKINT = 0

			SELECT @LINEFROM = @LINEFROM + 1

			--GET NEXT RECORD IN LINE TEMP
			SELECT
				@I_vID = I_vID, @I_vSOPTYPE = I_vSOPTYPE, @I_vSOPNUMBE = I_vSOPNUMBE, @I_vCUSTNMBR = I_vCUSTNMBR, 
				@I_vDOCDATE = I_vDOCDATE, @I_vDOCID = I_vDOCID, @I_vLNITMSEQ = I_vLNITMSEQ, 
				@I_vCMPNTSEQ = I_vCMPNTSEQ, @I_vITEMNMBR = I_vITEMNMBR, @I_vLOCNCODE = I_vLOCNCODE, 
				@I_vUNITPRCE = I_vUNITPRCE, @I_vXTNDPRCE = I_vXTNDPRCE, @I_vQUANTITY = I_vQUANTITY, 
				@I_vQTYCANCE = I_vQTYCANCE, @I_vATYALLOC = I_vATYALLOC, @I_vTOTALQTY = I_vTOTALQTY, 
				@I_vQTYTOINV = I_vQTYTOINV, @I_vReqShipDate = I_vReqShipDate, @I_vACTLSHIP = I_vACTLSHIP, 
				@I_vSHIPMTHD = I_vSHIPMTHD, @I_vPRSTADCD = I_vPRSTADCD, @I_vShipToName = I_vShipToName, 
				@I_vCNTCPRSN = I_vCNTCPRSN, @I_vADDRESS1 = I_vADDRESS1, @I_vADDRESS2 = I_vADDRESS2, 
				@I_vADDRESS3 = I_vADDRESS3, @I_vCITY = I_vCITY, @I_vSTATE = I_vSTATE, 
				@I_vZIPCODE = I_vZIPCODE, @I_vCOUNTRY = I_vCOUNTRY, @I_vPHONE1 = I_vPHONE1, 
				@I_vPHONE2 = I_vPHONE2, @I_vPHONE3 = I_vPHONE3, @I_vFAXNUMBR = I_vFAXNUMBR, 
				@I_vPrint_Phone_NumberGB = I_vPrint_Phone_NumberGB, @I_vCMPITUOM = I_vCMPITUOM, @I_vITMTSHID = I_vITMTSHID, 
				@I_vMRKDNAMT_HDR = I_vMRKDNAMT_HDR, @I_vMRKDNAMT = I_vMRKDNAMT_LINE, @I_vMRKDNPCT = I_vMRKDNPCT, 
				@RECORDLOCKED = RECORDLOCKED, @DEX_ROW_ID = DEX_ROW_ID
			FROM
				@SOP10200TEMP
			WHERE
				I_vID = @LINEFROM

		END -- LINE TEMP

	END -- RECORD NOT LOCKED

	SELECT @ORDFROM = @ORDFROM + 1

	--GET NEXT ORDER FROM ORDER TEMP TABLE
	SELECT @Cv_ID = Cv_ID, @Cv_SOPTYPE = Cv_SOPTYPE, @Cv_SOPNUMBE = Cv_SOPNUMBE, @Cv_MSTRNUMB = Cv_MSTRNUMB, @Cv_DEX_ROW_ID = Cv_DEX_ROW_ID 
	FROM @ORDSTEMP 
	WHERE Cv_ID = @ORDFROM
	
END -- ORDER TEMP
	
--RUN SUBSCRIPTION REPORT
if (@@servername = 'BLUDOT-SQL1')
begin
	EXEC [msdb].dbo.sp_start_job N'4DA7D6AA-C4AA-4F44-BFC8-856FE28983C1' ; 
end

/*	Jeremy removed this since the 'results' email always follows anyway
--SEND NOTIFICATION EMAIL
SELECT @NOTIFICATIONSSubject = N'SOP AUTO ALLOCATION ' + CONVERT(NVARCHAR(20),@SEQNUM) + N' HAS COMPLETED ' + CONVERT(NVARCHAR(20),@NOTIFICATIONSProcessTime)
SELECT @NOTIFICATIONSBody = N'SOP Auto Allocation has completed.'
EXEC msdb.dbo.sp_send_dbmail  
	@profile_name = @ProfileName,  
	@recipients = @Receipients,
	@body = @NOTIFICATIONSBody,
	@subject = @NOTIFICATIONSSubject
*/

--SEND DB MAIL
DECLARE @Subject VARCHAR(MAX) = 'SOP Auto Allocation Completion Results'
DECLARE @DBName VARCHAR(MAX) = DB_NAME()
DECLARE @HTML NVARCHAR(MAX)
DECLARE @QUERYTXT NVARCHAR(MAX)

--CREATE HTML BODY 
declare @currentSeq int
select @currentSeq = (SELECT MAX(SEQNUMBR) FROM bmiSOPAutoAllocationLog)
declare @header nvarchar(max)
declare @rows nvarchar(max)

set @header = 
(CAST((
	select
		(select 'Document Number' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Batch' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Customer Class' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Salesperson' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Item Number' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Original Quantity' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Backordered' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Allocated' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
set @header = @header +
(CAST((
	select
		(select 'Status Message' as th for xml path(''), type)
	) AS NVARCHAR(MAX))
)
	
set @rows = 
(CAST ((
	select
		(select RTRIM(SOPNUMBE) as 'td' for xml path(''), type),
		(select RTRIM(BACHNUMB) as 'td' for xml path(''), type),
		(select RTRIM(CUSTCLAS) as 'td' for xml path(''), type),
		(select RTRIM(SLPRSNID) as 'td' for xml path(''), type),
		(select RTRIM(ITEMNMBR) as 'td' for xml path(''), type),
		(select QUANTITY as 'td' for xml path(''), type),
		(select QTYTBAOR as 'td' for xml path(''), type),
		(select ATYALLOC as 'td' for xml path(''), type),
		(select CONVERT(VARCHAR(300),bmiErrorMessage) as 'td' for xml path(''), type)
	from
		-- bmiSOPAutoAllocationLog WITH (NOLOCK) where SEQNUMBR = @currentSeq AND (QTYTBAOR > 0 or CONVERT(VARCHAR(29), bmiErrorMessage) <> 'Record Processed Successfully')
		bmiSOPAutoAllocationLog WITH (NOLOCK) where SEQNUMBR = @currentSeq AND CONVERT(VARCHAR(29), bmiErrorMessage) not like 'Record Processed Successfully%'  -- Jeremy changed what is included in the results per customer service request.
	for
		xml path('tr')
	) as nvarchar(max))
)
	
SET @html = 
	'<style type="text/css">
		#box-table
		{
		font-family: "Verdana";
		font-size: 12px;
		text-align: center;
		border-collapse: collapse;
		border-top: 7px solid #29293d;
		border-bottom: 7px solid #29293d;
		}
		#box-table th
		{
		font-size: 13px;
		font-weight: normal;
		background: #e0e0eb;
		border-right: 1px solid #29293d;
		border-left: 1px solid #29293d;
		border-bottom: 2px solid #29293d;
		color: black;
		padding: 4px;
		text-align: left;
		}
		#box-table td
		{
		border-right: 1px solid #29293d;
		border-left: 1px solid #29293d;
		border-bottom: 1px solid #29293d;
		color: #669;
		padding: 4px;
		text-align: left;
		}
		#box-table tr
		{
		text-align: left;
		}
		</style>'
+ '<table id="box-table">'
+ @header
+ @rows
+ N'</table>'

--SET @HTML = 'SOP Auto Allocation Results Attached'
	
--SELECT @QUERYTXT = '
--	SELECT 
--		RTRIM(SOPNUMBE) [Document Number],
--			RTRIM(BACHNUMB) [Batch],
--			RTRIM(CUSTCLAS) [Customer Class],
--			RTRIM(SLPRSNID) [Salesperson],
--			RTRIM(ITEMNMBR) [Item Number],
--			QUANTITY [Original Quantity],
--			QTYTBAOR [Backordered],
--			ATYALLOC [Allocated],
--			CONVERT(VARCHAR(300),bmiErrorMessage) [Status Message]
--	FROM bmiSOPAutoAllocationLog WITH(NOLOCK) WHERE SEQNUMBR = (SELECT MAX(SEQNUMBR) FROM bmiSOPAutoAllocationLog WITH(NOLOCK)) AND (QTYTBAOR <> 0 OR CONVERT(VARCHAR(29),bmiErrorMessage) <> ''Record Processed Successfully'')'

EXEC msdb.dbo.sp_send_dbmail  
	@profile_name = @ProfileName,  
	@recipients = @Receipients,
	@body = @HTML,
	@body_format = 'HTML', 
	@query = @QUERYTXT,
	@execute_query_database = @DBName,
	--@query_attachment_filename = 'SOP Auto Allocation Log.txt',
	@subject = @Subject  
	--@attach_query_result_as_file = 1

IF @@ERROR <> 0
BEGIN
	DECLARE @SEQNUMEMAIL INT
	SELECT @SEQNUMEMAIL = MAX(SEQNUMBR) FROM bmiSOPAutoAllocationLog WITH(NOLOCK)

	--CREATE LOG RECORD OF EMAIL PROBLEM
	INSERT INTO bmiSOPAutoAllocationLog (USERID, SEQNUMBR, SOPTYPE, SOPNUMBE, LNITMSEQ, CMPNTSEQ, ITEMNMBR, bmiErrorMessage, DATE1, CUSTCLAS, SLPRSNID, BACHNUMB, QTYTBAOR, ATYALLOC, QUANTITY)
	VALUES (@I_vUserID, @SEQNUMEMAIL, 0, '', 0, 0, '', 'Error Sending Email, Error Code ' + CONVERT(VARCHAR(MAX),@@ERROR) + ' - ' + @DTVC, @EXECDATE, '', '', '', 0, 0, 0)
END

if (SUSER_NAME() = 'NT SERVICE\SQLSERVERAGENT')
begin
	-- Enable triggers that were originally disabled.  This won't
	-- execute as a regular user.
	alter table IV00101 ENABLE TRIGGER sc_trgen_IV00101_update_ONGOING_GP_ITEMS_TO_WOW_BLU_DOT
	alter table SOP10200 ENABLE TRIGGER zDT_SOP10200U
	alter table SOP10100 ENABLE TRIGGER sc_trgen_SOP10100_update_ONGOING_GO_SOP_TO_WOW_UPDATE_BLU_DOT
	alter table SOP10200 ENABLE TRIGGER sc_trgen_SOP10200_update_ONGOING_GO_SOP_TO_WOW_UPDATE_BLU_DOT

	-- While the smartconnect sop10100 trigger was off, orders wouldn't be triggered to update to WOW
	-- so this would bridge that gap since this will force the trigger to fire for orders updated
	-- during that time.

	-- If this logic below is changed, the SmartConnect maps need to be reviewed as well as a SQL Agent job
	-- that runs every morning at 5am.

	update sop10100
	set modifdt = modifdt
	where sopnumbe in 
	(	select sopnumbe
		from sop10100 so
		where  soptype = 2 and	DEX_ROW_TS > dateadd(hour,-1,GETUTCDATE()) -- go back an hour (AA takes about 15 to 30 minutes on average)
		and 
		getdate() >= case when 
								isdate(
							SUBSTRING(bachnumb, 5, 2) + '/' + SUBSTRING(bachnumb, 7, 2) + '/'
							 + CASE
									WHEN SUBSTRING(bachnumb, 5, 2) <> '12' AND DATEPART(mm, GETDATE()) = '12'
									THEN CAST(DATEPART(yy, GETDATE()) + 1 AS VARCHAR(10))
									ELSE CAST(DATEPART(yy, GETDATE()) AS VARCHAR(10))
								  END
							) = 1
					then
							cast(SUBSTRING(bachnumb, 5, 2) + '/' + SUBSTRING(bachnumb, 7, 2) + '/'
							 + CASE
									WHEN SUBSTRING(bachnumb, 5, 2) <> '12' AND DATEPART(mm, GETDATE()) = '12'
									THEN CAST(DATEPART(yy, GETDATE()) + 1 AS VARCHAR(10))
									ELSE CAST(DATEPART(yy, GETDATE()) AS VARCHAR(10))
								  END as datetime) 
					when BACHNUMB = 'READY TO SHIP'
						then ReqShipDate
					ELSE
						GETDATE() + 3
					END
		AND	DocID in ('STOREORDER','ORDER','WEBORD','B2BORD','SERVICE','RETURN','STSERVICE','B2CORDER','B2BORDER',
						'B2CORD','WEBORDER','STOREORD','TRANSFERORD','TRADEORD','TRADEORDER','STORDER','SCORD',
						'SCORDER','SCSERVICE')
	)
end
	
GO

