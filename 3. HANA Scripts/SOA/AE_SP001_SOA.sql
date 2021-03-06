CREATE PROCEDURE "AE_SP001_SOA"
	(IN BPFrom nvarchar(30), 
	 IN BPTo nvarchar(30),
	 IN ToDate timestamp) 
AS 

Header1 nvarchar(20);
Header2 nvarchar(20);
Header3 nvarchar(20);
Header4 nvarchar(20);
Header5 nvarchar(20);
LocCurr char(3);
ComName varchar(300);
CoRegNo nvarchar(40);
GSTRegNo nvarchar(40);
EMAIL nvarchar(300);
CompName nvarchar(5000);
Currency varchar(10);
COMPNYADDR nvarchar(100);
PHONE1 nvarchar(100);
FAX nvarchar(100);

BEGIN

Header1 := '0 to 30 Days';
Header2 := '31 to 60 Days';
Header3 := '61 to 90 Days';
Header4 := '91 to 120 Days';
Header5 := 'Above 120 Days';

SELECT
	 (SELECT
	 TOP 1 ifnull("PrintHeadr",
	 "CompnyName") 
	FROM"OADM") 
INTO ComName 
FROM DUMMY;

SELECT (SELECT "MainCurncy" FROM"OADM") INTO LocCurr FROM DUMMY;
SELECT (SELECT "TaxIdNum" FROM"OADM") INTO CoRegNo FROM DUMMY;
SELECT (SELECT "TaxIdNum2" FROM"OADM") INTO GSTRegNo FROM DUMMY;
SELECT (SELECT "Phone1" FROM"OADM") INTO PHONE1 FROM DUMMY;
SELECT (SELECT "Fax" FROM"OADM") INTO FAX FROM DUMMY;
SELECT (SELECT "CompnyAddr" FROM"OADM") INTO COMPNYADDR FROM DUMMY;
SELECT (SELECT "E_Mail" FROM"OADM") INTO EMAIL FROM DUMMY;
   
--SELECT "MainCurncy", "TaxIdNum", "TaxIdNum2","COMPNYADDR","PHONE1","FAX"
--INTO LocCurr, CoRegNo, GSTRegNo,COMPNYADDR,PHONE1,FAX FROM DUMMY;


/* Get Reconciliation Sum base on BP */
CREATE COLUMN TABLE RECON  ("TransId" INTEGER,"TransRowId" INTEGER,"IsCredit" NVARCHAR(1),
	"ReconDate" TIMESTAMP,"ReconSum" DECIMAL CS_DECIMAL_FLOAT,"ReconSumFC" DECIMAL CS_DECIMAL_FLOAT);
	
INSERT INTO RECON (SELECT ITR1."TransId", ITR1."TransRowId", "IsCredit", MAX("ReconDate") AS "ReconDate", 
	SUM(IFNULL(ITR1."ReconSum" * CASE WHEN "IsCredit" = 'C' THEN 1 ELSE -1 END, 0)) AS "ReconSum", 
	SUM(IFNULL(ITR1."ReconSumFC" * CASE WHEN "IsCredit" = 'C' THEN 1 ELSE -1 END, 0)) AS "ReconSumFC" 
FROM "OITR" OITR 
INNER JOIN "ITR1" ITR1 ON OITR."ReconNum" = ITR1."ReconNum" 
WHERE OITR."Canceled" <> 'Y' AND 
LEFT(TO_DATS(OITR."ReconDate"),6) < LEFT(TO_DATS(:ToDate),6)
--OITR."ReconDate" <= :ToDate 

AND OITR."ReconType" NOT IN (7) AND (IFNULL(:BPFrom, '') = '' OR ITR1."ShortName" >= :BPFrom) 
AND (IFNULL(:BPTo, '') = '' OR ITR1."ShortName" <= :BPTo) 
GROUP BY ITR1."TransId", ITR1."TransRowId", "IsCredit");

 CREATE COLUMN TABLE "SOA" ("TransId" INTEGER,"DocNum" INTEGER,"CustomerRef" NVARCHAR(100),
    "SERIES" NVARCHAR(10),"TransType" INTEGER,"DocEntry" INTEGER,"CardCode" NVARCHAR(20),"TempCardCode" NVARCHAR(20),"CardName" NVARCHAR(100),
    "PaymentTerms" NVARCHAR(50),"DueDate" TIMESTAMP,"DocDate" TIMESTAMP,"Debit" DECIMAL CS_DECIMAL_FLOAT,"DebitFC" DECIMAL CS_DECIMAL_FLOAT,
	"Credit" DECIMAL CS_DECIMAL_FLOAT,"CreditFC" DECIMAL CS_DECIMAL_FLOAT,"DocCurrency" NVARCHAR(5),
	"TransAmt" DECIMAL CS_DECIMAL_FLOAT,"CreditLimit" DECIMAL CS_DECIMAL_FLOAT,"currency" NVARCHAR(3),
	"Telephone" NVARCHAR(40),"Fax" NVARCHAR(40),"ContactPerson" NVARCHAR(180),"SlpName" NVARCHAR(180),"Canceled" NVARCHAR(1),
	"BillToAddress" NVARCHAR(100),"BillToStreet" NVARCHAR(100),"BillToBlock" NVARCHAR(100),"BillToCity" NVARCHAR(100),
	"BillBuilding" NVARCHAR(100),"BillToCountry" NVARCHAR(100),"BillToZipCode" NVARCHAR(100),"Balance" DECIMAL CS_DECIMAL_FLOAT,"Free_Text" NCLOB,
	"SOAMethod" NVARCHAR(10));
 
 CREATE COLUMN TABLE "SOA2" ("TransId" INTEGER,"DocNum" INTEGER,"CustomerRef" NVARCHAR(100),
    "SERIES" NVARCHAR(10),"TransType" INTEGER,"DocEntry" INTEGER,"CardCode" NVARCHAR(20),"TempCardCode" NVARCHAR(20),"CardName" NVARCHAR(100),
    "PaymentTerms" NVARCHAR(50),"DueDate" TIMESTAMP,"DocDate" TIMESTAMP,"Debit" DECIMAL CS_DECIMAL_FLOAT,"DebitFC" DECIMAL CS_DECIMAL_FLOAT,
	"Credit" DECIMAL CS_DECIMAL_FLOAT,"CreditFC" DECIMAL CS_DECIMAL_FLOAT,"DocCurrency" NVARCHAR(5),
	"TransAmt" DECIMAL CS_DECIMAL_FLOAT,"CreditLimit" DECIMAL CS_DECIMAL_FLOAT,"currency" NVARCHAR(3),
	"Telephone" NVARCHAR(40),"Fax" NVARCHAR(40),"ContactPerson" NVARCHAR(180),"SlpName" NVARCHAR(180),"Canceled" NVARCHAR(1),
	"BillToAddress" NVARCHAR(100),"BillToStreet" NVARCHAR(100),"BillToBlock" NVARCHAR(100),"BillToCity" NVARCHAR(100),
	"BillBuilding" NVARCHAR(100),"BillToCountry" NVARCHAR(100),"BillToZipCode" NVARCHAR(100),"Balance" DECIMAL CS_DECIMAL_FLOAT,"Free_Text" NCLOB,
	"SOAMethod" NVARCHAR(10));
 
  CREATE COLUMN TABLE "SOA3" ("CardCode" NVARCHAR(20),"CardName" NVARCHAR(100)
  		--,"Canceled" NVARCHAR(1),"TransAmt" DECIMAL CS_DECIMAL_FLOAT
  		,"Balance" DECIMAL CS_DECIMAL_FLOAT,"Free_Text" NCLOB
  		,"SOAMethod" NVARCHAR(10));
 

INSERT INTO SOA2 (SELECT T0."TransId",
	CASE WHEN T0."TransType" = 13 THEN T3."DocNum"
        WHEN T0."TransType" = 14 THEN T4."DocNum"
        WHEN T0."TransType" = 24 THEN T7."DocNum"
		WHEN T0."TransType" = 30 THEN OJDT."Number"
        ELSE NULL END AS "DocNum",  
    CASE WHEN T0."TransType" = 13 THEN IFNULL(T3."NumAtCard",T0."Ref2")
        WHEN T0."TransType" = 14 THEN T4."NumAtCard" 
        WHEN T0."TransType" = 24 THEN T7."CounterRef"
        ELSE T0."Ref2"
   END  AS "CustomerRef",
	CASE WHEN T0."TransType" = 13 THEN 'INV'
        WHEN T0."TransType" = 14 THEN 'CN'
        WHEN T0."TransType" = 24 THEN 'PYMT'
		WHEN T0."TransType" = 30 THEN 'JE'
        ELSE ''
  	 END AS "SERIES"
  ,T0."TransType" AS TransType
  ,T0."CreatedBy" AS DocEntry
  ,T0."ShortName" AS CardCode
  ,T0."ShortName" AS TempCardCode
  ,COALESCE(T1."CardName", T3."CardName", T4."CardName")AS CardName 
  ,OCTG."PymntGroup" AS PaymentTerms
  ,T0."DueDate" AS DueDate
  ,T0."RefDate" AS DocDate 
  ,IFNULL(T0."Debit",0) AS Debit
  ,IFNULL(T0."FCDebit",0) AS DebitFC
  ,IFNULL(T0."Credit",0) AS Credit
  ,IFNULL(T0."FCCredit",0) AS CreditFC
  ,CASE WHEN T0."RevSource"='F' OR IFNULL(T0."FCDebit",0)=0 AND  IFNULL(T0."FCCredit",0)=0 THEN :LocCurr ELSE IFNULL(UPPER(T0."FCCurrency"),:LocCurr) END AS DocCurrency
  -- Modified
  ,CASE WHEN LEFT(TO_DATS(T0."RefDate"),6)< LEFT(TO_DATS(:ToDate),6) THEN
		CASE WHEN IFNULL(T0."FCDebit",0)=0 AND  IFNULL(T0."FCCredit",0)=0
        THEN IFNULL(T0."Debit",0) - IFNULL(T0."Credit",0) + IFNULL(RECON."ReconSum",0)
        ELSE IFNULL(T0."FCDebit",0) - IFNULL(T0."FCCredit",0) + IFNULL(RECON."ReconSumFC",0) END
    ELSE
    	CASE WHEN IFNULL(T0."FCDebit",0)=0 AND  IFNULL(T0."FCCredit",0)=0
        THEN IFNULL(T0."Debit",0) - IFNULL(T0."Credit",0) 
        ELSE IFNULL(T0."FCDebit",0) - IFNULL(T0."FCCredit",0)  END 
   END AS TransAmt
   
  ,T1."CreditLine" AS CreditLimit
  ,T1."Currency" AS currency
  ,COALESCE(T1."Phone1",T1."Phone2",'') AS Telephone
  ,IFNULL(T1."Fax",'') AS Fax
  ,IFNULL(T1."CntctPrsn",'') AS ContactPerson
  ,IFNULL(T2."SlpName",'') AS SalesPerson
  ,CASE WHEN T0."TransType" = 13 THEN T3."CANCELED"
		WHEN T0."TransType" = 14 THEN T4."CANCELED"
		WHEN T0."TransType" = 24 THEN T7."Canceled"
		WHEN T0."TransType" = 18 THEN OPCH."CANCELED"
		WHEN T0."TransType" = 19 THEN ORPC."CANCELED"
		WHEN T0."TransType" = 46 THEN OVPM."Canceled"
		WHEN T0."TransType" in (30,-5) THEN 'N'
   END AS Canceled,
   IFNULL(X2."Address",'') AS BillToAddress
  ,IFNULL(X2."Street",'') AS BillToStreet--,IFNULL(X2."StreetNo",'') AS BillToStreetNo
  ,IFNULL(X2."Block",'') AS BillToBlock
  ,IFNULL(X2."City",'') AS BillToCity
  ,IFNULL(X2."Building",'') AS BillBuilding
  ,IFNULL((SELECT "Name" FROM"OCRY" WHERE "Code" = X2."Country"),'') AS BillToCountry
  ,IFNULL(X2."ZipCode",'') AS BillToZipCode  
   ,C1."Balance",C1."E_Mail" "Free_Text",C1."U_SOAMethod" "SOAMethod"
FROM "JDT1" T0
INNER JOIN "OJDT" OJDT ON T0."TransId"=OJDT."TransId"
INNER JOIN ("OCRD" T1 INNER JOIN"OSLP" T2 ON T1."SlpCode" = T2."SlpCode" 
	LEFT OUTER JOIN "CRD1" X2 ON T1."CardCode" = X2."CardCode" AND X2."Address" = T1."BillToDef" 
	AND X2."AdresType" = 'B')  ON T0."ShortName" = T1."CardCode" and T1."CardType" = 'C'
LEFT OUTER JOIN ("OINV" T3 INNER JOIN"OSLP" T5 ON T3."SlpCode" = T5."SlpCode" 
	LEFT OUTER JOIN "NNM1" NNM1_A ON NNM1_A."ObjectCode" = 13 AND T3."Series" = NNM1_A."Series") 
	ON T0."TransType" = 13 	AND T0."CreatedBy" = T3."DocEntry"    
LEFT OUTER JOIN ("ORIN" T4  INNER JOIN"OSLP" T6 ON T4."SlpCode" = T6."SlpCode" 
	LEFT OUTER JOIN "NNM1" nnm1_b on nnm1_b."ObjectCode" = 14 and T4."Series" = nnm1_b."Series") 
		ON T0."TransType" = 14 AND T0."CreatedBy" = T4."DocEntry"             
LEFT OUTER JOIN "ORCT" T7 ON T0."CreatedBy" = T7."DocEntry" AND T0."TransType" = 24 
LEFT OUTER JOIN "OPCH" ON T0."CreatedBy" = OPCH."DocEntry" AND T0."TransType" = 18
LEFT OUTER JOIN "ORPC" ON T0."CreatedBy" = ORPC."DocEntry" AND T0."TransType" = 19
LEFT OUTER JOIN "OVPM" ON T0."CreatedBy" = OVPM."DocEntry" AND T0."TransType" = 46
LEFT OUTER JOIN "OCRD" C1 ON C1."CardCode"=T0."ShortName" 
LEFT OUTER JOIN "OCTG" OCTG ON OCTG."GroupNum"=C1."GroupNum"
LEFT OUTER JOIN RECON ON RECON."TransId"=T0."TransId" AND T0."Line_ID"=RECON."TransRowId"
WHERE  (T0."ShortName" >= :BPFrom AND T0."ShortName" <= :BPTo) AND T0."RefDate" <= :ToDate
		AND T1."CardType"='C' and T0."TransType" in (13,14,24,18,19,30,46,-5));

--SELECT * FROMSOA WHERE "TransAmt"<>0 and "Canceled"<>'Y';

--DELETE FROM SOA WHERE "TransAmt"=0 AND MONTH("DocDate")<= ADD_MONTHS(:ToDate,-1);

--DELETE FROM SOA WHERE "TransAmt"=0 AND LEFT(TO_DATS("DocDate"),6)< LEFT(TO_DATS(:ToDate),6);


--ROUND IF DOC Currency is JPY


INSERT INTO SOA
(
	SELECT "TransId","DocNum","CustomerRef","SERIES","TransType","DocEntry","CardCode","TempCardCode","CardName","PaymentTerms"
  	,"DueDate","DocDate","Debit","DebitFC","Credit","CreditFC","DocCurrency"
  
  	,CASE WHEN "DocCurrency" = 'JPY' AND "TransType" IN (13,14) THEN ROUND("TransAmt",0) ELSE "TransAmt" END AS "TransAmt"
   
  	,"CreditLimit","currency","Telephone","Fax","ContactPerson","SlpName","Canceled","BillToAddress","BillToStreet","BillToBlock","BillToCity"
  	,"BillBuilding","BillToCountry","BillToZipCode","Balance","Free_Text","SOAMethod"
	FROM SOA2
);

INSERT INTO SOA3
(
	SELECT "CardCode","CardName"
	,"Balance","Free_Text","SOAMethod"
	FROM SOA  WHERE "TransAmt"<>0 and "Canceled"<>'Y' and "Canceled" <> 'C' 
);
SELECT Distinct "CardCode","CardName"
	,"Balance",TO_ALPHANUM("Free_Text") "Free_Text","SOAMethod" FROM SOA3 ;

  
DROP TABLE RECON;
DROP TABLE SOA;
DROP TABLE SOA2;
DROP TABLE SOA3;


END;