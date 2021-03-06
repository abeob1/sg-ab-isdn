--DROP PROCEDURE "AE_CUSTOMER_AGING_ISDN" 

--CALL "AE_CUSTOMER_AGING_ISDN" ('D','2017-01-03','C20000','C777777','I')

CREATE PROCEDURE "AE_CUSTOMER_AGING_ISDN" 
(
	IN Intervals Varchar(20),
	IN AgingDate TimeStamp,
	IN CustomerCodeFrom varchar(20),
	IN CustomerCodeTo varchar(20),
	IN InValid Varchar(20)
	 ) 
AS 

AgeColumn0 INTEGER;
AgeColumn1 INTEGER;
AgeColumn2 INTEGER;
AgeColumn3 INTEGER;
AgeColumn4 INTEGER;
AgeColumn5 INTEGER;
AgeColumn6 INTEGER;
AgeColumn7 INTEGER;
AgeColumn8 INTEGER;
AgeColumn9 INTEGER;
AgeColumn10 INTEGER;
HEADER1 varchar(50);
HEADER2 varchar(50);
HEADER3 varchar(50);
HEADER4 varchar(50);
HEADER5 varchar(50);
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
IF :Intervals = 'D' 
THEN
AgeColumn0 := 0;
AgeColumn1 := 90;
AgeColumn2 := 91;
AgeColumn3 := 180;
AgeColumn4 := 181;
AgeColumn5 := 270;
AgeColumn6 := 271;
AgeColumn7 := 300;
AgeColumn8 := 301;
HEADER1 := 'Current' ;
HEADER2 :='91-180 Days';
HEADER3 :='181-270 Days';
HEADER4 :='271-300 Days';
HEADER5 :='Above 300 Days';

ELSEIF :Intervals = 'M' 
THEN
AgeColumn0 := 0;
AgeColumn1 := 30;
AgeColumn2 := 31;
AgeColumn3 := 60;
AgeColumn4 := 61;
AgeColumn5 := 90;
AgeColumn6 := 91;
AgeColumn7 := 120;
AgeColumn8 := 121;
HEADER1 := 'Current' ;
HEADER2 :='2 Months';
HEADER3 :='3 Months';
HEADER4 :='4 Months';
HEADER5 :='5 Months and above';

ELSEIF :Intervals = 'Y' 
THEN
AgeColumn0 := 0;
AgeColumn1 := 365;
AgeColumn2 := 366;
AgeColumn3 := 730;
AgeColumn4 := 731;
AgeColumn5 := 1095;
AgeColumn6 := 1096;
AgeColumn7 := 1460;
AgeColumn8 := 1461;
HEADER1 := 'Current';
HEADER2 :='2 Years';
HEADER3 :='3 Years';
HEADER4 :='4 Years';
HEADER5 :='5 Years and above';
end if;

SELECT (SELECT TOP 1 ifnull("PrintHeadr", "CompnyName") FROM "OADM") INTO ComName FROM DUMMY;
SELECT (SELECT "MainCurncy" FROM"OADM") INTO Currency FROM DUMMY;
SELECT (SELECT "TaxIdNum" FROM"OADM") INTO CoRegNo FROM DUMMY;
SELECT (SELECT "TaxIdNum2" FROM"OADM") INTO GSTRegNo FROM DUMMY;
SELECT (SELECT "Phone1" FROM"OADM") INTO PHONE1 FROM DUMMY;
SELECT (SELECT "Fax" FROM"OADM") INTO FAX FROM DUMMY;
SELECT (SELECT "CompnyAddr" FROM"OADM") INTO COMPNYADDR FROM DUMMY;
SELECT (SELECT "E_Mail" FROM"OADM") INTO EMAIL FROM DUMMY;


Select A.*, B."AcctName",
 :ComName as CompanyName
       ,:CompName AS COMPNAME
	   ,:COMPNYADDR  AS COMPADD
	   ,:PHONE1 AS COMPTEL
	   ,:FAX AS COMPFAX
	   ,:CoRegNo AS COMPREGNO
	  , :GSTRegNo as GSTRegNo
	  ,:EMAIL As E_Mail
	  ,(select Top 1 "LogoImage" from"OADP") as LogoImage
 from (
		select T."CardCode", sum(T."FutureRemit") as "FutureRemit", sum(T."OverDue") as "OverDue", sum(T."AgeColumn1") 
		, sum(T."AgeColumn2") , sum("AgeColumn3"), sum("AgeColumn4"), sum("AgeColumn5")
		, max(S."CreditLine")
		, max(S."U_IncoTerm")
		, max(S."validFor") as "InValid"
		, max(S."Balance") as "Balance"
		, max(S."CardName") as "CardName"
		, max(T."PymntGroup")
		, max(V."AcctCode") as "AcctCode"
		, max(U."SlpName") as "SalesEmpName"
		, max(T."HEADER1") as "HEADER1"
		, max(T."HEADER2") as "HEADER2"
		, max(T."HEADER3") as "HEADER3"
		, max(T."HEADER4") as "HEADER4"
		, max(T."HEADER5") as "HEADER5"
		 from (
		SELECT M."CardCode", M."CardName", IFNULL(M."InvNo", '0') AS "InvNo", M."TaxDate", 
		M."DueDate", M."BalDueAmount",
		 DAYS_BETWEEN(m."DueDate", :AgingDate) "Aging",
		CASE WHEN M."DueDate" > :AgingDate THEN M."BalDueAmount" ELSE 0 END AS "FutureRemit", 
		CASE WHEN M."DueDate" <= :AgingDate THEN M."BalDueAmount" ELSE 0 END AS "OverDue"
	, Case when DAYS_BETWEEN(m."DueDate",:AgingDate) between :AgeColumn0 And :AgeColumn1 Then m."BalDueAmount" Else 0 End as "AgeColumn1"
	, Case when DAYS_BETWEEN(m."DueDate",:AgingDate) between :AgeColumn2 And :AgeColumn3 Then m."BalDueAmount" Else 0 End as "AgeColumn2"
	, Case when DAYS_BETWEEN(m."DueDate",:AgingDate) between :AgeColumn4 And :AgeColumn5 Then m."BalDueAmount" Else 0 End as "AgeColumn3"
	, Case when DAYS_BETWEEN(m."DueDate",:AgingDate) between :AgeColumn6 And :AgeColumn7 Then m."BalDueAmount" Else 0 End as "AgeColumn4"	
	, Case when DAYS_BETWEEN(m."DueDate",:AgingDate) >:AgeColumn8 Then m."BalDueAmount" Else 0 End as "AgeColumn5"
	 , PyC."PymntGroup"
	, :HEADER1 "HEADER1"
	, :HEADER2 "HEADER2"
	, :HEADER3 "HEADER3"
	, :HEADER4 "HEADER4"
	, :HEADER5 "HEADER5"
		   FROM 
		   (SELECT T0."TransId" AS "TransId", T0."Line_ID" AS "Line_ID", MAX(T0."Account") AS "Account",
		   MAX(T0."ShortName") AS "ShortName", MAX(T0."TransType") AS "TransType", MAX(T0."CreatedBy") AS "CreatedBy", MAX(T0."BaseRef") AS "BaseRef", 
		   MAX(T0."SourceLine") AS "SourceLine", MAX(T0."RefDate") AS "RefDate", MAX(T0."DueDate") AS "DueDate", MAX(T0."TaxDate") AS "TaxDate",
		    (MAX(T0."BalDueCred") + SUM(T1."ReconSum")) * -1 AS "BalDueAmount", MAX(T0."BalFcCred") + SUM(T1."ReconSumFC") AS "BalDueFC", 
		    MAX(T0."BalScCred") + SUM(T1."ReconSumSC") AS "BalDueSC", MAX(T0."LineMemo") AS "LineMemo", MAX(T3."FolioPref") AS "FolioPref", 
		    MAX(T3."FolioNum") AS "FolioNum", MAX(T0."Indicator") AS "Indicator", MAX(T4."CardName") AS "CardName", MAX(T5."CardCode") AS "CardCode", 
		    MAX(T5."CardName") AS "CardName1", MAX(T4."Balance") AS "BpBalance", MAX(T5."NumAtCard") AS "NumAtCard", MAX(T5."SlpCode") AS "SlpCode",
		     MAX(T0."Project") AS "Project", MAX(T0."Debit") + MAX(T0."Credit") AS "BalLC", MAX(T0."FCDebit") + MAX(T0."FCCredit") AS "BalFC", 
		     MAX(T0."SYSDeb") + MAX(T0."SYSCred") AS "BalSC", MAX(T4."PymCode") AS "PymCode", MAX(T5."BlockDunn") AS "BlockDunn",
		      MAX(T5."DunnLevel") AS "DunnLevel", MAX(T5."TransType") AS "TransType1", MAX(T5."IsSales") AS "IsSales", MAX(T4."Currency") AS "Currency", 
		      MAX(T0."FCCurrency") AS "FCCurr", MAX(T6."SlpName") AS "SlpName", MAX(T4."DunTerm") AS "DunTerm", MAX(T0."DunnLevel") AS "DunnLevel1",
		       MAX(T7."DocNum") AS "InvNo" 
		       FROM "JDT1" T0 INNER JOIN "ITR1" T1 ON T1."TransId" = T0."TransId" AND T1."TransRowId" = T0."Line_ID" 
		       INNER JOIN "OITR" T2 ON T2."ReconNum" = T1."ReconNum" 
		       INNER JOIN "OJDT" T3 ON T3."TransId" = T0."TransId" 
		       INNER JOIN "OCRD" T4 ON T4."CardCode" = T0."ShortName" 
		       INNER JOIN "OINV" T7 ON T7."DocNum" = T0."BaseRef" 
		       LEFT OUTER JOIN "B1_JournalTransSourceView" T5 ON T5."ObjType" = T0."TransType" AND T5."DocEntry" = T0."CreatedBy"
		        AND (T5."TransType" <> n'I' OR (T5."TransType" = n'I' AND T5."InstlmntID" = T0."SourceLine")) 
		        LEFT OUTER JOIN "OSLP" T6 ON T6."SlpCode" = T5."SlpCode" OR (T6."SlpName" = n'-No Sales Employee-' 
		        AND (T0."TransType" = n'30' OR T0."TransType" = n'321' OR T0."TransType" = n'-5' 
		        OR T0."TransType" = n'-2' OR T0."TransType" = n'-3' OR T0."TransType" = n'-4')) 
		        WHERE T0."RefDate" <= (:AgingDate) AND T0."RefDate" <= (:AgingDate) AND T4."CardType" = ('C') AND T2."ReconDate" > (:AgingDate) 
		        AND T1."IsCredit" = ('C') GROUP BY T0."TransId",
		         T0."Line_ID" HAVING MAX(T0."BalFcCred") <> -SUM(T1."ReconSumFC") OR MAX(T0."BalDueCred") <> -SUM(T1."ReconSum")
		          UNION ALL 
		          SELECT T0."TransId", T0."Line_ID", MAX(T0."Account") AS "Account", MAX(T0."ShortName") AS "ShortName", MAX(T0."TransType") AS "TransType", 
		          MAX(T0."CreatedBy"), MAX(T0."BaseRef"), MAX(T0."SourceLine"), MAX(T0."RefDate"), MAX(T0."DueDate"), MAX(T0."TaxDate"),
		           (-MAX(T0."BalDueDeb") - SUM(T1."ReconSum")) * -1, -MAX(T0."BalFcDeb") - SUM(T1."ReconSumFC"), -MAX(T0."BalScDeb") - SUM(T1."ReconSumSC"),
		            MAX(T0."LineMemo"),  MAX(T3."FolioPref"), MAX(T3."FolioNum"), MAX(T0."Indicator"), MAX(T4."CardName"), MAX(T5."CardCode"), MAX(T5."CardName"), 
		            MAX(T4."Balance"), MAX(T5."NumAtCard"), MAX(T5."SlpCode"), MAX(T0."Project"), MAX(T0."Debit") + MAX(T0."Credit"), 
		            MAX(T0."FCDebit") + MAX(T0."FCCredit"), MAX(T0."SYSDeb") + MAX(T0."SYSCred"), MAX(T4."PymCode"), MAX(T5."BlockDunn"), 
		            MAX(T5."DunnLevel"), MAX(T5."TransType"), MAX(T5."IsSales"), MAX(T4."Currency"), MAX(T0."FCCurrency"), MAX(T6."SlpName"),
		             MAX(T4."DunTerm"), MAX(T0."DunnLevel"), MAX(T7."DocNum") AS "Inv No" 
		             FROM "JDT1" T0 
		             INNER JOIN "ITR1" T1 ON T1."TransId" = T0."TransId" AND T1."TransRowId" = T0."Line_ID"
		              INNER JOIN "OITR" T2 ON T2."ReconNum" = T1."ReconNum" 
		              INNER JOIN "OJDT" T3 ON T3."TransId" = T0."TransId" 
		              INNER JOIN "OCRD" T4 ON T4."CardCode" = T0."ShortName" 
		              INNER JOIN "OINV" T7 ON T7."DocNum" = T0."BaseRef" 
		              LEFT OUTER JOIN "B1_JournalTransSourceView" T5 ON T5."ObjType" = T0."TransType" AND T5."DocEntry" = T0."CreatedBy" 
		              AND (T5."TransType" <> n'I' OR (T5."TransType" = n'I' AND T5."InstlmntID" = T0."SourceLine")) 
		              LEFT OUTER JOIN "OSLP" T6 ON T6."SlpCode" = T5."SlpCode" OR (T6."SlpName" = n'-No Sales Employee-'
		               AND (T0."TransType" = n'30' OR T0."TransType" = n'321' OR T0."TransType" = n'-5' OR T0."TransType" = n'-2' 
		               OR T0."TransType" = n'-3' OR T0."TransType" = n'-4')) 
		               WHERE T0."RefDate" <= (:AgingDate) AND T0."RefDate" <= (:AgingDate) 
		               AND T4."CardType" = ('C') AND T2."ReconDate" > (:AgingDate) AND T1."IsCredit" = ('D') 
		               GROUP BY T0."TransId", T0."Line_ID" HAVING MAX(T0."BalFcDeb") <> -SUM(T1."ReconSumFC") 
		               OR MAX(T0."BalDueDeb") <> -SUM(T1."ReconSum") UNION ALL SELECT T0."TransId", T0."Line_ID",
		                MAX(T0."Account"), MAX(T0."ShortName"), MAX(T0."TransType"), MAX(T0."CreatedBy"), MAX(T0."BaseRef"),
		                 MAX(T0."SourceLine"), MAX(T0."RefDate"), MAX(T0."DueDate"), MAX(T0."TaxDate"), 
		                 (MAX(T0."BalDueCred") - MAX(T0."BalDueDeb")) * -1, MAX(T0."BalFcCred") - MAX(T0."BalFcDeb"), 
		                 MAX(T0."BalScCred") - MAX(T0."BalScDeb"), MAX(T0."LineMemo"), MAX(T1."FolioPref"), MAX(T1."FolioNum"),
		                  MAX(T0."Indicator"), MAX(T2."CardName"), MAX(T3."CardCode"), MAX(T3."CardName"), MAX(T2."Balance"),
		                   MAX(T3."NumAtCard"), MAX(T3."SlpCode"), MAX(T0."Project"), MAX(T0."Debit") + MAX(T0."Credit"), 
		                   MAX(T0."FCDebit") + MAX(T0."FCCredit"), MAX(T0."SYSDeb") + MAX(T0."SYSCred"), MAX(T2."PymCode"),
		                    MAX(T3."BlockDunn"), MAX(T3."DunnLevel"), MAX(T3."TransType"), MAX(T3."IsSales"), MAX(T2."Currency"),
		                     MAX(T0."FCCurrency"), MAX(T4."SlpName"), MAX(T2."DunTerm"), MAX(T0."DunnLevel"), MAX(T7."DocNum") AS "Inv No" 
		                     FROM "JDT1" T0 INNER JOIN "OJDT" T1 ON T1."TransId" = T0."TransId" 
		                     INNER JOIN "OCRD" T2 ON T2."CardCode" = T0."ShortName" 
		                     INNER JOIN "OINV" T7 ON T7."DocNum" = T0."BaseRef" 
		                     LEFT OUTER JOIN "B1_JournalTransSourceView" T3 ON T3."ObjType" = T0."TransType" AND T3."DocEntry" = T0."CreatedBy"
		                      AND (T3."TransType" <> n'I' OR (T3."TransType" = n'I' AND T3."InstlmntID" = T0."SourceLine")) 
		                      LEFT OUTER JOIN "OSLP" T4 ON T4."SlpCode" = T3."SlpCode" OR (T4."SlpName" = n'-No Sales Employee-' 
		                      AND (T0."TransType" = n'30' OR T0."TransType" = n'321' OR T0."TransType" = n'-5' OR T0."TransType" = n'-2' 
		                      OR T0."TransType" = n'-3' OR T0."TransType" = n'-4')) WHERE T0."RefDate" <= (:AgingDate) AND T0."RefDate" <= (:AgingDate)
		                       AND T2."CardType" = ('C') AND (T0."BalDueCred" <> T0."BalDueDeb" OR T0."BalFcCred" <> T0."BalFcDeb")
		                        AND NOT EXISTS (SELECT U0."TransId", U0."TransRowId" FROM "ITR1" U0 INNER JOIN "OITR" U1 ON U1."ReconNum" = U0."ReconNum" 
		                        WHERE T0."TransId" = U0."TransId" AND T0."Line_ID" = U0."TransRowId" AND U1."ReconDate" > (:AgingDate) 
		                        GROUP BY U0."TransId", U0."TransRowId") GROUP BY T0."TransId", T0."Line_ID") AS M 
		                        LEFT OUTER JOIN OCRD Cards ON M."CardCode" = Cards."CardCode" 
		              LEFT OUTER JOIN OCTG PyC ON Cards."GroupNum" = PyC."GroupNum" 
		                        WHERE  
		                        (M."CardCode" >= :CustomerCodeFrom OR :CustomerCodeFrom = '') and 
									(M."CardCode" <= :CustomerCodeTo OR :CustomerCodeTo = '')
									)T 
									left outer join OCRD S on T."CardCode" = S."CardCode"
									left outer join OSLP U ON U."SlpCode" = S."SlpCode"
									left outer join (select distinct "CardCode", "AcctCode" from ACR3 ) V on V."CardCode" = S."CardCode"
									where "validFor"  Not IN (case 
							        when :Invalid = 'I' then 'Y' 
							        when :Invalid = 'V' then 'N' 
							        When :Invalid <> 'I' AND :Invalid <> 'V' then ''
							         end)
									Group by T."CardCode"
									) A left outer join OACT B on A."AcctCode" = B."AcctCode"
									;
     
		      End;