Create PROCEDURE "AE_OnTimeDelByCustomer"
(
 IN DateFrom DATE,
 IN DateTo DATE,
 IN BPFrom nvarchar(30), 
 IN BPTo nvarchar(30)

)
AS	
	FromDate VARCHAR(20);
	ToDate VARCHAR(20);
	
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
	FromDate:=TO_CHAR(:DateFrom ,'YYYY-MM-DD');
	ToDate:=TO_CHAR(:DateTo ,'YYYY-MM-DD');
	
SELECT (SELECT TOP 1 ifnull("PrintHeadr", "CompnyName") FROM "OADM") INTO ComName FROM DUMMY;
SELECT (SELECT "TaxIdNum" FROM"OADM") INTO CoRegNo FROM DUMMY;
SELECT (SELECT "TaxIdNum2" FROM"OADM") INTO GSTRegNo FROM DUMMY;
SELECT (SELECT "Phone1" FROM"OADM") INTO PHONE1 FROM DUMMY;
SELECT (SELECT "Fax" FROM"OADM") INTO FAX FROM DUMMY;
SELECT (SELECT "CompnyAddr" FROM"OADM") INTO COMPNYADDR FROM DUMMY;
SELECT (SELECT "E_Mail" FROM"OADM") INTO EMAIL FROM DUMMY;


SELECT 
 	    :ComName as CompanyName
       ,:CompName AS COMPNAME
	   ,:COMPNYADDR  AS COMPADD
	   ,:PHONE1 AS COMPTEL
	   ,:FAX AS COMPFAX
	   ,:CoRegNo AS COMPREGNO
	   ,:GSTRegNo as GSTRegNo
	   ,:EMAIL As E_Mail
	   ,(select Top 1 "LogoImage" from"OADP") as LogoImage,
	   row_number() over (order by T0."CardCode" DESC) as row_num,
T0."CardCode",T0."CardName",T0."DocNum" as DL_NO,T9."CfmDate",T0."DocDate",T5."DocDueDate"
 , case when DAYS_BETWEEN (IFNULL(T9."CfmDate",T5."DocDueDate"),T0."DocDate" ) <=0 then 100
 else 0 end as OnTime,
 case when DAYS_BETWEEN (IFNULL(T9."CfmDate",T5."DocDueDate"),T0."DocDate") <=0 then 1
 else 0 end as CountOnTime
  ,rank() OVER (PARTITION BY T0."CardCode" ORDER BY 
 sum(case when DAYS_BETWEEN (IFNULL(T9."CfmDate",T5."DocDueDate"),T0."DocDate" ) <=0 then 100
 else 0 end)/count(T0."DocNum") DESC ) AS Rank

FROM ODLN T0 INNER JOIN DLN1 T3 ON T3."DocEntry" = T0."DocEntry"
LEFT JOIN  INV1 T4 ON T4."BaseEntry" = T3."DocEntry" AND T4."BaseLine" = T3."LineNum"
LEFT JOIN OINV T1 ON T1."DocEntry" = T4."DocEntry"
LEFT outer JOIN ORDR T5 ON COALESCE(NULLIF(T3."BaseRef",''), '0') = T5."DocNum"
LEFT JOIN RDR1 T8 ON T8."DocEntry" = T5."DocEntry"  AND T3."BaseLine" = T8."LineNum"
Left JOIN OSLD T9 ON T8."DocEntry" = T9."DocEntry" AND T8."ItemCode" = T9."ItemCode" 
       			AND T9."DocLineNum" = T8."LineNum"
left join RIN1 T6 on T4."TrgetEntry" = T6."DocEntry" AND T6."LineNum" = T4."LineNum"
left join ORIN T7 on T7."DocEntry" = T6."DocEntry"
Where (T0."CANCELED" <> 'Y'  OR T1."CANCELED" <> 'Y' OR T1."CANCELED" <> 'Y')
AND T0."DocDate" between :FromDate AND :ToDate
AND T0."CardCode" >= :BPFrom AND T0."CardCode"<= :BPTo
group by T0."CardCode",T0."CardName",T0."DocNum",T9."CfmDate",T0."DocDate",T5."DocDueDate"
;
--AND T0."DocNum" in( 415,416);
END;