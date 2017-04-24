Create PROCEDURE "AE_OnTimeDelBySupplier"
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
T0."CardCode",T0."CardName",T0."DocNum" as GRPO_No
--, case when DAYS_BETWEEN (T3."ShipDate",IFNULL(T3."U_SupplierETD",T0."DocDueDate") ) <=0 then 100
-- else 0 end as OnTime,
-- case when DAYS_BETWEEN (T3."ShipDate",IFNULL(T3."U_SupplierETD",T0."DocDueDate") ) <=0 then 1
-- else 0 end as CountOnTime
--  ,rank() OVER (PARTITION BY T0."CardCode" ORDER BY 
-- sum(case when DAYS_BETWEEN (T3."ShipDate",IFNULL(T3."U_SupplierETD",T0."DocDueDate") ) <=0 then 100
-- else 0 end)/count(T0."DocNum") DESC ) AS Rank
 
 , case when DAYS_BETWEEN (IFNULL(T8."ShipDate",T5."DocDueDate"),T0."DocDate" ) <=0 then 100
 else 0 end as OnTime,
 case when DAYS_BETWEEN (IFNULL(T8."ShipDate",T5."DocDueDate"),T0."DocDate"  ) <=0 then 1
 else 0 end as CountOnTime
  ,rank() OVER (PARTITION BY T0."CardCode" ORDER BY 
 sum(case when DAYS_BETWEEN (IFNULL(T8."ShipDate",T5."DocDueDate"),T0."DocDate"  ) <=0 then 100
 else 0 end)/count(T0."DocNum") DESC ) AS Rank
FROM OPDN T0 INNER JOIN PDN1 T3 ON T3."DocEntry" = T0."DocEntry"
LEFT JOIN  PCH1 T4 ON T4."BaseEntry" = T3."DocEntry" AND T4."BaseLine" = T3."LineNum"
LEFT JOIN OPCH T1 ON T1."DocEntry" = T4."DocEntry"
LEFT JOIN OPOR T5 ON COALESCE(NULLIF(T3."BaseRef",''), '0')= T5."DocNum"
LEFT JOIN POR1 T8 ON T8."DocEntry" = T5."DocEntry" AND T3."BaseLine" = T8."LineNum"
left join RPC1 T6 on T4."TrgetEntry" = T6."DocEntry" AND T6."LineNum" = T4."LineNum"
LEFT join ORPC T7 on T7."DocEntry" = T6."DocEntry"
Where (T0."CANCELED" <> 'Y' OR T1."CANCELED" <> 'Y' or T7."CANCELED" <> 'Y')
AND T0."DocDate" between :FromDate AND :ToDate
AND T0."CardCode" >= :BPFrom AND T0."CardCode"<= :BPTo
group by T0."CardCode",T0."CardName",T0."DocNum",T0."DocDate",T5."DocDueDate"
,T8."ShipDate"
;
--AND T0."DocNum" in( 476,477,478);
END;