------DROP PROCEDURE "AE_STOCKAGING_ISDN_SUB"
CREATE PROCEDURE "AE_STOCKAGING_ISDN_SUB" 
(
	IN AgingDate TimeStamp,
	IN ItemCodeFrom varchar(20),
	IN ItemCodeTo varchar(20),
	IN WhsCodeFrom varchar(20),
	IN WhsCodeTo varchar(20)	
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


BEGIN

AgeColumn0 := 0;
AgeColumn1 := 365;
AgeColumn2 := 366;
AgeColumn3 := 730;
AgeColumn4 := 731;
AgeColumn5 := 1095;
AgeColumn6 := 1096;
AgeColumn7 := 1460;
AgeColumn8 := 1461;
HEADER1 := '0-1 Year';
HEADER2 :='2 Years';
HEADER3 :='3-4 Years';
HEADER4 :='4-5 Years';
HEADER5 :='5 Years and above';



SELECT
	 (SELECT
	 TOP 1 ifnull("PrintHeadr",
	 "CompnyName") 
	FROM"OADM") 
INTO ComName 
FROM DUMMY;

SELECT (SELECT "TaxIdNum" FROM"OADM") INTO CoRegNo FROM DUMMY;
SELECT (SELECT "TaxIdNum2" FROM"OADM") INTO GSTRegNo FROM DUMMY;

CREATE COLUMN TABLE OnHand  ("X" INTEGER ,"ItemCode" NVARCHAR(20),"Description" NVARCHAR(100),"Warehouse" NVARCHAR(8)
							  ,"OnHand" DECIMAL(36,3),"TransValue" DECIMAL(36,9),"AvgPrice" DECIMAL(36,15));
							  
CREATE COLUMN TABLE TEMP_RunTotal ("InQty1" DECIMAL(21,6),"TtlQty"  DECIMAL(21,6)  ,"ItemCode" NVARCHAR(20)
							   ,"DocDate" Date,"Warehouse" NVARCHAR(8),"InQty" DECIMAL(21,6),"TransType" NVARCHAR(20)
								 ,"BASE_REF" NVARCHAR(20),"CreatedBy" NVARCHAR(20),"TransNum" INTEGER,"TransValue" DECIMAL(36,9)
								 ,"CardCode" NVARCHAR(50),"CardName" NVARCHAR(200)
								 );
								 	
CREATE COLUMN TABLE RunTotal  ("InQty1" DECIMAL(21,6),"TtlQty"  DECIMAL(21,6)  ,"ItemCode" NVARCHAR(20)
								,"DocDate" Date,"Warehouse" NVARCHAR(8),"InQty" DECIMAL(21,6),"TransType" NVARCHAR(20)
								 ,"BASE_REF" NVARCHAR(20),"CreatedBy" NVARCHAR(20),"TransValue" DECIMAL(36,9)
								  ,"CardCode" NVARCHAR(50),"CardName" NVARCHAR(200));	
								 
CREATE COLUMN TABLE Temp1  ("ItemCode" NVARCHAR(20),"Description" NVARCHAR(100),"OnHand" DECIMAL(36,3),"TransValue" DECIMAL(36,9)
							,"DocDate" Date,"Warehouse" NVARCHAR(8),"Quantity" DECIMAL(21,6),"AvgPrice" DECIMAL(36,15)
							,"TransType" NVARCHAR(20),"BASE_REF" NVARCHAR(20),"CreatedBy" NVARCHAR(20)
							 ,"CardCode" NVARCHAR(50),"CardName" NVARCHAR(200));
							
CREATE COLUMN TABLE Temp  ("ItemCode" NVARCHAR(20),"ItemName" NVARCHAR(100),"OnHand" DECIMAL(36,3),"TransValue" DECIMAL(36,9)
							,"Warehouse" NVARCHAR(8),"WhsName" NVARCHAR(100),"ItmsGrpNam" NVARCHAR(20),"AgeColumn1" DECIMAL(21,6)
							,"AgeColumn2" DECIMAL(21,6),"AgeColumn3" DECIMAL(21,6)
							,"AgeColumn4" DECIMAL(21,6),"AgeColumn5" DECIMAL(21,6)
							,"AvgPrice" DECIMAL(36,15)
							,"TransType" NVARCHAR(20),"BASE_REF" NVARCHAR(20),"CreatedBy" NVARCHAR(20)
							 ,"CardCode" NVARCHAR(50),"CardName" NVARCHAR(200), "DocDate" Date
							);							
									
CREATE COLUMN TABLE Final  ("CompName" NVARCHAR(100),"CoRegNo" NVARCHAR(100),"GSTRegNo" NVARCHAR(100),"LogoImage" BLOB,
							"ItemCode" NVARCHAR(20),"ItemName" NVARCHAR(100),"OnHand" DECIMAL(36,3),"TransValue" DECIMAL(36,9)
						,"WhsCode" NVARCHAR(8),"WhsName" NVARCHAR(100),"ItmsGrpNam" NVARCHAR(20)
						,"AgeColumn1" DECIMAL(21,6),"AgeColumn2" DECIMAL(21,6),"AgeColumn3" DECIMAL(21,6)
							,"AgeColumn4" DECIMAL(21,6),"AgeColumn5" DECIMAL(21,6)
						,"AgeColumn1_Value" DECIMAL(36,15),"AgeColumn2_Value" DECIMAL(36,15),"AgeColumn3_Value" DECIMAL(36,15)
						,"AgeColumn4_Value" DECIMAL(36,15),"AgeColumn5_Value" DECIMAL(36,15)
						,"Header1" NVARCHAR(100) ,"Header2" NVARCHAR(100),"Header3" NVARCHAR(100)
						,"Header4" NVARCHAR(100),"Header5" NVARCHAR(100)
						,"AvgPrice" DECIMAL(36,15),"UOM" NVARCHAR(100)
						,"CardCode" NVARCHAR(50),"CardName" NVARCHAR(200), "DocDate" Date, "TransType" NVARCHAR(20)
						, "PONumber" NVARCHAR(20), "Aging" Integer
);								 							 						  

CREATE COLUMN TABLE Invoice_Details  ("ItemCode" NVARCHAR(50),"InvDate" Date );

INSERT INTO Invoice_Details
( 
select B."ItemCode", max(A."DocDate")
from OINV A INNER JOIN INV1 B ON A."DocEntry" = B."DocEntry" INNER JOIN OITM C ON C."ItemCode" = B."ItemCode"
AND (B."ItemCode" >= :ItemCodeFrom or :ItemCodeFrom = '') and 
	(B."ItemCode" <= :ItemCodeTo or :ItemCodeTo = '')
Group by B."ItemCode"
);	

--select * from Invoice_Details;	
			  
INSERT INTO OnHand
(
select 
SUM(A."TransValue") AS "X",
	A."ItemCode"
	, max(B."ItemName") as "Description"
	, A."Warehouse"
	, TO_DECIMAL(IFNULL(Sum(IFNull(A."InQty",0) - IFNULL(A."OutQty",0)),0.0),36,3) As "OnHand"
, TO_DECIMAL(IFNULL(Sum(A."TransValue"),0.0000000),36,9) As "TransValue"
, ROUND(ROUND(CASE WHEN IFNULL(Sum(IFNULL(A."InQty",0) - IFNULL(A."OutQty",0)),0)=0 THEN 0 ELSE 
	IFNULL(TO_DECIMAL(IFNULL(Sum(A."TransValue"),0.0000000),36,9)
	/TO_DECIMAL(IFNULL(Sum(IFNull(A."InQty",0) - IFNULL(A."OutQty",0)),0.0),36,3)
	,0) 
	END,16),4,ROUND_CEILING)  AS "AvgPrice"
From OINM A
	Join OITM B On A."ItemCode" = B."ItemCode"
	where A."DocDate" <= :AgingDate and 
	(A."ItemCode" >= :ItemCodeFrom or :ItemCodeFrom = '') and 
	(A."ItemCode" <= :ItemCodeTo or :ItemCodeTo = '') and 
	(A."Warehouse" >= :WhsCodeFrom or :WhsCodeFrom = '') and (A."Warehouse" <= :WhsCodeTO or :WhsCodeTo = '')
	Group By A."ItemCode", A."Warehouse"
		HAVING IFNULL(Sum(IFNULL(A."InQty",0) - IFNULL(A."OutQty",0)),0)<>0  OR  IFNULL(Sum(A."TransValue"),0)<>0
);		

--select * from OnHand;			  

INSERT INTO TEMP_RunTotal(						  
Select 
	A."InQty" as "InQty1"
		, 0 As "TtlQty"
		, A."ItemCode"
		, A."DocDate"
		, A."Warehouse"
		, A."InQty"
		, A."TransType"
		, A."BASE_REF"	
		, A."CreatedBy"				
		, A."TransNum"
		, A."TransValue"
		, A."CardCode"
		, A."CardName"

From OINM A INNER JOIN OITW ON
		A."ItemCode"=OITW."ItemCode" and A."Warehouse" = OITW."WhsCode"
	Where A."DocDate" <= :AgingDate and 
	(A."ItemCode" >= :ItemCodeFrom or :ItemCodeFrom = '') and 
	(A."ItemCode" <= :ItemCodeTo or :ItemCodeTo = '') and 
	(A."Warehouse" >= :WhsCodeFrom or :WhsCodeFrom = '') and 
		(A."Warehouse" <= :WhsCodeTo or :WhsCodeTo = '')
--	And A."TransType" In(-2, 14, 16, 18, 20, 59, 67) 
	And (IFNULL(A."InQty",0) <> 0)
Order By A."TransNum" Desc) ;

--select * from TEMP_RunTotal;

INSERT INTO RunTotal(						  
Select 
	A."InQty" as "InQty1"
	, (Select Sum("InQty") 
		From "OINM" 
		Where "DocDate" <= :AgingDate And IFNULL("InQty",0) <> 0 
			And "ItemCode" = A."ItemCode"  and "Warehouse" = A."Warehouse"
			And "TransNum" > A."TransNum"
--			And "TransType" In(-2, 14, 16, 18, 20, 59, 67) 
		group by "ItemCode", "Warehouse" 
		) As "TtlQty"
	, A."ItemCode"
	, A."DocDate"
	, A."Warehouse"
	, A."InQty"
	, A."TransType"
	, A."BASE_REF"		
	, A."CreatedBy"			
	, A."TransValue"
	, A."CardCode"
	, A."CardName"
	
From TEMP_RunTotal A 
Order By A."TransNum" Desc ) ;

--select * from RunTotal;

INSERT INTO Temp1(						  
Select 
	 A."ItemCode"
	, A."Description"
	, A."OnHand"
	, A."TransValue"
	, B."DocDate"
	, B."Warehouse"
	, Case When B."InQty" + IFNULL(B."TtlQty",0) > A."OnHand"
			Then Case When A."OnHand" - IFNULL(B."TtlQty",0) < 0 
				Then 0 Else A."OnHand" - IFNULL(B."TtlQty",0) End 
		Else B."InQty" 
		End As "Quantity"
	, A."AvgPrice"
	, B."TransType"
	, B."BASE_REF"		
	, B."CreatedBy"	
	, B."CardCode"
	, B."CardName"	
From OnHand A 
	Join RunTotal B On A."ItemCode" = B."ItemCode" AND A."Warehouse" = B."Warehouse"
order by A."ItemCode", A."Warehouse") ;

--Select * from Temp1 order by Temp1."DocDate" Desc;


INSERT INTO Temp(						  
Select 
	A."ItemCode"
	, max(A."Description") As "ItemName"
	, A."OnHand"
	, A."TransValue"
	, A."Warehouse" AS "WhsCode"
	, max(D."WhsName") AS "Whsname"
	, Max(H."ItmsGrpNam") As "ItmsGrpNam"
	, Sum (Case when DAYS_BETWEEN(A."DocDate",:AgingDate) between :AgeColumn0 And :AgeColumn1 Then "Quantity" Else 0 End) as "AgeColumn1"
	, Sum (Case when DAYS_BETWEEN(A."DocDate",:AgingDate) between :AgeColumn2 And :AgeColumn3 Then "Quantity" Else 0 End) as "AgeColumn2"
	, Sum (Case when DAYS_BETWEEN(A."DocDate",:AgingDate) between :AgeColumn4 And :AgeColumn5 Then "Quantity" Else 0 End) as "AgeColumn3"
	, Sum (Case when DAYS_BETWEEN(A."DocDate",:AgingDate) between :AgeColumn6 And :AgeColumn7 Then "Quantity" Else 0 End) as "AgeColumn4"	
	, Sum (Case when DAYS_BETWEEN(A."DocDate",:AgingDate) >:AgeColumn8 Then "Quantity" Else 0 End) as "AgeColumn5"
	, A."AvgPrice"
	, A."TransType"
	, A."BASE_REF"		
	, A."CreatedBy"	
	, A."CardCode"
	, A."CardName"	
	, A."DocDate"
	
From Temp1 A
	left outer Join OWHS D On A."Warehouse"=D."WhsCode"
	left outer join OITM F on A."ItemCode" = F."ItemCode"
	Left outer Join OITB H on H."ItmsGrpCod"=F."ItmsGrpCod"
	where A."Quantity" <> 0
	Group by 
	A."ItemCode"
	, A."Warehouse"
	, A."OnHand"
	,A."DocDate"
--	, Convert(Varchar(10),A.DocDate,112)
	, A."TransValue"
	, A."AvgPrice"
	, A."TransType"
	, A."BASE_REF"		
	, A."CreatedBy"
	, A."CardCode"
	, A."CardName"		
Order by A."ItemCode" desc) ;

--select * from Temp;

INSERT INTO Final(						  
select 
	:ComName as "CompName"
	,:CoRegNo as "CoRegNo"
	,:GSTRegNo as "GSTRegNo"
	,(select Top 1 "LogoImage" from OADP) as "LogoImage"
	,A."ItemCode"
	,A."ItemName"
	,A."OnHand"
	/*,A.TransValue as TransValue*/
--	 CAST(CAST(1.5 as DECIMAL(15,2)) as VARCHAR) 
	,CAST(CAST(MAX(A."AvgPrice")* (A."OnHand") as DECIMAL(19,2)) as VARCHAR) as "TransValue"  
--	,ROUND(MAX(A."AvgPrice")* (A."OnHand"),2) as "TransValue"  
--	,A."TransValue"  
	,A."Warehouse"
	,max(A."WhsName") as "WhsName"
	,A."ItmsGrpNam" 
	,sum(A."AgeColumn1") as "AgeColumn1"
	,sum(A."AgeColumn2") as "AgeColumn2" 
	,sum(A."AgeColumn3") as "AgeColumn3"
	,sum(A."AgeColumn4") as "AgeColumn4"
	,sum(A."AgeColumn5") as "AgeColumn5"
	, ROUND(max(A."AgeColumn1") * max(A."AvgPrice"),2) as "AgeColumn1_Value"
	, ROUND(max(A."AgeColumn2") * max(A."AvgPrice"),2) as "AgeColumn2_Value"
	, ROUND(max(A."AgeColumn3") * max(A."AvgPrice"),2) as "AgeColumn3_Value"
	, ROUND(max(A."AgeColumn4") * max(A."AvgPrice"),2) as "AgeColumn4_Value"
	, ROUND(max(A."AgeColumn5") * max(A."AvgPrice"),2) as "AgeColumn5_Value"
	, :HEADER1 "HEADER1"
	, :HEADER2 "HEADER2"
	, :HEADER3 "HEADER3"
	, :HEADER4 "HEADER4"
	, :HEADER5 "HEADER5"
	, MAX(A."AvgPrice") as "AvgPrice"
	, B."InvntryUom" as "UOM"
	, case when A."TransType" in ('20','22') then A."CardCode" else '' end "CardCode"
	, case when A."TransType" in ('20','22') then A."CardName" else '' end "CardName"	
	, case when A."TransType" in ('20','22') then A."DocDate" else '' end "DocDate"	
	, A."TransType"
	, case 
	when A."TransType" in ('20') then IFNULL(( SELECT max(P."BaseRef") FROM PDN1 P WHERE P."DocEntry" = A."CreatedBy" and P."ItemCode" = A."ItemCode" and P."BaseType" = '22'),'')  
	WHEN A."TransType" in ('22') THEN A."BASE_REF"	
	else '' end "PONo"
	,  DAYS_BETWEEN(A."DocDate",:AgingDate) as "Aging"
from Temp A
inner join OITM B on A."ItemCode" = B."ItemCode"
group by a."ItemCode",A."Warehouse",a."ItemName",a."OnHand",a."ItmsGrpNam",b."InvntryUom",A."TransValue",
 A."CardCode", A."CardName", A."DocDate", A."TransType"	, A."BASE_REF", A."CreatedBy"
order by a."ItemCode" asc
) ;



SELECT sum(A."AgeColumn3_Value") "3-4", sum(A."AgeColumn4_Value") "4-5", sum(A."AgeColumn5_Value")"5>"
 FROM Final A ;

--SELECT * FROM RunTotal;
DROP TABLE OnHand;
DROP TABLE TEMP_RunTotal;
DROP TABLE RunTotal;
DROP TABLE Temp1;
DROP TABLE Temp;
DROP TABLE Final;
DROP TABLE Invoice_Details;
END