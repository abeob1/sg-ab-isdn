CREATE PROCEDURE "AE_SP001_ETADateListings"
(IN PONO VARCHAR(20) , IN PRDNO VARCHAR(20))

AS
LPONO INT;
LPRDNO INT;

BEGIN

IF PONO = '' THEN
LPONO := 0;
ELSE
LPONO := TO_INTEGER(PONO);
END IF;

IF PRDNO = '' THEN
LPRDNO := 0;
ELSE
LPRDNO := TO_INTEGER(PRDNO);
END IF;

CREATE COLUMN TABLE ORDRINFO  ("ItemCode" NVARCHAR(20),"DocNum" INTEGER,"CardName" NVARCHAR(100)
							  ,"Quantity" DECIMAL(21,6),"DocDueDate" DATE,"DocEntry" INTEGER);
							  
CREATE COLUMN TABLE OPORINFO  ("DocNum" INTEGER,"CardName" NVARCHAR(100),"ItemCode" NVARCHAR(20),"Dscription" NVARCHAR(150)
							  ,"Quantity" DECIMAL(21,6),"U_SupplierETD" DATE,"U_OldPOETD" DATE,"CShipDate" DATE
							  ,"ShipDate" DATE,"DocEntry" INTEGER,"OSLDItemCode" NVARCHAR(20));

--ORDR information
INSERT INTO ORDRINFO (
SELECT TT1. "ItemCode",TT0."DocNum",TT0."CardName" ,TT1."Quantity",TT0."DocDueDate",TT0."DocEntry" FROM ORDR TT0 INNER join
   RDR1 TT1 ON TT0."DocEntry"=TT1."DocEntry"   
   WHERE TT0."DocEntry" not in(
     (select "DocEntry" from RDR1 where "LineNum" not in (
   
    SELECT T2."DocLineNum"  FROM ORDR TT0 INNER join
   RDR1 TT1 ON TT0."DocEntry"=TT1."DocEntry" 
   INNER   JOIN OSLD T2 ON TT1."DocEntry" = T2."DocEntry" --AND TT1."ItemCode" = T2."ItemCode"
   and T2."DocEntry" = TT1."DocEntry"
   where  T2."ObjType" = 17 ))
)
   and  TT1."LineNum"  in 
 (select T4."LineNum" from RDR1 T4 where T4."LineNum" 
not in (select "DocLineNum" from OSLD where "ObjType"  = 17 and "DocEntry" = TT1."DocEntry")
   and T4."DocEntry" = TT1."DocEntry"
   )and TT0."DocStatus" = 'O' 

);

--OPOR Information
INSERT INTO OPORINFO(
SELECT T0."DocNum" "PO No / Prod No", T0."CardName" "Supplier Name", T1."ItemCode" "Item Code",
   T1."Dscription" "Item Description", T1."Quantity", T1."U_SupplierETD" "New Supplier ETD", 
   T1."U_OldPOETD" "Initial PO ETA", T1."ShipDate" "Current PO ETA", 
   T1."ShipDate" "Initial Confirm Delivery Date"
  , T0."DocEntry", T2."ItemCode"
   FROM "OPOR"  T0 
   INNER JOIN POR1 T1 ON T0."DocEntry" = T1."DocEntry"
   LEFT OUTER JOIN OSLD T2 ON T1."DocEntry" = T2."DocEntry" AND T1."ItemCode" = T2."ItemCode"
   AND T2."DocLineNum" = T1."LineNum" AND T0."ObjType" = T2."ObjType"
   WHERE T0."DocNum" = :LPONO--220  and T1."U_OldPOETD" <> T1."ShipDate"
);


--  PO ETA Date listing

SELECT T0."DocNum" "PO No / Prod No", T0."CardName" "Supplier Name", T0."ItemCode" "Item Code",
   T0."Dscription" "Item Description", T0."Quantity", T0."U_SupplierETD" "New Supplier ETD", 
   T0."U_OldPOETD" "Initial PO ETA", T0."ShipDate" "Current PO ETA", 
  TT2."DocNum" "SO No", TT2."CardName" "Customer Name",
   T0."CShipDate" "Initial Confirm Delivery Date"
  , TT2."Quantity" "Confirmed Qty"
   FROM 
   "OPORINFO" T0 LEFT JOIN "ORDRINFO" TT2 on T0."ItemCode"=TT2."ItemCode"   
   AND TT2."ItemCode" = T0."OSLDItemCode"--"OSLDItemCode" 
   WHERE 
     T0."ShipDate" = 
  (
  SELECT  min(P1."ShipDate")
  FROM "OPOR"  P0 
  INNER JOIN POR1 P1 ON P0."DocEntry" = P1."DocEntry"
  LEFT OUTER JOIN OSLD P2 ON P1."DocEntry" = P2."DocEntry" AND P1."ItemCode" = P2."ItemCode"
  where P1."ItemCode" = T0."ItemCode" and P0."DocStatus" = 'O' 
  and (P1."Quantity" <= TT2."Quantity" or TT2."Quantity" <= P1."Quantity")
  )
  
  or T0."ShipDate" = 
(
 SELECT  P1."ShipDate"
   FROM "OPOR"  P0 
   INNER JOIN POR1 P1 ON P0."DocEntry" = P1."DocEntry"
   LEFT OUTER JOIN OSLD P2 ON P1."DocEntry" = P2."DocEntry" AND P1."ItemCode" = P2."ItemCode"
   where P1."ItemCode" = T0."ItemCode" and P0."DocStatus" = 'O' 
and P0."DocNum" = :LPONO--520 

   and ( P1."Quantity" < TT2."Quantity" )
   and P1."ShipDate"
   <>
   (SELECT  min(S1."ShipDate")
   FROM "OPOR"  S0 
   INNER JOIN POR1 S1 ON S0."DocEntry" = S1."DocEntry"
   LEFT OUTER JOIN OSLD S2 ON S1."DocEntry" = S2."DocEntry" AND S1."ItemCode" = S2."ItemCode"
   where S1."ItemCode" = T0."ItemCode" and S0."DocStatus" = 'O' )
)   

UNION ALL

--  PRD ETA Date listing

SELECT T1."DocNum" "PO No / Prod No", '' "Supplier Name" , T1."ItemCode" "Item Code",T4."ItemName" "Item Description",
 T1."PlannedQty" "Quantity", '' "New Supplier ETD" ,T1."U_OldPRETD" "Initial PO ETA",T1."DueDate" "Current PO ETA"
 ,TT2."DocNum" "SO No", TT2."CardName" "Customer Name",
T3."CfmDate" "Initial Confirm Delivery Date",
T3."CfmQty" "Confirmed Qty" 
FROM "OWOR" T1 
LEFT OUTER join "OSLD" T3 on T1."DocEntry"=T3."DocEntry" AND T1."ItemCode" = T3."ItemCode" AND T3."ObjType" = '202'
INNER JOIN "OITM" T4 on T4."ItemCode"=T1."ItemCode"
LEFT JOIN 
   (SELECT TT1. "ItemCode",TT0."DocNum",TT0."CardName"  FROM ORDR TT0 INNER join
   RDR1 TT1 ON TT0."DocEntry"=TT1."DocEntry")  TT2 on T1."ItemCode"=TT2."ItemCode"
WHERE T1."DocNum" = :LPRDNO--DocW  
AND T1."U_OldPRETD" <> T1."DueDate";
  
 DROP TABLE ORDRINFO;
 DROP TABLE OPORINFO;
 END;