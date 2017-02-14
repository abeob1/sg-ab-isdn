
--CALL "AE_SP001_ETADateListing"('1','1')

DROP PROCEDURE "AE_SP001_ETADateListing";
CREATE PROCEDURE "AE_SP001_ETADateListing"
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

--SELECT :LPONO FROM DUMMY;
--  PO ETA Date listing
SELECT T0."DocNum" "PO No / Prod No", T0."CardName" "Supplier Name", T1."ItemCode" "Item Code",
   T1."Dscription" "Item Description", T1."Quantity", T1."U_SupplierETD" "New Supplier ETD", 
   T1."U_OldPOETD" "Initial PO ETA", T1."ShipDate" "Current PO ETA", 
   TT2."DocNum" "SO No", TT2."CardName" "Customer Name",
   T2."CfmDate" "Initial Confirm Delivery Date", T2."CfmQty" "Confirmed Qty"
   FROM "OPOR"  T0 
   INNER JOIN POR1 T1 ON T0."DocEntry" = T1."DocEntry"
   LEFT OUTER JOIN OSLD T2 ON T1."DocEntry" = T2."DocEntry" AND T1."ItemCode" = T2."ItemCode"
   AND T2."DocLineNum" = T1."LineNum" AND T0."ObjType" = T2."ObjType" 
   LEFT JOIN 
   (SELECT TT1. "ItemCode",TT0."DocNum",TT0."CardName"  FROM ORDR TT0 INNER join
   RDR1 TT1 ON TT0."DocEntry"=TT1."DocEntry")  TT2 on T1."ItemCode"=TT2."ItemCode"
   WHERE T0."DocNum" = :LPONO and T1."U_OldPOETD" <> T1."ShipDate"
   
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
WHERE T1."DocNum" = :LPRDNO AND T1."U_OldPRETD" <> T1."DueDate";

END;







