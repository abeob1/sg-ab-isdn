CALL "AE_SP001_OutStandingSO"()

DROP PROCEDURE "AE_SP001_OutStandingSO";

CREATE PROCEDURE "AE_SP001_OutStandingSO"
( )
AS
BEGIN
SELECT T0."DocNum", T0."CardCode", T1."WhsCode", T0."NumAtCard", T0."U_ProjectNo" , 
T0."U_PositionNo" , T1."SubCatNum", T1."ItemCode", 
T1."Dscription", T1."Quantity", T1."DelivrdQty", T1."OpenQty", T1."ShipDate", 
null "Initial Confirm Delivery Date", null "Confirmed Qty", T0."U_RemarkSA" ,
T1."Price" 
FROM ORDR T0  INNER JOIN RDR1 T1 ON T0."DocEntry" = T1."DocEntry"
WHERE 
T0."DocEntry" not in (SELECT T0."DocEntry" FROM OSLD T0 WHERE T0."ObjType"  = 17)
AND T1."U_ETD" = '30000101';

END;
