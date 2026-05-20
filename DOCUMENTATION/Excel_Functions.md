## CUBEVALUE
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Current Month Trade Count]",Timeline_Date,Slicer_Analyst_Trader_Name,Slicer_Direction,Slicer_Strategy_Tag,
Slicer_Ticker_Symbol)

## CUBESET
=CUBESET(
"ThisWorkbookDataModel",
"[Calendar].[MonthYear].Children",
"All Months"
)//This saves all the Month-Years without the "All" at the top.

## CUBERANKEDMEMBER
=IFERROR(CUBERANKEDMEMBER(
"ThisWorkbookDataModel",
$G$1,
ROW(A1)
)," ")//This displays all the Month-Years without the "All" at the top. ROW(A1) acts as a counter.

## TopCount
//This stores the five most recent trades.
=CUBESET("ThisWorkbookDataModel", "TopCount([Sheet1].[Trade_ID].Children, 5, [Measures].[MaxExecutionDate])", "Recent Trades")

## INDEX AND MATCH
=PIVOTS!$I$1:INDEX(PIVOTS!$I$1:$I$12, MATCH(2, 1/(PIVOTS!$Q$1:$Q$12 > 0)))//The Lookup Value is 2: Because match_type is omitted, MATCH defaults to 1 
(less than or equal to).//The Trick: The lookup value (2) is deliberately larger than any valid number in the array (which only contains 1s and errors).
