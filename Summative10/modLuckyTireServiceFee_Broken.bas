Attribute VB_Name = "modLuckyTireServiceFee"
Option Compare Database
Option Explicit

' MIS342 Summative10 intentionally defective business-logic module.
' Known test:
' LaborAmount = 200
' PartsAmount = 300
' ShopFeeRate = 0.05
' Correct total = 525

Public Function CalcServiceTotal(LaborAmount As Currency, _
                                 PartsAmount As Currency, _
                                 ShopFeeRate As Double) As Currency

    ' Intentionally wrong: this calculates only the fee.
    CalcServiceTotal = (LaborAmount + PartsAmount) * ShopFeeRate

End Function
