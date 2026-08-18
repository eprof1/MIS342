Attribute VB_Name = "modLuckyTirePricing"
Option Compare Database
Option Explicit

' MIS342 Formative10 intentionally defective module.
' This file contains a business-logic error that students must diagnose,
' test, and correct. Do not assume that code is correct merely because it
' compiles successfully.

Public Function CalcDiscountedTotal(UnitPrice As Currency, _
                                    Quantity As Long, _
                                    DiscountRate As Double) As Currency

    ' Example business rule:
    ' $100 per tire * 2 tires with a 10% discount should equal $180.
    CalcDiscountedTotal = UnitPrice * Quantity * DiscountRate

End Function
