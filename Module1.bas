Attribute VB_Name = "Module1"
Public Sub test()
Sheets("хос").Select
Range("C2:D141") = Range("E2:F141").Value
Range("H2:I141") = Range("J2:K141").Value
Range("N2:N141") = Range("O2:O141").Value

Sheets("ндос").Select
Range("C4") = Range("C5").Value
Range("C8") = 0
Range("C9") = Range("C10").Value
Range("E8") = Range("E8") + 1

Sheets("рюпхтш").Select
Range("E3") = Range("E3").Value - Range("E6").Value
Range("E12") = 0

Sheets("оепепюяверш").Select
Range("E2:S141") = 0

Sheets("нокюрю").Select
Range("K2:K141") = Range("Q2:Q141").Value
Range("M2:M141") = 0
End Sub

