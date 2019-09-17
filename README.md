# Excel Avançado

# Macro 1

Sub Macro1()

' Aula 10/10/2019

Dim linha, contador As Integer
Dim juros As Double
linha = 7
Range("A8:A100").Clear
If Range("C5").Value = 1 Then
Range("A7:D100").Clear
 Range("A7").Value = "Desconto 10%"
 Range("B7").Value = 1
 Range("C7").FormulaLocal = "=C3 + 5"
 Range("C7").Value = CDate(Range("C7").Value)
 Range("D7").FormulaLocal = "=C4-(C4*10%)"
 Range("D7").Value = CCur(Range("=D7").Value)
 Dim a, b As Integer
 a = Range("A12").Value
 b = Range("A13").Value
 Range("A12:A13").Clear
 Range("A14").Value = a
 Range("A15").Value = b
 Range("C3").Font.Color = vbBlue
 Range("C4").Font.Color = vbRed
 Range("D7:D1000").Select
   With Selection.Font
                 .Color = RGB(10, 150, 30)
                 .Bold = "True"
                 .Size = 12
                 .Name = "Arial"
  End With
 
End If

If (Range("C5").Value > 1 And Range("C5").Value <= 5) Then
 Range("A7").Clear
 Range("A7:D100").Clear
 For contador = 1 To Range("C5").Value
 Range("B" & linha).Select
 ActiveCell.Formula = contador
 Range("C" & linha).Select
 ActiveCell.Formula = CDate(WorksheetFunction.EDate(Range("c3").Value, contador))
 Range("D" & linha).Select
 ActiveCell.Formula = "=C4/C5"
 Range("D" & linha).Value = CCur(Range("D" & linha).Value)
 linha = linha + 1
 Range("C3").Font.Color = vbBlue
 Range("C4").Font.Color = vbRed
 Range("D7:D1000").Select
   With Selection.Font
                 .Color = RGB(10, 150, 30)
                 .Bold = "True"
                 .Size = 12
                 .Name = "Arial"
                 
   End With
 Next contador
 
End

End If

If (Range("C5").Value >= 6 And Range("C5").Value <= 9) Then
 Range("A7:D100").Clear
 Range("A7").Clear
 For contador = 1 To Range("C5").Value
 Range("B" & linha).Select
 ActiveCell.Formula = contador
 Range("C" & linha).Select
 ActiveCell.Formula = CDate(WorksheetFunction.EDate(Range("c3").Value, contador))
 Range("D" & linha).Select
 ActiveCell.Formula = "=((C4/C5)*101%)"
 Range("D" & linha).Value = CCur(Range("D" & linha).Value)
 linha = linha + 1
 Range("C3").Font.Color = vbBlue
 Range("C4").Font.Color = vbRed
 Range("D7:D1000").Select
   With Selection.Font
                 .Color = RGB(10, 150, 30)
                 .Bold = "True"
                 .Size = 12
                 .Name = "Arial"
   End With
 Next contador

End

End If

If Range("C5").Value >= 10 Then
 juros = 0.01
 Range("A7").Clear
 Range("A7:D100").Clear
 For contador = 1 To Range("C5").Value
 Range("B" & linha).Select
 ActiveCell.Formula = contador
 Range("C" & linha).Select
 ActiveCell.Formula = CDate(WorksheetFunction.EDate(Range("c3").Value, contador))
 Range("D" & linha).Select
 ActiveCell.Formula = (Range("C4").Value \ Range("C5").Value) + (Range("C4") \ Range("C5").Value) * juros
 Range("D" & linha).Value = CCur(Range("D" & linha).Value)
 linha = linha + 1
 juros = juros + 0.01
 Range("C3").Font.Color = vbBlue
 Range("C4").Font.Color = vbRed
 Range("D7:D1000").Select
   With Selection.Font
                 .Color = RGB(10, 150, 30)
                 .Bold = "True"
                 .Size = 12
                 .Name = "Arial"
   End With
 Next contador

End

End If
                 
End Sub
 
