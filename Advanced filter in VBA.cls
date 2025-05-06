Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Range("M1:S1")) Is Nothing Then
    
Sheets("PROPOSAL").Unprotect
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim PosledniRadek As Integer
Dim CisloSloupce As Integer
Dim OznacOblast As Range
Dim i As Integer
Dim FiltrovaciBunka As String
Dim s As Integer
Dim WrdArray() As String
Dim SloupecOznacit As Integer

Range(Cells(1, 1), Cells(20000, 1)).EntireRow.Hidden = False
PosledniRadek = Cells(2, 14).End(xlDown).row

For CisloSloupce = 13 To 19
If Cells(1, CisloSloupce).Value = 0 Then
GoTo DALSI
Else
SloupecOznacit = CisloSloupce
End If

ReDim ZkoumanaBunka(3 To PosledniRadek) As String
For i = 3 To PosledniRadek
ZkoumanaBunka(i) = LCase(Cells(i, CisloSloupce))
Next i

WrdArray() = Split(LCase(Cells(1, CisloSloupce).Value))
For s = LBound(WrdArray) To UBound(WrdArray)
    With ActiveSheet
        For i = 3 To PosledniRadek
            If InStr(1, ZkoumanaBunka(i), WrdArray(s)) = 0 Then
                If OznacOblast Is Nothing Then
                    Set OznacOblast = .Rows(i)
                Else
                    Set OznacOblast = Union(OznacOblast, .Rows(i))
                End If
            End If
        Next i
     End With
Next s

DALSI:
Next CisloSloupce

With ActiveSheet
  If Not OznacOblast Is Nothing Then
    OznacOblast.EntireRow.Hidden = True
  End If
End With

If SloupecOznacit > 0 Then
Cells(1, SloupecOznacit).Select
Else
End If

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Sheets("PROPOSAL").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

End If
End Sub