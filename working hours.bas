Attribute VB_Name = "Module1"
Sub nowy_miesiac()
Attribute nowy_miesiac.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' Test Macro
'
' Keyboard Shortcut: Ctrl+n
'
Dim monthq As String
Dim yearq As String
Dim nazwa As String
Dim dzienq As String
Dim ws1 As Worksheet
monthq = InputBox("Podaj miesiac ")
yearq = InputBox("Podaj rok, (format yyyy)")
dzienq = InputBox("Podaj 1 dzieñ tygodnia w danym miesiacu, z duzej litery i z polskimi znakami")
If dzienq <> "Poniedzia³ek" And _
dzienq <> "Wtorek" And dzienq <> "Œroda" _
And dzienq <> "Czwartek" And dzienq <> "Pi¹tek" And dzienq _
<> "Sobota" And dzienq <> "Niedziela" Then
MsgBox ("Podales zle dzien tygodnia, usun nowo powstaly arkusz i sproboj jeszcze raz")
End If
' Last check :D if you did mistake , press no then it doesnt do anything
Msg = "Czy to s¹ poprawne dane: " & vbNewLine & "Rok: " _
& yearq & vbNewLine & "Miesi¹c: " _
& monthq & vbNewLine & "Dzieñ: " & dzienq & vbNewLine & "?"
ans = MsgBox(Msg, vbYesNo)
If ans = vbNo Then
    MsgBox "Wpisz dane jeszcze raz"
    Exit Sub
End If
If ans = vbYes Then
    MsgBox "Utworzono arkusz " & monthq & yearq
End If

Set ws1 = ThisWorkbook.Worksheets("templatka")
ws1.Copy ThisWorkbook.Sheets(Sheets.Count)
nazwa = monthq & yearq
ActiveSheet.Name = nazwa
ActiveSheet.Range("p4").Value = dzienq
'Autofill depends on 1st day name of month
Set SourceRange = ActiveSheet.Range("p4")
Set fillRange = ActiveSheet.Range("p4:p34")
SourceRange.AutoFill Destination:=fillRange

' i know that isn t a best way but im noob at vba :D
' Setting colours
' dla Poniedzialek
If ActiveSheet.Range("p4").Value = "Poniedzia³ek" Then
' Sobota
ActiveSheet.Range("A9:p9").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A16:p16").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a23:p23").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a30:p30").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a10:p10").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a17:p17").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a24:p24").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a31:p31").Interior.Color = RGB(255, 0, 0)
End If

' dla Wtorek
If ActiveSheet.Range("p4").Value = "Wtorek" Then
' Sobota
ActiveSheet.Range("A8:p8").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A15:p15").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a22:p22").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a29:p29").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a9:p9").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a16:p16").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a23:p23").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a30:p30").Interior.Color = RGB(255, 0, 0)
End If

' dla Œroda
If ActiveSheet.Range("p4").Value = "Œroda" Then
' Sobota
ActiveSheet.Range("A7:p7").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A14:p14").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a21:p21").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a28:p28").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a8:p8").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a15:p15").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a22:p22").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a29:p29").Interior.Color = RGB(255, 0, 0)
End If

' dla Czwartek
If ActiveSheet.Range("p4").Value = "Czwartek" Then
' Sobota
ActiveSheet.Range("A6:p6").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A13:p13").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a20:p20").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a27:p27").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a34:p34").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a7:p7").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a14:p14").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a21:p21").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a28:p28").Interior.Color = RGB(255, 0, 0)
End If

' dla Pi¹tek
If ActiveSheet.Range("p4").Value = "Pi¹tek" Then
' Sobota
ActiveSheet.Range("A5:p5").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A12:p12").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a19:p19").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a26:p26").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a33:p33").Interior.Color = RGB(255, 0, 0)
' Niedziela
ActiveSheet.Range("a6:p6").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a13:p13").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a20:p20").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a27:p27").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a34:p34").Interior.Color = RGB(255, 0, 0)

End If

' dla Sobota
If ActiveSheet.Range("p4").Value = "Sobota" Then
' Sobota
ActiveSheet.Range("A4:p4").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A11:p11").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a18:p18").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a25:p25").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a32:p32").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a5:p5").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a12:p12").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a19:p19").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a26:p26").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a33:p33").Interior.Color = RGB(255, 0, 0)
End If

' dla Niedziela
If ActiveSheet.Range("p4").Value = "Niedziela" Then
' Sobota
ActiveSheet.Range("A10:p10").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("A17:p17").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a24:p24").Interior.Color = RGB(153, 204, 255)
ActiveSheet.Range("a31:p31").Interior.Color = RGB(153, 204, 255)
' Niedziela
ActiveSheet.Range("a4:p4").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a11:p11").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a18:p18").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a25:p25").Interior.Color = RGB(255, 0, 0)
ActiveSheet.Range("a32:p32").Interior.Color = RGB(255, 0, 0)
End If

    Columns("Q:Q").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    Rows("39:39").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select

   ' Application.Goto Reference:="nowy_miesiac"
    
End Sub

Sub Dany_miesiac()
Attribute Dany_miesiac.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' Dany_miesiac Macro
'
' Keyboard Shortcut: Ctrl+d
'
Dim monthq As String
Dim yearq As String
Dim nazwa As String
monthq = InputBox("Podaj miesiac (Z duzej litery, bez polskich znakow)")
yearq = InputBox("Podaj rok, (format yyyy)")
nazwa = monthq & yearq
Sheets(nazwa).Select

End Sub



Sub zmien_wzor()
Attribute zmien_wzor.VB_ProcData.VB_Invoke_Func = "p\n14"

'Keyboard Shortcut: Ctrl+p
Sheets("templatka").Select

End Sub
Sub Test()
'
' test Macro
'
Msg = "Czy nazywasz siê " & Application.UserName & "?"
ans = MsgBox(Msg, vbYesNo)
If ans = vbNo Then MsgBox "sorry, niewazne"
If ans = vbYes Then MsgBox "ha , udalo mi sie"

End Sub

