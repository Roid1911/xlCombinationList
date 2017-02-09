'WorkSheet : Work, C2 -- InData Sheet, C3 -- Result Sheet, C4 -- Temp Sheet
Option Explicit

Sub Make()

On Error GoTo Errorcatch

Dim inSht As Worksheet
Dim tmpResSht As Worksheet
Dim resSht As Worksheet

Dim WS As Worksheet

'Check Input Data
If SheetExists(Worksheets("Work").Range("C2").value) Then
    Set inSht = ThisWorkbook.Worksheets(Worksheets("Work").Range("C2").value)
Else
    MsgBox ("입력 데이터가 들어있는 Sheet가 없음 !!")
    Exit Sub
End If

'Check Result Sheet
If SheetExists(Worksheets("Work").Range("C3").value) = False Then
    CreateSheet (Worksheets("Work").Range("C3").value)
End If
Set resSht = ThisWorkbook.Worksheets(Worksheets("Work").Range("C3").value)

'Check Temp Sheet
If SheetExists(Worksheets("Work").Range("C4").value) = False Then
    CreateSheet (Worksheets("Work").Range("C4").value)
End If
Set tmpResSht = ThisWorkbook.Worksheets(Worksheets("Work").Range("C4").value)

Dim startLine As Long
Dim endLine As Long

inSht.Activate

startLine = 2   'inSht input line
endLine = inSht.Cells(inSht.Rows.count, "A").End(xlUp).Row

'Sort
Call SortItem(inSht, startLine, endLine)


'Read & Combination & Count
Call MakeComb(inSht, startLine, endLine, resSht, tmpResSht)

'Last Beautify
Call MakeResult(tmpResSht, resSht)

resSht.Activate

Set inSht = Nothing
Set tmpResSht = Nothing
Set resSht = Nothing

Exit Sub

Errorcatch:

MsgBox err.Description

Set inSht = Nothing
Set tmpResSht = Nothing
Set resSht = Nothing

End Sub

Sub SortItem(sht As Worksheet, startLine As Long, endLine As Long)

Dim sortKey As Range
Dim sortRng As Range
Dim lastCols As Long
Dim i As Long

For i = startLine To endLine
    lastCols = sht.Cells(i, sht.Columns.count).End(xlToLeft).Column
       
    Set sortKey = sht.Range(sht.Cells(i, 1), sht.Cells(i, 1))
    Set sortRng = sht.Range(sht.Cells(i, 2), sht.Cells(i, lastCols))
    
    sortRng.sort Key1:=sortKey, Order1:=xlAscending, Orientation:=xlLeftToRight
Next i

End Sub

Sub MakeComb(iSht As Worksheet, startLine As Long, endLine As Long, oSht As Worksheet, tSht As Worksheet)

Dim i As Long
Dim lastCols As Long
Dim combRng As Range

Dim resArr As Variant
Dim hdCols As Collection

Set hdCols = New Collection

Dim j As Variant
Dim tmpVal As Long

For i = startLine To endLine
    lastCols = iSht.Cells(i, iSht.Columns.count).End(xlToLeft).Column
    
    Set combRng = iSht.Range(iSht.Cells(i, 2), iSht.Cells(i, lastCols))
    
    resArr = Util.CombRow(iSht, combRng, 2)
    
    For Each j In resArr
        
        'Debug.Print j
        
        If Util.Contains(j, hdCols) = False Then
            hdCols.Add Array(j, 1), j
        Else
            tmpVal = hdCols.Item(j)(1) + 1
            hdCols.Remove j
            hdCols.Add Array(j, tmpVal), j
        End If
        
    Next j
    
    'Debug.Print "----------------------------------"
    
    
    Call makeTemp(tSht, iSht, resArr, i, lastCols)
    
Next i

'Collections.sort hdCols

Dim index As Long
index = 2

For Each j In hdCols
    
    Debug.Print j(0) & " _ " & j(1)
    tSht.Cells(1, index).value = j(0)
    tSht.Cells(2, index).value = j(1)
    index = index + 1
    
Next j


Dim sortKey As Range
Dim sortRng As Range

lastCols = index
       
Set sortKey = tSht.Range(tSht.Cells(1, 1), tSht.Cells(1, 1))
Set sortRng = tSht.Range(tSht.Cells(1, 2), tSht.Cells(2, lastCols))
    
sortRng.sort Key1:=sortKey, Order1:=xlAscending, Orientation:=xlLeftToRight

Debug.Print "=============================================="


End Sub


Sub makeTemp(tSht As Worksheet, iSht As Worksheet, resArr As Variant, line As Long, cols As Long)

Dim i As Long
Dim lastHeadCol As Long

lastHeadCol = 0

' +2 is tSht disp line
tSht.Cells(line + 2, 1).value = iSht.Cells(line, 1).value

For i = 0 To UBound(resArr)
    'Debug.Print resArr(i)
    tSht.Cells(line + 2, 2 + i).value = resArr(i)
    
Next i


End Sub

Sub MakeResult(tSht As Worksheet, rSht As Worksheet)

Dim endLine As Long
endLine = tSht.Cells(tSht.Rows.count, "A").End(xlUp).Row

' Total Result
rSht.Cells(2, 2).value = tSht.Cells(4, 1).value
rSht.Cells(2, 3).value = "부터"
rSht.Cells(2, 4).value = tSht.Cells(endLine, 1).value
rSht.Cells(2, 5).value = "까지"

Dim index As Long
Dim lastCols As Long
Dim i As Long

index = 4
lastCols = tSht.Cells(1, tSht.Columns.count).End(xlToLeft).Column

' Index

Dim varName() As String

For i = 2 To lastCols
    
    rSht.Cells(index, 1) = index - 3
    varName = Split(tSht.Cells(1, i).value, "_")
    
    rSht.Cells(index, 2) = varName(0)
    rSht.Cells(index, 3) = varName(1)
    rSht.Cells(index, 4) = tSht.Cells(2, i).value
    
    index = index + 1

Next i


End Sub

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
End Function

Sub CreateSheet(sName As String)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = sName
    End With
End Sub
