Sub clan_phone()
    
    Dim i           As Long
    Dim j           As Long
    Dim li          As Long
    Dim new_colum      As Long
    
    Dim d1          As Long
    Dim d2          As Long
    
    i = 1
    
    j = 1        ' colum 1=A, 6= colum F (you need change it in your table)
    new_colum = j + 1
    
    ' -------------------------------------------------------------------------------------------------------
    ' copy the existing column
    
    Columns(new_colum).EntireColumn.Insert
    Columns(j).Copy Columns(new_colum)        ' j -  from, new_colum - to
    
    Dim Column      As Range: Set Column = Application.Columns(new_colum)
    ' -------------------------------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------------------------------
    'remove extra characters / text
    
    Column.Replace "(", ""
    Column.Replace ")", ""
    Column.Replace "-", ""
    Column.Replace "+", ""
    Column.Replace "_", ""
    Column.Replace ":", ""
    Column.Replace "~?", ""
    Column.Replace "~*", ""
    Column.Replace "@", ""
    Column.Replace """", ""
    Column.Replace "'", ""
    Column.Replace "`", ""
    'replace
    Column.Replace ",", ";"
    Column.Replace ".", ";"
    Column.Replace "/", ";"
    Column.Replace ",", ";"
    Column.Replace ", ", ";"
    Column.Replace ", ", ";"
    Column.Replace ",    ", ";"
    Column.Replace " ", ""
    Column.Replace " ", ""
    Column.Replace " ", ""
    
    Column.Replace "    ", ""
    
    'delete text  (ASCII see: https://www.ascii-codes.com/cp855.html )
    
    Column.Replace Chr(10), ""
    Column.Replace Chr(13), ""
    
    For li = 97 To 122        ' A-Z
        Column.Replace Chr(li), ""
    Next li
    
    For li = 65 To 90        'a-z
        Column.Replace Chr(li), ""
    Next li
    
    For li = 128 To 225        'Cyrillic
        Column.Replace Chr(li), ""
    Next li
    
    Del_Rus = sStr
    ' -------------------------------------------------------------------------------------------------------
    ' validation if 2 numbers
    
    r = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row        ' A - main colum in DB (primer key)
    
    ' --------
    
    For i = r To 2 Step -1
        
        d1 = InStr(1, Cells(i, new_colum), ";") - 1
        d2 = Len(Cells(i, new_colum)) - d1 - 1
        'Cells(i, new_colum + 5) = d1
        'Cells(i, new_colum + 6) = d2
        ' for d .
        
        If d1 >= 9 And d1 > -1 And d2 < 9 And d2 > -1 Then
            Cells(i, new_colum) = Left(Cells(i, new_colum), d1)
            
        ElseIf d1 < 9 And d1 > -1 And d2 >= 9 And d2 > -1 Then
            Cells(i, new_colum) = Right(Cells(i, new_colum), Len(Cells(i, new_colum)) - d1 - 1)
            
        ElseIf d1 >= 9 And d2 >= 9 Then
            Cells(i, new_colum) = Left(Cells(i, new_colum), d1)
            
        ElseIf d1 < 1 And d2 < 1 Then
            Cells(i, new_colum) = 0
            
        ElseIf d1 >= d2 Then
            
            Cells(i, new_colum) = Left(Cells(i, new_colum), d1)
        ElseIf d1 < d2 Then
            
            Cells(i, new_colum) = Right(Cells(i, new_colum), Len(Cells(i, new_colum)) - d1 - 1)
        Else
            Cells(i, new_colum) = 100
        End If
        
    Next
    ' -------------------------------------------------------------------------------------------------------
    
    ' del not valid nomber
    
    For i = r To 2 Step -1
        If Len(Cells(i, new_colum)) >= 9 Then
            Cells(i, new_colum) = Right(Cells(i, new_colum), 9)
        End If
        Cells(i, new_colum) = (380 & Cells(i, new_colum))
        Cells(i, new_colum).NumberFormat = "0"
        'Columns(new_colum+1).EntireColumn.Insert
        'Cells(i, new_colum+1) =  Cells(i, new_colum)* 1
        
    Next
    
    Cells(1, new_colum) = "valid_phone"
    
End Sub
