
Private Sub DelayMs(ms As Long)
    Debug.Print TimeValue(Now)
    Application.Wait (Now + (ms * 0.00000001))
    Debug.Print TimeValue(Now)
End Sub


Sub CLEAR()
Set aRange = Sheets("TLO BOT").Range("A5.ZZ50000")
aRange.ClearContents
End Sub


Sub TLO_Add()


Dim CurrentHost As Object
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession
Dim irow As Long
irow = 5
  file = Range("A" & irow).Value
  
  exclude = Array("-", "_", ",", "\", "/", ".")

Call Copy

Call DelayMs(200)

Call Parse_Addresses
  
  Do
  
    result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
  
    If Range("A" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
    
    If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
    CurrentHost.Output Range("A" & irow).Value & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output Range("A" & irow).Value & ChrW$(13)
    End If
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,LH,C,S,Dn,GC#,/,-,." Then
    CurrentHost.Output "2" & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output "2" & ChrW$(13)
    End If
 
    Call DelayMs(1000)
    CurrentHost.Output Range("T" & irow).Value & ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output Range("AD" & irow).Value & " " & Range("AE" & irow).Value
    Call DelayMs(200)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    
    'Street Parse Begins
    Range("U" & irow) = Trim(Range("U" & irow))
    Range("V" & irow) = Trim(Range("V" & irow))

If InStr(1, Range("U" & irow), "PO BOX", vbTextCompare) > 0 Then
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(2))
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
 GoTo END_PARSE
End If

If Not IsNumeric(Left(Range("U" & irow), 1)) Then
celltosplit = Range("V" & irow)
result = Trim(Split(celltosplit, " ")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
result = Trim(Split(celltosplit, " ")(1))
On Error Resume Next
result2 = Trim(Split(celltosplit, " ")(2))
result3 = Trim(Split(celltosplit, " ")(3))
result4 = Trim(Split(celltosplit, " ")(4))
result5 = result & " " & result1 & " " & result2 & " " & result3 & " " & result4
CurrentHost.Output result5
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
    GoTo END_PARSE
End If

If InStr(1, Range("U" & irow), " APT ", vbTextCompare) > 0 Then
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(1))
result = Trim(Split(result, "APT")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
     GoTo END_PARSE
End If

If InStr(1, Range("U" & irow), " LOT ", vbTextCompare) > 0 Then
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(1))
result = Trim(Split(result, "LOT")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
     GoTo END_PARSE
End If

If InStr(1, Range("U" & irow), " UNIT ", vbTextCompare) > 0 Then
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(1))
result = Trim(Split(result, "UNIT")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
     GoTo END_PARSE
End If

    Call DelayMs(800)
celltosplit = Range("U" & irow)
result = Trim(Split(celltosplit, " ")(0))
CurrentHost.Output result
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
celltosplit = Range("U" & irow)

On Error Resume Next

result = Trim(Split(celltosplit, " ")(1))
result2 = Trim(Split(celltosplit, " ")(2))
result3 = Trim(Split(celltosplit, " ")(3))
result4 = Trim(Split(celltosplit, " ")(4))
result5 = result & " " & result1 & " " & result2 & " " & result3 & " " & result4
CurrentHost.Output result5
    CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        CurrentHost.Output ChrW$(13)
    Call DelayMs(800)
        result = ""
    result2 = ""
    result3 = ""
    result4 = ""
    result5 = ""
    'Parse Ends
END_PARSE:
    
    
    CurrentHost.Output Range("W" & irow).Value & ChrW$(13)
    Call DelayMs(800)
    CurrentHost.Output Range("X" & irow).Value & ChrW$(13)
    Call DelayMs(800)
    
    If Len(Range("Y" & irow).Value) = 3 Then
    CurrentHost.Output "00" & Range("Y" & irow).Value
    Call DelayMs(200)
    CurrentHost.Output ChrW$(13)
    ElseIf Len(Range("Y" & irow).Value) = 4 Then
    CurrentHost.Output "0" & Range("Y" & irow).Value
    Call DelayMs(200)
    CurrentHost.Output ChrW$(13)
    Else
    CurrentHost.Output Range("Y" & irow).Value & ChrW$(13)
    End If
    
    Call DelayMs(800)
    CurrentHost.Output Left(Range("Z" & irow).Value, 3) & "-" & Mid(Range("Z" & irow).Value, 4, 3) & "-" & Right(Range("Z" & irow).Value, 4)
    Call DelayMs(200)
    
    CurrentHost.Output ChrW$(13)
    
    Call DelayMs(800)
    CurrentHost.Output "//" & ChrW$(13)
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 24) = "OK TO FILE  (CR=Y,/,/nn)" Then
    CurrentHost.Output ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output ChrW$(13)
    End If
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,LH,C,S,Dn,GC#,/,-,." Then
    CurrentHost.Output "4" & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output "4" & ChrW$(13)
    End If
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 17) = "ENTER WHAT (nn,X)" Then
    CurrentHost.Output "18" & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output "18" & ChrW$(13)
    End If
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 16) = "ENTER WHO (nn,/)" Then
    CurrentHost.Output "17" & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output "17" & ChrW$(13)
    End If
    
    Call DelayMs(800)
    CurrentHost.Output "TLO POE Verifation Notes:" & " " & Range("AG" & irow).Value
    Call DelayMs(800)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(100)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(100)
    CurrentHost.Output ChrW$(13)
    Call DelayMs(100)
    CurrentHost.Output ChrW$(13)
       
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,LH,C,S,Dn,GC#,/,-,." Then
    CurrentHost.Output "/" & ChrW$(13)
    Else
    Call DelayMs(600)
    CurrentHost.Output "/" & ChrW$(13)
    End If
    
    Call DelayMs(200)
    
    If CurrentHost.GetText(0, 22, 19) = "ENTER WHAT (nn)" Then
    CurrentHost.Output "16" & ChrW$(13)
    End If
    
    If CurrentHost.GetText(0, 22, 16) = "ENTER WHO (nn,/)" Then
    CurrentHost.Output "17" & ChrW$(13)
    End If
    
    If CurrentHost.GetText(0, 22, 19) = "ENTER RESULT (nn,/)" Then
    CurrentHost.Output "12" & ChrW$(13)
    End If
    
    Call DelayMs(600)
    CurrentHost.Output "/" & ChrW$(13)
    
    Call DelayMs(1000)
    
    If CurrentHost.GetText(0, 22, 17) = "ENTER RESULT (nn)" Then
            CurrentHost.Output "12" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "12" & ChrW$(13)
            End If

    

    irow = irow + 1
    Loop

End Sub


Sub Parse_Addresses()
     
    Dim sSplitAddr() As String
    Dim vPart As Variant
    Dim vStreets As Variant
    Dim i As Integer
    Dim iFound As Integer
    Dim vType As Variant
    Dim validStreet As Variant
    Dim sBuildAddress As String
    Dim rProcessCell As Range
     
     'This code will handle well formed addresses and
     ' highlight deficient ones and PO Boxes
     
    Set rProcessCell = Sheets("Sheet1").Range("AI5") 'Sheet 2 is a working copy of sheet 1
    Do While rProcessCell.Value <> ""
        sSplitAddr = Split(Trim(rProcessCell.Value), " ")
        If UBound(sSplitAddr) = 0 Then GoTo DontProcessFurther
         
         'Check whether Direction is in address, if not set blank position
        If Not HasDirn(sSplitAddr(1)) Then
            ReDim Preserve sSplitAddr(UBound(sSplitAddr) + 1)
            For i = UBound(sSplitAddr) To 2 Step -1
                sSplitAddr(i) = sSplitAddr(i - 1)
            Next
            sSplitAddr(1) = ""
        End If
         
         ' Check whether any street type suffixes are in the address
        vStreets = Array("ST", "TER", "DR", "LN", "RD", "CT", "AVE")
        iFound = 0
        For Each vType In vStreets
            For i = 3 To UBound(sSplitAddr)
                If sSplitAddr(i) = vType Then
                    validStreet = True
                    iFound = i
                    Exit For
                End If
            Next i
        Next
         
        If iFound > 3 Then
             'Street has a two or more word name then combine name and contract array
            sBuildAddress = ""
            For i = 2 To iFound - 1
                sBuildAddress = sBuildAddress & " " & sSplitAddr(i)
            Next i
            sBuildAddress = Mid(sBuildAddress, 2, Len(sBuildAddress) - 1)
            sSplitAddr(2) = sBuildAddress
            For i = iFound To UBound(sSplitAddr)
                sSplitAddr(i - iFound + 3) = sSplitAddr(i)
            Next i
            ReDim Preserve sSplitAddr(UBound(sSplitAddr) + 3 - iFound)
        End If
         
         'check last address part for # and remove
        If Left(sSplitAddr(UBound(sSplitAddr)), 1) = "#" Then
            sSplitAddr(UBound(sSplitAddr)) = Right(sSplitAddr(UBound(sSplitAddr)), Len(sSplitAddr(UBound(sSplitAddr))) - 1)
        End If
         'check last address part for apt and remove
        If Left(sSplitAddr(UBound(sSplitAddr)), 3) = "APT" Then
            sSplitAddr(UBound(sSplitAddr)) = Right(sSplitAddr(UBound(sSplitAddr)), Len(sSplitAddr(UBound(sSplitAddr))) - 3)
        End If
         
DontProcessFurther:
        rProcessCell.Offset(0, 1).Resize(, UBound(sSplitAddr) + 1).Value = sSplitAddr
         'check for badly formed address
        Highlight_Bad_Addresses rProcessCell, vStreets
        Set rProcessCell = rProcessCell.Offset(1, 0)
    Loop
End Sub
 
Function HasDirn(checkDirn As String)
     
    Dim test As Variant
    HasDirn = False
    Dim sDirn() As Variant
     
    sDirn = Array("N", "NE", "E", "E", "S", "SW", "W", "NW")
    For Each test In sDirn
        If checkDirn = test Then HasDirn = True
    Next
     
End Function
 
Sub Highlight_Bad_Addresses(rProcessCell As Range, vStreets As Variant)
     'if Column E doesn't contain a valid street type it will be considered
     ' to be a badly formed address and highlihgted for manual attention
    Range(rProcessCell.Offset(0, 1), rProcessCell.Offset(0, 7)).Interior.ColorIndex = 0
    If UBound(Filter(vStreets, rProcessCell.Offset(0, 4).Value)) < 0 Or rProcessCell.Offset(0, 4).Value = "" Then
        Range(rProcessCell.Offset(0, 1), rProcessCell.Offset(0, 7)).Interior.ColorIndex = 3
    End If
     
End Sub

Sub Copy()
Range("U5:U10000").Copy Destination:=Range("AI5")

End Sub


Function AlphaNumericOnly(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 32, 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult
End Function


Sub Lowhorn()

Dim irow As Long
irow = 5

Do

If Range("A" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
    
    Range("U" & irow).Value = AlphaNumericOnly(Range("U" & irow).Value)

    irow = irow + 1
Loop



End Sub


