Attribute VB_Name = "Module1"
Option Explicit

Public DB As Database
Public rsDRH As Recordset
Public rsDRD As Recordset
Public rsCF As Recordset
Public rsINV As Recordset
Public rsEF As Recordset

Public Sub setDBaseTable()
    Set DB = OpenDatabase(App.Path & "\IdealHome.mdb")
    Set rsDRH = DB.OpenRecordset("DRHEADER")
    Set rsDRD = DB.OpenRecordset("DRDETAIL")
    Set rsCF = DB.OpenRecordset("CUSTOMERFILE")
    Set rsINV = DB.OpenRecordset("INVENTORY")
    Set rsEF = DB.OpenRecordset("EMPLOYEEFILE")
End Sub

Public Function checkDRHEADERFILE(num As String) As Boolean
    If Not rsDRH.BOF Then
        rsDRH.MoveFirst
    End If
    rsDRH.Index = "ndxDRHEADER"
    rsDRH.Seek "=", Trim(num)
    If rsDRH.NoMatch Then
        checkDRHEADERFILE = True
    Else
        checkDRHEADERFILE = False
    End If
End Function

Public Function checkDRDETAILFILE(num As String, num1 As String) As Boolean
    If Not rsDRD.BOF Then
        rsDRD.MoveFirst
    End If
    rsDRD.Index = "ndxDRDETAIL"
    rsDRD.Seek "=", Trim(num), Trim(num1)
    If rsDRD.NoMatch Then
        checkDRDETAILFILE = True
    Else
        checkDRDETAILFILE = False
    End If
End Function

Public Function checkCUSTOMERFILE(num As String) As Boolean
    If Not rsCF.BOF Then
        rsCF.MoveFirst
    End If
    rsCF.Index = "ndxCUSTOMERFILE"
    rsCF.Seek "=", Trim(num)
    If rsCF.NoMatch Then
        checkCUSTOMERFILE = True
    Else
        checkCUSTOMERFILE = False
    End If
End Function

Public Function checkINVENTORYFILE(num As String) As Boolean
    If Not rsINV.BOF Then
        rsINV.MoveFirst
    End If
    rsINV.Index = "ndxINVENTORY"
    rsINV.Seek "=", Trim(num)
    If rsINV.NoMatch Then
        checkINVENTORYFILE = True
    Else
        checkINVENTORYFILE = False
    End If
End Function

Public Function checkEMPLOYEEFILE(num As String) As Boolean
    If Not rsEF.BOF Then
        rsEF.MoveFirst
    End If
    rsEF.Index = "ndxEMPLOYEEFILE"
    rsEF.Seek "=", Trim(num)
    If rsEF.NoMatch Then
        checkEMPLOYEEFILE = True
    Else
        checkEMPLOYEEFILE = False
    End If
End Function

Public Function isLetter(str As String) As Boolean
    Dim i As Integer
    Dim schr As String
    Dim lchr As String
    str = Trim(str)
    For i = 1 To Len(str)
        schr = Mid(str, i, 1)
        lchr = asc(schr)
        Select Case lchr
            Case 65 To 90, 97 To 122
                isLetter = True
            Case Else
                isLetter = False
                Exit For
        End Select
    Next i
End Function

Public Function checkLetter(str As String) As Boolean
    Dim i As Integer
    Dim schr As String
    Dim lchr As String
    str = Trim(str)
    For i = 1 To Len(str)
        schr = Mid(str, i, 1)
        lchr = asc(schr)
        Select Case lchr
            Case 65 To 90, 97 To 122
                checkLetter = True
            Case Else
                checkLetter = False
                Exit For
        End Select
    Next i
End Function

Public Function checkSpecialChar(str As String) As Boolean
    Dim i As Integer
    Dim schr As String
    Dim lchr As Integer
    str = Trim(str)
    For i = 1 To Len(str)
        schr = Mid(str, i, 1)
        lchr = asc(schr)
        Select Case lchr
            Case 33 To 47, 58 To 64
                checkSpecialChar = True
            Case Else
                checkSpecialChar = False
                Exit For
        End Select
    Next i
End Function

Public Function isIntegral(str As String) As Boolean
    Dim i As Integer
    Dim schr As String
    Dim lchr As Integer
    str = Trim(str)
    For i = 1 To Len(str)
        schr = Mid(str, i, 1)
        lchr = asc(schr)
        Select Case lchr
            Case 48 To 57
                isIntegral = True
            Case Else
                isIntegral = False
                Exit For
        End Select
    Next i
End Function

 


