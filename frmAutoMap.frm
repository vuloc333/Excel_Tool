VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoMap 
   Caption         =   "UserForm1"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "frmAutoMap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutoMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long, n As Long
    lstHeaders.Clear
    If AC_SelectGroupColumn Then
        Me.Caption = "Select Grouping Column"
        lstHeaders.MultiSelect = 0
        ReDim AC_ListColumnIndex(1 To UBound(AC_HeaderList) - LBound(AC_HeaderList) + 1)
        For i = LBound(AC_HeaderList) To UBound(AC_HeaderList)
            n = n + 1
            AC_ListColumnIndex(n) = i
            lstHeaders.AddItem CStr(AC_HeaderList(i))
        Next i
    Else
        Me.Caption = "Select Columns to Remove"
        lstHeaders.MultiSelect = 1
        If UBound(AC_HeaderList) > LBound(AC_HeaderList) Then
            ReDim AC_ListColumnIndex(1 To UBound(AC_HeaderList) - LBound(AC_HeaderList))
        Else
            Erase AC_ListColumnIndex
        End If
        For i = LBound(AC_HeaderList) To UBound(AC_HeaderList)
            If i <> AC_GroupIndex Then
                n = n + 1
                AC_ListColumnIndex(n) = i
                lstHeaders.AddItem CStr(AC_HeaderList(i))
            End If
        Next i
    End If
End Sub

Private Sub btnOK_Click()
    Dim i As Long, n As Long
    If AC_SelectGroupColumn Then
        For i = 0 To lstHeaders.ListCount - 1
            If lstHeaders.Selected(i) Then
                AC_GroupIndex = AC_ListColumnIndex(i + 1)
                AC_Cancelled = False
                Me.Hide
                Exit Sub
            End If
        Next i
        MsgBox "Select one grouping column.", vbExclamation
        Exit Sub
    End If

    For i = 0 To lstHeaders.ListCount - 1
        If lstHeaders.Selected(i) Then n = n + 1
    Next i

    AC_RemoveCount = n
    If n > 0 Then
        ReDim AC_RemoveColumns(1 To n)
        n = 0
        For i = 0 To lstHeaders.ListCount - 1
            If lstHeaders.Selected(i) Then
                n = n + 1
                AC_RemoveColumns(n) = AC_ListColumnIndex(i + 1)
            End If
        Next i
    End If
    AC_Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    AC_Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then AC_Cancelled = True
End Sub

