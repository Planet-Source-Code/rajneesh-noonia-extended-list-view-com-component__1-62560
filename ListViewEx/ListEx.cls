VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListEx"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum CompareTypes
    lvDate = 1
    lvLong = 2
    lvString = 3
    lvCustom = 4
End Enum

Dim m_ListViewCtrls As New lvCtrls

Public Property Let ShowHeaderIcon(ByVal Hwnd As Long, ByVal Value As Boolean)
    GetListViewCtrl(Hwnd).ShowIcon = Value
End Property

Private Function GetListViewCtrl(ByVal Hwnd As Long) As lvCtrl
    Set GetListViewCtrl = m_ListViewCtrls.Item("H" & Hwnd)
End Function

Public Sub SortColumn(ByVal Hwnd As Long, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call m_ListViewCtrls.Item("H" & Hwnd).ColumnSort(ColumnHeader)
End Sub
Public Sub Attach(ByRef Client As Object, ByRef ListViewCtrl As Object, ByVal ColIndex As Integer, ByVal CompType As CompareTypes, Optional ByVal ImgUPArrowIndex As Long = -1, Optional ByVal ImgDownArrowIndex As Long = -1)
    Dim pListViewCtrl As lvCtrl
    Dim plvColumn As lvColumn
    Dim pblnGridFound As Boolean
    Dim pblnColumnFound As Boolean
    Dim plvNewColumns As lvColumns
    If Not (TypeOf ListViewCtrl Is ListView) Then
        MsgBox "This is only implemented for ListView Control", vbCritical, "Invalid Parameter Passed"
        Exit Sub
    End If
    For Each pListViewCtrl In m_ListViewCtrls
        If (pListViewCtrl.ListViewCtrl.Hwnd = ListViewCtrl.Hwnd) Then
            pblnGridFound = True
            For Each plvColumn In pListViewCtrl.ListViewlvColumns
                If (plvColumn.lvColumnIndex = ColIndex) Then
                    pblnColumnFound = True
                    plvColumn.SortType = CompType
                    Exit For
                End If
            Next
            If Not (pblnColumnFound) Then
                pListViewCtrl.ListViewlvColumns.Add ColIndex, CompType, "H" & ListViewCtrl.Hwnd & "C" & ColIndex
                Exit Sub
            End If
        End If
    Next
    
    If Not (pblnGridFound) Then
        Set plvNewColumns = New lvColumns
        plvNewColumns.Add ColIndex, CompType
        If Not (TypeOf Client Is ICompare) Then
            m_ListViewCtrls.Add ListViewCtrl, Nothing, plvNewColumns, "H" & ListViewCtrl.Hwnd, ImgUPArrowIndex, ImgDownArrowIndex
        Else
            m_ListViewCtrls.Add ListViewCtrl, Client, plvNewColumns, "H" & ListViewCtrl.Hwnd, ImgUPArrowIndex, ImgDownArrowIndex
        End If
        
    End If
End Sub
