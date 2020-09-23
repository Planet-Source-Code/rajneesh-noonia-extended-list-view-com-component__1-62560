VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lvCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public ListViewlvColumns As lvColumns
Public WithEvents ListViewCtrl As ListView
Attribute ListViewCtrl.VB_VarHelpID = -1
Public lvClient As ICompare
Public SortAsc As Boolean
Public ShowIcon As Boolean
Public ImgUPArrowIndex As Long
Public ImgDownArrowIndex As Long

Private Sub ShowHeaderIcon(colNo As Long, imgIconNo As Long, showImage As Boolean)

   Dim hHeader As Long
   Dim HD As HD_ITEM
   Dim pOldHeaderImgLst As ImageList
   
  'get a handle to the listview header component
   hHeader = SendMessage(ListViewCtrl.Hwnd, LVM_GETHEADER, 0, ByVal 0)
   
  'set up the required structure members
   With HD
      .mask = HDI_IMAGE Or HDI_FORMAT
      .pszText = ListViewCtrl.ColumnHeaders(colNo + 1).Text
      
       If showImage Then
         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
         .iImage = imgIconNo
       Else
         .fmt = HDF_STRING
      End If


   End With
   
  'modify the header
   Set pOldHeaderImgLst = ListViewCtrl.ColumnHeaderIcons
   If (pOldHeaderImgLst Is Nothing) Then
      Set ListViewCtrl.ColumnHeaderIcons = frmResource.ImageList1
   End If
   Call SendMessage(hHeader, HDM_SETITEM, colNo, HD)
   'Set ListViewCtrl.ColumnHeaderIcons = pOldHeaderImgLst
End Sub


Private Sub Class_Initialize()
    ImgUPArrowIndex = -1
    ImgDownArrowIndex = -1
End Sub

Public Sub ColumnSort(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim plvCol As lvColumn
    Dim pblnFound As Boolean
    Static LastSelection As Long
    
    If (LastSelection = ColumnHeader.Index Or LastSelection = 0) Then
        SortAsc = Not SortAsc
    End If
    
    Call SetProp(ListViewCtrl.Hwnd, "ColumnIndex", ColumnHeader.Index - 1)
    If Not (lvClient Is Nothing) Then
        Call SetProp(ListViewCtrl.Hwnd, "ClientRef", ObjPtr(lvClient))
    End If
    
    Call SetProp(ListViewCtrl.Hwnd, "SortOrder", SortAsc)
    
    
    If (LastSelection <> 0) Then Call ShowHeaderIcon(LastSelection - 1, 0, False)
    LastSelection = ColumnHeader.Index
    
    If (SortAsc) Then
        If (ImgUPArrowIndex > -1 And Not (ListViewCtrl.ColumnHeaderIcons Is Nothing)) Then
            Call ShowHeaderIcon(ColumnHeader.Index - 1, ImgUPArrowIndex, ShowIcon)
        Else
            Call ShowHeaderIcon(ColumnHeader.Index - 1, 0, ShowIcon)
        End If
    Else
        If (ImgUPArrowIndex > -1 And Not (ListViewCtrl.ColumnHeaderIcons Is Nothing)) Then
            Call ShowHeaderIcon(ColumnHeader.Index - 1, ImgDownArrowIndex, ShowIcon)
        Else
            Call ShowHeaderIcon(ColumnHeader.Index - 1, 1, ShowIcon)
        End If
    End If
    
    For Each plvCol In ListViewlvColumns
        If (plvCol.lvColumnIndex = ColumnHeader.Index) Then
            Select Case plvCol.SortType
                Case CompareTypes.lvCustom:
                    ListViewCtrl.Sorted = False
                    Call SetProp(ListViewCtrl.Hwnd, "CheckIntImp", True)
                    SendMessage ListViewCtrl.Hwnd, LVM_SORTITEMS, ListViewCtrl.Hwnd, ByVal FARPROC(AddressOf CompareCustom)
                Case CompareTypes.lvDate:
                    ListViewCtrl.Sorted = False
                    SendMessage ListViewCtrl.Hwnd, LVM_SORTITEMS, ListViewCtrl.Hwnd, ByVal FARPROC(AddressOf CompareDates)
                Case CompareTypes.lvLong:
                    ListViewCtrl.Sorted = False
                    SendMessage ListViewCtrl.Hwnd, LVM_SORTITEMS, ListViewCtrl.Hwnd, ByVal FARPROC(AddressOf CompareValues)
                Case CompareTypes.lvString
                    ListViewCtrl.SortKey = ColumnHeader.Index - 1
                    If (SortAsc) Then
                        ListViewCtrl.SortOrder = lvwAscending
                    Else
                        ListViewCtrl.SortOrder = lvwDescending
                    End If
                    ListViewCtrl.Sorted = True
            End Select
        Exit For
        End If
    Next
    
End Sub

Private Sub ListViewCtrl_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ColumnSort(ColumnHeader)
  
End Sub

