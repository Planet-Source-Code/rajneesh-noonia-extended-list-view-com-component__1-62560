Attribute VB_Name = "modCompare"
Option Explicit

Private objFind As LV_FINDINFO
Private objItem As LV_ITEM
  
Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Private Const LVM_FIRST As Long = &H1000

Public Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2

Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_IMAGE = &H800
Public Const HDF_STRING = &H4000

Private Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)
Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

Public Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   iImage As Long
   iOrder As Long
End Type

 
 
'Constants
Private Const LVFI_PARAM As Long = &H1
Private Const LVIF_TEXT As Long = &H1


Private Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
     
'API declarations

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal Hwnd As Long, ByVal lpString As String) As Long


Public Function CompareCustom(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal Hwnd As Long) As Long
    Dim pClientCompClass As Long
    Dim iwp As ICompare
    Dim iwpT As ICompare
    Dim pColIndex As Long
    Dim dValue1 As Variant
    Dim dValue2 As Variant
    Dim pOrder As Boolean
    Dim pblnCheckInterfaceImp As Boolean
    pColIndex = GetProp(Hwnd, "ColumnIndex")
    pOrder = GetProp(Hwnd, "SortOrder")
    pClientCompClass = GetProp(Hwnd, "ClientRef")
    pblnCheckInterfaceImp = GetProp(Hwnd, "CheckIntImp")
    
    Call SetProp(Hwnd, "CheckIntImp", False)
    
    If (pClientCompClass <> 0) Then
        CopyMemory iwpT, pClientCompClass, 4
        Set iwp = iwpT
        CopyMemory iwpT, 0&, 4
    End If
 
  'Obtain the item names and dates corresponding to the
  'input parameters

   dValue1 = CVar(GetItemData(Hwnd, pColIndex, lParam1))
   dValue2 = CVar(GetItemData(Hwnd, pColIndex, lParam2))
     
   If Not (iwp Is Nothing) Then
        CompareCustom = iwp.Compare(Hwnd, pColIndex, pOrder, dValue1, dValue2)
   Else
       If (pblnCheckInterfaceImp) Then
        MsgBox "Implement Interface 'ICompare' For Custom Sorting", vbExclamation, "Interface Not Implemented"
       End If
   End If
   
End Function


Public Function CompareDates(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal Hwnd As Long) As Long
    Dim pColIndex As Long
    Dim pOrder As Boolean
    
    pColIndex = GetProp(Hwnd, "ColumnIndex")
    pOrder = GetProp(Hwnd, "SortOrder")
   'CompareDates: This is the sorting routine that gets passed to the
   'ListView control to provide the comparison test for date values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than

   Dim dDate1 As Date
   Dim dDate2 As Date
     
  'Obtain the item names and dates corresponding to the
  'input parameters

   dDate1 = CDate(GetItemData(Hwnd, pColIndex, lParam1))
   dDate2 = CDate(GetItemData(Hwnd, pColIndex, lParam2))
     
   
  'based on the Public variable sOrder set in the
  'ColumnHeader click sub, sort the dates appropriately:
  
   Select Case pOrder
      Case True 'sort descending
            
            If dDate1 < dDate2 Then
               CompareDates = 0
            ElseIf dDate1 = dDate2 Then
               CompareDates = 1
            Else
               CompareDates = 2
            End If
      
      Case Else 'sort ascending
   
            If dDate1 > dDate2 Then
               CompareDates = 0
            ElseIf dDate1 = dDate2 Then
               CompareDates = 1
            Else
               CompareDates = 2
            End If
   
   End Select

End Function


Public Function CompareValues(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal Hwnd As Long) As Long

    Dim pColIndex As Long
    Dim pOrder As Boolean
    Dim val1 As Long
    Dim val2 As Long
   
    pColIndex = GetProp(Hwnd, "ColumnIndex")
    pOrder = GetProp(Hwnd, "SortOrder")
    
  'CompareValues: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
   
     
  'Obtain the item names and values corresponding
  'to the input parameters
   val1 = CLng(GetItemData(Hwnd, pColIndex, lParam1))
   val2 = CLng(GetItemData(Hwnd, pColIndex, lParam2))
     
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the values appropriately:
   Select Case pOrder
      Case True 'sort descending
            
            If val1 < val2 Then
               CompareValues = 0
            ElseIf val1 = val2 Then
               CompareValues = 1
            Else
               CompareValues = 2
            End If
      
      Case Else 'sort ascending
   
            If val1 > val2 Then
               CompareValues = 0
            ElseIf val1 = val2 Then
               CompareValues = 1
            Else
               CompareValues = 2
            End If
   
   End Select

End Function


Private Function GetItemData(Hwnd As Long, ByVal pColIndex As Long, lParam As Long) As String
  
   Dim hIndex As Long
   Dim r As Long
   'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(Hwnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.mask = LVIF_TEXT
   objItem.iSubItem = pColIndex
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'get the string at subitem 1
  'and convert it into a date and exit
   r = SendMessage(Hwnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      GetItemData = CStr(Left$(objItem.pszText, r))
   End If
  
End Function

Public Function FARPROC(ByVal pfn As Long) As Long
  
 'A procedure that receives and returns the value of the AddressOf operator.
 'This workaround is needed as you can't assign AddressOf directly to an API when you are also
 'passing the value ByVal in the statement (as is being done with SendMessage)
 FARPROC = pfn

End Function






