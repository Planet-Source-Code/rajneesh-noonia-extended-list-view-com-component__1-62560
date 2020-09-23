VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum CompareResult
    LessThan = 0
    Equal = 1
    GreaterThan = 2
End Enum

Public Function Compare(ByVal ListViewHwnd As Long, ByVal lvColumnIndex As Single, ByVal SortAsc As Boolean, ByVal Value1 As Variant, ByVal Value2 As Variant) As CompareResult
'
End Function
