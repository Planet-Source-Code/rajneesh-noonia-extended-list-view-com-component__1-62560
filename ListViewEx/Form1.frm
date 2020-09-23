VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3645
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6429
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3645
      Left            =   5280
      TabIndex        =   1
      Top             =   300
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6429
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ICompare



Private Sub ListViewInitilize(ByVal lvControl As ListView)

   Dim itmX As ListItem
     
  'Add three Column Headers to the control
   lvControl.ColumnHeaders.Add , , "Name"
   lvControl.ColumnHeaders.Add , , Text:="Date"
   lvControl.ColumnHeaders.Add , , "Value"
   
   'Add information regarding Sorting
   Attach Me, lvControl, 1, lvString
   Attach Me, lvControl, 2, lvDate
   Attach Me, lvControl, 3, lvCustom
     
   ShowHeaderIcon(lvControl.hWnd) = True
   
  ' Set lvControl.ColumnHeaderIcons = ImageList1
  'Set the ListView to Report view
   lvControl.View = lvwReport
     
  'Add some data to the ListView control
   Set itmX = lvControl.ListItems.Add(Text:="Ritu")
   itmX.SubItems(1) = "05/07/97"
   itmX.SubItems(2) = "44"

   Set itmX = lvControl.ListItems.Add(Text:="Rajneesh")
   itmX.SubItems(1) = "04/08/1999"
   itmX.SubItems(2) = "15"

   Set itmX = lvControl.ListItems.Add(Text:="Ravi")
   itmX.SubItems(1) = "05/29/2000"
   itmX.SubItems(2) = "1"

   Set itmX = lvControl.ListItems.Add(Text:="Sujan")
   itmX.SubItems(1) = "03/17/2002"
   itmX.SubItems(2) = "11"

   Set itmX = lvControl.ListItems.Add(Text:="Nittin")
   itmX.SubItems(1) = "07/01/2003"
   itmX.SubItems(2) = "20"

   Set itmX = lvControl.ListItems.Add(Text:="Anoop")
   itmX.SubItems(1) = "04/01/2004"
   itmX.SubItems(2) = "21"

   Set itmX = lvControl.ListItems.Add(Text:="Aashish")
   itmX.SubItems(1) = "12/25/2004"
   itmX.SubItems(2) = "176"

   Set itmX = lvControl.ListItems.Add(Text:="Paul")
   itmX.SubItems(1) = "11/23/2006"
   itmX.SubItems(2) = "113"

   Set itmX = lvControl.ListItems.Add(Text:="Maria")
   itmX.SubItems(1) = "02/01/2005"
   itmX.SubItems(2) = "567"
   'Sort programatically
   Call SortColumn(lvControl.hWnd, lvControl.ColumnHeaders(1))
End Sub

Private Sub Form_Load()
    Call ListViewInitilize(ListView1)
    Call ListViewInitilize(ListView2)
End Sub

Private Function ICompare_Compare(ByVal ListViewHwnd As Long, ByVal lvColumnIndex As Single, ByVal SortAsc As Boolean, ByVal Value1 As Variant, ByVal Value2 As Variant) As ListViewEx.CompareResult
    If ListView1.hWnd = ListViewHwnd Then
        Select Case lvColumnIndex
            Case 2:
                If (SortAsc) Then
                    If (CLng(Value1) > CLng(Value2)) Then
                        ICompare_Compare = GreaterThan
                    ElseIf (CLng(Value1) = CLng(Value2)) Then
                        ICompare_Compare = Equal
                    Else
                        ICompare_Compare = LessThan
                    End If
                Else
                    If (CLng(Value1) > CLng(Value2)) Then
                        ICompare_Compare = LessThan
                    ElseIf (CLng(Value1) = CLng(Value2)) Then
                        ICompare_Compare = Equal
                    Else
                        ICompare_Compare = GreaterThan
                    End If
                End If
            
        End Select
    End If
End Function


Private Function ICompare_CompareDates(ByVal Date1 As Date, ByVal Date2 As Date) As ListViewEx.CompareResult

End Function
