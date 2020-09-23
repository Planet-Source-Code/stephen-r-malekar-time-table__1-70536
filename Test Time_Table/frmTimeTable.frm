VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTimeTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Table"
   ClientHeight    =   9540
   ClientLeft      =   3210
   ClientTop       =   1080
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   9465
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      Picture         =   "frmTimeTable.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   9495
      TabIndex        =   27
      Top             =   0
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Current Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   29
         Top             =   1200
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxGrd_TT 
         Height          =   4095
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         ForeColor       =   4194368
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox lst_Temp_Per_No 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   23
         Top             =   7560
         Width           =   975
      End
      Begin VB.ListBox lst_Temp_Teach_Sub 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         TabIndex        =   24
         Top             =   7560
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox lstP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   930
         Left            =   4440
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ListBox lstCl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   930
         Left            =   3480
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.ListBox lstS 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   930
         Left            =   2040
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ListBox lstN 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   930
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdExcl 
         Caption         =   "&Send to MS-Excel for Printing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   15
         Top             =   7850
         Width           =   2775
      End
      Begin VB.CommandButton cmdGen 
         Caption         =   "&Generate the Time-table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   2950
         Width           =   2055
      End
      Begin VB.VScrollBar VScrl_Tot 
         Height          =   375
         Left            =   8400
         Max             =   6000
         Min             =   1
         TabIndex        =   13
         Top             =   2350
         Value           =   1
         Width           =   375
      End
      Begin VB.TextBox txttot_periods 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   12
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add to List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.VScrollBar VScrl_Peri 
         Height          =   375
         Left            =   5205
         Max             =   25
         Min             =   1
         TabIndex        =   9
         Top             =   465
         Value           =   1
         Width           =   375
      End
      Begin VB.TextBox txtPeriods 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtCls 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbSub 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTimeTable.frx":29406
         Left            =   2040
         List            =   "frmTimeTable.frx":29422
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbTeacher 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTimeTable.frx":29472
         Left            =   120
         List            =   "frmTimeTable.frx":2948E
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Periods:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   9240
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "Patterns/Groups :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   20
         Top             =   1920
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   9240
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Caption         =   "Total No. of Periods per day:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   0
         X2              =   9240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No of Periods in a Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4380
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class-Div"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subject taught"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teachers Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmTimeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Collect_Period As Collection, Collect_Teach_Sub As Collection
Dim to_Add As Boolean, Added As Boolean
Dim lstR As Integer, lstC As Integer

Private Sub cmdAdd_Click()
Me.lstN.AddItem Me.cmbTeacher.Text
Me.lstS.AddItem Me.cmbSub.Text
Me.lstCl.AddItem Me.txtCls.Text
Me.lstP.AddItem Me.txtPeriods.Text
'Me.lstD.AddItem Me.txtdays.Text
If Me.Label7.Caption = "-" Then
    Me.Label7.Caption = Me.txtPeriods.Text
Else
    Me.Label7.Caption = CInt(Me.Label7.Caption) + CInt(Me.txtPeriods.Text)
End If
End Sub

Private Sub cmdClear_Click()
Me.lst_Temp_Per_No.Clear
Me.lst_Temp_Teach_Sub.Clear
Me.lstCl.Clear
Me.lstN.Clear
Me.lstP.Clear
Me.lstS.Clear
Me.MSFlxGrd_TT.Clear
Me.Label7.Caption = "-"
Me.txttot_periods.Text = ""
End Sub

Private Sub cmdDelete_Click()
Dim ind As Integer
ind = Me.lstN.ListIndex
Me.lstN.RemoveItem ind
Me.lstCl.RemoveItem ind
Me.lstP.RemoveItem ind
Me.lstS.RemoveItem ind
If Me.Label7.Caption = "-" Then
    Me.Label7.Caption = Me.txtPeriods.Text
Else
    Me.Label7.Caption = CInt(Me.Label7.Caption) - CInt(Me.txtPeriods.Text)
End If

End Sub

Private Sub cmdExcl_Click()
'Define the required variable
Dim Data_Row As Integer, Data_Col As Integer

Dim Excel As Excel.Application ' This is the excel program
Dim ExcelWBk As Excel.Workbook ' This is the work book
Dim ExcelWS As Excel.Worksheet ' This is the sheet

If Not Excel Is Nothing Then Set Excel = Nothing
Set Excel = CreateObject("Excel.Application") 'Create Excel Object.

Set ExcelWBk = Excel.Workbooks.Add 'Add this Workbook to Excel.
Set ExcelWS = ExcelWBk.Worksheets(1) ' Add this sheet to this Workbook

'Fill the Excel Sheet
For Data_Row = 0 To Me.MSFlxGrd_TT.Rows - 1
    For Data_Col = 1 To Me.MSFlxGrd_TT.Cols - 1
        Me.MSFlxGrd_TT.Row = Data_Row
        Me.MSFlxGrd_TT.Col = Data_Col
        ExcelWS.Cells(Data_Row + 1, Data_Col) = Me.MSFlxGrd_TT.Text
    Next
Next
Me.CommonDialog1.ShowSave
If Len(Me.CommonDialog1.FileName) > 0 Then ExcelWBk.SaveAs Me.CommonDialog1.FileName
' Close the WorkBook
ExcelWBk.Close
' Quit Excel app
Excel.Quit
Set Excel = Nothing
End Sub

Private Sub cmdGen_Click()
Dim start_Val  As Integer, tot_peri_dys As Integer
Dim no_peri As Integer, filler As Integer

Me.lst_Temp_Per_No.Clear
Me.lst_Temp_Teach_Sub.Clear

Set Collect_Period = New Collection
Set Collect_Teach_Sub = New Collection

If Len(Me.txttot_periods.Text) > 0 Then
    tot_peri_dys = CInt(Me.txttot_periods.Text)
Else
    Exit Sub
End If

If CInt(Me.Label7.Caption) > ((tot_peri_dys * 6)) * _
((Me.lstN.ListCount - 1) / (Me.cmbTeacher.ListCount - 1)) Then
    If MsgBox("The machine may hang while designing," & vbCrLf & "Do you want to continue?", vbCritical + vbOKCancel, "Problem...") = vbOK Then
        'Nothing much to do
        'Try to design the Time-table
        'If mach. hangs its not the prog. problem.
    Else
        MsgBox "Try increasing the no of periods in a week, and click again.", , "Problem..."
        Exit Sub
    End If
End If


'Start Enumerating through the list
For start_Val = 0 To lstN.ListCount - 1
    lstN.ListIndex = start_Val
    lstCl.ListIndex = start_Val
    lstS.ListIndex = start_Val
    lstP.ListIndex = start_Val
    
    Fill_Collection tot_peri_dys, lstN.Text & " " & lstCl.Text & " " & lstS.Text, CInt(lstP.Text)
Next

Dim c As Integer, t As Integer

Me.MSFlxGrd_TT.Clear
Me.MSFlxGrd_TT.Rows = tot_peri_dys + 1
Me.MSFlxGrd_TT.Cols = 7
'Marks the Periods
For c = 1 To tot_peri_dys
    Me.MSFlxGrd_TT.Row = c
    Me.MSFlxGrd_TT.Col = 0
    Me.MSFlxGrd_TT.Text = "# " & c
Next
'Fill the FlexGrid with the Weekdays.
'Monday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 1
Me.MSFlxGrd_TT.Text = "Monday"
'Tuesday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 2
Me.MSFlxGrd_TT.Text = "Tueday"
'Wednesday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 3
Me.MSFlxGrd_TT.Text = "Wednesday"
'Thursday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 4
Me.MSFlxGrd_TT.Text = "Thursday"
'Friday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 5
Me.MSFlxGrd_TT.Text = "Friday"
'Saturday
Me.MSFlxGrd_TT.Row = 0
Me.MSFlxGrd_TT.Col = 6
Me.MSFlxGrd_TT.Text = "Saturday"

'For filler = 1 To Collect_Period.Count
'
'    For c = 1 To (tot_peri_dys * 6) Step 6
'        If Collect_Period.Item(filler) >= c And Collect_Period.Item(filler) <= (c + 5) Then
'            Me.MSFlxGrd_TT.Row = (c + 5) \ 6
'            Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'            Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'        End If
'        If CInt(Me.lst_Temp_Per_No.List(filler)) >= c And CInt(Me.lst_Temp_Per_No.List(filler)) <= (c + 5) Then
'            Me.MSFlxGrd_TT.Row = (c + 5) \ 6
'            Me.MSFlxGrd_TT.Col = CInt(Me.lst_Temp_Per_No.List(filler)) Mod 6 + 1
'            Me.MSFlxGrd_TT.Text = Me.lst_Temp_Teach_Sub.List(filler)
'        End If
'
'
'    Next
'Next

For filler = 0 To CInt(Me.lst_Temp_Per_No.ListCount) - 1
    For c = 1 To (tot_peri_dys * 6) Step 6
        If CInt(Me.lst_Temp_Per_No.List(filler)) >= c And CInt(Me.lst_Temp_Per_No.List(filler)) <= (c + 5) Then
            Me.MSFlxGrd_TT.Row = (c + 5) \ 6
            Me.MSFlxGrd_TT.Col = CInt(Me.lst_Temp_Per_No.List(filler)) Mod 6 + 1
            Me.MSFlxGrd_TT.Text = Me.lst_Temp_Teach_Sub.List(filler)
        End If
    Next

    
'    If Collect_Period.Item(filler) >= 1 And Collect_Period.Item(filler) <= 6 Then
'        Me.MSFlxGrd_TT.Row = 1
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 7 And Collect_Period.Item(filler) <= 12 Then
'        Me.MSFlxGrd_TT.Row = 2
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 13 And Collect_Period.Item(filler) <= 18 Then
'        Me.MSFlxGrd_TT.Row = 3
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 19 And Collect_Period.Item(filler) <= 24 Then
'        Me.MSFlxGrd_TT.Row = 4
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 25 And Collect_Period.Item(filler) <= 30 Then
'        Me.MSFlxGrd_TT.Row = 5
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 31 And Collect_Period.Item(filler) <= 36 Then
'        Me.MSFlxGrd_TT.Row = 6
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 37 And Collect_Period.Item(filler) <= 42 Then
'        Me.MSFlxGrd_TT.Row = 7
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    ElseIf Collect_Period.Item(filler) >= 43 And Collect_Period.Item(filler) <= 48 Then
'        Me.MSFlxGrd_TT.Row = 8
'        Me.MSFlxGrd_TT.Col = Collect_Period.Item(filler) Mod 6 + 1
'        Me.MSFlxGrd_TT.Text = Collect_Teach_Sub.Item(filler)
'    End If
Next

Set Collect_Period = Nothing
Set Collect_Teach_Sub = Nothing
End Sub

'Creates the Collection and list.

Private Sub Fill_Collection(tot_peri_dys As Integer, s As String, no_peri As Integer)
Dim i As Integer, Tot_Periods As Integer, pos As Integer
Dim flag As Boolean, j As Integer
Tot_Periods = 6 * tot_peri_dys
to_Add = True

For i = 0 To no_peri - 1
    If Collect_Period.Count > 0 Then
        pos = GetPos(Tot_Periods)
        Do While 1
            If NoDuplicate(pos, s) Then
                Collect_Period.Add pos
                Collect_Teach_Sub.Add s
                If Added = False Then
                    If to_Add = True Then
                        Me.lst_Temp_Per_No.AddItem pos
                        Me.lst_Temp_Teach_Sub.AddItem s
                        to_Add = False
                    End If
                End If
                Exit Do
            Else
                 If Added = True Then Exit Do
                 pos = GetPos(Tot_Periods)
            End If
        Loop
    Else
        pos = GetPos(Tot_Periods)
        If pos <> 0 Then
            Collect_Period.Add pos
            Collect_Teach_Sub.Add s
            Me.lst_Temp_Per_No.AddItem pos
            Me.lst_Temp_Teach_Sub.AddItem s
        End If
    End If
    Added = False
Next

End Sub

'The No Duplication function prevents duplication of values, selected @ random.
Private Function NoDuplicate(N As Integer, CL_Sub As String) As Boolean
    Dim i As Integer, flag As Boolean, temp As Integer, garb As String
    Dim ext, a, X As Integer, Y As Integer
    Dim s1 As String, s2 As String
    Dim e, d
    flag = True
    For i = 1 To Collect_Period.Count
        temp = Collect_Period.Item(i)
        If N = 0 Then
            flag = False
            Exit For
        End If
        If N = temp Then
            garb = Me.lst_Temp_Teach_Sub.List(i - 1)
            ext = Split(garb, " ")
            a = Split(CL_Sub, " ")
            
'Checks if a teacher/ other teacher or class gets repeated for the same period.
            For Each d In a
                For Each e In ext
                    If e = d Then
                        X = 0
                        flag = False
                        to_Add = False
                        NoDuplicate = False
                        Exit Function
                    Else
                        X = 1
                    End If
                Next
            Next
            

            For Each e In ext
                If e = a(1) Then
                    Y = 0
                    flag = False
                    to_Add = False
                    NoDuplicate = False
                    Exit Function
                Else
                    Y = 1
                End If
            Next
            
            If X > 0 And Y > 0 Then
                Me.lst_Temp_Teach_Sub.List(i - 1) = _
                Me.lst_Temp_Teach_Sub.List(i - 1) & " " & CL_Sub
                flag = False
                Added = True
                Exit For
            End If
        Else
            flag = True
        End If
    Next
    If flag = True And X = 0 Then to_Add = True
    NoDuplicate = flag
End Function

'This extracts the pos @ random, very messy and time-taking,
'Sometimes things get repeated.
Private Function GetPos(v As Integer) As Integer
    Randomize
    GetPos = Rnd * v
End Function

Private Sub Form_Load()
lstR = 1
lstC = 1

'Clear Temp List Boxes
Me.lst_Temp_Per_No.Clear
Me.lst_Temp_Teach_Sub.Clear
Me.MSFlxGrd_TT.WordWrap = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Collect_Period = Nothing
Set Collect_Teach_Sub = Nothing
End Sub

Private Sub lst_Temp_Per_No_Click()
Me.lst_Temp_Teach_Sub.Selected(Me.lst_Temp_Per_No.ListIndex) = True
End Sub

Private Sub lst_Temp_Per_No_GotFocus()
Me.lst_Temp_Per_No.ToolTipText = Me.lst_Temp_Per_No.Text
End Sub

Private Sub lst_Temp_Teach_Sub_Click()
Me.lst_Temp_Per_No.Selected(Me.lst_Temp_Teach_Sub.ListIndex) = True
End Sub

Private Sub lst_Temp_Teach_Sub_GotFocus()
Me.lst_Temp_Teach_Sub.ToolTipText = Me.lst_Temp_Teach_Sub.Text
End Sub

Private Sub lstCl_Click()
Me.lstS.Selected(Me.lstCl.ListIndex) = True
Me.lstN.Selected(Me.lstCl.ListIndex) = True
Me.lstP.Selected(Me.lstCl.ListIndex) = True
End Sub

Private Sub lstN_Click()
Me.lstS.Selected(Me.lstN.ListIndex) = True
Me.lstCl.Selected(Me.lstN.ListIndex) = True
Me.lstP.Selected(Me.lstN.ListIndex) = True
End Sub

Private Sub lstP_Click()
Me.lstS.Selected(Me.lstP.ListIndex) = True
Me.lstCl.Selected(Me.lstP.ListIndex) = True
Me.lstN.Selected(Me.lstP.ListIndex) = True
End Sub

Private Sub lstS_Click()
Me.lstN.Selected(Me.lstS.ListIndex) = True
Me.lstCl.Selected(Me.lstS.ListIndex) = True
Me.lstP.Selected(Me.lstS.ListIndex) = True
End Sub

Private Sub VScrl_Peri_Change()
Me.txtPeriods.Text = VScrl_Peri.Value
End Sub

Private Sub VScrl_Tot_Change()
Me.txttot_periods.Text = VScrl_Tot.Value
End Sub


