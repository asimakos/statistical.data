VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form Form1 
   Caption         =   "������������     �������������     �����������"
   ClientHeight    =   6585
   ClientLeft      =   2790
   ClientTop       =   1830
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9360
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog6 
      Left            =   1800
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog5 
      Left            =   1680
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   3840
      TabIndex        =   8
      Top             =   3360
      Width           =   5415
      _Version        =   524288
      _ExtentX        =   9551
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2009
      Month           =   1
      Day             =   30
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog4 
      Left            =   1080
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   360
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   2
            TextSave        =   "11:10 ��"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   960
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "�������   �����"
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "��������   ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "�����  ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   1065
   End
   Begin VB.OLE OLE3 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   45
   End
   Begin VB.OLE OLE2 
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   45
   End
   Begin VB.Menu file 
      Caption         =   "������"
      Begin VB.Menu file_open 
         Caption         =   "�������  �������  xlsx"
      End
      Begin VB.Menu Clean_data 
         Caption         =   "����������   ���������"
      End
      Begin VB.Menu Exit 
         Caption         =   "������"
      End
   End
   Begin VB.Menu result 
      Caption         =   "����������  �����"
      Begin VB.Menu sum 
         Caption         =   "��������  ����������  �����  xls"
      End
      Begin VB.Menu dif 
         Caption         =   "�������  ����������  �����  xls"
      End
      Begin VB.Menu analog 
         Caption         =   "��������  ����������  �����  xls"
      End
      Begin VB.Menu derivative 
         Caption         =   "������  ���������  ����������   �����  xls"
      End
   End
   Begin VB.Menu month 
      Caption         =   "�������"
      Begin VB.Menu dif_month 
         Caption         =   "�������  �����  xls"
      End
      Begin VB.Menu analog_month 
         Caption         =   "��������  �����  xls"
      End
      Begin VB.Menu derivat_month 
         Caption         =   "������  ���������  �����  xls"
      End
   End
   Begin VB.Menu comp 
      Caption         =   "����������  ��������"
      Begin VB.Menu comp_result 
         Caption         =   "����������  �����"
         Begin VB.Menu comp_dif 
            Caption         =   "�������  ����������  �����  doc"
         End
         Begin VB.Menu comp_analog 
            Caption         =   "��������  ����������  �����  doc"
         End
         Begin VB.Menu comp_derivative 
            Caption         =   "������  ���������  ���������� �����   doc"
         End
      End
      Begin VB.Menu comp_month 
         Caption         =   "�������"
         Begin VB.Menu comp_dif_month 
            Caption         =   "�������   �����   doc"
         End
         Begin VB.Menu comp_analog_month 
            Caption         =   "��������  �����  doc"
         End
         Begin VB.Menu comp_derivat_month 
            Caption         =   "������  ���������  �����  doc"
         End
      End
   End
   Begin VB.Menu open 
      Caption         =   "�������"
      WindowList      =   -1  'True
      Begin VB.Menu open_sum 
         Caption         =   "xls  �����������  ����������  �����"
      End
      Begin VB.Menu open_statistic 
         Caption         =   "xls  ����������  �����"
      End
      Begin VB.Menu open_month 
         Caption         =   "xls  �����"
      End
      Begin VB.Menu About 
         Caption         =   "�������  �� ..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim axls(30, 14) As Double, adiv(30, 14) As Double
Dim bxls(30, 14) As Double, bdiv(30, 14) As Double
Dim perife(17) As String, testfilenew As String
Dim testfile As String, xint As Integer, filestr As String
Dim xint1 As Integer, xint2 As Integer, testfileadd As String
Dim xdiv1 As Integer, xdiv2 As Integer
Sub CleanData()
Dim i As Integer, j As Integer

For i = 1 To GetNum()
For j = 1 To GetCol()
axls(i, j) = 0
bxls(i, j) = 0
Next j
Next i
For i = 1 To GetNumnew()
For j = 1 To GetColnew()
adiv(i, j) = 0
bdiv(i, j) = 0
Next j
Next i

End Sub
Function Months(x As Integer) As String

Select Case x
 Case 1
 Months = "����������"
 Case 2
 Months = "����������, " & "�����������"
 Case 3
 Months = "����������, " & "�����������, " & "�������"
 Case 4
 Months = "����������, " & "�����������, " & "�������, " & "��������"
 Case 5
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����"
 Case 6
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������"
 Case 7
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������"
 Case 8
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������, " & "���������"
 Case 9
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������, " & "���������, " & "�����������"
 Case 10
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������, " & "���������, " & "�����������, " & "���������"
 Case 11
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������, " & "���������, " & "�����������, " & "���������, " & "���������"
 Case 12
 Months = "����������, " & "�����������, " & "�������, " & "��������, " & "�����, " & "�������, " & vbNewLine & "�������, " & "���������, " & "�����������, " & "���������, " & "���������, " & "����������"
 End Select

 End Function
Function Monthval(x As Integer) As String

Select Case x
 Case 1
 Monthval = "����������"
 Case 2
 Monthval = "�����������"
 Case 3
 Monthval = "�������"
 Case 4
 Monthval = "��������"
 Case 5
 Monthval = "�����"
 Case 6
 Monthval = "�������"
 Case 7
 Monthval = "�������"
 Case 8
 Monthval = "���������"
 Case 9
 Monthval = "�����������"
 Case 10
 Monthval = "���������"
 Case 11
 Monthval = "���������"
 Case 12
 Monthval = "����������"
 End Select

 End Function

Private Sub About_Click()

frm.Caption = "�������   ��   ������������   �������������   �����������"
frm.lblTitle.Caption = "������������   �������������  �����������"
frm.Show

End Sub

Private Sub analog_Click()

Label5.Caption = "�������� ����������..."
Call CleanData
Call Printdivisionxls


MsgBox "� ���������� ��������� ��� " & filestr & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Private Sub analog_month_Click()

Label5.Caption = "�������� ����������..."
Call CleanData
Call Printmonthdivisionxls
'Call Printderivativexls

MsgBox "� ���������� ��������� ��� " & filestr & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Private Sub Clean_data_Click()

Label1.Caption = Empty
Label2.Caption = Empty
Label3.Caption = Empty
Label4.Caption = Empty
analog.Enabled = False
derivative.Enabled = False
analog_month.Enabled = False
derivat_month.Enabled = False
Text1.Text = Empty

End Sub
Function GetNumnew() As Integer

Dim i As Integer
Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label4.Caption
'Set xlSheet = xlApp.Sheets(1)
Set xlSheet = xlApp.Sheets.Item(1)

i = 3
Do
i = i + 1
If IsEmpty(xlSheet.Cells(i, 3).Value) Then
Exit Do
End If
Loop While i <> 0
xlApp.Workbooks.Close
Set xlApp = Nothing
Set xlSheet = Nothing

GetNumnew = i - 1
End Function

Function GetColnew() As Integer

Dim i As Integer
Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label4.Caption
'Set xlSheet = xlApp.Sheets(1)
Set xlSheet = xlApp.Sheets.Item(1)

i = 0
Do
i = i + 1
If IsEmpty(xlSheet.Cells(10, i).Value) Then
Exit Do
End If
Loop While i <> 0
xlApp.Workbooks.Close
Set xlApp = Nothing
Set xlSheet = Nothing

GetColnew = i - 1
End Function
Function GetNum() As Integer

Dim i As Integer
Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label2.Caption
'Set xlSheet = xlApp.Sheets(1)
Set xlSheet = xlApp.Sheets.Item(1)

i = 3
Do
i = i + 1
If IsEmpty(xlSheet.Cells(i, 3).Value) Then
Exit Do
End If
Loop While i <> 0
xlApp.Workbooks.Close
Set xlApp = Nothing
Set xlSheet = Nothing

GetNum = i - 1
End Function
Function GetCol() As Integer

Dim i As Integer
Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label2.Caption
'Set xlSheet = xlApp.Sheets(1)
Set xlSheet = xlApp.Sheets.Item(1)

i = 0
Do
i = i + 1
If IsEmpty(xlSheet.Cells(10, i).Value) Then
Exit Do
End If
Loop While i <> 0
xlApp.Workbooks.Close
Set xlApp = Nothing
Set xlSheet = Nothing

GetCol = i - 1
End Function

Sub CopyDatadif()

Dim xlApp As Excel.Application
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet
Dim i As Integer, j As Integer

Set xlApp = CreateObject("Excel.Application")


xlApp.Workbooks.open Label2.Caption
Set xlSheet1 = xlApp.Sheets.Item(xint1)

For i = 1 To GetNum()
 For j = 1 To GetCol()
If IsNumeric(xlSheet1.Cells(i, j).Value) Then
xlSheet1.Cells(i, j).NumberFormat = "General"
axls(i, j) = axls(i, j) + xlSheet1.Cells(i, j).Value
Else
axls(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNum()
If CLng((axls(i, 3)) >= 0) Then
axls(i, 1) = i
End If
Next i
xint1 = xint1 + 1

'xlApp.Workbooks.Close
xlApp.Workbooks.open Label4.Caption
Set xlSheet2 = xlApp.Sheets.Item(xint2)

For i = 1 To GetNumnew()
 For j = 1 To GetColnew()
If IsNumeric(xlSheet2.Cells(i, j).Value) Then
xlSheet2.Cells(i, j).NumberFormat = "General"
bxls(i, j) = bxls(i, j) + xlSheet2.Cells(i, j).Value
Else
bxls(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNumnew()
If CLng((bxls(i, 3)) >= 0) Then
bxls(i, 1) = i
End If
Next i
xint2 = xint2 + 1


End Sub

Sub CopyDatadivis()

Dim xlApp As Excel.Application
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet
Dim i As Integer, j As Integer

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label2.Caption
Set xlSheet1 = xlApp.Sheets.Item(xdiv1)

For i = 1 To GetNum()
 For j = 1 To GetCol()
If IsNumeric(xlSheet1.Cells(i, j).Value) Then
xlSheet1.Cells(i, j).NumberFormat = "General"
adiv(i, j) = adiv(i, j) + xlSheet1.Cells(i, j).Value
Else
adiv(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNum()
If CLng((adiv(i, 3)) >= 0) Then
adiv(i, 1) = i
End If
Next i
xdiv1 = xdiv1 + 1

'xlApp.Workbooks.Close
xlApp.Workbooks.open Label4.Caption
Set xlSheet2 = xlApp.Sheets.Item(xdiv2)

For i = 1 To GetNumnew()
 For j = 1 To GetColnew()
If IsNumeric(xlSheet2.Cells(i, j).Value) Then
xlSheet2.Cells(i, j).NumberFormat = "General"
bdiv(i, j) = bdiv(i, j) + xlSheet2.Cells(i, j).Value
Else
bdiv(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNumnew()
If CLng((bdiv(i, 3)) >= 0) Then
bdiv(i, 1) = i
End If
Next i
xdiv2 = xdiv2 + 1

End Sub

Sub CopyDatamonthdif()

Dim xlApp As Excel.Application
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet
Dim i As Integer, j As Integer

Set xlApp = CreateObject("Excel.Application")


xlApp.Workbooks.open Label2.Caption
Set xlSheet1 = xlApp.Sheets.Item(xint1)

For i = 1 To GetNum()
 For j = 1 To GetCol()
If IsNumeric(xlSheet1.Cells(i, j).Value) Then
xlSheet1.Cells(i, j).NumberFormat = "General"
axls(i, j) = xlSheet1.Cells(i, j).Value
Else
axls(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNum()
If CLng((axls(i, 3)) >= 0) Then
axls(i, 1) = i
End If
Next i
xint1 = xint1 + 1

'xlApp.Workbooks.Close
xlApp.Workbooks.open Label4.Caption
Set xlSheet2 = xlApp.Sheets.Item(xint2)

For i = 1 To GetNumnew()
 For j = 1 To GetColnew()
If IsNumeric(xlSheet2.Cells(i, j).Value) Then
xlSheet2.Cells(i, j).NumberFormat = "General"
bxls(i, j) = xlSheet2.Cells(i, j).Value
Else
bxls(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNumnew()
If CLng((bxls(i, 3)) >= 0) Then
bxls(i, 1) = i
End If
Next i
xint2 = xint2 + 1


End Sub

Sub CopyDatamonthdivis()

Dim xlApp As Excel.Application
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet
Dim i As Integer, j As Integer

Set xlApp = CreateObject("Excel.Application")

xlApp.Workbooks.open Label2.Caption
Set xlSheet1 = xlApp.Sheets.Item(xdiv1)

For i = 1 To GetNum()
 For j = 1 To GetCol()
If IsNumeric(xlSheet1.Cells(i, j).Value) Then
xlSheet1.Cells(i, j).NumberFormat = "General"
adiv(i, j) = xlSheet1.Cells(i, j).Value
Else
adiv(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNum()
If CLng((adiv(i, 3)) >= 0) Then
adiv(i, 1) = i
End If
Next i
xdiv1 = xdiv1 + 1

'xlApp.Workbooks.Close
xlApp.Workbooks.open Label4.Caption
Set xlSheet2 = xlApp.Sheets.Item(xdiv2)

For i = 1 To GetNumnew()
 For j = 1 To GetColnew()
If IsNumeric(xlSheet2.Cells(i, j).Value) Then
xlSheet2.Cells(i, j).NumberFormat = "General"
bdiv(i, j) = xlSheet2.Cells(i, j).Value
Else
bdiv(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNumnew()
If CLng((bdiv(i, 3)) >= 0) Then
bdiv(i, 1) = i
End If
Next i
xdiv2 = xdiv2 + 1

End Sub

Private Sub comp_analog_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xdiv1 = 1
xdiv2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� �������� ���������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatadivis
'thisRange.InsertAfter "�������� " & resul & " - MHN��"
thisRange.InsertAfter Months(resul)
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "���������� �������� %(��������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(adiv(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bdiv(4 + xline, z), "##,##0.0#")
If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
wtable.Cell(z, 4).Range.Text = "�������� �� �����"
Else
'wtable.Cell(z, 4).Range.Text = (bdiv(4 + xline, z) / adiv(4 + xline, z)) * 100
wtable.Cell(z, 4).Range.Text = Format((bdiv(4 + xline, z) / adiv(4 + xline, z)), "##,##0.0#%")
End If

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing

End Sub

Private Sub comp_analog_month_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xdiv1 = 1
xdiv2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� �������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdivis
'thisRange.InsertAfter "�������� " & resul & " - MHNA"
thisRange.InsertAfter Monthval(resul)
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "���������� �������� %(��������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(adiv(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bdiv(4 + xline, z), "##,##0.0#")

If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
wtable.Cell(z, 4).Range.Text = "�������� �� �����"
Else
'wtable.Cell(z, 4).Range.Text = (bdiv(4 + xline, z) / adiv(4 + xline, z)) * 100
wtable.Cell(z, 4).Range.Text = Format((bdiv(4 + xline, z) / adiv(4 + xline, z)), "##,##0.0#%")
End If

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing


End Sub

Private Sub comp_derivat_month_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xdiv1 = 1
xdiv2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� ����� ��������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdivis
'thisRange.InsertAfter "�������� " & resul & " - MHNA"
thisRange.InsertAfter Monthval(resul)
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "���������� �������� %(������ ���������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(adiv(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bdiv(4 + xline, z), "##,##0.0#")

If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
wtable.Cell(z, 4).Range.Text = "�������� �� �����"
Else
'wtable.Cell(z, 4).Range.Text = ((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)) * 100
wtable.Cell(z, 4).Range.Text = Format(((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)), "##,##0.0#%")
End If

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing


End Sub

Private Sub comp_derivative_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xdiv1 = 1
xdiv2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� ����� ��������� ���������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatadivis
'thisRange.InsertAfter "�������� " & resul & " - MHN��"
thisRange.InsertAfter Months(resul)
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "���������� �������� %(������ ���������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(adiv(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bdiv(4 + xline, z), "##,##0.0#")
If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
wtable.Cell(z, 4).Range.Text = "�������� �� �����"
Else
'wtable.Cell(z, 4).Range.Text = ((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)) * 100
wtable.Cell(z, 4).Range.Text = Format(((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)), "##,##0.0#%")
End If

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing

End Sub

Private Sub comp_dif_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xint1 = 1
xint2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� ������� ���������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatadif
'thisRange.InsertAfter "�������� " & resul & " - MHN��"
thisRange.InsertAfter Months(resul)
'thisRange.InsertAfter vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
'thisRange.InsertAfter vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "�������� (�������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(axls(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bxls(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 4).Range.Text = Format((bxls(4 + xline, z) - axls(4 + xline, z)), "##,##0.0#")

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing

End Sub

Private Sub comp_dif_month_Click()

Dim x As String, thisDoc As Word.Document, wtable As Word.Table
Dim thisRange As Word.Range, AppWord As Word.Application
Dim resul As Integer, i As Integer, j As Integer
Dim z As Integer, xline As Integer, stats(10) As String

x = Empty
CommonDialog6.FileName = Empty

CommonDialog6.Filter = "Word|*.doc"
CommonDialog6.ShowSave
x = CommonDialog6.FileName
If (Len(x) = 0) Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

Label5.Caption = "�������� ���������� ..."
Call CleanData

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

stats(1) = "���,��� ��� ������ ��������������� �����"
stats(2) = "���������� ������� - ��������� (�������)"
stats(3) = "���������� ������� - ��������� (���� ���������)"
stats(4) = "������������ �������� �� ���"
stats(5) = "���������� ���� ��������� �� ���"
stats(6) = "������������ �������� �� ��������"
stats(7) = "���������� ���� ��������� �� ��������"
stats(8) = "�������� ���������� ���� ��������� (��������)"
stats(9) = "�������� ���������� ���� ��������� & ���"

Set AppWord = New Word.Application
Set thisDoc = AppWord.Documents.Add
'thisDoc.Range.InsertBefore "Document Title" & vbCrLf
Set thisRange = thisDoc.Paragraphs(1).Range

xint1 = 1
xint2 = 1

thisRange.Font.Size = 10
thisRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
thisRange.InsertAfter "���� ������ ����������� ���������� (�.�.�.�)" & vbCrLf & vbCrLf
thisRange.InsertAfter "�/��� ���������� & ����������� �������" & vbCrLf
thisRange.InsertAfter "���������� ���������� �������� �������������� ������� ��� ������� �����" & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf
'thisRange.InsertAfter Space(10) & Label2.Caption & Space(5) & Label4.Caption & Space(5) & "��������"
'thisRange.InsertAfter vbCrLf & vbCrLf

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdif
'thisRange.InsertAfter "�������� " & resul & " - MHNA"
thisRange.InsertAfter Monthval(resul)
thisRange.InsertAfter vbCrLf
xline = 0

For i = 4 To GetNumnew()

thisRange.InsertAfter "������������ ��������� " & perife(i - 3)
thisRange.InsertAfter vbCrLf
Set wtable = thisDoc.Tables.Add(thisDoc.Bookmarks("\endofdoc").Range, 10, 4)
'wtable.Cell(1, 1).Range.Text = Space(10)
'wtable.Cell(1, 2).Range.Text = Label2.Caption
'wtable.Cell(1, 3).Range.Text = Label4.Caption
wtable.Cell(1, 2).Range.Text = Mid(Trim(Label2.Caption), InStr(Trim(Label2.Caption), "2"), 4)
wtable.Cell(1, 3).Range.Text = Mid(Trim(Label4.Caption), InStr(Trim(Label4.Caption), "2"), 4)
wtable.Cell(1, 4).Range.Text = "�������� (�������)"

For z = 2 To GetColnew()

wtable.Cell(z, 1).Range.Text = stats(z - 1)
wtable.Cell(z, 2).Range.Text = Format(axls(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 3).Range.Text = Format(bxls(4 + xline, z), "##,##0.0#")
wtable.Cell(z, 4).Range.Text = Format(bxls(4 + xline, z) - axls(4 + xline, z), "##,##0.0#")

Next z

wtable.Rows(1).Range.Font.Bold = True
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf
thisRange.InsertAfter vbCrLf

If (i <> GetNumnew()) Then
thisRange.InsertAfter vbCrLf
End If

xline = xline + 1

Next i
'thisRange.InsertAfter "*****************************************************************"
'thisRange.InsertAfter vbCrLf & vbCrLf & vbCrLf

Next resul

thisDoc.SaveAs x
MsgBox "� ���������� ��������� ���" & x & "������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

AppWord.Quit
'AppWord.Visible = True
'AppWord.Activate

Set thisDoc = Nothing
Set AppWord = Nothing

End Sub

Private Sub derivat_month_Click()

Label5.Caption = "�������� ����������..."

Call CleanData
Call Printmonthderivativexls

MsgBox "� ���������� ��������� ��� " & filestr & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Private Sub derivative_Click()

Label5.Caption = "�������� ����������..."

Call CleanData
Call Printderivativexls

MsgBox "� ���������� ��������� ��� " & filestr & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Sub Printmonthderivativexls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.open(testfileadd)

Set xlSheet = xlBook.Worksheets.Item(3)

'Set xlSheet = xlApp.Sheets.Item(2)
monthcounter = 1
linecounter = 0
xdiv1 = 1
xdiv2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "������  ���������   �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdivis

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - ���A"
xlSheet.Cells(1 + linecounter, 3).Value = Monthval(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If

If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
xlSheet.Cells(3 + j + linecounter, z).Value = "�������� �� �����"
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
Else
xlSheet.Cells(3 + j + linecounter, z).Value = Format(((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)), "##,##0.0#%")
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10

Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

'xlBook.SaveAs testfileadd, FileFormat:=-4143, CreateBackup:=False
xlSheet.SaveAs testfileadd
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub

Sub Printmonthdivisionxls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.open(testfileadd)

Set xlSheet = xlBook.Worksheets.Item(2)

'Set xlSheet = xlApp.Sheets.Item(2)
monthcounter = 1
linecounter = 0
xdiv1 = 1
xdiv2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "��������  �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdivis

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - ���A"
xlSheet.Cells(1 + linecounter, 3).Value = Monthval(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If
If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
xlSheet.Cells(3 + j + linecounter, z).Value = "�������� �� �����"
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
Else
xlSheet.Cells(3 + j + linecounter, z).Value = Format((bdiv(4 + xline, z) / adiv(4 + xline, z)), "##,##0.0#%")
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10

Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

'xlBook.SaveAs testfilenew, FileFormat:=-4143, CreateBackup:=False
xlSheet.SaveAs testfileadd
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub
Sub Printmonthdifxls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.Add

Set xlSheet = xlBook.Worksheets.Item(1)

'Set xlSheet = xlApp.Sheets.Item(1)
monthcounter = 1
linecounter = 0
xint1 = 1
xint2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "������� �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatamonthdif

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - ����"
xlSheet.Cells(1 + linecounter, 3).Value = Monthval(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If
If (bxls(4 + xline, z) - axls(4 + xline, z) < 0) Then
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10
xlSheet.Cells(3 + j + linecounter, z).Value = Format(bxls(4 + xline, z) - axls(4 + xline, z), "##,##0.0#")
Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

xlBook.SaveAs testfileadd, FileFormat:=-4143, CreateBackup:=False
'xlSheet.SaveAs testfilenew
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub

Private Sub dif_Click()

testfilenew = Empty
CommonDialog4.FileName = Empty

CommonDialog4.Filter = "Excel|*.xls"
CommonDialog4.ShowSave
testfilenew = CommonDialog4.FileName
filestr = CommonDialog4.FileTitle
If Len(testfilenew) = 0 Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

analog.Enabled = True
derivative.Enabled = True
Label5.Caption = "�������� ����������..."
Call CleanData
Call Printdifxls


MsgBox "� ���������� ��������� ��� " & CommonDialog4.FileTitle & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Private Sub dif_month_Click()

testfileadd = Empty
CommonDialog5.FileName = Empty

CommonDialog5.Filter = "Excel|*.xls"
CommonDialog5.ShowSave
testfileadd = CommonDialog5.FileName
filestr = CommonDialog5.FileTitle
If Len(testfileadd) = 0 Then
Exit Sub
End If

If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If

analog_month.Enabled = True
derivat_month.Enabled = True
Label5.Caption = "�������� ����������..."
Call CleanData
Call Printmonthdifxls


MsgBox "� ���������� ��������� ��� " & CommonDialog5.FileTitle & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub

Private Sub Exit_Click()
End
End Sub

Sub Printdivisionxls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.open(testfilenew)

Set xlSheet = xlBook.Worksheets.Item(2)

'Set xlSheet = xlApp.Sheets.Item(2)
monthcounter = 1
linecounter = 0
xdiv1 = 1
xdiv2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "�������� ���������� �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatadivis

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - �����"
xlSheet.Cells(1 + linecounter, 3).Value = Months(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If
If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
xlSheet.Cells(3 + j + linecounter, z).Value = "�������� �� �����"
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
Else
xlSheet.Cells(3 + j + linecounter, z).Value = Format((bdiv(4 + xline, z) / adiv(4 + xline, z)), "##,##0.0#%")
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10

Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

'xlBook.SaveAs testfilenew, FileFormat:=-4143, CreateBackup:=False
xlSheet.SaveAs testfilenew
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub
Sub CopyData()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim i As Integer, j As Integer

Set xlApp = CreateObject("Excel.Application")


xlApp.Workbooks.open Label2.Caption
Set xlSheet = xlApp.Sheets.Item(xint)

For i = 1 To GetNum()
 For j = 1 To GetCol()
If IsNumeric(xlSheet.Cells(i, j).Value) Then
xlSheet.Cells(i, j).NumberFormat = "General"
axls(i, j) = axls(i, j) + xlSheet.Cells(i, j).Value
Else
axls(i, j) = -100
 End If
 Next j
Next i

For i = 1 To GetNum()
If CLng((axls(i, 3)) >= 0) Then
axls(i, 1) = i
End If
Next i
xint = xint + 1

End Sub

Sub Printderivativexls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.open(testfilenew)

Set xlSheet = xlBook.Worksheets.Item(3)

'Set xlSheet = xlApp.Sheets.Item(2)
monthcounter = 1
linecounter = 0
xdiv1 = 1
xdiv2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "������  ���������  ���������� �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatadivis

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - �����"
xlSheet.Cells(1 + linecounter, 3).Value = Months(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If

If (StrComp(CStr(adiv(4 + xline, z)), "0") = 0) Then
xlSheet.Cells(3 + j + linecounter, z).Value = "�������� �� �����"
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
Else
xlSheet.Cells(3 + j + linecounter, z).Value = Format(((bdiv(4 + xline, z) - adiv(4 + xline, z)) / adiv(4 + xline, z)), "##,##0.0#%")
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10

Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

'xlBook.SaveAs testfilenew, FileFormat:=-4143, CreateBackup:=False
xlSheet.SaveAs testfilenew
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub
Sub Printdifxls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.Add

Set xlSheet = xlBook.Worksheets.Item(1)

'Set xlSheet = xlApp.Sheets.Item(1)
monthcounter = 1
linecounter = 0
xint1 = 1
xint2 = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "������� ���������� �����"
xlSheet.Cells(1 + linecounter, 7).Interior.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyDatadif

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - �����"
xlSheet.Cells(1 + linecounter, 3).Value = Months(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNumnew() + linecounter)
j = j + 1
For z = 2 To GetColnew()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If
If (bxls(4 + xline, z) - axls(4 + xline, z) < 0) Then
xlSheet.Cells(3 + j + linecounter, z).Font.ColorIndex = 3
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10
xlSheet.Cells(3 + j + linecounter, z).Value = Format(bxls(4 + xline, z) - axls(4 + xline, z), "##,##0.0#")
Next z
xline = xline + 1

Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

xlBook.SaveAs testfilenew, FileFormat:=-4143, CreateBackup:=False
'xlSheet.SaveAs testfilenew
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub
Sub Printsumxls()

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook As Excel.Workbook
Static monthcounter As Integer
Static linecounter As Integer
Dim i As Integer, j As Integer, z As Integer
Dim resul As Integer, xline As Integer

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.Add

Set xlSheet = xlBook.Worksheets.Item(1)

'Set xlSheet = xlApp.Sheets.Item(1)
monthcounter = 1
linecounter = 0
xint = 1

perife(1) = "�������"
perife(2) = "������� ����������"
perife(3) = "������������"
perife(4) = "������� �������"
perife(5) = "������ �����"
perife(6) = "������ �������"
perife(7) = "������� �������"
perife(8) = "��������� ����������"
perife(9) = "���������� ���������� & ������"
perife(10) = "���� ������"
perife(11) = "������� �������"
perife(12) = "���������"
perife(13) = "�������"
perife(14) = "���� ������������"
perife(15) = "������"
perife(16) = "�������� ����"

xlSheet.Cells(1 + linecounter, 7).Value = "���������� ���������� �����"
xlSheet.Cells(1 + linecounter, 7).Font.ColorIndex = 6
xlSheet.Cells(1 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 7).Font.Size = 14

For resul = 1 To CInt(Text1.Text)

Call CopyData

xlSheet.Cells(1 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(1 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(1 + linecounter, 3).Font.Size = 12
'xlSheet.Cells(1 + linecounter, 3).Value = "�������� " & monthcounter & " - �����"
xlSheet.Cells(1 + linecounter, 3).Value = Months(resul)

xlSheet.Cells(2 + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 1).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 1).Font.Size = 10
xlSheet.Cells(2 + linecounter, 1).Value = "������������ ���������"

xlSheet.Cells(2 + linecounter, 2).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 2).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 2).Font.Size = 10
xlSheet.Cells(2 + linecounter, 2).Value = "���, ��� ��� ������ �����/����� �����"

xlSheet.Cells(2 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 3).Font.Size = 10
xlSheet.Cells(2 + linecounter, 3).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 4).Font.Size = 10
xlSheet.Cells(2 + linecounter, 4).Value = "���������� ������� - ���������"

xlSheet.Cells(2 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 5).Font.Size = 10
xlSheet.Cells(2 + linecounter, 5).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 6).Font.Size = 10
xlSheet.Cells(2 + linecounter, 6).Value = "������������ �������� �� ���"

xlSheet.Cells(2 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 7).Font.Size = 10
xlSheet.Cells(2 + linecounter, 7).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 8).Font.Size = 10
xlSheet.Cells(2 + linecounter, 8).Value = "������������ �������� �� ��������"

xlSheet.Cells(2 + linecounter, 9).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 9).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 9).Font.Size = 10
xlSheet.Cells(2 + linecounter, 9).Value = "�������� ���������� ���� ��������� (��� ��������)"

xlSheet.Cells(2 + linecounter, 10).NumberFormat = "@"
xlSheet.Cells(2 + linecounter, 10).Font.FontStyle = "Bold"
xlSheet.Cells(2 + linecounter, 10).Font.Size = 10
xlSheet.Cells(2 + linecounter, 10).Value = "�������� ���������� ���� ��������� + ���"


xlSheet.Cells(3 + linecounter, 3).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 3).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 3).Font.Size = 10
xlSheet.Cells(3 + linecounter, 3).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 4).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 4).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 4).Font.Size = 10
xlSheet.Cells(3 + linecounter, 4).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 5).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 5).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 5).Font.Size = 10
xlSheet.Cells(3 + linecounter, 5).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 6).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 6).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 6).Font.Size = 10
xlSheet.Cells(3 + linecounter, 6).Value = "���������� ���� ���������"

xlSheet.Cells(3 + linecounter, 7).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 7).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 7).Font.Size = 10
xlSheet.Cells(3 + linecounter, 7).Value = "������� ��������"

xlSheet.Cells(3 + linecounter, 8).NumberFormat = "@"
xlSheet.Cells(3 + linecounter, 8).Font.FontStyle = "Bold"
xlSheet.Cells(3 + linecounter, 8).Font.Size = 10
xlSheet.Cells(3 + linecounter, 8).Value = "���������� ���� ���������"

j = 0
For i = (4 + linecounter) To (GetNum() + linecounter)
j = j + 1
If (j = 16) Then
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "bold"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
Else
xlSheet.Cells(3 + j + linecounter, 1).NumberFormat = "@"
xlSheet.Cells(3 + j + linecounter, 1).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, 1).Font.Size = 12
End If
xlSheet.Cells(3 + j + linecounter, 1).Value = perife(j)
Next i

j = 0
xline = 0
For i = (4 + linecounter) To (GetNum() + linecounter)
j = j + 1
For z = 2 To GetCol()
If (xline + 4 = 19) Then
xlSheet.Cells(3 + j + linecounter, z).Interior.ColorIndex = 4
End If
xlSheet.Cells(3 + j + linecounter, z).NumberFormat = "General"
xlSheet.Cells(3 + j + linecounter, z).Font.FontStyle = "Normal"
xlSheet.Cells(3 + j + linecounter, z).Font.Size = 10
xlSheet.Cells(3 + j + linecounter, z).Value = Format(axls(4 + xline, z), "##,##0.0#")
Next z
xline = xline + 1
Next i

monthcounter = monthcounter + 1
linecounter = linecounter + 23

Next resul

xlBook.SaveAs testfile, FileFormat:=-4143, CreateBackup:=False
xlBook.Close False
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

End Sub

Private Sub file_open_Click()

If (Len(Label2.Caption) <> 0) Then
CommonDialog3.Filter = "Excel|*.xlsx"
CommonDialog3.ShowOpen
Label3.Caption = "2. ������" & " : " & CommonDialog3.FileName
Label4.Caption = CommonDialog3.FileName
Else
CommonDialog1.Filter = "Excel|*.xlsx"
CommonDialog1.ShowOpen
Label1.Caption = "1. ������" & " : " & CommonDialog1.FileName
Label2.Caption = CommonDialog1.FileName

End If

End Sub

Private Sub Form_Load()

Label2.Caption = Empty
analog.Enabled = False
derivative.Enabled = False
analog_month.Enabled = False
derivat_month.Enabled = False

End Sub

Private Sub open_month_Click()

OLE3.CreateLink (testfileadd)
OLE3.DoVerb

End Sub

Private Sub open_statistic_Click()

OLE2.CreateLink (testfilenew)
OLE2.DoVerb

End Sub

Private Sub open_sum_Click()

OLE1.CreateLink (testfile)
OLE1.DoVerb

End Sub

Private Sub sum_Click()

testfile = Empty
CommonDialog2.FileName = Empty

CommonDialog2.Filter = "Excel|*.xls"
CommonDialog2.ShowSave
testfile = CommonDialog2.FileName
If Len(testfile) = 0 Then
Exit Sub
End If
If Len(Text1.Text) = 0 Then
MsgBox "�������� �������� ��� ������ ��� �����!"
Exit Sub
End If
Label5.Caption = "�������� ����������..."
Call CleanData
Call Printsumxls
MsgBox "� ���������� ��������� ��� " & CommonDialog2.FileTitle & " ������������!"
Label5.Caption = "� ����������� ��� ������������� ��������!"

End Sub
