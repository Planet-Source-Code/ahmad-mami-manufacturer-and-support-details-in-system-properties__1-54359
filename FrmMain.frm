VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Manufacturer and Support Details "
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1710
      Left            =   120
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   1650
      ScaleWidth      =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2700
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1440
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Logo Picture"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Data"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmMain.frx":F0BA
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "if you like it please vote "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DDC539&
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub Command1_Click()
Dim SysDir As String
Dim leng As Integer
Dim Info As String
SysDir = Space(260)
leng = GetSystemDirectory(SysDir, 260)
SysDir = Mid(SysDir, 1, leng) & "\OEMINFO.INI"
Info = Text1.Text
Open SysDir For Binary Access Write As #1
Put #1, , Info
Close #1
SavePicture Picture1.Image, Mid(SysDir, 1, leng) & "\OEMLOGO.BMP"
MsgBox "Done :)"
End Sub

Private Sub Command2_Click()
CD.ShowOpen
Picture1.Picture = LoadPicture(CD.FileName)
End Sub

'Do You Want To Do It Manually? ok here is how

'To add the manufacturer and support information you _
need to create two new files in the Windows system _
directly, normally 'c:\windows\system' for Windows 9x _
and 'c:\winnt\system32' for Windows NT/2000/xp.

'The first file is a text file called 'OEMINFO.INI'. _
To create the file open notepad and copy the template _
below, make any changes and save the file in the System _
directory.

';----------------------------
'OEMINFO.INI Template
'[General]
'Manufacturer=Enter the Company Name Here
'Model=Enter the Computer Model Name Here

'[Support Information]
'Line1=first line of support information
'Line2=second line
'Line3=third line
'Line4=fourth line
';------------------------------

'Create as many lines as you need by incrementing the _
Line number.

'The other file you need to create is a logo file. _
This is a standard Windows bitmap file, but it must be _
saved as 'OEMLOGO.BMP' in the System directory.

'Once you 've created both these files, open System _
Properties from Control Panel, and your logo and company _
name will be listed.

'The details you entered in the [Support Information] _
section will show up when you click on the Support _
Information button.

