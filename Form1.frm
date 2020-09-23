VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Image Viewer"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13185
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   879
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   7920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   6690
      Left            =   10440
      TabIndex        =   10
      Top             =   1215
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   11800
      _Version        =   393216
      LargeChange     =   100
      Orientation     =   1179648
      SmallChange     =   100
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   11040
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   885
      Width           =   5175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15162
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1566B
            Key             =   "d"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A7DD
            Key             =   "a"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F94F
            Key             =   "b"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":54AC1
            Key             =   "c"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69C33
            Key             =   "e"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EDA5
            Key             =   "back"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7F2B1
            Key             =   "up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   1429
      ButtonWidth     =   1244
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            ImageKey        =   "a"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            ImageKey        =   "b"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Up"
            ImageKey        =   "c"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageKey        =   "d"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            ImageKey        =   "e"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   0
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   5
      Top             =   5640
      Width           =   3975
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   4080
      ScaleHeight     =   446
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   423
      TabIndex        =   3
      Top             =   1200
      Width           =   6375
      Begin VB.PictureBox PicAll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7815
         Left            =   0
         ScaleHeight     =   521
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   489
         TabIndex        =   4
         Top             =   0
         Width           =   7335
         Begin VB.PictureBox picSingle 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   0
            Left            =   0
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   12
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   10920
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   10920
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   10920
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   0
      Top             =   6240
      Width           =   2775
   End
   Begin MSComctlLib.TreeView DirTree 
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7646
      _Version        =   393217
      Indentation     =   471
      Style           =   7
      ImageList       =   "img"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList img 
      Left            =   10920
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7F755
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80099
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8284D
            Key             =   "fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":82AA3
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":82EBF
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":832F5
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8348D
            Key             =   "open"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83612
            Key             =   "remote"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblImgInfo 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   353
      X2              =   353
      Y1              =   530
      Y2              =   540
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   352
      X2              =   352
      Y1              =   530
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   128
      X2              =   128
      Y1              =   530
      Y2              =   550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   129
      X2              =   129
      Y1              =   530
      Y2              =   550
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Image imgStatus 
      Height          =   375
      Left            =   12840
      Picture         =   "Form1.frx":8376E
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape shpFolders 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   840
      Width           =   3975
   End
   Begin VB.Shape shpPath 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4080
      Top             =   840
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuToolbars 
      Caption         =   "&Toolbars"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name: Fabio
' Email: italiangoose@yahoo.com
' My Image Viewer App
' Date Started: sometime in Ausust 2004
' Funcionality: an image viewer similar to ACDSee
'-------------------------------------------------------
' I started creating an image viewer because i thought that
' there isnt really a decent free image viewer out there
' besides XnView (which i discovered later). So I began on
' this one.
' Im quite caught up with time, work and studies, and so i
' decided to share this code with everyone, in the hope that
' this program will continue to grow and improve.
'
' If you do decided to under take this project then let me know.
' I would love to see whats new.
'
' Thanks and Enjoy
'-------------------------------------------------------

Private Const WM_USER = &H400
Private Const CCM_FIRST       As Long = &H2000&
Private Const CCM_SETBKCOLOR  As Long = (CCM_FIRST + 1)

'set progressbar backcolor in IE3 or later
Private Const PBM_SETBKCOLOR  As Long = CCM_SETBKCOLOR

'set progressbar barcolor in IE4 or later
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Dim ThumbW, ThumbH As Integer
Dim cols As Integer
Dim prefile As Integer

Private nNode As Node
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub LoadImages()
On Error Resume Next
Dim x, a As Integer
Dim Y, z, c, r As Integer
    
    Me.MousePointer = 11
    z = 0
    Y = 0
    c = 0
    r = 20
    PicAll.Cls
    PicAll.Height = ((File1.ListCount / cols) + 1) * (ThumbH + 30)
    ProgressBar.Min = 1
    ProgressBar.Max = File1.ListCount
    lblInfo.Caption = "Total " & File1.ListCount & " Files"
    
    For a = 0 To prefile
        Unload picSingle(a)
    Next a
    
    If File1.ListCount <> 0 Then
        For x = 1 To File1.ListCount
                        
            picMain.Cls
            picThumb.Cls
            'Call drawline
            picMain.Picture = LoadPicture(File1.path & "/" & File1.List(x - 1))
             
            If picMain.Width > picMain.Height Then
                picThumb.PaintPicture picMain.Picture, 5, (ThumbH - ((ThumbH - 12) / (picMain.Width / picMain.Height))) / 2, (ThumbH - 12), (ThumbH - 12) / (picMain.Width / picMain.Height)
            Else
                picThumb.PaintPicture picMain.Picture, (ThumbW - ((ThumbW - 12) / (picMain.Height / picMain.Width))) / 2, 5, (ThumbW - 12) / (picMain.Height / picMain.Width), (ThumbW - 12)
            End If
            Call DrawLine
            
            Load picSingle(x)
            picSingle(x).Width = picThumb.Width
            picSingle(x).Height = picThumb.Height + 20
            picSingle(x).Tag = File1.List(x - 1)
            
            picSingle(x).PaintPicture picThumb.Image, 0, 0 ', (c * ThumbW) + (c + 1) * 10, r
            picSingle(x).Top = r + 10 '+ ThumbH + 5
            picSingle(x).Left = (10 + (c * (ThumbW + 30))) ' + (ThumbW - (Len(File1.List(x - 1)) * 5)) / 2) + c * 10
            
            picSingle(x).CurrentY = picThumb.Height + 3
            picSingle(x).Print File1.List(x - 1)
            picSingle(x).ToolTipText = File1.List(x - 1)
            
            c = c + 1
            If c = cols Then
                c = 0
                r = r + ThumbH + 30
            End If
            picSingle(x).Visible = True
            ProgressBar.Value = x + 1
        Next x
    End If
    prefile = x - 1
    
    PicAll.Height = r + ThumbH + 40
    If PicAll.Height < picBack.Height Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
    VScroll1.Min = 10
    VScroll1.Max = PicAll.Height - VScroll1.Height
    ProgressBar.Value = 1
    'VScroll1.LargeChange = PicAll.Height
    Me.MousePointer = 0
End Sub

Private Sub LoadPreview(imageid As Integer)
    picPreview.Cls
    picMain.Cls
    picMain.Picture = LoadPicture(File1.path & "/" & File1.List(imageid))
             
    If picMain.Width > picMain.Height Then
        picPreview.PaintPicture picMain.Picture, 0, (picPreview.Height - ((picMain.Height / picMain.Width) * picPreview.Width)) / 2, picPreview.Width, (picMain.Height / picMain.Width) * picPreview.Width
    Else
        picPreview.PaintPicture picMain.Picture, (picPreview.Width - ((picMain.Width / picMain.Height) * picPreview.Height)) / 2, 0, (picMain.Width / picMain.Height) * picPreview.Height, picPreview.Height
        'MsgBox (picPreview.Width - ((picMain.Width / picMain.Height) * picPreview.Height)) / 2
    End If
    lblImgInfo.Caption = picMain.Width & " x " & picMain.Height
End Sub

Private Sub Form_Load()
    ThumbW = 120
    ThumbH = 120
    picThumb.Width = ThumbW
    picThumb.Height = ThumbH
    Call DrawLine
    Call CalcCols
    
    LoadTreeView
    DirTree.Nodes.Item(2).Expanded = True
    SetProgressBarColour ProgressBar.hwnd, RGB(0, 0, 0)
End Sub

Private Sub CalcCols()
Dim x, r, c As Integer

    cols = PicAll.ScaleWidth \ (ThumbW + 30)
    r = 20
    c = 0
    
    For x = 1 To File1.ListCount
        picSingle(x).Top = r + 10
        picSingle(x).Left = (10 + (c * (ThumbW + 30)))
        c = c + 1
        If c = cols Then
            c = 0
            r = r + ThumbH + 30
        End If
    Next x
    PicAll.Height = r + ThumbH + 40
    If PicAll.Height < picBack.Height Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
    VScroll1.Min = 10
    VScroll1.Max = PicAll.Height - VScroll1.Height
    'Me.Caption = cols
End Sub

Private Sub DrawLine()

    'draw lines
    picThumb.ForeColor = &H8000000C
    picThumb.Line (0, 0)-(ThumbW, 0)
    picThumb.Line (0, 0)-(0, ThumbH)
    picThumb.Line (0, ThumbH - 3)-(ThumbW, ThumbH - 3)
    picThumb.Line (ThumbW - 3, 0)-(ThumbW - 3, ThumbH)
End Sub
'
Private Sub Form_Resize()
    Call ResizeForm
End Sub

Private Sub picSingle_Click(Index As Integer)
    LoadPreview (Index - 1)
End Sub

Private Sub VScroll1_Change()
    PicAll.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    PicAll.Top = -VScroll1.Value
End Sub
'
Private Sub ResizeForm()
    imgStatus.Left = Me.ScaleWidth - imgStatus.Width
    imgStatus.Top = Me.ScaleHeight - imgStatus.Height
    
    shpPath.Width = Me.ScaleWidth - shpPath.Left
    VScroll1.Left = Me.ScaleWidth - VScroll1.Width
    picBack.Width = Me.ScaleWidth - picBack.Left
    PicAll.Width = picBack.Width
    picBack.Height = Me.ScaleHeight - picBack.Top - imgStatus.Height
    VScroll1.Height = picBack.Height - 1
    
    picPreview.Height = picPreview.Width
    picPreview.Top = (picBack.Height + picBack.Top) - picPreview.Height
    
    DirTree.Height = picPreview.Top - DirTree.Top - 5
    
    ProgressBar.Left = picBack.Left + 100
    ProgressBar.Top = imgStatus.Top + 2
    ProgressBar.Height = imgStatus.Height - 2
    ProgressBar.Width = picBack.Width - 100 - imgStatus.Width
    
    lblInfo.Top = imgStatus.Top + 7
    lblImgInfo.Top = lblInfo.Top
    lblImgInfo.Width = ProgressBar.Left - 20 - lblImgInfo.Left
    
    Line1.Y1 = imgStatus.Top + 2
    Line1.Y2 = Line1.Y1 + imgStatus.Height
    Line2.Y1 = imgStatus.Top + 2
    Line2.Y2 = Line2.Y1 + imgStatus.Height
    
    Line3.Y1 = imgStatus.Top + 2
    Line3.Y2 = Line3.Y1 + imgStatus.Height
    Line4.Y1 = imgStatus.Top + 2
    Line4.Y2 = Line4.Y1 + imgStatus.Height
    
    Line3.X1 = ProgressBar.Left - 10
    Line3.X2 = ProgressBar.Left - 10
    Line4.X1 = ProgressBar.Left - 9
    Line4.X2 = ProgressBar.Left - 9
    Call CalcCols
End Sub

'---------------start of dirtree--------------------------

Private Sub DisplayDir(Pth, Parent)
Dim j As Integer
    On Error Resume Next
    Pth = Pth & "\"
    tmp = Dir(Pth, vbDirectory)
    Do Until tmp = ""
        If tmp <> "." And tmp <> ".." Then
            If GetAttr(Pth & tmp) And vbDirectory Then
                List1.AddItem tmp
            End If
        End If
        tmp = Dir
    Loop
    'Add sorted directory names to TreeView
    For j = 1 To List1.ListCount
        Set nNode = DirTree.Nodes.Add(Parent, tvwChild, , List1.List(j - 1), "folder", "open")
    Next j
    List1.Clear
End Sub

Private Sub LoadTreeView()
    DirTree.Nodes.Clear
    Dim DriveNum As String
    Dim DriveType As Long
    DriveNum = 64
    On Error Resume Next
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
            Case 0: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 2: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, "(" & Chr$(DriveNum) & ":)", "remove")
            Case 3: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 4: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 5: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 6: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
        End Select
    Loop
End Sub

Private Sub DirTree_Expand(ByVal Node As MSComctlLib.Node)
Dim j As Integer

    For j = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        
        path = Mid(DirTree.Nodes(j).FullPath, InStr(1, DirTree.Nodes(j).FullPath, ":") - 1, 2) & Mid(DirTree.Nodes(j).FullPath, InStr(1, DirTree.Nodes(j).FullPath, ":") + 2)
        If DirTree.Nodes(j).Children > 0 Then
            If Right(path, 1) <> "\" Then path = path & "\"
            Exit Sub
        End If
        DisplayDir path, DirTree.Nodes(j).Index
    
    Next j
    Node.Selected = True
End Sub

Private Sub DirTree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then DirTree.Nodes.Clear: LoadTreeView
End Sub

Private Sub DirTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim path As String
    
    path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    
    If Node.Children > 0 Then
        If Right(path, 1) <> "\" Then path = path & "\"
            Text1.Text = path
            File1.path = Text1.Text
            Call LoadImages
        Exit Sub
    End If
    DisplayDir path, Node.Index
    Text1.Text = path
    File1.path = Text1.Text
    Call LoadImages
End Sub

'--------------end if dirtree---------------

'--------------change statusbar color-------

Private Sub SetProgressBarColour(hwndProgBar As Long, ByVal clrref As Long)
    Call SendMessage(hwndProgBar, PBM_SETBARCOLOR, 0&, ByVal clrref)
End Sub

