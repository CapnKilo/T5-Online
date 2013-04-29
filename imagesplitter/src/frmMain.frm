VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Splitter"
   ClientHeight    =   5790
   ClientLeft      =   4620
   ClientTop       =   2715
   ClientWidth     =   4455
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   Begin VB.Frame Frame3 
      Caption         =   "BitBlt images (Do not show)"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   4215
      Begin VB.PictureBox picConversion 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1080
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   4215
      Begin VB.CommandButton cmdSplit 
         Caption         =   "Split Image"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         Caption         =   "Width: N/A"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblSprites 
         AutoSize        =   -1  'True
         Caption         =   "Sprites:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2235
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtEvents 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original Image"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Image"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtStartAt 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   1725
         Width           =   1935
      End
      Begin VB.TextBox txtSizeY 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Text            =   "32"
         Top             =   1725
         Width           =   1935
      End
      Begin VB.FileListBox flList 
         Height          =   1065
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Start At:"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblFrames 
         Caption         =   "Size Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()

    picConversion.Height = 1
    picConversion.Width = 1
    
    If flList.ListIndex < 0 Then
        MsgBox "You must select a file to convert first!", , "Error"
        Exit Sub
    End If
    
    FileName = Split(frmMain.flList.List(frmMain.flList.ListIndex), ".", , vbTextCompare)
    Extension = "." & FileName(1)
    
    Select Case Extension
    
        Case ".bmp"
        Case ".gif"
        Case ".jpg"
        Case Else
            MsgBox "This does not support original saved as " & UCase$(Extension) & "! Reselect your image to convert.", , "Error"
            flList.ListIndex = -1
            Exit Sub
            
    End Select
    
    picImage.Picture = LoadPicture(App.Path & "\original\" & flList.List(flList.ListIndex))
    AddToLog "Loaded " & flList.List(flList.ListIndex) & "."
    UpdateProperties
    
    Me.Height = 6195
End Sub

Private Sub cmdSplit_Click()
    SplitImage
End Sub

Private Sub Form_Load()

    frmMain.Height = 4635

    If Dir$(App.Path & "\converted", vbDirectory) <> "converted" Then
        MkDir App.Path & "\converted"
    End If
    
    If Dir$(App.Path & "\original", vbDirectory) <> "original" Then
        MkDir App.Path & "\original"
    End If
    
    flList.Path = App.Path & "\original"
    
    If flList.ListCount = 0 Then
        MsgBox "No original found in folder \original\! Aborting!", , "Error"
        End
    End If
    
    flList.ListIndex = 0
    
    txtEvents.Text = "-Eclipse Origins image splitter-"
    
End Sub

