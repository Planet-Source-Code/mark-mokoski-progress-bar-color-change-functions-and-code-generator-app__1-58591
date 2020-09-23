VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPB_Demo 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress Bar Color Change Functions Demo"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frmPB_Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPB_Demo.frx":1272
   ScaleHeight     =   7800
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Generate Code"
      ForeColor       =   &H00FFFF00&
      Height          =   4695
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   6735
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy Code to Clipboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         MouseIcon       =   "frmPB_Demo.frx":157C
         MousePointer    =   99  'Custom
         Picture         =   "frmPB_Demo.frx":1886
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenCode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Generate Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3480
         MouseIcon       =   "frmPB_Demo.frx":1CC8
         MousePointer    =   99  'Custom
         Picture         =   "frmPB_Demo.frx":1FD2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox rtbCode 
         Height          =   2415
         Left            =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4260
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   32768
         TextRTF         =   $"frmPB_Demo.frx":2414
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtControlName 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add this code to your form ""Form_Load ( )"" event"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Your Control Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPB_Demo.frx":248B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.CheckBox chkBorder 
      BackColor       =   &H00808000&
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3765
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CheckBox chkBorder 
      BackColor       =   &H00808000&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   1935
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CheckBox chkAppearance 
      BackColor       =   &H00808000&
      Caption         =   "Flat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   3765
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1020
   End
   Begin VB.CheckBox chkAppearance 
      BackColor       =   &H00808000&
      Caption         =   "3D"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   1935
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1245
   End
   Begin VB.Timer countTimer 
      Interval        =   150
      Left            =   6840
      Top             =   600
   End
   Begin VB.CheckBox chkScroll 
      BackColor       =   &H00808000&
      Caption         =   "Standard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   1935
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2115
      Width           =   1260
   End
   Begin VB.CheckBox chkScroll 
      BackColor       =   &H00808000&
      Caption         =   "Smooth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3765
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Caption         =   "Color Functions"
      ForeColor       =   &H00FFFF80&
      Height          =   1335
      Left            =   90
      TabIndex        =   1
      Top             =   645
      Width           =   4815
      Begin VB.CommandButton cmdDefaltColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set to Default Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         MouseIcon       =   "frmPB_Demo.frx":2532
         MousePointer    =   99  'Custom
         Picture         =   "frmPB_Demo.frx":283C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBackColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Change BackColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1680
         MouseIcon       =   "frmPB_Demo.frx":2C7E
         MousePointer    =   99  'Custom
         Picture         =   "frmPB_Demo.frx":2F88
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdForeColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Change ForeColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MouseIcon       =   "frmPB_Demo.frx":33CA
         MousePointer    =   99  'Custom
         Picture         =   "frmPB_Demo.frx":36D4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5240
      MouseIcon       =   "frmPB_Demo.frx":3B16
      MousePointer    =   99  'Custom
      Picture         =   "frmPB_Demo.frx":3E20
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   880
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   225
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblBackColor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current BackColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblForeColor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   28
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current ForeColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   2
      Left            =   3300
      TabIndex        =   18
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Border"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2700
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   1
      Left            =   3300
      TabIndex        =   13
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   345
      TabIndex        =   11
      Top             =   2420
      Width           =   1395
   End
   Begin VB.Label labelPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   180
      Left            =   150
      TabIndex        =   10
      Top             =   30
      Width           =   6720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   0
      Left            =   3300
      TabIndex        =   8
      Top             =   2130
      Width           =   270
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrolling Style"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   345
      TabIndex        =   7
      Top             =   2130
      Width           =   1410
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Selected Text"
      End
   End
End
Attribute VB_Name = "frmPB_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    '***********************************************************
    '
    '   Progress Bar Color API Demo
    '
    '   Mark Mokoski
    '   31-JAN-2005
    '   markm@cmtelephone.com
    '   http://www.rjillc.com
    '
    '   Demo application to show the use of API calls
    '   to change foreground and background colors
    '   for the standard MS Common Controls Progress Bar.
    '   See modProgBarColor for function code and
    '   API declairs / constants / types.
    '
    '   With the Demo Application comes a code Generator.
    '   Once you get the Demo Progress Bar to look like
    '   what you want on your form, enter your control name then
    '   click the "Generate Code" button.
    '   Then copy the code to the Clipboard and
    '   paste into your application's form "Form_Load ( )" event.
    '   You can use the provided "Copy Code to Clipboard" button,
    '   or do the normal highlight/select and right click
    '   to bring up the copy menu.
    '
    '   All you need to do is place the control on your form,
    '   size it and set the Min/Max limits. The pasted code
    '   will set the style and colors.
    '
    '***********************************************************
    Option Explicit

    Dim pbCurrentForeColor            As Long
    Dim pbCurrentBackColor            As Long
    Dim counter                       As Integer


Private Sub chkAppearance_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        Select Case Index
            Case 0  '3D Appearance select
                chkAppearance(0).Value = 1
                chkAppearance(1).Value = 0
                ProgressBar.Appearance = cc3D

                
            Case 1  'Flat Appearance selected
                chkAppearance(0).Value = 0
                chkAppearance(1).Value = 1
                ProgressBar.Appearance = ccFlat

        End Select

    pbForeColor ProgressBar, pbCurrentForeColor
    pbBackColor ProgressBar, pbCurrentBackColor
  
    cmdEnd.SetFocus


End Sub

Private Sub chkBorder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        Select Case Index
            Case 0  'Border 'Yes" selected
                chkBorder(0).Value = 1
                chkBorder(1).Value = 0
                ProgressBar.BorderStyle = ccFixedSingle

                
            Case 1  'Border "None" selected
                chkBorder(0).Value = 0
                chkBorder(1).Value = 1
                ProgressBar.BorderStyle = ccNone

        End Select

    pbForeColor ProgressBar, pbCurrentForeColor
    pbBackColor ProgressBar, pbCurrentBackColor
    cmdEnd.SetFocus

End Sub

Private Sub chkScroll_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        Select Case Index
            Case 0  'Standard Scroll (segmented) selected
                chkScroll(0).Value = 1
                chkScroll(1).Value = 0
                ProgressBar.Scrolling = ccScrollingStandard

                
            Case 1  'Solid (smooth) scroll selected
                chkScroll(0).Value = 0
                chkScroll(1).Value = 1
                ProgressBar.Scrolling = ccScrollingSmooth

        End Select

    pbForeColor ProgressBar, pbCurrentForeColor
    pbBackColor ProgressBar, pbCurrentBackColor
    cmdEnd.SetFocus

End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear
    rtbCode.SelStart = 0
    'Select all text
    rtbCode.SelLength = Len(rtbCode.Text) + 1
    'Copy it into the Clipboard
    Clipboard.SetText (rtbCode.SelText)
    'Reset the selected text
    rtbCode.SelStart = 0

    cmdEnd.SetFocus

End Sub

Private Sub cmdDefaltColor_Click()

    'Set the Progress Bar to the default colors
    pbDefaultColor ProgressBar
    pbCurrentForeColor = CLR_DEFAULT
    pbCurrentBackColor = CLR_DEFAULT
    lblForeColor.Caption = "Default Color"
    lblBackColor.Caption = "Default Color"

End Sub

Private Sub cmdEnd_Click()

    Unload Me

End Sub

Private Sub cmdGenCode_Click()

    'Clear any text in the Code RTB Control
    rtbCode.Text = ""

    'Start writing at the top
    rtbCode.SelStart = 0
    'Write the header info in comment green like VB IDE
    rtbCode.SelColor = &H8000&     'Dark Green
    rtbCode.SelText = "'******** Code Added - " & Now & " ********" & vbCrLf & _
    "'********  Progress Bar Color By Mark Mokoski   ********" & vbCrLf
    rtbCode.SelStart = Len(rtbCode.Text) + 1
    'Change to black text
    rtbCode.SelColor = vbBlack
    rtbCode.SelText = "Me.Visible = "
    rtbCode.SelStart = Len(rtbCode.Text) + 1
    'Change highlight to blue for keywords like VB IDE
    rtbCode.SelColor = vbBlue
    rtbCode.SelText = "True" & vbCrLf & vbCrLf
    rtbCode.SelStart = Len(rtbCode.Text) + 1
    'Change back to black for the rest of the code
    rtbCode.SelColor = vbBlack
    
    'Write Progress Bar Color code
    'If the current color is the default, don't write code for it

        If pbCurrentForeColor <> CLR_DEFAULT Then 'If not default color
            rtbCode.SelText = "pbForeColor " & txtControlName.Text & ", " & "&H" & Hex(pbCurrentForeColor) & "&" & vbCrLf
            rtbCode.SelStart = Len(rtbCode.Text) + 1
        End If
        
    'If the current color is the default, don't write code for it

        If pbCurrentBackColor <> CLR_DEFAULT Then 'If not default color
            rtbCode.SelColor = vbBlack
            rtbCode.SelText = "pbBackColor " & txtControlName.Text & ", " & "&H" & Hex(pbCurrentBackColor) & "&" & vbCrLf
            rtbCode.SelStart = Len(rtbCode.Text) + 1
        End If

    rtbCode.SelColor = vbBlack
    
    'Scroll code

        Select Case chkScroll(0).Value
            Case 0  'Smooth (solid) scrolling
                rtbCode.SelText = txtControlName.Text & ".Scrolling = ccScrollingSmooth" & vbCrLf
            Case 1  'Standard (segmented) scrolling
                rtbCode.SelText = txtControlName.Text & ".Scrolling = ccScrollingStandard" & vbCrLf
        End Select

    rtbCode.SelStart = Len(rtbCode.Text) + 1

    rtbCode.SelColor = vbBlack
    
    'Appearance code

        Select Case chkAppearance(0).Value
            Case 0  'Flat Appearance
                rtbCode.SelText = txtControlName.Text & ".Appearance = ccFlat" & vbCrLf
            Case 1  '3D Appearance
                rtbCode.SelText = txtControlName.Text & ".Appearance = cc3D" & vbCrLf
        End Select

    rtbCode.SelStart = Len(rtbCode.Text) + 1

    rtbCode.SelColor = vbBlack

    'Border code

        Select Case chkBorder(0).Value
            Case 0  'No Border
                rtbCode.SelText = txtControlName.Text & ".BorderStyle = ccNone" & vbCrLf
            Case 1  'Has Border
                rtbCode.SelText = txtControlName.Text & ".BorderStyle = ccFixedSingle" & vbCrLf
        End Select

    rtbCode.SelStart = Len(rtbCode.Text) + 1
        
    'Write the footer info in comment green like VB IDE
    rtbCode.SelColor = &H8000&    'Dark Green
    rtbCode.SelText = "'******** End Code Add ********"

    cmdCopy.SetFocus

End Sub

Private Sub Form_Load()

    '*******************************************************
    'I used the "Generate Code" feature of this application
    'to set the starting appearance and colors of the Demo
    'Progress Bar. See below
    '*******************************************************
    
    '******** Code Added - 1/31/2005 10:58:43 PM ********
    '********  PB Color Demo By Mark Mokoski   ********
    Me.Visible = True

    '    pbForeColor ProgressBar, &H80FF&
    '    pbBackColor ProgressBar, &HC08000
    '    ProgressBar.Scrolling = ccScrollingStandard
    '    ProgressBar.Appearance = cc3D
    '    ProgressBar.BorderStyle = ccFixedSingle
    '    '******** End Code Add ********

    
    '    pbCurrentForeColor = &H80FF&
    '    pbCurrentBackColor = &HC08000

    pbCurrentForeColor = CLR_DEFAULT
    pbCurrentBackColor = CLR_DEFAULT

    '    lblForeColor.Caption = "&&H" & Hex(pbCurrentForeColor) & "&&"
    '    lblBackColor.Caption = "&&H" & Hex(pbCurrentBackColor) & "&&"

    lblForeColor.Caption = "Default Color"
    lblBackColor.Caption = "Default Color"
    
    'Read current style of Demo Progress Bar
    'and set the proper check boxes
    
        Select Case ProgressBar.Scrolling
            Case 0  'Standard (segmented) scrolling
                chkScroll(0).Value = 1
                chkScroll(1).Value = 0
            Case 1  'Smooth (solid) scrolling
                chkScroll(0).Value = 0
                chkScroll(1).Value = 1
        End Select
        
        Select Case ProgressBar.Appearance
            Case 0  'Flat appearance
                chkAppearance(0).Value = 0
                chkAppearance(1).Value = 1
            Case 1  '3D appearance
                chkAppearance(0).Value = 1
                chkAppearance(1).Value = 0
        End Select
        
        Select Case ProgressBar.BorderStyle
            Case 0  'No border
                chkBorder(0).Value = 0
                chkBorder(1).Value = 1
            Case 1  'Has border
                chkBorder(0).Value = 1
                chkBorder(1).Value = 0
        End Select

    'Disable Code Generate and Code Copy buttons
    'until there is something to do with them
    cmdGenCode.Enabled = False
    cmdGenCode.BackColor = vbButtonFace
    cmdCopy.Enabled = False
    cmdCopy.BackColor = vbButtonFace

End Sub

Private Sub cmdForeColor_Click()

    'Set new Progress Bar forecolor
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.Flags = cdlCCRGBInit
    'CommonDialog1.Flags = cdlCCFullOpen
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the Progress Bar foreground color to selected color
    pbCurrentForeColor = CommonDialog1.Color
    pbForeColor ProgressBar, pbCurrentForeColor
    lblForeColor.Caption = "&&H" & Hex(pbCurrentForeColor) & "&&"

    Exit Sub

ErrHandler:
    ' User pressed the Cancel button

End Sub

Private Sub cmdBackColor_Click()

    'Set new Progress Bar backcolor
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.Flags = cdlCCRGBInit
    'CommonDialog1.Flags = cdlCCFullOpen
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the ProgressBar background color to selected color
    pbCurrentBackColor = CommonDialog1.Color
    pbBackColor ProgressBar, pbCurrentBackColor
    lblBackColor.Caption = "&&H" & Hex(pbCurrentBackColor) & "&&"

    Exit Sub

ErrHandler:
    ' User pressed the Cancel button

End Sub

Private Sub countTimer_Timer()

    'Just counts from 0 to 100 for demo of a Progress Bar
    counter = counter + 1

        If counter = 101 Then counter = 0
        
    ProgressBar.Value = counter
    'Show the count as percent (just for effect)
    labelPercent.Caption = counter & " %"
    
End Sub

Private Sub mnuCopy_Click()

    Clipboard.Clear
    'Copy selected text to Clipboard
    Clipboard.SetText (rtbCode.SelText)
    'Reset the selected text
    rtbCode.SelStart = 0

End Sub

Private Sub rtbCode_Change()

    'Enable Copy Code button if text is not null

        If rtbCode.Text <> "" Then
            cmdCopy.Enabled = True
            cmdCopy.BackColor = &HC0C0C0
        Else
            cmdCopy.Enabled = False
            cmdCopy.BackColor = vbButtonFace
        End If

End Sub

Private Sub rtbCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If there is no code, or no text selected (highlighted), exit sub

        If rtbCode.Text = "" Or rtbCode.SelText = "" Then Exit Sub

    'If the mouse button is the right button, show copy menu

        If Button = 2 Then

            PopupMenu Me.mnuEdit
        End If

        

End Sub

Private Sub txtControlName_Change()

    'Enable Code Generate button if text is not null

        If txtControlName.Text <> "" Then
            cmdGenCode.Enabled = True
            cmdGenCode.BackColor = &HC0C0C0
        Else
            cmdGenCode.Enabled = False
            cmdGenCode.BackColor = vbButtonFace
        End If

End Sub
