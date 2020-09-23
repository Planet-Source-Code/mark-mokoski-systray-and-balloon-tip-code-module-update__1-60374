VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5190
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Shell and Common Controls Version"
      ForeColor       =   &H000000C0&
      Height          =   960
      Left            =   15
      TabIndex        =   13
      Top             =   1815
      Width           =   5175
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Version of Shell32.DLL is"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         TabIndex        =   17
         Top             =   240
         Width           =   2835
      End
      Begin VB.Label lblShell32Ver 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblShell32Ver"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2985
         TabIndex        =   16
         Top             =   165
         Width           =   2100
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Version of Comctl32.DLL is"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   15
         Top             =   570
         Width           =   2835
      End
      Begin VB.Label lblComctl32Ver 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblComctl32Ver"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2985
         TabIndex        =   14
         Top             =   555
         Width           =   2100
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pop Up Baloonn Tip in Tray"
      ForeColor       =   &H000000C0&
      Height          =   1965
      Left            =   20
      TabIndex        =   5
      Top             =   2760
      Width           =   5160
      Begin VB.CommandButton cmdTrayTip 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pop Up Systray Tip"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form1.frx":1272
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   2325
      End
      Begin VB.OptionButton IconOption 
         Caption         =   "No Icon"
         Height          =   210
         Index           =   0
         Left            =   2595
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton IconOption 
         Caption         =   "INFO Icon"
         Height          =   240
         Index           =   1
         Left            =   2595
         TabIndex        =   10
         Top             =   525
         Width           =   2025
      End
      Begin VB.OptionButton IconOption 
         Caption         =   "WARNING Icon"
         Height          =   240
         Index           =   2
         Left            =   2595
         TabIndex        =   9
         Top             =   765
         Width           =   2025
      End
      Begin VB.OptionButton IconOption 
         Caption         =   "ERROR Icon"
         Height          =   240
         Index           =   3
         Left            =   2595
         TabIndex        =   8
         Top             =   1005
         Width           =   2025
      End
      Begin VB.OptionButton IconOption 
         Caption         =   "GUID Icon (Shell Ver 6.x Only"
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   2595
         TabIndex        =   7
         Top             =   1245
         Width           =   2490
      End
      Begin VB.CheckBox chkMultiTip 
         Caption         =   "Allow Multiple Balloon Tips (Queuing) Shell32.dll Verion 5.x Only"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   1620
         Width           =   4965
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   20
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":19BE
      Top             =   0
      Width           =   5160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Code by"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   4695
      Width           =   5085
   End
   Begin VB.Label websiteLabel 
      Alignment       =   2  'Center
      Caption         =   "www.cmtelephone.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   30
      MouseIcon       =   "Form1.frx":1BF7
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to goto my web site"
      Top             =   5580
      Width           =   5085
   End
   Begin VB.Label emailLabel 
      Alignment       =   2  'Center
      Caption         =   "markm@cmtelephone.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   30
      MouseIcon       =   "Form1.frx":1F01
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Click to send mail to Mark Mokoski"
      Top             =   5265
      Width           =   5085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mark Mokoski"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   5010
      Width           =   5085
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Form to Screen"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMclose 
         Caption         =   "Close Menu"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End Program"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Addtional Information"
      Begin VB.Menu mnuFormCode 
         Caption         =   "Form Code"
      End
      Begin VB.Menu mnuMSDN1 
         Caption         =   "MSDN Ref 1"
      End
      Begin VB.Menu mnuMSDN2 
         Caption         =   "MSDN Ref 2"
      End
      Begin VB.Menu mnuMSDN3 
         Caption         =   "MSDN Ref 3"
      End
      Begin VB.Menu mnuMSDN4 
         Caption         =   "MSDN Ref 4"
      End
      Begin VB.Menu mnuMSDN5 
         Caption         =   "MSDN Ref 5"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    '********************************************************************
    '
    'Systray, Balloon Tool Tip add-in code to the form
    '
    'Mark Mokoski
    'markm@cmtelephone.com
    'www.cmtelephone.com
    '
    '6-NOV-2004
    '
    'See Systray Form Code.txt in the ZIP file for form add-in's to make it all work
    '
    'Also see Microsoft Knowledge base http://support.microsoft.com/default.aspx?scid=kb;en-us;149276
    'for more information.
    '
    'This code is based on the Microsoft Knowledge Base code.
    '********************************************************************
    
    Dim MessageIcon            As Integer 'Icon used for popup balloon test
    
    

Private Sub chkMultiTip_Click()

        If chkMultiTip.Value = Checked Then
            multiTip = True
        Else
            multiTip = False
        End If

End Sub

Private Sub cmdTrayTip_Click()

    '
    'Icons,
    '0 = none
    '1 = Info
    '2 = Warning
    '3 = Error
    '4 = GUID
    '

    Call PopupBalloon(Me, "This is a Tray Balloon Message" & vbCrLf & "at " & Time$ & " on " & Date$, "Systray Test Message", MessageIcon)

End Sub

Private Sub Form_Load()

    '
    'Icons,
    '0 = none
    '1 = Info
    '2 = Warning
    '3 = Error
    '4 = GUID
    '

    Dim WinNt            As Object
    Dim systemfolder

    multiTip = False
    MessageIcon = 0
    '********************************************************************
    'Get version info for %systemroot%/system32/shell32.dll
    
    Set WinNt = CreateObject("Scripting.FileSystemObject")
    systemfolder = WinNt.GetSpecialFolder(0)

    lblShell32Ver.Caption = FileInfo(systemfolder & "/system32/shell32.dll").ProductVersion
    
    'If shell32.dll is Ver6 (WinXP), allow GUID Icon

        If Val(FileInfo.ProductVersion) > 5 Then
            IconOption(4).Enabled = True
            chkMultiTip.Enabled = False
            multiTip = False
        End If
        
    lblComctl32Ver.Caption = FileInfo(systemfolder & "/system32/comctl32.dll").ProductVersion
    
    'If you want the form to be in the tray on startup add this
    Call SystrayOn(Me, "The Form is visible on the screen")
    Call PopupBalloon(Me, "Put your Message Here !", "Balloon Tool Tip", 1)







    Set WinNt = Nothing
End Sub

Private Sub Form_Resize()

    '********************************************************************
    'Add this to resize event to hide in tray on minimize
    '
    'Icons,
    '0 = none
    '1 = Info
    '2 = Warning
    '3 = Error
    '4 = GUID
    '

        If Me.WindowState = vbMinimized Then
            Call SystrayOn(Me, "Double Click to Restore Me back to the screen")
            Call ChangeSystrayToolTip(Me, "Double Click to Restore Me back to the screen")
            Call PopupBalloon(Me, "App is now hidden in the Systray !" + vbCrLf + "Double click Icon to restore", "Balloon Tool Tip", 1)
            Me.Hide
        End If

End Sub

Private Sub Form_Terminate()

    '********************************************************************
    'If you don't remove icon from tray on double click show, add this
    'good idea

    Call SystrayOff(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SystrayOff(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*********************************************************************
    'Add this event code to repond to mouse over and clicks on icon in the tray

    Static lngMsg            As Long
    Dim blnflag              As Boolean

    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then

            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup menu on right-click
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        'Reference the menu object of the form below for popup
                        PopupMenu Me.mnuFile

                    Case WM_LBUTTONDBLCLK   'SHow form on left-dblclick
                        'Use line below if you want to remove tray icon on dbclick show form.
                        'If not, be sure to put Systrayoff in form unload and terminate events.
                        'Call SystrayOff(Me)
                        Call ChangeSystrayToolTip(Me, "The Form is visible on the screen")
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        Me.WindowState = vbNormal
                        Me.Show
                        Me.SetFocus
            
                End Select
        
            blnflag = False
        
        End If
    
End Sub

Private Sub emailLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:name@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute hWnd, vbNullString, "mailto:markm@cmtelephone.com?Subject=Questions or Comments on Systray Code Module. %09 ", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub IconOption_Click(Index As Integer)

    MessageIcon = Index
 
End Sub

Private Sub mnuEnd_Click()

    Unload Me

End Sub

Private Sub mnuFormCode_Click()

    'Addtional code need in form events for Systray to work OK
    ShellExecute hWnd, vbNullString, App.Path & "\systray form code.txt", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN1_Click()

    'How To Use Icons with the Windows 95.htm
    ShellExecute hWnd, vbNullString, App.Path & "\html\How To Use Icons with the Windows 95.htm", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN2_Click()

    'How To Manipulate Icons in the System Tray with Visual Basic.htm
    ShellExecute hWnd, vbNullString, App.Path & "\html\How To Manipulate Icons in the System Tray with Visual Basic.htm", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN3_Click()

    ShellExecute hWnd, vbNullString, "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/versions.asp", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN4_Click()

    ShellExecute hWnd, vbNullString, "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/structures/notifyicondata.asp", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN5_Click()

    ShellExecute hWnd, vbNullString, "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shell_notifyicon.asp", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuRestore_Click()

    Me.WindowState = vbNormal
    Me.Show

End Sub

Private Sub Option1_Click(Index As Integer)

    MessageIcon = Index

End Sub

Private Sub websiteLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "http://www.domain.com", vbNullString, vbNullString, vbNormalFocus
    ShellExecute hWnd, vbNullString, "http://www.rjillc.com", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub
