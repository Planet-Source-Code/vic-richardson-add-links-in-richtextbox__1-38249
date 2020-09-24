VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Versalink RichTextBox Example"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12105
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   855
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   7200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link Builder"
      Height          =   3375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   375
         Left            =   4320
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Text            =   "1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "GoTo"
         Height          =   375
         Left            =   4320
         TabIndex        =   30
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse"
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   375
         Left            =   5400
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   4815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":00F3
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Launchable Application"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Link Destination"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Link Title"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   4575
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11456
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":01D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Up to 99 links"
      Height          =   255
      Left            =   7200
      TabIndex        =   33
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   $"Form1.frx":02C9
      Height          =   2055
      Left            =   8760
      TabIndex        =   29
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Current"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Total Links"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Versalink Application"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   7320
      Width           =   3975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Title,  Versalink Destination"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Lnk Start    End"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Csr Pos"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   6600
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load .vlt Files"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save .vlt Files"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import .rtf .txt Files"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New .vlt File"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Application"
      End
   End
   Begin VB.Menu mnuExittext 
      Caption         =   "Exit TextView"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuExitimage 
      Caption         =   "Exit ImageView"
      Visible         =   0   'False
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuNewlink 
         Caption         =   "Make link"
      End
      Begin VB.Menu mnuEditlink 
         Caption         =   "Edit link"
      End
      Begin VB.Menu mnuDeletelink 
         Caption         =   "Delete Link"
      End
      Begin VB.Menu break 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Prototype for VersaLinks embedded in RTF textbox
'  DazyWeb Labs       build   23 - August - 2002
'  email:    vbalthezr@earthlink.net
'
'
'  Idea inspired by code on PSC by:  mike15@blueyonder.co.uk
'
'  (original method was to use the mouse position to determine
'   if the link corordinates matched the mouse coordinates, text
'   editing constantly moved the link coordinates and made keeping
'   track of all positions a laborous task)
'
'  New Methodology is to use the mouse position to define a start point
'  look for underlined text there, then search to each side for the
'  end point of the underlined text and do a string compare of that
'  text to a known list of versalink titles. Advantage is that the same
'  link may be used multiple times in a document and housekeeping is minimal.
'
'
'  Versalinks, as opposed to just hyperlinks and maillinks allow the launch of
'  an application only, an application associated with a file extension and that
'  extension automatically loaded or built in ImageView picturebox or TextView
'  textbox viewers with the destination file displayed. With additional coding,
'  data could be passed to an external app and the result displayed locally...
'  thus the versa in VersaLinks.



Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type tLinkInfo
LinkText As String
LinkDest As String
LinkApp As String
End Type

Private Links(99) As tLinkInfo 'upto 100 links.
Private LnkCnt As Integer ' a counter.
Dim I As Integer
Dim wasI As Integer
Dim isIold As Integer
Dim xx As Integer
Dim lastlength As Long
Dim pastelength As Long
Dim tempStart As Long
Dim Ret As Variant
Dim temppath As String
Dim temppath2 As String
Dim fnum1
Dim temptxt As String
Dim keyword As String
Dim gotalock As Integer
Dim linkstart As Long
Dim linkend As Long
Dim maxcount As Long
Dim linktodel As Long
Dim tmpTT As Long
Dim templink As Integer
Dim setfocuslock As Integer
Dim setfocuslock2 As Integer

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'sndPlaySound Constants
Const SND_ALIAS = &H10000
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_MEMORY = &H4
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_SYNC = &H0

Dim SoundFile As String
Dim playtoggle As Integer


Function GetPosOverCursor(Box As Object, X, Y) As Long
Dim tmpPnt As POINTAPI, tmpPos As Long
tmpPnt.X = X / Screen.TwipsPerPixelX
tmpPnt.Y = Y / Screen.TwipsPerPixelY
tmpPos = SendMessage(Box.hWnd, &HD7, 0&, tmpPnt)
GetPosOverCursor = tmpPos
Label1.Caption = CStr(tmpPos)

End Function


Private Sub Command5_Click() 'goto link to edit

templink = Val(Text4.Text)
If templink < 1 Or templink > LnkCnt Then
templink = 1
Text4.Text = "1"
End If

Text1.Text = Links(templink).LinkText
Text2.Text = Links(templink).LinkDest
Text3.Text = Links(templink).LinkApp


End Sub


Private Sub Command6_Click() 'save editted link info

templink = Val(Text4.Text)
Links(templink).LinkText = Trim(Text1.Text)
Links(templink).LinkDest = Text2.Text
Links(templink).LinkApp = Text3.Text

End Sub


Private Sub Form_Load()

LnkCnt = 0
rtb.SelText = "             Hello World!" + vbCrLf
rtb.SelText = vbCrLf + "  Examples of versalinks embedded in RichTextBox controls."
rtb.SelText = vbCrLf + " "
rtb.SelText = vbCrLf + "  You can insert Microsoft Excel or Word files, drawing files that open"
rtb.SelText = vbCrLf + "  in Paint or CAD programs, hyperlinks, mail to's or use the built in Apps"
rtb.SelText = vbCrLf + "  TextView, ImageView and PlayWav to view text, images and play sounds."
rtb.SelText = vbCrLf + " " + vbCrLf
AddLink "Planet Source Code", "http://www.pscode.com", ""
rtb.SelText = " - Home of great source code." & vbCrLf + vbCrLf
AddLink "Email the author", "mailto:vrbalthezr@earthlink.net", ""
rtb.SelText = " - Vic Richardson"
rtb.SelText = " " + vbCrLf + vbCrLf
AddLink "Windows Ding Sound", "C:/Windows/Media/ding.wav", "PlayWav"
rtb.SelText = " - Heard this before?" & vbCrLf
rtb.SelText = vbCrLf + ""
AddLink "Windows ini Textfile", "C:/Windows/win.ini", "TextView"
rtb.SelText = " - I guessed this might be on all pc's."
rtb.SelText = vbCrLf + " " + vbCrLf
AddLink "Windows Cloud Image", "C:/Windows/Clouds.bmp", "ImageView"
rtb.SelText = " - View from Microsoft campus?"
rtb.SelText = vbCrLf + " " + vbCrLf
rtb.SelText = "  This example cannot launch an app that is not associated with the" + vbCrLf
rtb.SelText = "  file extension in the Link Destination and get that app to automatically" + vbCrLf
rtb.SelText = "  load and display that file (using SendKey in a carefully orchestrated manner" + vbCrLf
rtb.SelText = "  works as long as the new application Window doesn't lose focus). As a convenience" + vbCrLf
rtb.SelText = "  the Destination filename is at least pasted to the Windows clipboard to have it" + vbCrLf
rtb.SelText = "  ready if the destination app supports cut and paste in the file load window." + vbCrLf
rtb.SelText = " " + vbCrLf
rtb.SelText = "  Future additions could include links on embedded pictures in rtf files (for viewing" + vbCrLf
rtb.SelText = "  enlargements of that pix). Note: This method of using links is not compliant" + vbCrLf
rtb.SelText = "  with the Rich Text Format Standard regarding embedded objects and hyperlinks but" + vbCrLf
rtb.SelText = "  is meant as an enhancement to the RTB control for custom applications such as help files." + vbCrLf
rtb2.Left = rtb.Left
rtb2.Top = rtb.Top
rtb2.Width = rtb.Width
rtb2.Height = rtb.Height
rtb2.BackColor = &HDDDDDD

Picture1.Left = rtb.Left
Picture1.Top = rtb.Top

End Sub


Private Sub mnuImport_Click() 'import rtf file

On Error Resume Next
dlgfile.Filter = "Text Files |*.*;"
dlgfile.FileName = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0
 
temppath = dlgfile.FileName


If temppath <> "" Then
rtb.LoadFile temppath
End If

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)

If Button = 2 Then 'delete image
Picture1.Visible = False
End If

End Sub


Private Sub rtb2_MouseDown(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)


If Button = 2 Then 'delete image
rtb2.Visible = False
End If

End Sub


Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If rtb.SelLength = 0 And setfocuslock2 <> 2 Then
   tempStart = rtb.SelStart
   On Error Resume Next
     If IsCursorOverLink(rtb, CLng(X), CLng(Y)) = True Then
     GetLinkInfo
     End If
   rtb.SelStart = tempStart
End If

If setfocuslock > 0 Then 'allow for mouse twitches when doubleclicking link so link will maintain focus
  setfocuslock = setfocuslock + 1
    If setfocuslock = 10 Then
    setfocuslock = 0
    End If
End If

End Sub


Function IsCursorOverLink(Box As Object, X, Y) As Boolean
Dim tmpT As Long
maxcount = 0

gotalock = 0
IsCursorOverLink = False
On Error Resume Next
tmpT = GetPosOverCursor(Box, X, Y)
rtb.SelStart = tmpT
tmpTT = tmpT

'add routine to recover text on both sides underlined and blue
If rtb.SelUnderline = True Then  'and trb.selcolor = vbblue


   Do While rtb.SelUnderline = True
    rtb.SelStart = rtb.SelStart - 1
     If rtb.SelStart = 0 Then
     Exit Do
     End If
        maxcount = maxcount + 1
     If maxcount = 50 Then
     maxcount = 0
     Exit Do
     End If
   Loop
 
 linkstart = rtb.SelStart + 0
 Label3.Caption = CStr(linkstart)
 rtb.SelStart = rtb.SelStart + 2
 
    Do While rtb.SelUnderline = True
    rtb.SelStart = rtb.SelStart + 1
    If rtb.SelStart = Len(rtb.Text) Then
    Exit Do
    End If
       maxcount = maxcount + 1
    If maxcount = 50 Then
    maxcount = 0
    Exit Do
    End If
   Loop

 linkend = rtb.SelStart - 1
 Label4.Caption = CStr(linkend)
 
 rtb.SelStart = linkstart
 rtb.SelLength = linkend - linkstart
 keyword = rtb.SelText
 rtb.SelLength = 0
 If setfocuslock = 0 Then
 List1.SetFocus 'prevents rapid cursor scan over link during keyword check
 End If
  
Else 'not over a versalink
keyword = ""
gotalock = 0
If setfocuslock = 0 Then  'holdoff if a link is activated so it gets focus
rtb.SetFocus
End If
End If


For I = 0 To LnkCnt
If Links(I).LinkText = keyword Then
gotalock = I
Label16.Caption = CStr(I)
IsCursorOverLink = True
'List1.ListIndex = (gotalock)  'uncomment to autohighlight active link in listbox
Exit For
Else
gotalock = 0
Label16.Caption = " "
End If
Next I



End Function


Sub AddLink(Text As String, Dest As String, App As String) 'used in FormLoad only

LnkCnt = LnkCnt + 1
On Error Resume Next
Links(LnkCnt).LinkDest = Dest
Links(LnkCnt).LinkText = Text
Links(LnkCnt).LinkApp = App

rtb.SelStart = Len(rtb.Text)
rtb.SelText = "  "
rtb.SelUnderline = True
rtb.SelColor = vbBlue
rtb.SelText = Text
rtb.SelStart = Len(rtb.Text)
rtb.SelColor = vbBlack
rtb.SelUnderline = False
rtb.SelText = "  "
Update

End Sub


Function GetLinkInfo() 'for display at bottom of form only
     
     If gotalock <> 0 Then
     Label2.Caption = Links(gotalock).LinkText & " - " & Links(gotalock).LinkDest
     Label12.Caption = Links(gotalock).LinkApp
     Else
     Label2.Caption = ""
     Label12.Caption = ""
     Label3.Caption = ""
     Label4.Caption = ""
     End If
     
End Function


Private Sub rtb_dblclick() 'launch the link application

  On Error Resume Next

  If gotalock <> 0 Then
    If Links(gotalock).LinkApp = "" Then
    setfocuslock = 1
    Ret = ShellExecute(Me.hWnd, "open", Links(gotalock).LinkDest, "", "", 5)
    
    Else
      If gotalock <> 0 Then
        If Links(gotalock).LinkApp = "ImageView" Then
        Picture1.Picture = LoadPicture(Links(gotalock).LinkDest)
        Picture1.Visible = True
        mnuExitimage.Visible = True
        Else
          If Links(gotalock).LinkApp = "TextView" Then
          rtb2.LoadFile (Links(gotalock).LinkDest)
          rtb2.Visible = True
          mnuExittext.Visible = True
          Else
              If Links(gotalock).LinkApp = "PlayWav" Then
              PlayMe
              Else
               Clipboard.SetText Links(gotalock).LinkDest
               Ret = Shell(Links(gotalock).LinkApp, 1)
               setfocuslock = 1
              End If
          End If
        End If
      End If
    End If
  End If

End Sub


Private Sub PlayMe()

On Error Resume Next

SoundFile = Links(gotalock).LinkDest
If SoundFile <> "" And Right(SoundFile, 3) = "wav" Then
sndPlaySound SoundFile, SND_ASYNC Or SND_FILENAME
End If

End Sub


Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

If Button = 2 Then  'right click edit menu
PopupMenu Form1.mnupopup
End If

End Sub


Private Sub mnuNewlink_click()

tempStart = rtb.SelStart

If Frame1.Visible = False Then
Frame1.Visible = True
setfocuslock2 = 2

Text1.Text = Links(LnkCnt + 1).LinkText
Text2.Text = Links(LnkCnt + 1).LinkDest
Text3.Text = Links(LnkCnt + 1).LinkApp
Text4.Text = CStr(LnkCnt + 1)
Else
Frame1.Visible = False
End If

End Sub

Private Sub mnuEditlink_click()

tempStart = rtb.SelStart

If Frame1.Visible = False Then
Text1.Text = Links(gotalock).LinkText
Text2.Text = Links(gotalock).LinkDest
Text3.Text = Links(gotalock).LinkApp
Text4.Text = CStr(gotalock)
templink = gotalock
Frame1.Visible = True
setfocuslock2 = 2
Else
Frame1.Visible = False
End If

End Sub


Private Sub Command2_Click() 'exit link maker

Frame1.Visible = False
setfocuslock2 = 0

End Sub


Private Sub Command1_Click() 'insert link info

templink = Val(Text4.Text)
  If templink > 99 Then
  templink = 99
  Text4.Text = "99"
  End If
   If templink < 1 Or templink > LnkCnt + 1 Then
   MsgBox ("Not a valid link number! Must be between 1 and (Total Links + 1).")
   Else
     If Trim(Text1.Text) = "" Then
     MsgBox ("No Link Title. Cannot insert without one!")
     Else
     On Error Resume Next
        If rtb.SelUnderline = False Then 'inserting brand new link
        rtb.SelStart = tempStart
        rtb.SelText = "  "
        LnkCnt = templink
        Links(templink).LinkText = Trim(Text1.Text)
        Links(templink).LinkDest = Text2.Text
        Links(templink).LinkApp = Text3.Text
        rtb.SelColor = vbBlue
        rtb.SelUnderline = True
        rtb.SelText = Text1.Text
        rtb.SelColor = vbBlack
        rtb.SelUnderline = False
        rtb.SelText = "  "
        Else 'already on a link so replace with editted link title
        rtb.SelStart = linkstart
        rtb.SelLength = linkend - linkstart
        Form1.rtb.SelText = Trim(Text1.Text)
        Links(templink).LinkText = Trim(Text1.Text)
        Links(templink).LinkDest = Text2.Text
        Links(templink).LinkApp = Text3.Text
        End If

     Frame1.Visible = False
     setfocuslock2 = 0
     Update
     End If
   End If


End Sub


Private Sub mnuDeletelink_click()

     On Error Resume Next
     
     linktodel = gotalock
     tempStart = rtb.SelStart
     
     rtb.SelStart = linkstart
     rtb.SelLength = Len(Links(linktodel).LinkText)
     Label2.Caption = " "
     Label3.Caption = " "
     Label4.Caption = " "
     
     Form1.rtb.SelText = vbNullString
     
    
    

  lastlength = Len(rtb.Text) 'keep track of current length
     
     If linktodel < LnkCnt And linktodel > 0 Then
     For xx = linktodel To LnkCnt - 1
     Links(xx).LinkText = Links(xx + 1).LinkText
     Links(xx).LinkDest = Links(xx + 1).LinkDest
     Links(xx).LinkApp = Links(xx + 1).LinkApp
     Next xx
     Else
     Links(linktodel).LinkText = ""
     Links(linktodel).LinkDest = ""
     Links(linktodel).LinkApp = ""
     End If
     
     LnkCnt = LnkCnt - 1
       If LnkCnt < 0 Then
       LnkCnt = 0
       End If
     rtb.SelStart = tempStart
     rtb.SelLength = 0
    
     Update
  
End Sub


Private Sub Command3_Click() 'browse for app path

On Error Resume Next
dlgfile.Filter = ""
dlgfile.FileName = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0
 
temppath = dlgfile.FileName

If temppath <> "" Then
Text3.Text = temppath
End If

End Sub


Private Sub Command4_Click() 'browse for file path

On Error Resume Next
dlgfile.Filter = ""
dlgfile.FileName = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0
 
temppath = dlgfile.FileName

If temppath <> "" Then
Text2.Text = temppath
End If

End Sub


Private Sub mnuLoad_Click() 'load rtf file with versalinks

On Error Resume Next
dlgfile.Filter = "Rich Text Files with Links|*.vlt;"
dlgfile.FileName = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0
 
temppath = dlgfile.FileName

If temppath <> "" Then
  fnum1 = FreeFile
On Error Resume Next
Open temppath For Input As #fnum1 'save link info
On Error Resume Next
For xx = 0 To 99
Input #fnum1, Links(xx).LinkText
Input #fnum1, Links(xx).LinkDest
Input #fnum1, Links(xx).LinkApp
Next xx
Input #fnum1, LnkCnt
Close fnum1

'save rtf main body

temppath2 = Left(temppath, Len(temppath) - 3) + "rtf"
rtb.LoadFile (temppath2)
Update
End If

End Sub


Private Sub mnuSave_Click() 'save rtf file with versalinks

'  vlt = versa link text files

On Error Resume Next
dlgfile.Filter = "Rich Text Files with Links|*.vlt;"
dlgfile.FileName = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0
 
temppath = dlgfile.FileName


  fnum1 = FreeFile
On Error Resume Next
Open temppath For Output As #fnum1 'write .vlf with link info
On Error Resume Next
For xx = 0 To 99
Write #fnum1, Links(xx).LinkText
Write #fnum1, Links(xx).LinkDest
Write #fnum1, Links(xx).LinkApp
Next xx
Write #fnum1, LnkCnt
Close fnum1

'write rtf main body (can be opened in any rtf editor)
temppath2 = Left(temppath, Len(temppath) - 3) + "rtf"
rtb.SaveFile (temppath2)

End Sub


Private Sub mnuNew_Click() 'start new file
 
On Error Resume Next
rtb.Text = ""
For xx = 0 To 99
Links(xx).LinkText = ""
Links(xx).LinkDest = ""
Links(xx).LinkApp = ""
Next xx
LnkCnt = 0
Update

End Sub


Private Sub mnuExit_Click() 'exit app

Unload Me

End Sub


Private Sub mnuPaste_Click() 'paste rtf from clipboard

     On Error Resume Next
     Form1.rtb.SelRTF = Clipboard.GetText
     
End Sub


Private Sub mnuCut_Click() 'cut rtf to clipboard

     On Error Resume Next
     Clipboard.SetText rtb.SelRTF
     Form1.rtb.SelText = vbNullString
     
End Sub


Private Sub mnuCopy_Click() 'copy rtf to clipboard

     On Error Resume Next
     Clipboard.SetText rtb.SelRTF
     
End Sub


Private Sub Update() 'refresh listbox of versalinks

List1.Clear
For xx = 0 To LnkCnt
If xx = 0 Then
List1.AddItem ("_____Versalinks_____"), xx
Else
List1.AddItem (CStr(xx) + "   " + Links(xx).LinkText), xx
End If
Next xx
Label14.Caption = CStr(LnkCnt)

End Sub


Private Sub mnuExittext_Click()

     On Error Resume Next
     rtb2.Visible = False
     mnuExittext.Visible = False
     
End Sub


Private Sub mnuExitimage_Click()

     On Error Resume Next
     Picture1.Visible = False
     mnuExitimage.Visible = False
     
End Sub

