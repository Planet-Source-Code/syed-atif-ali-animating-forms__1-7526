VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| | | W I N D O W    A N I M A T I O N | | |"
   ClientHeight    =   3480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5430
   FillColor       =   &H008080FF&
   FillStyle       =   7  'Diagonal Cross
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   15  'Size All
   ScaleHeight     =   3480
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "MOVING THIS WINDOW..."
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   480
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
      Begin VB.OptionButton Option3 
         Caption         =   "Make this window move horizontally as well as vertically when moved"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   4095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Make this window bounce vertically when moved"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Make this window bounce horizontally when moved"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   4095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   2520
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "http://www.planet-source-code.com"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   2640
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "icq: 60278335"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   1440
      MousePointer    =   10  'Up Arrow
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "hutell@hotmail.com"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   0
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5564
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   $"Form1.frx":030A
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   0
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   120
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   5564
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|--------------------------------------|
'|--------------------------------------|
'|--------------------------------------|
'HUTELL (HUman inTELLigence) Productions'
'|--------------------------------------|
'|------------HELLO EVERYONE!-----------|
'|--------------------------------------|
'|--OKAY,THIS CODE IS SPECIFICALLY FOR--|
'|----ANIMATING OR WHAT I WOULD CALL----|
'|-----------"FORM BOUNCING".-----------|
'|--------------------------------------|
'|--BY SYED ATIF ALI (owner of HUTELL)--|
'|----EMAIL ME AT hutell@hotmail.com----|
'|--------------------------------------|
'HUTELL (HUman inTELLigence) Productions'
'|--------------------------------------|
'|--------------------------------------|
'|--------------------------------------|

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'NOTES: Okay, the main thing is within the timer, so you
'may want to skip the following at go to the timer code.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'WARNING: I, Syed Atif Ali, should and could not be held
'responsible or liable for any kind of damages or anything
'bad arising from the use of this. That means, use it at
'your own risk.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'DISCLAIMER: You could, without permission, include this
'code in your non-commercial applications given that you
'mention me and my company somewhere in the credits
'section of your non-commercial software. However, you
'DO need to obtain written permission of mine before you
'include this code within your commercial software. YOU
'COULD NOT(!), however, reproduce, publicise, or throw
'this code (into a web-site) without the given permission
'of Syed Atif Ali.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'COPYRIGHT: This source code is a copyright of Syed Atif
'Ali of HUTELL© Productions.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

Option Explicit


Private Sub Label2_Click(index As Integer)
Select Case index
    Case 0
        Dim ComeToMe
        ComeToMe = Shell("start.exe mailto:hutell@hotmail.com", vbNormalFocus)
        '^this is a very nice and easy way I discovered while...
        'playing around how to open default email programs to...
        'send mail. If you want to know how to open the default...
        'browser to surf a site with, replace from...
        '"mailto" to "com" and put in a web site's address...
        'example:
        '"start.exe http://www.planet-source-code.com"^
    Case 1
        Dim GoToCoolSite
        GoToCoolSite = Shell("start.exe http://www.planet-source-code.com", vbNormalFocus)
    Case 2
        Clipboard.Clear
        '^clear any old clipboard data.^
        Clipboard.SetText "60278335"
        '^set my ICQ number to be pasted anywhere.^
        MsgBox "The ICQ number has been copied to the clipboard. You may open your ICQ, click on Add Users, paste this number and find for this number for inquiries or chatting." & vbNewLine & vbNewLine & "Thank You.", vbOKOnly + vbInformation, "Number Copied"
End Select
End Sub

Private Sub Label2_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(index).ForeColor = vbRed
Label2(index).BorderStyle = 1
End Sub

Private Sub Label2_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(index).ForeColor = &HFFC0C0
Label2(index).BorderStyle = 0
End Sub

Private Sub mnuAbout_Click()
MsgBox "Made by Hutell Productions..." + vbNewLine + "Hutell Productions is owned and run by the 'one-kid force' by Syed Atif Ali." + vbNewLine + vbNewLine + "You can reach him at hutell@hotmail.com" + vbNewLine + vbNewLine + vbNewLine + "---Hutell Productions---" + vbNewLine + "HUman inTELLigence" + vbNewLine + "---------we are technology...", vbOKOnly + vbInformation, "About Window Animation..."
End Sub

Private Sub Timer1_Timer()
'------------------------------------------------------------------
'---OKAY, THIS IS THE MAIN PART------------------------------------
'------------------------------------------------------------------

If Form1.WindowState = 0 Then
'^The form can't be moved in Max. or Min. mode^
    If Option1.Value = True Then
    '^If the horizontal bouncing is enabled then...^
    Form1.Left = (Screen.Height - Form1.Left) / 2
    '^try to visualize this because it can't be explained...
    'you can also draw it on a paper to help you
    'understand what it actually is. +_+
    ElseIf Option2.Value = True Then
    '^if the vertical bouncing is enabled then...^
    Form1.Top = (Screen.Width - Form1.Top) / 2
    '^try to visualize om your own.^
    ElseIf Option3.Value = True Then
    '^if the horizontal and vertical bouncing are enabled then...^
    Form1.Left = (Screen.Height - Form1.Left) / 2
    Form1.Top = (Screen.Width - Form1.Top) / 2
    '^same as above. Try to visualize.^
    End If
End If

'------------------------------------------------------------------
'------------------------------------------------------------------
'------------------------------------------------------------------
End Sub

