VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "msblast.worm - [Removal]"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "eXIT"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdDetect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detect"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unreal © 2003"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1980
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10:52PM 10/14/2003
'Written by Kyriakos Nicola
'eMAIL: kyriakosnicola@yahoo.com

'This is not a professional removal so please do not
'criticise my work or even compare it with other removals.

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub cmdDetect_Click()

On Error Resume Next

Dim fso, WScript, ReadKey, DirSystem, msBlastRegString, msBlastDefaultPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set DirSystem = fso.getspecialfolder(1)
Set WScript = CreateObject("WScript.Shell")

msBlastRegString = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\windows auto update"
msBlastDefaultPath = DirSystem & "\msblast.exe"
ReadKey = WScript.RegRead(msBlastRegString)

If ReadKey = msBlastRegString And (fso.FileExists(msBlastDefaultPath)) Then
    MsgBox "There is 95% chance that you might be infected!", vbExclamation, "WARNING!"
    cmdRemove.Enabled = True
    Exit Sub
End If

If ReadKey = msBlastRegString Xor (fso.FileExists(msBlastDefaultPath)) Then
    MsgBox "There is 60% chance that you might be infected!", vbExclamation, "WARNING!"
    MsgBox "Click 'Remove' button just to be sure.", vbInformation, "Tip..."
    cmdRemove.Enabled = True
    Exit Sub
End If

MsgBox "msblast worm was NOT detected!", vbInformation, "Area Clear!"

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdRemove_Click()

On Error GoTo Err

Dim WScript, fso, DirSystem, Win&

Set fso = CreateObject("Scripting.FileSystemObject")
Set DirSystem = fso.getspecialfolder(1)
Set WScript = CreateObject("WScript.Shell")

WScript.RegDelete ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\windows auto update")

'i'm not sure if the caption is correct so don't
'count on that. it may not work. (possible name "msblast")
'anyway the worm must be shutdown else it won't be deleted.
'ctrl+alt+del --> processes (tab) --> msblast.exe endtask this entry.
Win = FindWindow("ThunderRT6FormDC", "windows auto update")
SendMessage Win, &H10, 0, 0                  '^--- caption goes here.

Kill (DirSystem & "\msblast.exe")

MsgBox "msblast worm removed!", vbInformation, "Removed.."
cmdRemove.Enabled = False

Err:
MsgBox "Either the virus was not found or there was a critical error!", vbExclamation, "Error!"

End Sub

