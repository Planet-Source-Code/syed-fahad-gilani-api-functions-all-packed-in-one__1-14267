VERSION 5.00
Begin VB.Form Startup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3240
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Call MagicBox.Show
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call MagicBox.Show
Unload Me
End Sub
