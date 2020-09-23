VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8655
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTFText 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      _Version        =   393217
      BackColor       =   12648447
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "D:\My code\VB6\Sleep\sleep.rtf"
      TextRTF         =   $"frmAbout.frx":08CA
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

