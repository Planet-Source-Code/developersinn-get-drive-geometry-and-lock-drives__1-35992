VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Direct Disk Access Examples"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnLock 
      Caption         =   "Unlock Drive"
      Height          =   495
      Left            =   2235
      TabIndex        =   5
      Top             =   4905
      Width           =   1095
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock Drive"
      Height          =   495
      Left            =   2235
      TabIndex        =   4
      Top             =   4425
      Width           =   1095
   End
   Begin VB.ListBox lstDrives 
      Height          =   1035
      Left            =   585
      TabIndex        =   2
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox txtDiskG 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1290
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Created by Sameers, theAngrycodeR@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   945
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5625
      Width           =   4515
   End
   Begin VB.Label Label5 
      Caption         =   "NOTE: This Programme will work under Win NT, 2K or WinXP only"
      Height          =   330
      Left            =   585
      TabIndex        =   8
      Top             =   630
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Drive Geometry Reader and Locker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   315
      TabIndex        =   7
      Top             =   135
      Width           =   5550
   End
   Begin VB.Label Label3 
      Caption         =   "As much time you will lock your drive, you must have to unlock that for the same no. of time"
      Height          =   690
      Left            =   3375
      TabIndex        =   6
      Top             =   4500
      Width           =   3075
   End
   Begin VB.Label Label2 
      Caption         =   "Removeable Drives"
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
      Left            =   600
      TabIndex        =   3
      Top             =   4140
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Disk Geometries"
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
      TabIndex        =   1
      Top             =   1050
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLock_Click()
    If lstDrives.ListIndex <> -1 Then
        mdlAPIs.LockDrive lstDrives.Text, True
    End If
End Sub

Private Sub cmdUnLock_Click()
    If lstDrives.ListIndex <> -1 Then
        mdlAPIs.LockDrive lstDrives.Text, False
    End If
End Sub

Private Sub Form_Load()
    mdlAPIs.GetDisksAndProfiles
    mdlAPIs.GetRemoveableDrives
End Sub

