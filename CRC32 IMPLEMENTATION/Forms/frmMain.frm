VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CRC 32 Algorithm Test"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblChecksum 
      Caption         =   "Checksum: 0x00000000"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lblData 
      Caption         =   "Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngInitializeValue As Long

Private Sub Form_Load()
    lngInitializeValue = InitializeCRC32Table
End Sub

Private Sub txtData_Change()
    Dim lngCRC32 As Long
    
    lngCRC32 = GenerateCRC32(txtData.Text, lngInitializeValue)
    lblChecksum.Caption = "Checksum: 0x" & Hex$(lngCRC32)
End Sub
