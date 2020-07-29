VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Perbedaan Fungsi Date dan Date$"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  MsgBox DateTime.Date '--> Menghasilkan tanggal hari
  'ini, sesuai dengan setting format tanggal di
  'komputer 'Anda.
  'Contoh: Jika tgl hari ini = 22 Januari 2002 dan
  'format Short Date Style di Regional Setting =
  '"dd/mm/yyyy", akan menghasilkan: 22/01/2002

  MsgBox DateTime.Date$ '--> Menghasilkan tanggal hari
  'ini dengan format tanggal Standar Internasional,
  'yaitu: "mm-dd-yyyy"
  'Contoh: (sama dengan di atas), maka akan
  'menghasilkan: 01/22/2002
End Sub

