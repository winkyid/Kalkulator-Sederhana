VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14070
   Icon            =   "WInkyKurniatama_Kalkulator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleMode       =   0  'User
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   13680
      Top             =   6600
   End
   Begin VB.PictureBox header_box 
      BackColor       =   &H00000020&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14055
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.Label txt_close 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13680
         TabIndex        =   29
         Top             =   -10
         Width           =   255
      End
      Begin VB.Label txt_minimaze 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13260
         TabIndex        =   28
         Top             =   -30
         Width           =   285
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   300
         Left            =   13260
         Shape           =   2  'Oval
         Top             =   50
         Width           =   300
      End
      Begin VB.Shape btn_close 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   300
         Left            =   13680
         Shape           =   2  'Oval
         Top             =   50
         Width           =   300
      End
      Begin VB.Label txt_versionApps 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0 (Winky Kurniatama)"
         BeginProperty Font 
            Name            =   "Montserrat Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   45
         Width           =   3255
      End
      Begin VB.Label txt_titleApps 
         BackStyle       =   0  'Transparent
         Caption         =   "Kalkulator Basic Simple"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   40
         Width           =   2535
      End
      Begin VB.Label txt_logoApps 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Heydings Icons"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   30
         Width           =   255
      End
   End
   Begin VB.Label txt_closeinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x Tutup"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6480
      TabIndex        =   39
      Top             =   -9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape bg_closeinfo 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   -9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label txt_web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web : WinkyID.github.io"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   4440
      TabIndex        =   38
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_fakultas 
      BackStyle       =   0  'Transparent
      Caption         =   "Fakultas : Bisnis && Informatika"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   37
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_semester 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester : 4 (Genap)"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   4440
      TabIndex        =   36
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_prodi 
      BackStyle       =   0  'Transparent
      Caption         =   "Prodi : Sistem Informasi"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   35
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_matkul 
      BackStyle       =   0  'Transparent
      Caption         =   "Matkul : Pemrogaman Beorientasi Objek"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   4440
      TabIndex        =   34
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_nim 
      BackStyle       =   0  'Transparent
      Caption         =   "Nim : 23.54.027135"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   33
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label txt_nama 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama : Winky Kurniatama "
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   4440
      TabIndex        =   32
      Top             =   -9000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label judul_aboutme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Me :)"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Shape bg_aboutme 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5175
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   -9000
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Image bg_infoblur 
      Height          =   6720
      Left            =   0
      Picture         =   "WInkyKurniatama_Kalkulator.frx":10CA
      Stretch         =   -1  'True
      Top             =   -9000
      Visible         =   0   'False
      Width           =   13995
   End
   Begin VB.Label info_hitungan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   13200
      TabIndex        =   30
      Top             =   2400
      Width           =   615
   End
   Begin VB.Shape bg_eksekusi 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   13200
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kalkulator Basic Simple"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Heydings Icons"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   255
   End
   Begin VB.Label txt_about1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Heydings Icons"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   22
      Top             =   675
      Width           =   615
   End
   Begin VB.Label txt_hapus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "< BackSpace"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   11160
      TabIndex        =   21
      Top             =   3045
      Width           =   2535
   End
   Begin VB.Label txt_samadengan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   20
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label txt_bagi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   5700
      TabIndex        =   19
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label txt_kali 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   4035
      TabIndex        =   18
      Top             =   4185
      Width           =   1215
   End
   Begin VB.Label txt_kurang 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   5580
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label txt_tambah 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   4035
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label txt_clearAll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1440
      TabIndex        =   15
      Top             =   5980
      Width           =   1815
   End
   Begin VB.Label txt_0 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label txt_3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   13
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label txt_2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   12
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label txt_1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label txt_6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label txt_5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label txt_4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label txt_9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label txt_8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label txt_7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape btn_backspace 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Shape btn_samadengan 
      BackColor       =   &H00008214&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Shape btn_bagi 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Shape btn_kali 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Shape btn_kurang 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape btn_tambah 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape btn_clearAll 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   2040
   End
   Begin VB.Shape btn_0 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   960
   End
   Begin VB.Shape btn_3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   960
   End
   Begin VB.Shape btn_2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   960
   End
   Begin VB.Shape btn_1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   960
   End
   Begin VB.Shape btn_6 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   960
   End
   Begin VB.Shape btn_5 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   960
   End
   Begin VB.Shape btn_4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   960
   End
   Begin VB.Shape btn_9 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   960
   End
   Begin VB.Shape btn_8 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   960
   End
   Begin VB.Shape btn_7 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   960
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label txt_display 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Montserrat Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   13335
   End
   Begin VB.Shape bg_display 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   13575
   End
   Begin VB.Label txt_setting 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Heydings Icons"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   930
      TabIndex        =   3
      Top             =   670
      Width           =   255
   End
   Begin VB.Shape btn_settings 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   415
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   415
   End
   Begin VB.Label txt_menu 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Byom Icons"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   670
      Width           =   255
   End
   Begin VB.Shape btn_menu 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   420
   End
   Begin VB.Label txt_about2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Me"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12360
      TabIndex        =   1
      Top             =   645
      Width           =   1455
   End
   Begin VB.Shape btn_about 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama       : Winky Kurniatam
'Nim        : 23.54.027135
'Matkul     : Pemrogaman Berorientasi Objek
'Prodi      : Sistem Informasi
'Semester   : 4 (genap)

' deklarasi varible global ==============================================================================
Dim nilai1 As Double
Dim nilai2 As Double
Dim operator_aktif As String
Dim baruInput As Boolean
Dim sudahOperator As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'kode form loader =====================================================================================

Private Sub Form_Load()
txt_display.Caption = "0"
baruInput = True
sudahOperator = False
End Sub

'header apps customize ==================================================================================

Private Sub header_box_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub txt_versionApps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub txt_titleApps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub txt_logoApps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub txt_minimaze_Click()
    Me.WindowState = vbMinimized
End Sub
Private Sub txt_close_Click()
    Unload Me
End Sub

'kode operator tombol-tombol 0 - 9 =========================================================================
Private Sub angka_ditekan(teks As String)
    txt_display.ForeColor = RGB(0, 0, 0)
    If txt_display.Caption = "0" Or baruInput = True Then
        txt_display.Caption = teks
    Else
        txt_display.Caption = txt_display.Caption & teks
    End If
    baruInput = False
End Sub

Private Sub txt_0_Click()
    angka_ditekan "0"
    btn_0.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_1_Click()
    angka_ditekan "1"
    btn_1.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_2_Click()
    angka_ditekan "2"
    btn_2.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_3_Click()
    angka_ditekan "3"
    btn_3.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_4_Click()
    angka_ditekan "4"
    btn_4.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_5_Click()
    angka_ditekan "5"
    btn_5.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_6_Click()
    angka_ditekan "6"
    btn_6.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_7_Click()
    angka_ditekan "7"
    btn_7.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_8_Click()
    angka_ditekan "8"
    btn_8.BackColor = RGB(255, 170, 170)
End Sub

Private Sub txt_9_Click()
    angka_ditekan "9"
    btn_9.BackColor = RGB(255, 170, 170)
End Sub

'kode operator penghitungan (tambah,kurang, kali bagi)===============================================
Private Sub operator_ditekan(op As String)
    If sudahOperator = True Then Exit Sub

    nilai1 = Val(txt_display.Caption)
    operator_aktif = op
    baruInput = True
    sudahOperator = True

    Select Case op
        Case "+"
            info_hitungan.Caption = "+"
            txt_display.ForeColor = RGB(215, 215, 215)
        Case "-"
            info_hitungan.Caption = "-"
            txt_display.ForeColor = RGB(215, 215, 215)
        Case "*"
            info_hitungan.Caption = "x"
            txt_display.ForeColor = RGB(215, 215, 215)
        Case "/"
            info_hitungan.Caption = ":"
            txt_display.ForeColor = RGB(215, 215, 215)
    End Select
End Sub


Private Sub txt_tambah_Click()
    operator_ditekan "+"
End Sub

Private Sub txt_kurang_Click()
    operator_ditekan "-"
End Sub

Private Sub txt_kali_Click()
    operator_ditekan "*"
End Sub

Private Sub txt_bagi_Click()
    operator_ditekan "/"
End Sub

'kode untuk fungsi sama dengan ==================================================================
Private Sub txt_samadengan_Click()
    btn_samadengan.BackColor = RGB(39, 255, 0)
    
    nilai2 = Val(txt_display.Caption)
    Select Case operator_aktif
        Case "+"
            txt_display.Caption = nilai1 + nilai2
        Case "-"
            txt_display.Caption = nilai1 - nilai2
        Case "*"
            txt_display.Caption = nilai1 * nilai2
        Case "/"
            If nilai2 <> 0 Then
                txt_display.Caption = nilai1 / nilai2
            Else
                txt_display.Caption = "Error! tidak bisa dibagi 0"
            End If
    End Select
    baruInput = True
    info_hitungan.Caption = "="
    sudahOperator = False
End Sub

'kode untuk menghapus atau reset kalkulator ke 0 ==================================================
Private Sub txt_clearAll_Click()
    btn_clearAll.BackColor = RGB(215, 0, 0)

    txt_display.Caption = "0"
    nilai1 = 0
    nilai2 = 0
    operator_aktif = ""
    info_hitungan.Caption = "..."
    baruInput = True
End Sub

'kode untuk menghapus angka satu-persatu ==========================================================
Private Sub txt_hapus_Click()
    btn_backspace.BackColor = RGB(255, 100, 100)

    If Len(txt_display.Caption) > 1 Then
        txt_display.Caption = Left(txt_display.Caption, Len(txt_display.Caption) - 1)
    Else
        txt_display.Caption = "0"
    End If
End Sub

' kode untuk eksekusi reset default color agar ada animasi klik nya
Private Sub Timer1_Timer()
    btn_clearAll.BackColor = &H40&
    btn_samadengan.BackColor = &H8214&
    btn_0.BackColor = &HE0E0E0
    btn_1.BackColor = &HE0E0E0
    btn_2.BackColor = &HE0E0E0
    btn_3.BackColor = &HE0E0E0
    btn_4.BackColor = &HE0E0E0
    btn_5.BackColor = &HE0E0E0
    btn_6.BackColor = &HE0E0E0
    btn_7.BackColor = &HE0E0E0
    btn_8.BackColor = &HE0E0E0
    btn_9.BackColor = &HE0E0E0
    btn_0.BackColor = &HE0E0E0
    btn_backspace.BackColor = &HE0E0E0
End Sub

' tempat hide and show informasi diri
Private Sub ShowAboutMe()
    bg_infoblur.Top = 360
    bg_aboutme.Top = 960
    txt_nama.Top = 1680
    txt_nim.Top = 2040
    txt_matkul.Top = 2400
    txt_prodi.Top = 3180
    txt_semester.Top = 3580
    txt_fakultas.Top = 3960
    txt_web.Top = 4320
    txt_closeinfo.Top = 5280
    bg_closeinfo.Top = 5280

    bg_infoblur.Visible = True
    bg_aboutme.Visible = True
    txt_nama.Visible = True
    txt_nim.Visible = True
    txt_matkul.Visible = True
    txt_prodi.Visible = True
    txt_semester.Visible = True
    txt_fakultas.Visible = True
    txt_web.Visible = True
    txt_closeinfo.Visible = True
    bg_closeinfo.Visible = True
End Sub

Private Sub HideAboutMe()
    Dim hiddenTop As Integer
    hiddenTop = -9000

    bg_infoblur.Top = hiddenTop
    bg_aboutme.Top = hiddenTop
    txt_nama.Top = hiddenTop
    txt_nim.Top = hiddenTop
    txt_matkul.Top = hiddenTop
    txt_prodi.Top = hiddenTop
    txt_semester.Top = hiddenTop
    txt_fakultas.Top = hiddenTop
    txt_web.Top = hiddenTop
    txt_closeinfo.Top = hiddenTop
    bg_closeinfo.Top = hiddenTop

    ' Sembunyikan semua
    bg_infoblur.Visible = False
    bg_aboutme.Visible = False
    txt_nama.Visible = False
    txt_nim.Visible = False
    txt_matkul.Visible = False
    txt_prodi.Visible = False
    txt_semester.Visible = False
    txt_fakultas.Visible = False
    txt_web.Visible = False
    txt_closeinfo.Visible = False
    bg_closeinfo.Visible = False
End Sub

Private Sub txt_about1_Click()
    Call ShowAboutMe
End Sub

Private Sub txt_about2_Click()
    Call ShowAboutMe
End Sub

Private Sub txt_closeinfo_Click()
    Call HideAboutMe
End Sub

