VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "多点温度监控系统上位机——By WillianChan"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   16920
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   7.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   28.862
   ScaleMode       =   0  'User
   ScaleWidth      =   140.875
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12240
      Top             =   9360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12720
      Top             =   9360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13200
      Top             =   9360
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13680
      Top             =   9360
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   14160
      Top             =   9360
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "总控制台"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8775
      Left            =   14760
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin VB.Frame Frame16 
         BackColor       =   &H00000000&
         Caption         =   "窗体透明度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton PelluciditySub 
            BackColor       =   &H00808080&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton PellucidityAdd 
            BackColor       =   &H00808080&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblPellucidity 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "20%"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   480
            TabIndex        =   72
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00000000&
         Caption         =   "通信端口选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   1815
         Begin VB.CommandButton cmdOpenCOM 
            BackColor       =   &H00808080&
            Caption         =   "打开端口"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   960
            Width           =   1575
         End
         Begin VB.ComboBox cmbCOM 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            ItemData        =   "frmMain.frx":08CA
            Left            =   120
            List            =   "frmMain.frx":08CC
            TabIndex        =   55
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "端口未打开"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   68
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00000000&
         Caption         =   "曲线刷新时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   50
         Top             =   3720
         Width           =   1815
         Begin VB.CommandButton cmdChangeTime 
            BackColor       =   &H00808080&
            Caption         =   "更改时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox DelayValueText 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   51
            Text            =   "1200"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "ms"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1320
            TabIndex        =   52
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00808080&
         Caption         =   "全部清屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6480
         Width           =   1815
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H00808080&
         Caption         =   "退出程序"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7560
         Width           =   1815
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00808080&
         Caption         =   "全部开始测温"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5400
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "D测温点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   7440
      TabIndex        =   6
      Top             =   4560
      Width           =   7215
      Begin VB.CommandButton cmdClear4 
         BackColor       =   &H00808080&
         Caption         =   "清屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart4 
         BackColor       =   &H00808080&
         Caption         =   "开始测温"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00000000&
         Caption         =   "当前温度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   5520
         TabIndex        =   40
         Top             =   360
         Width           =   1455
         Begin VB.Label lblValue4 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "+00.00℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00000000&
         Caption         =   "Y轴显示范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   5520
         TabIndex        =   35
         Top             =   1320
         Width           =   1455
         Begin VB.TextBox YMinText4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox YMaxText4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            TabIndex        =   37
            Text            =   "60"
            Top             =   375
            Width           =   495
         End
         Begin VB.CommandButton cmdChange4 
            BackColor       =   &H00808080&
            Caption         =   "更改范围"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.PictureBox picVoltage4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   600
         ScaleHeight     =   3615
         ScaleWidth      =   4710
         TabIndex        =   34
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label YMidLabel4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label YMinLabel4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label YMaxLabel4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "B测温点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdStart2 
         BackColor       =   &H00808080&
         Caption         =   "开始测温"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear2 
         BackColor       =   &H00808080&
         Caption         =   "清屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00000000&
         Caption         =   "Y轴显示范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   5520
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
         Begin VB.CommandButton cmdChange2 
            BackColor       =   &H00808080&
            Caption         =   "更改范围"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox YMaxText2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   840
            TabIndex        =   31
            Text            =   "60"
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox YMinText2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Text            =   "0"
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   33
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00000000&
         Caption         =   "当前温度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   5520
         TabIndex        =   27
         Top             =   360
         Width           =   1455
         Begin VB.Label lblValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "+00.00℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.PictureBox picVoltage2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   600
         ScaleHeight     =   3615
         ScaleWidth      =   4710
         TabIndex        =   26
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label YMidLabel2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label YMinLabel2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label YMaxLabel2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "C测温点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   7215
      Begin VB.CommandButton cmdStart3 
         BackColor       =   &H00808080&
         Caption         =   "开始测温"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear3 
         BackColor       =   &H00808080&
         Caption         =   "清屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         Caption         =   "当前温度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   5520
         TabIndex        =   16
         Top             =   360
         Width           =   1455
         Begin VB.Label lblValue3 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "+00.00℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Caption         =   "Y轴显示范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   5520
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
         Begin VB.CommandButton cmdChange3 
            BackColor       =   &H00808080&
            Caption         =   "更改范围"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox YMaxText3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   840
            TabIndex        =   13
            Text            =   "60"
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox YMinText3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   12
            Text            =   "0"
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.PictureBox picVoltage3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   600
         ScaleHeight     =   3615
         ScaleWidth      =   4710
         TabIndex        =   4
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label YMidLabel3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label YMinLabel3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label YMaxLabel3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "A测温点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdClear1 
         BackColor       =   &H00808080&
         Caption         =   "清屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart1 
         BackColor       =   &H00808080&
         Caption         =   "开始测温"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
         Caption         =   "Y轴显示范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   5520
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
         Begin VB.TextBox YMinText1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Text            =   "0"
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox YMaxText1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   840
            TabIndex        =   23
            Text            =   "60"
            Top             =   375
            Width           =   495
         End
         Begin VB.CommandButton cmdChange1 
            BackColor       =   &H00808080&
            Caption         =   "更改范围"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Caption         =   "当前温度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   5520
         TabIndex        =   19
         Top             =   360
         Width           =   1455
         Begin VB.Label lblValue1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "+00.00℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.PictureBox picVoltage1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   600
         ScaleHeight     =   3615
         ScaleWidth      =   4710
         TabIndex        =   18
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label YMidLabel1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label YMinLabel1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label YMaxLabel1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   375
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14640
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "欢迎使用【多点无线温度采集系统上位机】！这里是提示栏，您可以在这里获得帮助提示。请先选择通信端口。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   16695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NowX1, NowX2, NowX3, NowX4 As Integer 'ABCD测温点现在的X轴位置
Dim MaxPlotNo As Long  '最长的X轴范围
Dim YMax1, YMax2, YMax3, YMax4 As Long '温度显示范围上限值
Dim YMin1, YMin2, YMin3, YMin4 As Long '温度显示范围下限值
Dim ValueStr1, ValueStr2, ValueStr3, ValueStr4 As Single 'ABCD测温点当前的测量值
Dim PreValue1, PreValue2, PreValue3, PreValue4 As Single 'ABCD测温点前一个测量值
Dim Buf As String  '存放串口通讯数据
Dim Flag As Integer  '用于控制总测温开始与结束的标志位
Dim DelayTimeValueMS As Integer  '延时时长值
Dim SaveFileTimeStar1, SaveFileTimeStar2, SaveFileTimeStar3, SaveFileTimeStar4 As String  '保存图片文件时，ABCD测温点温度曲线开始的时间
Dim SaveFileTimeEnd1, SaveFileTimeEnd2, SaveFileTimeEnd3, SaveFileTimeEnd4 As String  '保存图片文件时，ABCD测温点温度曲线结束的时间
Dim cmbCOMWidthSize As Single  '用于存放cmbCOM控件的宽度缩放比例
Dim cmbCOMLeftSize As Single  '用于存放cmbCOM控件的左边位置缩放比例
'窗口透明API
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'窗口透明常数
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2  '使用此参数，透明度有效，透明颜色无效
Const LWA_COLORKEY = &H1  '使用此参数，透明度无效，透明颜色有效
Dim Pellucidity As Integer  '透明度（完全透明0~不透明255）

'更改A测温点Y轴显示范围
Private Sub cmdChange1_Click()
    Dim I%
    If Val(YMinText1.Text) = Val(YMaxText1.Text) Then
        MsgBox "温度显示范围上下值不应相同，请重新设置。", vbCritical + vbOKOnly, "系统信息"
    Else
        YMin1 = Val(YMinText1.Text)
        YMax1 = Val(YMaxText1.Text)
        YMidLabel1.Caption = (Val(YMinText1.Text) + Val(YMaxText1.Text)) / 2
        If Val(YMinText1.Text) < Val(YMaxText1.Text) Then
            YMinLabel1.Caption = YMinText1.Text
            YMaxLabel1.Caption = YMaxText1.Text
            picVoltage1.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage1.Scale (0, YMax1)-(MaxPlotNo, YMin1)
            '设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage1.CurrentX = 0  '设定起点
            picVoltage1.CurrentY = 0  '设定起点
            picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage1.Line (0, (Abs(YMax1 - YMin1) / 10) * I + YMin1)-(MaxPlotNo, (Abs(YMax1 - YMin1) / 10) * I + YMin1), RGB(255, 255, 255)
            Next I
         Else
            YMinLabel1.Caption = YMaxText1.Text
            YMaxLabel1.Caption = YMinText1.Text
            'YMax1和YMin1交换植
            YMin1 = Val(YMaxText1.Text)
            YMax1 = Val(YMinText1.Text)
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage1.Cls '清除图形
            picVoltage1.Scale (0, YMin1)-(MaxPlotNo, YMax1)
            '设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage1.CurrentX = 0  '设定起点
            picVoltage1.CurrentY = 0  '设定起点
            picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage1.Line (0, (Abs(YMax1 - YMin1) / 10) * I + YMin1)-(MaxPlotNo, (Abs(YMax1 - YMin1) / 10) * I + YMin1), RGB(255, 255, 255)
            Next I
         End If
        NowX1 = 0
        '画曲线需要使用两个像素宽度的画笔
        picVoltage1.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage1.PSet (0, 0)  '设定起点
    End If
End Sub

'更改B测温点Y轴显示范围
Private Sub cmdChange2_Click()
    Dim I%
    If Val(YMinText2.Text) = Val(YMaxText2.Text) Then
        MsgBox "温度显示范围上下值不应相同，请重新设置。", vbCritical + vbOKOnly, "系统信息"
    Else
        YMin2 = Val(YMinText2.Text)
        YMax2 = Val(YMaxText2.Text)
        YMidLabel2.Caption = (Val(YMinText2.Text) + Val(YMaxText2.Text)) / 2
        If Val(YMinText2.Text) < Val(YMaxText2.Text) Then
            YMinLabel2.Caption = YMinText2.Text
            YMaxLabel2.Caption = YMaxText2.Text
            picVoltage2.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage2.Scale (0, YMax2)-(MaxPlotNo, YMin2)
            picVoltage2.Cls '清除图形
            '设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage2.CurrentX = 0  '设定起点
            picVoltage2.CurrentY = 0  '设定起点
            picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage2.Line (0, (Abs(YMax2 - YMin2) / 10) * I + YMin2)-(MaxPlotNo, (Abs(YMax2 - YMin2) / 10) * I + YMin2), RGB(255, 255, 255)
            Next I
        Else
            YMinLabel2.Caption = YMaxText2.Text
            YMaxLabel2.Caption = YMinText2.Text
            'YMax2和YMin2交换植
            YMin2 = Val(YMaxText2.Text)
            YMax2 = Val(YMinText2.Text)
            picVoltage2.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage2.Scale (0, YMax2)-(MaxPlotNo, YMin2)
            '设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage2.CurrentX = 0  '设定起点
            picVoltage2.CurrentY = 0  '设定起点
            picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage2.Line (0, (Abs(YMax2 - YMin2) / 10) * I + YMin2)-(MaxPlotNo, (Abs(YMax2 - YMin2) / 10) * I + YMin2), RGB(255, 255, 255)
            Next I
        End If
        NowX2 = 0
        '画曲线需要使用两个像素宽度的画笔
        picVoltage2.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage2.PSet (0, 0)  '设定起点
    End If
End Sub

'更改C测温点Y轴显示范围
Private Sub cmdChange3_Click()
    Dim I%
    If Val(YMinText3.Text) = Val(YMaxText3.Text) Then
        MsgBox "温度显示范围上下值不应相同，请重新设置。", vbCritical + vbOKOnly, "系统信息"
    Else
        YMin3 = Val(YMinText3.Text)
        YMax3 = Val(YMaxText3.Text)
        YMidLabel3.Caption = (Val(YMinText3.Text) + Val(YMaxText3.Text)) / 2
        If Val(YMinText3.Text) < Val(YMaxText3.Text) Then
            YMinLabel3.Caption = YMinText3.Text
            YMaxLabel3.Caption = YMaxText3.Text
            picVoltage3.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage3.Scale (0, YMax3)-(MaxPlotNo, YMin3)
            '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage3.CurrentX = 0  '设定起点
            picVoltage3.CurrentY = 0  '设定起点
            picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage3.Line (0, (Abs(YMax3 - YMin3) / 10) * I + YMin3)-(MaxPlotNo, (Abs(YMax3 - YMin3) / 10) * I + YMin3), RGB(255, 255, 255)
            Next I
        Else
            YMinLabel3.Caption = YMaxText3.Text
            YMaxLabel3.Caption = YMinText3.Text
            'YMax1和YMin1交换植
            YMin1 = Val(YMaxText1.Text)
            YMax1 = Val(YMinText1.Text)
            picVoltage3.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage3.Scale (0, YMin3)-(MaxPlotNo, YMax3)
            '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage3.CurrentX = 0  '设定起点
            picVoltage3.CurrentY = 0  '设定起点
            picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage3.Line (0, (Abs(YMax3 - YMin3) / 10) * I + YMin3)-(MaxPlotNo, (Abs(YMax3 - YMin3) / 10) * I + YMin3), RGB(255, 255, 255)
            Next I
        End If
        NowX3 = 0
        '画曲线需要使用两个像素宽度的画笔
        picVoltage3.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage3.PSet (0, 0)  '设定起点
    End If
End Sub

'更改D测温点Y轴显示范围
Private Sub cmdChange4_Click()
    Dim I%, YMaxYMinChange%
    If Val(YMinText4.Text) = Val(YMaxText4.Text) Then
        MsgBox "温度显示范围上下值不应相同，请重新设置。", vbCritical + vbOKOnly, "系统信息"
    Else
        YMin4 = Val(YMinText4.Text)
        YMax4 = Val(YMaxText4.Text)
        YMidLabel4.Caption = (Val(YMinText4.Text) + Val(YMaxText4.Text)) / 2
        If Val(YMinText4.Text) < Val(YMaxText4.Text) Then
            YMinLabel4.Caption = YMinText4.Text
            YMaxLabel4.Caption = YMaxText4.Text
            picVoltage4.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage4.Scale (0, YMax4)-(MaxPlotNo, YMin4)
            '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage4.CurrentX = 0  '设定起点
            picVoltage4.CurrentY = 0  '设定起点
            picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage4.Line (0, (Abs(YMax4 - YMin4) / 10) * I + YMin4)-(MaxPlotNo, (Abs(YMax4 - YMin4) / 10) * I + YMin4), RGB(255, 255, 255)
            Next I
        Else
            YMinLabel4.Caption = YMaxText4.Text
            YMaxLabel4.Caption = YMinText4.Text
            'YMax4和YMin4交换植
            YMin4 = Val(YMaxText4.Text)
            YMax4 = Val(YMinText4.Text)
            picVoltage4.Cls '清除图形
            '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
            picVoltage4.Scale (0, YMin4)-(MaxPlotNo, YMax4)
            '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
            picVoltage4.CurrentX = 0  '设定起点
            picVoltage4.CurrentY = 0  '设定起点
            picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
            '画坐标网络的横线
            For I = 1 To 9 Step 1
                picVoltage4.Line (0, (Abs(YMax4 - YMin4) / 10) * I + YMin4)-(MaxPlotNo, (Abs(YMax4 - YMin4) / 10) * I + YMin4), RGB(255, 255, 255)
            Next I
        End If
        NowX4 = 0
        '画曲线需要使用两个像素宽度的画笔
        picVoltage4.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage4.PSet (0, 0)  '设定起点
    End If
End Sub

'更改曲线刷新时间
Private Sub cmdChangeTime_Click()
    DelayTimeValueMS = DelayValueText.Text
    lblMsg.Caption = "提示：曲线刷新时间已更改为" & DelayValueText.Text & "毫秒"
End Sub

'使用End命令将系统结束
Private Sub cmdEnd_Click()
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False          '关闭通信端口
    End If
    End  '退出程序
End Sub

'将MSComm控件的参数设置好，并打开
Private Sub cmdOpenCOM_Click()
    If MSComm1.PortOpen = False Then
        '判断端口号码是否落在1--16之间
        If cmbCOM.ListIndex >= 0 And cmbCOM.ListIndex <= 16 Then
            MSComm1.CommPort = cmbCOM.ListIndex + 1
        Else
            MsgBox "指定通信端口时发生错误！", vbCritical + vbOKOnly, "系统信息"
            Exit Sub
        End If
        '激活错误检测机制
        On Error GoTo comErr
        MSComm1.Settings = "9600,n,8,1"  '设定通信参数
        MSComm1.PortOpen = True          '打开通信端口
        'cmdOpenCOM.Enabled = False       '将此按钮设为禁用状态
        cmdStart.Enabled = True          '激活【全部开始检测】按钮
        cmdStart1.Enabled = True          '激活【开始检测】按钮
        cmdStart2.Enabled = True          '激活【开始检测】按钮
        cmdStart3.Enabled = True          '激活【开始检测】按钮
        cmdStart4.Enabled = True          '激活【开始检测】按钮
        lblMsg.Caption = "提示：已打开通信端口，可单击【全部开始测温】按钮，执行测温的工作。"
        cmdOpenCOM.Caption = "关闭端口"
        Shape1.BackColor = &HFF00&
        If Val(cmbCOM.ListIndex + 1) < 10 Then
            Label7.Caption = "COM" & Val(cmbCOM.ListIndex + 1) & "已打开"
        Else
            Label7.Caption = "COM" & Val(cmbCOM.ListIndex + 1) & "已开"
        End If
        Exit Sub
comErr:
        MsgBox "打开通信端口时发生错误！请确定通信端口存在且正常。", vbCritical + vbOKOnly, "系统信息"
    Else
        MSComm1.PortOpen = False          '关闭通信端口
        cmdStart.Enabled = False          '禁用【全部开始检测】按钮
        cmdStart1.Enabled = False          '禁用【开始检测】按钮
        cmdStart2.Enabled = False          '禁用【开始检测】按钮
        cmdStart3.Enabled = False          '禁用【开始检测】按钮
        cmdStart4.Enabled = False          '禁用【开始检测】按钮
        Timer1.Enabled = False                      '关闭定时器
        Timer2.Enabled = False                      '关闭定时器
        Timer3.Enabled = False                      '关闭定时器
        Timer4.Enabled = False                      '关闭定时器
        cmdStart.Caption = "全部开始测温"
        cmdStart1.Caption = "开始测温"
        cmdStart2.Caption = "开始测温"
        cmdStart3.Caption = "开始测温"
        cmdStart4.Caption = "开始测温"
        Flag = 0  '标志位停止
        lblMsg.Caption = "提示：已关闭通讯端口并停止测温"
        cmdOpenCOM.Caption = "打开端口"
        Shape1.BackColor = &HC0C0C0
    Label7.Caption = "端口已关闭"
    End If
End Sub

'全部清屏按纽
Private Sub cmdClear_Click()
    NowX1 = 0
    picVoltage1.Cls '清除图形
    NowX2 = 0
    picVoltage2.Cls '清除图形
    NowX3 = 0
    picVoltage3.Cls '清除图形
    NowX4 = 0
    picVoltage4.Cls '清除图形
    
    '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
    Dim I%
    picVoltage1.CurrentX = 0  '设定起点
    picVoltage1.CurrentY = 0  '设定起点
    picVoltage2.CurrentX = 0  '设定起点
    picVoltage2.CurrentY = 0  '设定起点
    picVoltage3.CurrentX = 0  '设定起点
    picVoltage3.CurrentY = 0  '设定起点
    picVoltage4.CurrentX = 0  '设定起点
    picVoltage4.CurrentY = 0  '设定起点
    picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage1.Line (0, ((YMax1 - YMin1) / 10) * I)-(MaxPlotNo, ((YMax1 - YMin1) / 10) * I), RGB(255, 255, 255)
        picVoltage2.Line (0, ((YMax2 - YMin2) / 10) * I)-(MaxPlotNo, ((YMax2 - YMin2) / 10) * I), RGB(255, 255, 255)
        picVoltage3.Line (0, ((YMax3 - YMin3) / 10) * I)-(MaxPlotNo, ((YMax3 - YMin3) / 10) * I), RGB(255, 255, 255)
        picVoltage4.Line (0, ((YMax4 - YMin4) / 10) * I)-(MaxPlotNo, ((YMax4 - YMin4) / 10) * I), RGB(255, 255, 255)
    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage1.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage2.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage3.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage4.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage1.PSet (0, 0)  '设定起点
    picVoltage2.PSet (0, 0)  '设定起点
    picVoltage3.PSet (0, 0)  '设定起点
    picVoltage4.PSet (0, 0)  '设定起点
    End Sub

'A测温点清屏按纽
Private Sub cmdClear1_Click()
    NowX1 = 0
    picVoltage1.Cls '清除图形
    
    '设置坐标网络，画网格需要使用一个像素宽度的画笔
    Dim I%
    picVoltage1.CurrentX = 0  '设定起点
    picVoltage1.CurrentY = 0  '设定起点
    picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage1.Line (0, (Abs(YMax1 - YMin1) / 10) * I + YMin1)-(MaxPlotNo, (Abs(YMax1 - YMin1) / 10) * I + YMin1), RGB(255, 255, 255)
    Next I
    '画坐标网络的竖线
'    For I = 1 To 9 Step 1
'        picVoltage1.Line (((0 + MaxPlotNo) / 10) * I, YMax1)-(((0 + MaxPlotNo) / 10) * I, YMin1), RGB(255, 255, 255)
'    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage1.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage1.PSet (0, 0)  '设定起点
End Sub

'B测温点清屏按纽
Private Sub cmdClear2_Click()
    NowX2 = 0
    picVoltage2.Cls '清除图形
    
    '设置坐标网络，画网格需要使用一个像素宽度的画笔
    Dim I%
    picVoltage2.CurrentX = 0  '设定起点
    picVoltage2.CurrentY = 0  '设定起点
    picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage2.Line (0, (Abs(YMax2 - YMin2) / 10) * I + YMin2)-(MaxPlotNo, (Abs(YMax2 - YMin2) / 10) * I + YMin2), RGB(255, 255, 255)
    Next I
    '画坐标网络的竖线
'    For I = 1 To 9 Step 1
'        picVoltage2.Line (((0 + MaxPlotNo) / 10) * I, YMax2)-(((0 + MaxPlotNo) / 10) * I, YMin2), RGB(255, 255, 255)
'    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage2.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage2.PSet (0, 0)  '设定起点
End Sub

'C测温点清屏按纽
Private Sub cmdClear3_Click()
    NowX3 = 0
    picVoltage3.Cls '清除图形
    
    '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
    Dim I%
    picVoltage3.CurrentX = 0  '设定起点
    picVoltage3.CurrentY = 0  '设定起点
    picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage3.Line (0, (Abs(YMax3 - YMin3) / 10) * I + YMin3)-(MaxPlotNo, (Abs(YMax3 - YMin3) / 10) * I + YMin3), RGB(255, 255, 255)
    Next I
    '画坐标网络的竖线
'    For I = 1 To 9 Step 1
'        picVoltage3.Line (((0 + MaxPlotNo) / 10) * I, YMax3)-(((0 + MaxPlotNo) / 10) * I, YMin3), RGB(255, 255, 255)
'    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage3.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage3.PSet (0, 0)  '设定起点
End Sub

'D测温点清屏按纽
Private Sub cmdClear4_Click()
    NowX4 = 0
    picVoltage4.Cls '清除图形
    
    '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
    Dim I%
    picVoltage4.CurrentX = 0  '设定起点
    picVoltage4.CurrentY = 0  '设定起点
    picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage4.Line (0, (Abs(YMax4 - YMin4) / 10) * I + YMin4)-(MaxPlotNo, (Abs(YMax4 - YMin4) / 10) * I + YMin4), RGB(255, 255, 255)
    Next I
    '画坐标网络的竖线
'    For I = 1 To 9 Step 1
'        picVoltage4.Line (((0 + MaxPlotNo) / 10) * I, YMax4)-(((0 + MaxPlotNo) / 10) * I, YMin4), RGB(255, 255, 255)
'    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage4.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage4.PSet (0, 0)  '设定起点
End Sub

'将定时器激活或关闭，并显示对应的文字在按钮上，以指示用户操作
Private Sub cmdStart_Click()
    Flag = Not Flag  '用于控制总测温开始与结束的标志位
    If Flag Then
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
        Timer4.Enabled = True
        cmdStart.Caption = "全部停止测温"
        cmdStart1.Caption = "停止测温"
        cmdStart2.Caption = "停止测温"
        cmdStart3.Caption = "停止测温"
        cmdStart4.Caption = "停止测温"
        lblMsg.Caption = "提示：已全部开始测温"
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        cmdStart.Caption = "全部开始测温"
        cmdStart1.Caption = "开始测温"
        cmdStart2.Caption = "开始测温"
        cmdStart3.Caption = "开始测温"
        cmdStart4.Caption = "开始测温"
        lblMsg.Caption = "提示：已全部停止测温"
    End If
End Sub

'A测温点开始按纽
Private Sub cmdStart1_Click()
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled Then
        cmdStart1.Caption = "停止测温"
    Else
        cmdStart1.Caption = "开始测温"
    End If
End Sub

'B测温点开始按纽
Private Sub cmdStart2_Click()
    Timer2.Enabled = Not Timer2.Enabled
    If Timer2.Enabled Then
        cmdStart2.Caption = "停止测温"
    Else
        cmdStart2.Caption = "开始测温"
    End If
End Sub

'C测温点开始按纽
Private Sub cmdStart3_Click()
    Timer3.Enabled = Not Timer3.Enabled
    If Timer3.Enabled Then
        cmdStart3.Caption = "停止测温"
    Else
        cmdStart3.Caption = "开始测温"
    End If
End Sub

'D测温点开始按纽
Private Sub cmdStart4_Click()
    Timer4.Enabled = Not Timer4.Enabled
    If Timer4.Enabled Then
        cmdStart4.Caption = "停止测温"
    Else
        cmdStart4.Caption = "开始测温"
    End If
End Sub

'窗体的Load事件
'输入图形暂时设为灰色，表示无状态信息进入
'将通讯端口号码及站号填入Combo控件；并默认二者的选项是第一个
Private Sub Form_Load()
 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Pellucidity = 255 * 20 * 0.01 '透明度（完全透明0~不透明255）
    '以下部分代码用于设置窗口透明
    Pellucidity = 255 - (Val(lblPellucidity.Caption) * 0.01 * 255)
    Dim rtn As Long
'    Me.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), Pellucidity, LWA_ALPHA  'RGB(0, 0, 0)参数就是要透明掉的颜色，后面那个数值改变透明度大小（完全透明0~不透明255）
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cmbCOMWidthSize = cmbCOM.Width / Form1.Width 'cmbCOM控件宽度与窗口宽度的比值，用于当窗口缩放时cmbCOM控制控件的宽度随之跟着变化
    cmbCOMLeftSize = cmbCOM.Left / Form1.Width  'cmbCOM控件左边位置与窗口宽度的比值，用于当窗口缩放时cmbCOM控制控件的左边位置随之跟着变化
    Form1.Caption = "多点温度监控系统上位机——By WillianChan" & "——当前时间" & Time$  '在窗口状态栏显示当前系统时间，实时刷新代码在timer5控件处
    Call ResizeInit(Me)  '确保窗体改变时控件随之改变
    Dim I%
    MaxPlotNo = 100  '最长的X轴范围
    DelayTimeValueMS = 1200  '延时时长为1200毫秒
    Shape1.BackColor = &HC0C0C0  '端口打开提示灯颜色
    Label7.Caption = "端口已关闭"
    cmbCOM.Clear
     cmbCOM.AddItem "COM1"
    cmbCOM.AddItem "COM2"
     cmbCOM.AddItem "COM3"
    cmbCOM.AddItem "COM4"
     cmbCOM.AddItem "COM5"
    cmbCOM.AddItem "COM6"
     cmbCOM.AddItem "COM7"
    cmbCOM.AddItem "COM8"
     cmbCOM.AddItem "COM9"
    cmbCOM.AddItem "COM10"
     cmbCOM.AddItem "COM11"
    cmbCOM.AddItem "COM12"
     cmbCOM.AddItem "COM13"
    cmbCOM.AddItem "COM14"
     cmbCOM.AddItem "COM15"
    cmbCOM.AddItem "COM16"
    cmbCOM.ListIndex = 0
    cmdStart.Enabled = False
    cmdStart1.Enabled = False
    cmdStart2.Enabled = False
    cmdStart3.Enabled = False
    cmdStart4.Enabled = False
    
    '以下设定绘图范围，(Xmin,YMax)-(XMax,YMin)
    YMax1 = Val(YMaxText1.Text)
    YMax2 = Val(YMaxText2.Text)
    YMax3 = Val(YMaxText3.Text)
    YMax4 = Val(YMaxText4.Text)
    YMin1 = Val(YMinText1.Text)
    YMin2 = Val(YMinText2.Text)
    YMin3 = Val(YMinText3.Text)
    YMin4 = Val(YMinText4.Text)
    picVoltage1.Scale (0, YMax1)-(MaxPlotNo, YMin1)
    picVoltage2.Scale (0, YMax2)-(MaxPlotNo, YMin2)
    picVoltage3.Scale (0, YMax3)-(MaxPlotNo, YMin3)
    picVoltage4.Scale (0, YMax4)-(MaxPlotNo, YMin4)
    
    '以下设置坐标网络，画网格需要使用一个像素宽度的画笔
    picVoltage1.CurrentX = 0  '设定起点
    picVoltage1.CurrentY = 0  '设定起点
    picVoltage2.CurrentX = 0  '设定起点
    picVoltage2.CurrentY = 0  '设定起点
    picVoltage3.CurrentX = 0  '设定起点
    picVoltage3.CurrentY = 0  '设定起点
    picVoltage4.CurrentX = 0  '设定起点
    picVoltage4.CurrentY = 0  '设定起点
    picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
    picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
    '画坐标网络的横线
    For I = 1 To 9 Step 1
        picVoltage1.Line (0, (Abs(YMax1 - YMin1) / 10) * I + YMin1)-(MaxPlotNo, (Abs(YMax1 - YMin1) / 10) * I + YMin1), RGB(255, 255, 255)
        picVoltage2.Line (0, (Abs(YMax2 - YMin2) / 10) * I + YMin2)-(MaxPlotNo, (Abs(YMax2 - YMin2) / 10) * I + YMin2), RGB(255, 255, 255)
        picVoltage3.Line (0, (Abs(YMax3 - YMin3) / 10) * I + YMin3)-(MaxPlotNo, (Abs(YMax3 - YMin3) / 10) * I + YMin3), RGB(255, 255, 255)
        picVoltage4.Line (0, (Abs(YMax4 - YMin4) / 10) * I + YMin4)-(MaxPlotNo, (Abs(YMax4 - YMin4) / 10) * I + YMin4), RGB(255, 255, 255)
    Next I
    '画坐标网络的竖线
'    For I = 1 To 9 Step 1
'        picVoltage1.Line (((MaxPlotNo - 0) / 10) * I, YMax1)-(((MaxPlotNo - 0) / 10) * I, YMin1), RGB(255, 255, 255)
'        picVoltage2.Line (((MaxPlotNo - 0) / 10) * I, YMax1)-(((MaxPlotNo - 0) / 10) * I, YMin1), RGB(255, 255, 255)
'        picVoltage3.Line (((MaxPlotNo - 0) / 10) * I, YMax1)-(((MaxPlotNo - 0) / 10) * I, YMin1), RGB(255, 255, 255)
'        picVoltage4.Line (((MaxPlotNo - 0) / 10) * I, YMax1)-(((MaxPlotNo - 0) / 10) * I, YMin1), RGB(255, 255, 255)
'    Next I
    '画曲线需要使用两个像素宽度的画笔
    picVoltage1.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage2.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage3.DrawWidth = 2 '使用两个像素宽度的画笔
    picVoltage4.DrawWidth = 2 '使用两个像素宽度的画笔
    
    picVoltage1.CurrentX = 0  '设定起点
    picVoltage1.CurrentY = 0  '设定起点
    picVoltage2.CurrentX = 0  '设定起点
    picVoltage2.CurrentY = 0  '设定起点
    picVoltage3.CurrentX = 0  '设定起点
    picVoltage3.CurrentY = 0  '设定起点
    picVoltage4.CurrentX = 0  '设定起点
    picVoltage4.CurrentY = 0  '设定起点
End Sub

'增加透明度，上限70%
Private Sub PellucidityAdd_Click()
    If Val(lblPellucidity.Caption) < 70 Then
        lblPellucidity.Caption = Val(lblPellucidity.Caption) + 10 & "%"
    Else
        lblPellucidity.Caption = Val(lblPellucidity.Caption) & "%"  '达到上限之后不再变化
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '以下部分代码用于设置窗口透明
    Pellucidity = 255 - (Val(lblPellucidity.Caption) * 0.01 * 255)
    Dim rtn As Long
'    Me.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), Pellucidity, LWA_ALPHA  'RGB(0, 0, 0)参数就是要透明掉的颜色，后面那个数值改变透明度大小（完全透明0~不透明255）(LWA_COLORKEY,LWA_ALPHA)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

'降低透明度，下限0%（不透明）
Private Sub PelluciditySub_Click()
    If Val(lblPellucidity.Caption) <> 0 Then
        lblPellucidity.Caption = Val(lblPellucidity.Caption) - 10 & "%"
    Else
        lblPellucidity.Caption = Val(lblPellucidity.Caption) & "%"  '达到下限之后不再变化
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '以下部分代码用于设置窗口透明
    Pellucidity = 255 - (Val(lblPellucidity.Caption) * 0.01 * 255)
    Dim rtn As Long
'    Me.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), Pellucidity, LWA_ALPHA  'RGB(0, 0, 0)参数就是要透明掉的颜色，后面那个数值改变透明度大小（完全透明0~不透明255）(LWA_COLORKEY,LWA_ALPHA)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

'定时器的Timer事件引发后，就不断地执行其中的程序。
'将模拟读值命令送出，再取得返回字符串并判断。
Private Sub Timer1_Timer()
    Buf = MSComm1.Input  '读取变量，存放串口通讯数据
    TimeDelay (DelayTimeValueMS)
    
    ValueStr1 = Val(Mid(Buf, 2, 6))  '分离出正号以后的数值
    lblValue1.Caption = ""                    '清空上一次数据
    lblValue1.Caption = "+" & Format(ValueStr1, "00.00") & "℃"  '显示在画面上
     
    If NowX1 = 0 Then
        SaveFileTimeStar1 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，A测温点温度曲线开始的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        picVoltage1.Cls '清除图形
        '设置坐标网络，画网格需要使用一个像素宽度的画笔
        Dim I%
        picVoltage1.CurrentX = 0  '设定起点
        picVoltage1.CurrentY = 0  '设定起点
        picVoltage1.DrawWidth = 1 '使用一个像素宽度的画笔
        '画坐标网络的横线
        For I = 1 To 9 Step 1
            picVoltage1.Line (0, (Abs(YMax1 - YMin1) / 10) * I + YMin1)-(MaxPlotNo, (Abs(YMax1 - YMin1) / 10) * I + YMin1), RGB(255, 255, 255)
        Next I
        '画坐标网络的竖线
    '    For I = 1 To 9 Step 1
    '        picVoltage1.Line (((0 + MaxPlotNo) / 10) * I, YMax1)-(((0 + MaxPlotNo) / 10) * I, YMin1), RGB(255, 255, 255)
    '    Next I
        '画曲线需要使用两个像素宽度的画笔
        picVoltage1.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage1.PSet (0, ValueStr1)  '设定起点
    Else
        '以下判断现在的读值是否大于前一次的读值，若是，则以红色绘线
        '若否，则以绿色绘线
        '若不变，则以蓝色绘线
        If ValueStr1 > PreValue1 Then
            picVoltage1.Line -(NowX1, ValueStr1), RGB(255, 0, 0) '由上一次的位置画至此点
        End If
        If ValueStr1 = PreValue1 Then
            picVoltage1.Line -(NowX1, ValueStr1), RGB(0, 0, 255) '由上一次的位置画至此点
        End If
        If ValueStr1 < PreValue1 Then
            picVoltage1.Line -(NowX1, ValueStr1), RGB(0, 255, 0) '由上一次的位置画至此点
        End If
    End If
     
    PreValue1 = ValueStr1  '当曲线准备画下一个点时，当前值变成上一次的值
     
    NowX1 = NowX1 + 1 '位置加1
    If NowX1 > MaxPlotNo Then
        NowX1 = 0  '超过范围则数值归零
        SaveFileTimeEnd1 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，A测温点温度曲线结束的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        If Dir(App.Path & "\多点无线温度采集系统上位机测温点数据", vbDirectory) <> "" Then  '判断路径是否存在，vbNormal是普通文件，vbHidden是隐藏文件，vbDirectory是文件夹
            '将图片保存到文件。picVoltage1.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage1.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "A测温点数据" & "——时间" & SaveFileTimeStar1 & "~" & SaveFileTimeEnd1 & ".jpg"
        Else
            MkDir (App.Path & "\多点无线温度采集系统上位机测温点数据")  '若路径不存在则新建此路径
            '将图片保存到文件。picVoltage1.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage1.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "A测温点数据" & "——时间" & SaveFileTimeStar1 & "~" & SaveFileTimeEnd1 & ".jpg"
        End If
'        picVoltage1.BackColor = &H808080    '闪一下，提示图片数据已保存
'        TimeDelay (50)  '延时50ms
'        picVoltage1.BackColor = &H404040   '恢复原色
    End If
End Sub

'定时器的Timer事件引发后，就不断地执行其中的程序。
'将模拟读值命令送出，再取得返回字符串并判断。
Private Sub Timer2_Timer()
    Buf = MSComm1.Input  '读取变量，存放串口通讯数据
    TimeDelay (DelayTimeValueMS)
    
    ValueStr2 = Val(Mid(Buf, 10, 6)) '分离出正号以后的数值
    lblValue2.Caption = ""                    '清空上一次数据
    lblValue2.Caption = "+" & Format(ValueStr2, "00.00") & "℃"  '显示在画面上
     
    If NowX2 = 0 Then
        SaveFileTimeStar2 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，B测温点温度曲线开始的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        picVoltage2.Cls '清除图形
        '设置坐标网络，画网格需要使用一个像素宽度的画笔
        Dim I%
        picVoltage2.CurrentX = 0  '设定起点
        picVoltage2.CurrentY = 0  '设定起点
        picVoltage2.DrawWidth = 1 '使用一个像素宽度的画笔
        '画坐标网络的横线
        For I = 1 To 9 Step 1
            picVoltage2.Line (0, (Abs(YMax2 - YMin2) / 10) * I + YMin2)-(MaxPlotNo, (Abs(YMax2 - YMin2) / 10) * I + YMin2), RGB(255, 255, 255)
        Next I
        '画坐标网络的竖线
    '    For I = 1 To 9 Step 1
    '        picVoltage2.Line (((0 + MaxPlotNo) / 10) * I, YMax2)-(((0 + MaxPlotNo) / 10) * I, YMin2), RGB(255, 255, 255)
    '    Next I
        '画曲线需要使用两个像素宽度的画笔
        picVoltage2.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage2.PSet (0, ValueStr2)  '设定起点
    Else
        '以下判断现在的读值是否大于前一次的读值，若是，则以红色绘线
        '若否，则以绿色绘线
        '若不变，则以蓝色绘线
        If ValueStr2 > PreValue2 Then
            picVoltage2.Line -(NowX2, ValueStr2), RGB(255, 0, 0) '由上一次的位置画至此点
        End If
        If ValueStr2 = PreValue2 Then
            picVoltage2.Line -(NowX2, ValueStr2), RGB(0, 0, 255) '由上一次的位置画至此点
        End If
        If ValueStr2 < PreValue2 Then
            picVoltage2.Line -(NowX2, ValueStr2), RGB(0, 255, 0) '由上一次的位置画至此点
        End If
    End If
     
    PreValue2 = ValueStr2  '当曲线准备画下一个点时，当前值变成上一次的值
     
    NowX2 = NowX2 + 1 '位置加1
    If NowX2 > MaxPlotNo Then
        NowX2 = 0  '超过范围则数值归零
        SaveFileTimeEnd2 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，B测温点温度曲线结束的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        If Dir(App.Path & "\多点无线温度采集系统上位机测温点数据", vbDirectory) <> "" Then  '判断路径是否存在，vbNormal是普通文件，vbHidden是隐藏文件，vbDirectory是文件夹
            '将图片保存到文件。picVoltage2.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage2.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "B测温点数据" & "——时间" & SaveFileTimeStar2 & "~" & SaveFileTimeEnd2 & ".jpg"
        Else
            MkDir (App.Path & "\多点无线温度采集系统上位机测温点数据")  '若路径不存在则新建此路径
            '将图片保存到文件。picVoltage2.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage2.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "B测温点数据" & "——时间" & SaveFileTimeStar2 & "~" & SaveFileTimeEnd2 & ".jpg"
        End If
'        picVoltage2.BackColor = &H808080    '闪一下，提示图片数据已保存
'        TimeDelay (50)  '延时50ms
'        picVoltage2.BackColor = &H404040   '恢复原色
    End If
End Sub

'定时器的Timer事件引发后，就不断地执行其中的程序。
'将模拟读值命令送出，再取得返回字符串并判断。
Private Sub Timer3_Timer()
    Buf = MSComm1.Input  '读取变量，存放串口通讯数据
    TimeDelay (DelayTimeValueMS)
    
    ValueStr3 = Val(Mid(Buf, 18, 6)) '分离出正号以后的数值
    lblValue3.Caption = ""                    '清空上一次数据
    lblValue3.Caption = "+" & Format(ValueStr3, "00.00") & "℃"  '显示在画面上
     
    If NowX3 = 0 Then
        SaveFileTimeStar3 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，C测温点温度曲线开始的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        picVoltage3.Cls '清除图形
        '设置坐标网络，画网格需要使用一个像素宽度的画笔
        Dim I%
        picVoltage3.CurrentX = 0  '设定起点
        picVoltage3.CurrentY = 0  '设定起点
        picVoltage3.DrawWidth = 1 '使用一个像素宽度的画笔
        '画坐标网络的横线
        For I = 1 To 9 Step 1
            picVoltage3.Line (0, (Abs(YMax3 - YMin3) / 10) * I + YMin3)-(MaxPlotNo, (Abs(YMax3 - YMin3) / 10) * I + YMin3), RGB(255, 255, 255)
        Next I
        '画坐标网络的竖线
    '    For I = 1 To 9 Step 1
    '        picVoltage3.Line (((0 + MaxPlotNo) / 10) * I, YMax3)-(((0 + MaxPlotNo) / 10) * I, YMin3), RGB(255, 255, 255)
    '    Next I
        '画曲线需要使用两个像素宽度的画笔
        picVoltage3.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage3.PSet (0, ValueStr3)  '设定起点
    Else
        '以下判断现在的读值是否大于前一次的读值，若是，则以红色绘线
        '若否，则以绿色绘线
        '若不变，则以蓝色绘线
        If ValueStr3 > PreValue3 Then
            picVoltage3.Line -(NowX3, ValueStr3), RGB(255, 0, 0) '由上一次的位置画至此点
        End If
        If ValueStr3 = PreValue3 Then
            picVoltage3.Line -(NowX3, ValueStr3), RGB(0, 0, 255) '由上一次的位置画至此点
        End If
        If ValueStr3 < PreValue3 Then
            picVoltage3.Line -(NowX3, ValueStr3), RGB(0, 255, 0) '由上一次的位置画至此点
        End If
    End If
     
    PreValue3 = ValueStr3  '当曲线准备画下一个点时，当前值变成上一次的值
     
    NowX3 = NowX3 + 1 '位置加1
    If NowX3 > MaxPlotNo Then
        NowX3 = 0  '超过范围则数值归零
        SaveFileTimeEnd3 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，C测温点温度曲线结束的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        If Dir(App.Path & "\多点无线温度采集系统上位机测温点数据", vbDirectory) <> "" Then  '判断路径是否存在，vbNormal是普通文件，vbHidden是隐藏文件，vbDirectory是文件夹
            '将图片保存到文件。picVoltage3.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage3.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "C测温点数据" & "——时间" & SaveFileTimeStar3 & "~" & SaveFileTimeEnd3 & ".jpg"
        Else
            MkDir (App.Path & "\多点无线温度采集系统上位机测温点数据")  '若路径不存在则新建此路径
            '将图片保存到文件。picVoltage3.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage3.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "C测温点数据" & "——时间" & SaveFileTimeStar3 & "~" & SaveFileTimeEnd3 & ".jpg"
        End If
'        picVoltage3.BackColor = &H808080    '闪一下，提示图片数据已保存
'        TimeDelay (50)  '延时50ms
'        picVoltage3.BackColor = &H404040   '恢复原色
    End If
End Sub

'定时器的Timer事件引发后，就不断地执行其中的程序。
'将模拟读值命令送出，再取得返回字符串并判断。
Private Sub Timer4_Timer()
    Buf = MSComm1.Input  '读取变量，存放串口通讯数据
    TimeDelay (DelayTimeValueMS)
    
    ValueStr4 = Val(Mid(Buf, 26, 6)) '分离出正号以后的数值
    lblValue4.Caption = ""                    '清空上一次数据
    lblValue4.Caption = "+" & Format(ValueStr4, "00.00") & "℃" '显示在画面上
    
    If NowX4 = 0 Then
        SaveFileTimeStar4 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，D测温点温度曲线开始的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        picVoltage4.Cls '清除图形
        '设置坐标网络，画网格需要使用一个像素宽度的画笔
        Dim I%
        picVoltage4.CurrentX = 0  '设定起点
        picVoltage4.CurrentY = 0  '设定起点
        picVoltage4.DrawWidth = 1 '使用一个像素宽度的画笔
        '画坐标网络的横线
        For I = 1 To 9 Step 1
            picVoltage4.Line (0, (Abs(YMax4 - YMin4) / 10) * I + YMin4)-(MaxPlotNo, (Abs(YMax4 - YMin4) / 10) * I + YMin4), RGB(255, 255, 255)
        Next I
        '画坐标网络的竖线
    '    For I = 1 To 9 Step 1
    '        picVoltage4.Line (((0 + MaxPlotNo) / 10) * I, YMax4)-(((0 + MaxPlotNo) / 10) * I, YMin4), RGB(255, 255, 255)
    '    Next I
        '画曲线需要使用两个像素宽度的画笔
        picVoltage4.DrawWidth = 2 '使用两个像素宽度的画笔
        picVoltage4.PSet (0, ValueStr4)  '设定起点
    Else
        '以下判断现在的读值是否大于前一次的读值，若是，则以红色绘线
        '若否，则以绿色绘线
        '若不变，则以蓝色绘线
        If ValueStr4 > PreValue4 Then
            picVoltage4.Line -(NowX4, ValueStr4), RGB(255, 0, 0) '由上一次的位置画至此点
        End If
        If ValueStr4 = PreValue4 Then
            picVoltage4.Line -(NowX4, ValueStr4), RGB(0, 0, 255) '由上一次的位置画至此点
        End If
        If ValueStr4 < PreValue4 Then
            picVoltage4.Line -(NowX4, ValueStr4), RGB(0, 255, 0) '由上一次的位置画至此点
        End If
    End If
     
    PreValue4 = ValueStr4  '当曲线准备画下一个点时，当前值变成上一次的值
     
    NowX4 = NowX4 + 1 '位置加1
    If NowX4 > MaxPlotNo Then
        NowX4 = 0  '超过范围则数值归零
        SaveFileTimeEnd4 = Format(Now, "yyyy年mm月dd日hh：mm：ss")  '保存图片文件时，D测温点温度曲线结束的时间，注意！文件名不能包含以下英文字符\/:*?"<>|
        If Dir(App.Path & "\多点无线温度采集系统上位机测温点数据", vbDirectory) <> "" Then  '判断路径是否存在，vbNormal是普通文件，vbHidden是隐藏文件，vbDirectory是文件夹
            '将图片保存到文件。picVoltage4.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage4.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "D测温点数据" & "——时间" & SaveFileTimeStar4 & "~" & SaveFileTimeEnd4 & ".jpg"
        Else
            MkDir (App.Path & "\多点无线温度采集系统上位机测温点数据")  '若路径不存在则新建此路径
            '将图片保存到文件。picVoltage4.AutoRedraw属性要设置成true，注意！文件名不能包含以下英文字符\/:*?"<>|
            SavePicture picVoltage4.Image, App.Path & "\多点无线温度采集系统上位机测温点数据" & "\" & "D测温点数据" & "——时间" & SaveFileTimeStar4 & "~" & SaveFileTimeEnd4 & ".jpg"
        End If
'        picVoltage4.BackColor = &H808080    '闪一下，提示图片数据已保存
'        TimeDelay (50)  '延时50ms
'        picVoltage4.BackColor = &H404040   '恢复原色
    End If
End Sub

'不断读取变量，存放串口通讯数据
Private Sub Timer5_Timer()
    If MSComm1.PortOpen Then
        Buf = Buf + MSComm1.Input  '读取变量，存放串口通讯数据
    End If
    Form1.Caption = "多点温度监控系统上位机——By WillianChan" & "——当前时间" & Time$  '在窗口状态栏实时刷新当前系统时间
End Sub

'按比例改变表单内各元件的大小,在调用ReSizeForm前先调用ReSizeInit函数
Private Sub Form_Resize()
    Call ResizeForm(Me)  '窗口改变大小控件随之改变
    If Form1.ScaleWidth <> 0 Or Form1.ScaleHeight <> 0 Then  '防止窗口最小化的时候报错
        Call cmdChange1_Click
        Call cmdChange2_Click
        Call cmdChange3_Click
        Call cmdChange4_Click
    End If
    '因为ResizeForm()函数对这个控件无效（不知道为什么），所以只能手动设置其缩放
    cmbCOM.Width = cmbCOMWidthSize * Form1.Width  '当窗口缩放时cmbCOM控制控件的宽度随之跟着变化
    cmbCOM.Left = cmbCOMLeftSize * Form1.Width  '当窗口缩放时cmbCOM控制控件的左边位置随之跟着变化
End Sub
