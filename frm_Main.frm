VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantone Conversion Utility"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   6495
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Convert"
      TabPicture(0)   =   "frm_Main.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPantoneColor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPantone"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCMYK"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdConvert"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_Y"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_C"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_M"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraRGB"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraHex"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdDisplay"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraNavigation"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboPantone"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdConvertCurrent"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtRGB"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Colors1"
      TabPicture(1)   =   "frm_Main.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Picture1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Picture1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Picture1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Picture1(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Picture1(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Picture1(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Picture1(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Picture1(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Picture1(11)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Picture1(12)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Picture1(13)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Picture1(14)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Picture1(15)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Picture1(16)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Picture1(17)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Picture1(18)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Picture1(19)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Picture1(20)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Picture1(21)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Picture1(22)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Picture1(23)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Picture1(24)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Picture1(25)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Picture1(26)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Picture1(27)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Picture1(28)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Picture1(29)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Picture1(30)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Picture1(31)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Picture1(32)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Picture1(33)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Picture1(34)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Picture1(35)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Picture1(36)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Picture1(37)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Picture1(38)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Picture1(39)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Picture1(40)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Picture1(41)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Picture1(42)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Picture1(43)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Picture1(44)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Picture1(45)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Picture1(46)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Picture1(47)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Picture1(48)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Picture1(49)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Picture1(50)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Picture1(51)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Picture1(52)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Picture1(53)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Picture1(54)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Picture1(55)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Picture1(5)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).ControlCount=   56
      TabCaption(2)   =   "Colors2"
      TabPicture(2)   =   "frm_Main.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1(56)"
      Tab(2).Control(1)=   "Picture1(57)"
      Tab(2).Control(2)=   "Picture1(58)"
      Tab(2).Control(3)=   "Picture1(59)"
      Tab(2).Control(4)=   "Picture1(60)"
      Tab(2).Control(5)=   "Picture1(61)"
      Tab(2).Control(6)=   "Picture1(62)"
      Tab(2).Control(7)=   "Picture1(63)"
      Tab(2).Control(8)=   "Picture1(64)"
      Tab(2).Control(9)=   "Picture1(65)"
      Tab(2).Control(10)=   "Picture1(66)"
      Tab(2).Control(11)=   "Picture1(67)"
      Tab(2).Control(12)=   "Picture1(68)"
      Tab(2).Control(13)=   "Picture1(69)"
      Tab(2).Control(14)=   "Picture1(70)"
      Tab(2).Control(15)=   "Picture1(71)"
      Tab(2).Control(16)=   "Picture1(72)"
      Tab(2).Control(17)=   "Picture1(73)"
      Tab(2).Control(18)=   "Picture1(74)"
      Tab(2).Control(19)=   "Picture1(75)"
      Tab(2).Control(20)=   "Picture1(76)"
      Tab(2).Control(21)=   "Picture1(77)"
      Tab(2).Control(22)=   "Picture1(78)"
      Tab(2).Control(23)=   "Picture1(79)"
      Tab(2).Control(24)=   "Picture1(80)"
      Tab(2).Control(25)=   "Picture1(81)"
      Tab(2).Control(26)=   "Picture1(82)"
      Tab(2).Control(27)=   "Picture1(83)"
      Tab(2).Control(28)=   "Picture1(84)"
      Tab(2).Control(29)=   "Picture1(85)"
      Tab(2).Control(30)=   "Picture1(86)"
      Tab(2).Control(31)=   "Picture1(87)"
      Tab(2).Control(32)=   "Picture1(88)"
      Tab(2).Control(33)=   "Picture1(89)"
      Tab(2).Control(34)=   "Picture1(90)"
      Tab(2).Control(35)=   "Picture1(91)"
      Tab(2).Control(36)=   "Picture1(92)"
      Tab(2).Control(37)=   "Picture1(93)"
      Tab(2).Control(38)=   "Picture1(94)"
      Tab(2).Control(39)=   "Picture1(95)"
      Tab(2).Control(40)=   "Picture1(96)"
      Tab(2).Control(41)=   "Picture1(97)"
      Tab(2).Control(42)=   "Picture1(98)"
      Tab(2).Control(43)=   "Picture1(99)"
      Tab(2).Control(44)=   "Picture1(100)"
      Tab(2).Control(45)=   "Picture1(101)"
      Tab(2).Control(46)=   "Picture1(102)"
      Tab(2).Control(47)=   "Picture1(103)"
      Tab(2).Control(48)=   "Picture1(104)"
      Tab(2).Control(49)=   "Picture1(105)"
      Tab(2).Control(50)=   "Picture1(106)"
      Tab(2).Control(51)=   "Picture1(107)"
      Tab(2).Control(52)=   "Picture1(108)"
      Tab(2).Control(53)=   "Picture1(109)"
      Tab(2).Control(54)=   "Picture1(110)"
      Tab(2).Control(55)=   "Picture1(111)"
      Tab(2).ControlCount=   56
      TabCaption(3)   =   "Colors3"
      TabPicture(3)   =   "frm_Main.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1(112)"
      Tab(3).Control(1)=   "Picture1(113)"
      Tab(3).Control(2)=   "Picture1(114)"
      Tab(3).Control(3)=   "Picture1(115)"
      Tab(3).Control(4)=   "Picture1(116)"
      Tab(3).Control(5)=   "Picture1(117)"
      Tab(3).Control(6)=   "Picture1(118)"
      Tab(3).Control(7)=   "Picture1(119)"
      Tab(3).Control(8)=   "Picture1(120)"
      Tab(3).Control(9)=   "Picture1(121)"
      Tab(3).Control(10)=   "Picture1(122)"
      Tab(3).Control(11)=   "Picture1(123)"
      Tab(3).Control(12)=   "Picture1(124)"
      Tab(3).Control(13)=   "Picture1(125)"
      Tab(3).Control(14)=   "Picture1(126)"
      Tab(3).Control(15)=   "Picture1(127)"
      Tab(3).Control(16)=   "Picture1(128)"
      Tab(3).Control(17)=   "Picture1(129)"
      Tab(3).Control(18)=   "Picture1(130)"
      Tab(3).Control(19)=   "Picture1(131)"
      Tab(3).Control(20)=   "Picture1(132)"
      Tab(3).Control(21)=   "Picture1(133)"
      Tab(3).Control(22)=   "Picture1(134)"
      Tab(3).Control(23)=   "Picture1(135)"
      Tab(3).Control(24)=   "Picture1(136)"
      Tab(3).Control(25)=   "Picture1(137)"
      Tab(3).Control(26)=   "Picture1(138)"
      Tab(3).Control(27)=   "Picture1(139)"
      Tab(3).Control(28)=   "Picture1(140)"
      Tab(3).Control(29)=   "Picture1(141)"
      Tab(3).Control(30)=   "Picture1(142)"
      Tab(3).Control(31)=   "Picture1(143)"
      Tab(3).Control(32)=   "Picture1(144)"
      Tab(3).Control(33)=   "Picture1(145)"
      Tab(3).Control(34)=   "Picture1(146)"
      Tab(3).Control(35)=   "Picture1(147)"
      Tab(3).Control(36)=   "Picture1(148)"
      Tab(3).Control(37)=   "Picture1(149)"
      Tab(3).Control(38)=   "Picture1(150)"
      Tab(3).Control(39)=   "Picture1(151)"
      Tab(3).Control(40)=   "Picture1(152)"
      Tab(3).Control(41)=   "Picture1(153)"
      Tab(3).Control(42)=   "Picture1(154)"
      Tab(3).Control(43)=   "Picture1(155)"
      Tab(3).Control(44)=   "Picture1(156)"
      Tab(3).Control(45)=   "Picture1(157)"
      Tab(3).Control(46)=   "Picture1(158)"
      Tab(3).Control(47)=   "Picture1(159)"
      Tab(3).Control(48)=   "Picture1(160)"
      Tab(3).Control(49)=   "Picture1(161)"
      Tab(3).Control(50)=   "Picture1(162)"
      Tab(3).Control(51)=   "Picture1(163)"
      Tab(3).Control(52)=   "Picture1(164)"
      Tab(3).Control(53)=   "Picture1(165)"
      Tab(3).Control(54)=   "Picture1(166)"
      Tab(3).Control(55)=   "Picture1(167)"
      Tab(3).ControlCount=   56
      Begin VB.TextBox txtRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1905
         Left            =   270
         TabIndex        =   194
         Top             =   3735
         Width           =   5865
      End
      Begin VB.CommandButton cmdConvertCurrent 
         Caption         =   "Convert Current"
         Height          =   285
         Left            =   3150
         TabIndex        =   192
         Top             =   2025
         Width           =   1500
      End
      Begin VB.ComboBox cboPantone 
         Height          =   315
         Left            =   2025
         Style           =   2  'Dropdown List
         TabIndex        =   189
         Top             =   3285
         Width           =   4110
      End
      Begin VB.Frame fraNavigation 
         Caption         =   "Navigation"
         Height          =   1950
         Left            =   4770
         TabIndex        =   183
         Top             =   1110
         Width           =   1365
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   285
            Left            =   225
            TabIndex        =   187
            Top             =   1530
            Width           =   960
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            Height          =   285
            Left            =   225
            TabIndex        =   186
            Top             =   1125
            Width           =   960
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Previous"
            Height          =   285
            Left            =   225
            TabIndex        =   185
            Top             =   720
            Width           =   960
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   285
            Left            =   225
            TabIndex        =   184
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   5
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   167
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   182
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   166
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   181
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   165
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   180
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   164
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   179
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   163
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   178
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   162
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   177
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   161
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   176
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   160
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   175
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   159
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   174
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   158
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   173
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   157
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   172
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   156
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   171
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   155
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   170
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   154
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   169
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   153
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   168
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   152
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   167
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   151
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   166
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   150
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   165
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   149
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   164
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   148
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   163
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   147
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   162
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   146
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   161
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   145
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   160
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   144
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   159
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   143
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   158
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   142
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   157
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   141
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   156
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   140
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   155
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   139
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   154
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   138
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   153
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   137
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   152
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   136
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   151
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   135
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   150
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   134
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   149
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   133
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   148
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   132
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   147
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   131
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   146
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   130
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   145
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   129
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   144
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   128
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   143
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   127
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   142
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   126
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   141
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   125
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   140
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   124
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   139
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   123
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   138
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   122
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   137
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   121
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   136
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   120
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   135
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   119
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   134
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   118
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   133
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   117
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   132
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   116
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   131
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   115
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   130
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   114
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   129
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   113
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   128
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   112
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   127
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   111
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   126
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   110
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   125
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   109
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   124
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   108
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   123
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   107
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   122
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   106
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   121
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   105
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   120
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   104
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   119
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   103
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   118
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   102
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   117
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   101
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   116
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   100
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   115
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   99
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   114
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   98
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   113
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   97
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   112
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   96
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   111
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   95
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   110
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   94
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   109
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   93
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   108
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   92
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   107
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   91
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   106
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   90
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   105
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   89
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   104
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   88
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   103
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   87
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   102
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   86
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   101
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   85
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   100
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   84
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   99
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   83
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   98
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   82
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   97
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   81
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   96
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   80
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   95
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   79
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   94
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   78
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   93
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   77
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   92
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   76
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   91
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   75
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   90
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   74
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   89
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   73
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   88
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   72
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   87
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   71
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   86
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   70
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   85
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   69
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   84
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   68
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   83
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   67
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   82
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   66
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   81
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   65
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   80
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   64
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   79
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   63
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   78
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   62
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   77
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   61
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   76
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   60
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   75
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   59
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   74
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   58
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   73
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   57
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   72
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   55
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   69
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   54
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   68
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   53
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   67
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   52
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   66
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   51
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   65
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   50
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   64
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   49
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   63
         Top             =   5520
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   48
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   62
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   47
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   61
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   46
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   60
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   45
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   59
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   44
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   58
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   43
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   57
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   42
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   56
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   41
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   55
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   40
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   54
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   39
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   53
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   38
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   52
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   37
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   51
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   36
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   50
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   35
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   49
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   34
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   48
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   33
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   47
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   32
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   46
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   31
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   45
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   30
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   44
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   29
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   43
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   28
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   42
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   27
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   26
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   40
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   25
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   39
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   24
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   38
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   23
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   37
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   22
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   36
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   21
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   35
         Top             =   2640
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   20
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   34
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   19
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   18
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   17
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   16
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   30
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   15
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   29
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   14
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   28
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   13
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   12
         Left            =   -70560
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   11
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   10
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   9
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   8
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   7
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   6
         Left            =   -69720
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   4
         Left            =   -71400
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   3
         Left            =   -72240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   2
         Left            =   -73080
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   1
         Left            =   -73920
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   0
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Index           =   56
         Left            =   -74760
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Refresh Colors"
         Height          =   285
         Left            =   3150
         TabIndex        =   12
         Top             =   2745
         Width           =   1500
      End
      Begin VB.Frame fraHex 
         Caption         =   "HEX"
         Height          =   855
         Left            =   3180
         TabIndex        =   10
         Top             =   1110
         Width           =   1455
         Begin VB.TextBox txt_Hex 
            DataField       =   "HEX"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraRGB 
         Caption         =   "RGB"
         Height          =   1935
         Left            =   1710
         TabIndex        =   5
         Top             =   1110
         Width           =   1335
         Begin VB.TextBox txt_ID 
            DataField       =   "ID"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_B 
            DataField       =   "B"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txt_G 
            DataField       =   "G"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txt_R 
            DataField       =   "R"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_M 
         DataField       =   "M"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1830
         Width           =   1095
      End
      Begin VB.TextBox txt_C 
         DataField       =   "C"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txt_Y 
         DataField       =   "Y"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   2190
         Width           =   1095
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Convert Recordset"
         Height          =   285
         Left            =   3150
         TabIndex        =   1
         Top             =   2380
         Width           =   1500
      End
      Begin VB.Frame fraCMYK 
         Caption         =   "CMYK"
         Height          =   1935
         Left            =   240
         TabIndex        =   70
         Top             =   1110
         Width           =   1335
         Begin VB.TextBox txt_K 
            DataField       =   "K"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         Caption         =   $"frm_Main.frx":04B2
         Height          =   510
         Left            =   225
         TabIndex        =   193
         Top             =   5760
         Width           =   5910
      End
      Begin VB.Label lblPantone 
         Height          =   285
         Left            =   1665
         TabIndex        =   191
         Top             =   630
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "Pantone Color:"
         Height          =   240
         Left            =   270
         TabIndex        =   190
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label lblPantoneColor 
         Caption         =   "Pantone Color Search:"
         Height          =   240
         Left            =   315
         TabIndex        =   188
         Top             =   3375
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I wrote this for a company who needed to know the hex and rgb values
'for the Pantone colors.  After about a week of research I came up with
'this code.  The CMYK values in the database actually came from Pantone.
'It seems pretty simple and is actually pretty close if you
'compare it with a Pantone color pallette.  Pantone already has the info
'available for the CMYK conversion.  So all that needed to be done was to
'convert it to RGB then HEX.

'You can modify this code only if asked via email at jesse@sageglobal.net.
'This code can not be sold in any way form or fashion.  You have agreed to
'these conditions if you are reading this.

Option Explicit
Public rs As New ADODB.Recordset 'holds variable for recordset connection



Private Sub cboPantone_Click()
    '************************************
    'This sub finds items selected in combo
    '************************************
    
    On Error Resume Next
    
    rs.MoveFirst
    rs.Find "PANTONENAM= '" & cboPantone & "'"
    Call LoadFields
    Call SetColor
    
End Sub

Private Sub cmdConvert_Click()
    '************************************
    'This sub loops through and converts all
    '************************************
    
    On Error GoTo ErrorHandler
    
    MousePointer = vbHourglass
    
    '************************************
    'loop through all data in database
    '************************************
    rs.MoveFirst
    Do Until rs.EOF
        Call LoadFields
        Call CMYKtoRGB
        Call RGBtoHex
        Call UpdateFields
        rs.MoveNext
    Loop
    
    rs.MoveFirst
    
    MousePointer = vbDefault
    
Exit Sub
ErrorHandler:
    MsgBox "Error looping through recordset!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

          
Sub CMYKtoRGB()
    '************************************
    'This sub converts CMYK to RGB
    '************************************
    Dim int_Red As Integer, int_Green As Integer, int_Blue As Integer, strRGB As String
    Dim int_C As Integer, int_M As Integer, int_Y As Integer, int_K As Integer
    Dim int_Kfactor
    Dim yPercent As Double
    
    On Error Resume Next
    
    '************************************
    'convert CMYK to an 8-bit value
    '************************************
    int_C = txt_C * 2.55
    int_M = txt_M * 2.55
    int_Y = txt_Y * 2.55
    int_K = txt_K * 2.55
        
    '************************************
    'convert 8 bit cmyk to rgb
    '************************************
    int_Red = (255 - int_C - int_K)
    int_Green = (255 - int_M - int_K)
    int_Blue = (255 - int_Y - int_K)

    '************************************
    'convert - values to + values if -
    '************************************
    If int_Red < 0 Then txt_R = int_Red / -1 Else txt_R = int_Red
    If int_Green < 0 Then txt_G = int_Green / -1 Else txt_G = int_Green
    If int_Blue < 0 Then txt_B = int_Blue / -1 Else txt_B = int_Blue
    

End Sub

Private Sub cmdConvertCurrent_Click()
    '************************************
    'This sub only the current Pantone
    '************************************
        
    On Error Resume Next
    
    Call CMYKtoRGB
    Call RGBtoHex
    Call UpdateFields
    Call SetColor
        
End Sub

Private Sub cmdDisplay_Click()
    '************************************
    'This sub displays all colors in tabs
    '************************************
    Dim int_Count As Integer, i As Integer, x As Integer, strHex
    
    rs.MoveFirst
    On Error Resume Next
    
    '************************************
    'assign each value to picture box
    '************************************
    For i = int_Count - 1 To int_Count + 166
        strHex = rs.Fields("HEX")
        Picture1(x).BackColor = strHex
        x = x + 1
        rs.MoveNext
    Next

Exit Sub
ErrorHandler:
    MsgBox "Error moving in recordset!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub


Sub RGBtoHex()
    '************************************
    'This sub converts RGB to Hex
    '************************************
    Dim tmp_R, tmp_G, tmp_B, tmp_Hex
        
    On Error Resume Next
    
    '************************************
    'Convert RGB to Hex Values
    '************************************
    tmp_R = "0" & Hex(txt_R)
    tmp_G = "0" & Hex(txt_G)
    tmp_B = "0" & Hex(txt_B)
    
    tmp_R = Right(tmp_R, 2)
    tmp_G = Right(tmp_G, 2)
    tmp_B = Right(tmp_B, 2)
       
    txt_Hex = "&H" & tmp_B & tmp_G & tmp_R

End Sub

Private Sub cmdFirst_Click()
    '************************************
    'This sub moves rs to first record
    '************************************
    
    On Error Resume Next
    
    rs.MoveFirst
    Call LoadFields
    Call SetColor
    
End Sub

Private Sub cmdLast_Click()
    '************************************
    'This sub moves rs to last record
    '************************************
    
    On Error Resume Next
    
    rs.MoveLast
    Call LoadFields
    Call SetColor
    
End Sub

Private Sub cmdNext_Click()
    '************************************
    'This sub moves rs to previous record
    '************************************
    
    On Error GoTo ErrorHandler
    
    rs.MoveNext
    Call LoadFields
    Call SetColor
    
Exit Sub
ErrorHandler:
    MsgBox "You are at the end of this recordset!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub cmdPrevious_Click()
    '************************************
    'This sub moves rs to previous record
    '************************************
    
    On Error GoTo ErrorHandler
    
    rs.MovePrevious
    Call LoadFields
    Call SetColor
    
Exit Sub
ErrorHandler:
    MsgBox "You are at the begining of this recordset!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    '************************************
    'This sub creates rs connection
    '************************************
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    '************************************
    'Establish connection to data source
    '************************************
    sql = "SELECT * FROM tblPantoneConvTable"
    rs.Open sql, gConn, adOpenKeyset, adLockOptimistic
    
    Call LoadPantoneCbo
    Call LoadFields
    Call SetColor
    
Exit Sub
ErrorHandler:
    MsgBox "Error connecting to data source!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub SetColor()
    '************************************
    'This sub set color for RGB
    '************************************
    
    On Error Resume Next
    
    txtRGB.BackColor = txt_Hex
    txtRGB.Refresh

End Sub

Private Sub UpdateFields()
    '************************************
    'update fields with value in texbox
    '************************************
    
    On Error Resume Next
    
    'I put the "" in there in case it's null. Just a safe guard.
    'You can avoid this if your db is set up to allow null values. Just FYI
    rs("R") = "" & txt_R
    rs("G") = "" & txt_G
    rs("B") = "" & txt_B
    rs("HEX") = "" & txt_Hex
    rs.Update

End Sub

Private Sub LoadFields()
    '************************************
    'Load fields with values in rs
    '************************************
    
    On Error Resume Next
    
    txt_C = "" & rs("C")
    txt_M = "" & rs("M")
    txt_Y = "" & rs("Y")
    txt_K = "" & rs("K")
    txt_R = "" & rs("R")
    txt_G = "" & rs("G")
    txt_B = "" & rs("B")
    txt_ID = "" & rs("ID")
    txt_Hex = "" & rs("HEX")
    lblPantone.Caption = "" & rs("PantoneNam")

End Sub

Public Sub LoadPantoneCbo()
    '************************************
    'This sub loads Pantone names in combo
    '************************************
    Dim rs1 As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    'establish recordset connection
    sql = "Select PANTONENAM from tblPantoneConvTable"
    rs1.Open sql, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText

    'populate combo box
    Do Until rs1.EOF
        cboPantone.AddItem rs1("Pantonenam")
        rs1.MoveNext
    Loop
    
    'release rs
    rs1.Close
    Set rs1 = Nothing

Exit Sub
ErrorHandler:
    MsgBox "Error loading Pantone combo box!" & vbNewLine & Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '************************************
    'This closes the rs and db connection
    '************************************
    
    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    
    gConn.Close
    Set gConn = Nothing
    
    
End Sub

