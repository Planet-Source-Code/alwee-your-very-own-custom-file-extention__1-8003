VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmExt 
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Ext. by:Alwee"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicLoaded 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4215
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Lbl3D 
      BackColor       =   &H80000014&
      Height          =   4215
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuInstr 
         Caption         =   "Instructions"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy To Clipboard"
      End
   End
End
Attribute VB_Name = "FrmExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declaring my variables
'Declaring alot of things as Intenger instead of Long could work
'But this is a 32 but application, and in most 32 bit applications they use long
Dim HeightOfLine, frmWidth As Long 'All the "," only make it so that I can declare them all as long at ONCE
Dim R, B As Long
Dim Filter As String

Private Sub Form_Activate()
'Nothing more than a quick formfade
'Just to make it look good, and make it look like
'I know what I'm doing
HeightOfLine = FrmExt.Height 'HeightOfLine equals the forms height
R = 255 'The value for red later on in the sub
B = 255 'The value for blue later on in the sub
    Do Until R = 0 'Keeps the sub going on over and over until red equals 0
        HeightOfLine = HeightOfLine - FrmExt.Height / 255 'this tells the lines where to end at
        FrmExt.Line (0, 0)-(FrmExt.Width, HeightOfLine), RGB(R, 0, B), BF 'Makes a box from cordinate to the other(BF)
        R = R - 1 'Substract 1, to make it go from dark to light, and to make it end the sub
    Loop
End Sub

Private Sub Form_Load()
    frmWidth = FrmExt.Width 'The variable equals the forms width
End Sub

Private Sub Form_Resize()
On Error GoTo ResizeError: 'Stops errors from ocurring(usually when minimized)

If FrmExt.Width <> frmWidth Then 'If the forms height goes up or down then
    HeightOfLine = FrmExt.Height 'And you know what this does
    R = 255
    B = 255
        Do Until R = 0
            HeightOfLine = HeightOfLine - FrmExt.Height / 255
            FrmExt.Line (0, 0)-(FrmExt.Width, HeightOfLine), RGB(R, 0, B), BF
            R = R - 1
        Loop
End If
'P.S. it's common sense that if you resize the form, that the if will happen
'And without this sub, if the form got resized the old fade would be left, and it would look ugly

'This is make it so that when a form's size is changed
'The controls Picloaded and Lbl3D get
'There sizes changed too
PicLoaded.Width = FrmExt.Width - 1200
Lbl3D.Width = FrmExt.Width - 1200
PicLoaded.Height = FrmExt.Height - 1650
Lbl3D.Height = FrmExt.Height - 1650

ResizeError:
    Exit Sub
End Sub


Private Sub mnuCopy_Click()
Clipboard.Clear 'Clears whatever is currently on the clip board
        Clipboard.SetData PicLoaded.Picture 'Copies the picture in the PicLoaded picture box
End Sub

Private Sub mnuExit_Click()
    Unload Me 'Unload the form and return system resources
    End 'End program
End Sub

Private Sub mnuInstr_Click()
    Call MsgBox("1) Just open a .jpg file and save it as .alw  2) Then try to open that .alw file  3) It should work 4) now just read the code to learn my secret", vbOKOnly + vbInformation, "HOW DO I DO IT?")
End Sub

Private Sub mnuOpen_Click()
On Error GoTo OpenError: 'Stops errors, like opening a false file path
    
    Filter = "Alwee Files (*.alw)|*.alw;|" 'This is the normal way of opening a file with your own ext.
    Filter = Filter + "JPEG Files (*.jpg)|*.jpg;|" 'Shows JPG Files(This is how you normally open a file)
    Filter = Filter + "All Formats(*.*)|*.alw;*.jpg;|" 'Show both formats at once
    
    CommonDialog1.Filter = Filter 'This is how you make the filter show in the filter section
    CommonDialog1.ShowOpen 'Show the dialog now
    CommonDialog1.FilterIndex = 1 'Makes the *.alw extention come up first as default
    
    PicLoaded = LoadPicture(CommonDialog1.filename) 'The file from the dialog is loaded in the picture box
    
OpenError:
    Exit Sub
End Sub

Private Sub mnuSaveas_Click()
On Error GoTo SaveError 'Stops errors from occuring
    
    If PicLoaded.Picture = "" Then 'If the picture box has nothin in it then
        Call MsgBox("You MUST Load a picture before you can save one", vbOKOnly + vbInformation, "ERROR") 'give the user an error telling them to load a picture first
    End If
    
    Filter = "Alwee Files (*.alw)|*.alw;|" 'This is the normal way of saveing a file with your own ext.
    Filter = Filter + "JPEG Files (*.jpg)|*.jpg;|" 'Shows JPG Files(This is how you normally save a file)
    Filter = Filter + "All Formats(*.*)|*.alw;*.jpg;|" 'Show both formats
    
    CommonDialog1.Filter = Filter 'This is how you make the filter show in the filter section
    CommonDialog1.ShowSave 'Show the dialog now
    CommonDialog1.FilterIndex = 1 'Makes the *.alw extention come up first as default
    
    Call SavePicture(PicLoaded, CommonDialog1.filename) 'The file from the dialog is loaded in the picture box

SaveError:
    Exit Sub
End Sub

