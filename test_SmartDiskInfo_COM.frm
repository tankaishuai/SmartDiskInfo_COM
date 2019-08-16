VERSION 5.00
Object = "{D01648D4-2109-40F4-BF12-BCE8C3080E19}#1.0#0"; "EvalExprCtrl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin EVALEXPRCTRLLib.EvalExprCtrl EvalExprCtrl1 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   4260
      _StockProps     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error Resume Next
    
    Dim sd As New SmartDevice
    Dim sr As SmartResult
    Dim si As SmartItem
    
    Set sr = sd.get_drive_info("C:")
    Set si = sr.root_item()
    strtext = si.item_by_key("serial_number").to_string
    MsgBox strtext, , "serial_number of C:"
    
    Set sr = sd.get_drive_info("D:")
    Set si = sr.root_item()
    strtext = si.item_by_key("serial_number").to_string
    MsgBox strtext, , "serial_number of D:"
    
    Set sr = sd.get_drive_info("E:")
    Set si = sr.root_item()
    strtext = si.item_by_key("serial_number").to_string
    MsgBox strtext, , "serial_number of E:"
    
    Set sr = sd.get_drive_info("F:")
    Set si = sr.root_item()
    strtext = si.item_by_key("serial_number").to_string
    MsgBox strtext, , "serial_number of F:"
End Sub
