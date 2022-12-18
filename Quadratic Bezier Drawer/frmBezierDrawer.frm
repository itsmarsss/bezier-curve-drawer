VERSION 5.00
Begin VB.Form frmBezierDrawer 
   Caption         =   "Quadratic Bezier Curve"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   Icon            =   "frmBezierDrawer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   909
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   11040
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdDrawSel 
      Caption         =   "Draw Sel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   11040
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   315
      Left            =   11040
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CheckBox chkGds 
      Caption         =   "Guides"
      Height          =   255
      Left            =   11040
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdRmv 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   315
      Left            =   11040
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ListBox lstCurves 
      Height          =   1815
      Left            =   11040
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox picBezier 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   727
      TabIndex        =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label lblCoords 
      Caption         =   "(0, 0)"
      Height          =   255
      Left            =   11040
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
End
Attribute VB_Name = "frmBezierDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cs As Integer
Public bez As New Collection
Public self As Boolean
Public sel As Integer
Public temp As String
Public c1 As Integer
Public v1 As Integer
Public c2 As Integer
Public v2 As Integer
Public c3 As Integer
Public v3 As Integer

Private Sub picBezier_Click()
    If cs = 3 Then
        Dim t As Double
        For t = 0 To 1 Step 0.001
            Dim X As Integer
            Dim Y As Integer
            
            X = (1 - t) * (1 - t) * (c1) + (2) * (1 - t) * (t) * (c2) + (t) * (t) * (c3)
            Y = (1 - t) * (1 - t) * (v1) + (2) * (1 - t) * (t) * (v2) + (t) * (t) * (v3)
            
            picBezier.PSet (X, Y), vbBlack
        Next t
        cs = 0
    End If
End Sub

Private Sub picBezier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempS As String
    If cs = 0 Then
        c1 = X
        v1 = Y
        
        temp = X & ":" & Y
    ElseIf cs = 1 Then
        c2 = X
        v2 = Y
        
        temp = temp & "-" & X & ":" & Y
        If chkGds.Value = 1 Then
            picBezier.Line (c1, v1)-(X, Y), vbRed
        End If
    ElseIf cs = 2 Then
        c3 = X
        v3 = Y
        
        temp = temp & "-" & X & ":" & Y
        
        bez.Add (temp)
        
        lstCurves.AddItem bez(bez.Count)
        
        If chkGds.Value = 1 Then
            picBezier.Line (c1, v1)-(c2, v2), vbRed
            picBezier.Line (c2, v2)-(c3, v3), vbRed
        End If
    End If
    If chkGds.Value = 1 Then
        picBezier.Circle (X, Y), 5, vbBlue
    End If
    
    cs = cs + 1
End Sub

Private Sub chkGds_Click()
    cmdDraw.Value = 1
End Sub

Private Sub cmdClear_Click()
    If self = True Then
        picBezier.Cls
        self = False
    Else
        Dim res As String
        res = MsgBox("Clear list aswell?", vbYesNoCancel + vbQuestion, "Clear Bezier Curves")
    
        If res = 6 Then
            picBezier.Cls
            lstCurves.Clear

            Dim i As Integer
            For i = bez.Count - 1 To 1 Step 1
                bez.Remove (1)
            Next i
            cmdRmv.Enabled = False
            cmdDrawSel.Enabled = False
        End If
        If res = 7 Then
            picBezier.Cls
        End If
    End If
End Sub

Private Sub cmdDraw_Click()
    self = True
    cmdClear.Value = 1
    
    Dim i As Integer
    
    For i = 1 To bez.Count Step 1
        Dim raw() As String
        raw = Split(bez(i), "-")
    
        Dim a() As String
        Dim b() As String
        Dim c() As String
        a = Split(raw(0), ":")
        b = Split(raw(1), ":")
        c = Split(raw(2), ":")
        
        Dim c1l As Integer
        Dim v1l As Integer
        Dim c2l As Integer
        Dim v2l As Integer
        Dim c3l As Integer
        Dim v3l As Integer
        c1l = Val(a(0))
        v1l = Val(a(1))
        c2l = Val(b(0))
        v2l = Val(b(1))
        c3l = Val(c(0))
        v3l = Val(c(1))
        
        If chkGds.Value = 1 Then
            picBezier.Circle (c1l, v1l), 5, vbBlue
            picBezier.Circle (c2l, v2l), 5, vbBlue
            picBezier.Circle (c3l, v3l), 5, vbBlue
            
            picBezier.Line (c1l, v1l)-(c2l, v2l), vbRed
            picBezier.Line (c2l, v2l)-(c3l, v3l), vbRed
        End If
        
        Dim t As Double
        For t = 0 To 1 Step 0.001
            Dim X As Integer
            Dim Y As Integer
            
            X = (1 - t) * (1 - t) * (c1l) + (2) * (1 - t) * (t) * (c2l) + (t) * (t) * (c3l)
            Y = (1 - t) * (1 - t) * (v1l) + (2) * (1 - t) * (t) * (v2l) + (t) * (t) * (v3l)
            
            picBezier.PSet (X, Y), vbBlack
        Next t
    Next i
End Sub

Private Sub cmdDrawSel_Click()
    Dim raw() As String
    raw = Split(bez(sel), "-")

    Dim a() As String
    Dim b() As String
    Dim c() As String
    a = Split(raw(0), ":")
    b = Split(raw(1), ":")
    c = Split(raw(2), ":")
    
    Dim c1 As Integer
    Dim v1 As Integer
    Dim c2 As Integer
    Dim v2 As Integer
    Dim c3 As Integer
    Dim v3 As Integer
    c1 = Val(a(0))
    v1 = Val(a(1))
    c2 = Val(b(0))
    v2 = Val(b(1))
    c3 = Val(c(0))
    v3 = Val(c(1))
    
    If chkGds.Value = 1 Then
        picBezier.Circle (c1, v1), 5, vbBlue
        picBezier.Circle (c2, v2), 5, vbBlue
        picBezier.Circle (c3, v3), 5, vbBlue
        
        picBezier.Line (c1, v1)-(c2, v2), vbRed
        picBezier.Line (c2, v2)-(c3, v3), vbRed
    End If
    
    Dim t As Double
    For t = 0 To 1 Step 0.001
        Dim X As Integer
        Dim Y As Integer
        
        X = (1 - t) * (1 - t) * (c1) + (2) * (1 - t) * (t) * (c2) + (t) * (t) * (c3)
        Y = (1 - t) * (1 - t) * (v1) + (2) * (1 - t) * (t) * (v2) + (t) * (t) * (v3)
        
        picBezier.PSet (X, Y), vbBlack
    Next t
End Sub

Private Sub cmdRmv_Click()
    bez.Remove (sel)
    lstCurves.RemoveItem (sel - 1)
    sel = -1
    cmdDraw = 1
    cmdRmv.Enabled = False
    cmdDrawSel.Enabled = False
End Sub

Private Sub Form_Load()
    chkGds.Value = 1
End Sub

Private Sub Form_Resize()
    Dim w As Long
    Dim h As Long
    w = frmBezierDrawer.Width / Screen.TwipsPerPixelX
    h = frmBezierDrawer.Height / Screen.TwipsPerPixelX
    If w < 750 Then
        w = 750
        frmBezierDrawer.Width = w * Screen.TwipsPerPixelX
    End If
    
    If h < 500 Then
        h = 500
        frmBezierDrawer.Height = h * Screen.TwipsPerPixelX
    End If

    picBezier.Top = 0
    picBezier.Left = 0
    
    picBezier.Width = w - 200
    picBezier.Height = h - 40
    
    cmdClear.Left = picBezier.Width + 10
    lstCurves.Left = picBezier.Width + 10
    cmdRmv.Left = picBezier.Width + 10
    chkGds.Left = picBezier.Width + 10
    cmdDraw.Left = picBezier.Width + 10
    cmdDrawSel.Left = picBezier.Width + 10
    lblCoords.Left = picBezier.Width + 10
End Sub

Private Sub lstCurves_Click()
    sel = lstCurves.ListIndex + 1
    cmdRmv.Enabled = True
    cmdDrawSel.Enabled = True
End Sub

Private Sub picBezier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCoords.Caption = "(" & X & ", " & Y & ")"
End Sub

