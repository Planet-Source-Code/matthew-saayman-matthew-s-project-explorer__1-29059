VERSION 5.00
Begin VB.UserControl MSplitter 
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   MouseIcon       =   "MSplitter.ctx":0000
   MousePointer    =   99  'Custom
   PropertyPages   =   "MSplitter.ctx":0152
   ScaleHeight     =   4215
   ScaleWidth      =   705
   ToolboxBitmap   =   "MSplitter.ctx":0163
End
Attribute VB_Name = "MSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SplitterWidth As Integer

Private m_MinSize1 As Integer
Private m_MinSize2 As Integer

Private m_bHResize As Boolean

Private m_oldW As Integer
Private m_oldH As Integer
Private m_DeltaH As Integer
Private m_SavedPos As Integer

Private m_LeftObj As Object
Private m_RightObj As Object

Private m_ControlName1 As String
Private m_ControlName2 As String

Private m_ControlVisible1 As Boolean
Private m_ControlVisible2 As Boolean

Private m_bMoving As Boolean

Public Event Moving()

Public Property Get ControlVisible1() As Boolean
    ControlVisible1 = m_ControlVisible1
End Property
Public Property Let ControlVisible1(vData As Boolean)
    On Error Resume Next
    m_ControlVisible1 = vData
    If vData = True Then
        m_LeftObj.Visible = True
        Extender.Left = m_SavedPos
    Else
        m_LeftObj.Visible = False
    End If
    m_bHResize = False
    Resize
    RaiseEvent Moving
End Property
Public Property Get ControlVisible2() As Boolean
    ControlVisible2 = m_ControlVisible2
End Property
Public Property Let ControlVisible2(vData As Boolean)
    On Error Resume Next
    m_ControlVisible2 = vData
    If vData = True Then
        m_RightObj.Visible = True
        Extender.Left = m_SavedPos
    Else
        m_RightObj.Visible = False
    End If
    m_bHResize = False
    Resize
    RaiseEvent Moving
End Property

Public Property Get SplitterWidth() As Integer
    SplitterWidth = m_SplitterWidth
End Property
Public Property Let SplitterWidth(vData As Integer)
    m_SplitterWidth = vData
    Extender.Width = vData
End Property

Private Sub InitCtrl1()
    Dim ctrl As Object
    
    For Each ctrl In ParentControls
        If ctrl.Name = m_ControlName1 Then
            Set m_LeftObj = ctrl
            Exit For
        End If
    Next
End Sub
Private Sub InitCtrl2()
    Dim ctrl As Object
    
    For Each ctrl In ParentControls
        If ctrl.Name = m_ControlName2 Then
            Set m_RightObj = ctrl
            Exit For
        End If
    Next
End Sub

Public Property Let ControlName1(vData As String)
    m_ControlName1 = vData

    InitCtrl1
    
    SizeObjects
    
    PropertyChanged "ControlName1"
End Property
Public Property Get ControlName1() As String
    ControlName1 = m_ControlName1
End Property
Public Property Let ControlName2(vData As String)
    m_ControlName2 = vData
    
    InitCtrl2
    SizeObjects
    
    PropertyChanged "ControlName2"
End Property
Public Property Get ControlName2() As String
    ControlName2 = m_ControlName2
End Property

Friend Function FormControls() As Object
    Set FormControls = ParentControls
End Function
Public Property Get Name() As String
    Name = UserControl.Extender.Name
End Property
Friend Function mahParent() As Object
    Set mahParent = Extender.Parent
End Function





Private Sub UserControl_Initialize()
    m_bMoving = False
    m_bHResize = True
    m_MinSize1 = 300
    m_MinSize2 = 300
    m_SplitterWidth = 75
    m_ControlVisible1 = True
    m_ControlVisible2 = True
End Sub




Public Sub Resize()
    On Error GoTo theEnd
    Dim Self As Object
    Set Self = UserControl.Extender
    If Self Is Nothing Then Exit Sub
    Dim pos As Integer
    
    If Not m_ControlVisible1 Then
        pos = -m_SplitterWidth
    End If
    If Not m_ControlVisible2 Then
        pos = Self.Parent.Width
    End If
    
    If m_ControlVisible1 And m_ControlVisible2 Then
        pos = Self.Left / m_oldW * Extender.Parent.Width
        If pos < m_MinSize1 Then
            pos = m_MinSize1
        ElseIf Self.Parent.ScaleWidth - pos - Self.Width < m_MinSize2 Then
            pos = Self.Parent.ScaleWidth - m_MinSize2 - Self.Width
        End If
        m_SavedPos = pos
    End If
    
    Self.Left = pos
    
    Dim hght As Integer
    
    If m_bHResize Then
        hght = Extender.Parent.Height - m_DeltaH
        Self.Height = IIf(hght > 0, hght, 0)
    End If
    
    m_bHResize = True
    
    
    m_oldH = Extender.Parent.Height
    m_oldW = Extender.Parent.Width
    
    SizeObjects
    
    Exit Sub
    
theEnd:
'    MsgBox "Err:" & Err.Number & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub SizeObjects(Optional X As Single = 0)
    On Error Resume Next
        
    Dim ctrlParent As Object
    Dim ctrl As Object
    
    Set ctrl = UserControl.Extender
    Set ctrlParent = UserControl.Parent
    
    If ctrl Is Nothing Then Exit Sub
    If ctrlParent Is Nothing Then Exit Sub
    
    ctrl.Left = ctrl.Left + X
    If m_LeftObj Is Nothing Then InitCtrl1
    If m_RightObj Is Nothing Then InitCtrl2
    
    If m_ControlVisible2 Then
        If Not m_RightObj Is Nothing Then
            Dim itemp As Integer
            itemp = ctrlParent.ScaleWidth - ctrl.Left - ctrl.Width
            If X <= 0 Then
                m_RightObj.Width = IIf(itemp > m_MinSize2, _
                                       itemp, m_MinSize2)
                m_RightObj.Left = ctrl.Left + ctrl.Width
            Else
                m_RightObj.Left = ctrl.Left + ctrl.Width
                m_RightObj.Width = IIf(itemp > m_MinSize2, _
                                       itemp, m_MinSize2)
            End If
            m_RightObj.Height = ctrl.Height
            m_RightObj.Top = ctrl.Top
        End If
    ElseIf m_ControlVisible1 Then
        m_LeftObj.Width = ctrlParent.ScaleWidth
        m_LeftObj.Height = ctrl.Height
    End If
    
    If m_ControlVisible1 Then
        If Not m_LeftObj Is Nothing Then
            m_LeftObj.Left = 0
            m_LeftObj.Top = ctrl.Top
            m_LeftObj.Width = ctrl.Left - m_LeftObj.Left
            m_LeftObj.Height = ctrl.Height
        End If
    ElseIf m_ControlVisible2 Then
        m_RightObj.Left = 0
        m_RightObj.Width = ctrlParent.ScaleWidth
        m_RightObj.Height = ctrl.Height
    End If
    
    RaiseEvent Moving
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_bMoving = True
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bMoving Then
        Dim pos As Integer
        pos = Extender.Left + X
        If pos < m_MinSize1 Then
            Extender.Left = m_MinSize1
            X = 0
        ElseIf Extender.Parent.ScaleWidth - pos - Extender.Width < m_MinSize2 Then
            Extender.Left = Extender.Parent.ScaleWidth - m_MinSize2 - Extender.Width
            X = 0
        End If
        m_SavedPos = Extender.Left
        SizeObjects X
        RaiseEvent Moving
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    m_bMoving = False
    SizeObjects
End Sub


Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Width = m_SplitterWidth
    m_oldW = UserControl.Parent.Width
    m_oldH = UserControl.Parent.Height
    
    m_DeltaH = Extender.Parent.Height - Extender.Height
    
    If m_LeftObj Is Nothing Or m_RightObj Is Nothing Then Exit Sub
    
    SizeObjects
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "ControlName1", m_ControlName1
    PropBag.WriteProperty "ControlName2", m_ControlName2
    PropBag.WriteProperty "SplitterWidth", m_SplitterWidth
    PropBag.WriteProperty "ControlVisible1", m_ControlVisible1
    PropBag.WriteProperty "ControlVisible2", m_ControlVisible2
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    m_ControlName1 = PropBag.ReadProperty("ControlName1")
    m_ControlName2 = PropBag.ReadProperty("ControlName2")
    m_SplitterWidth = PropBag.ReadProperty("SplitterWidth")
    m_ControlVisible1 = PropBag.ReadProperty("ControlVisible1")
    m_ControlVisible2 = PropBag.ReadProperty("ControlVisible2")
End Sub

