VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDragList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Klasse zur Erstellung von DragListBoxen"
Option Explicit
'#############################################################
'# Class: clsDragList                                        #
'# Author: Jonas Wolz (jwolzvb@yahoo.de)                     #
'# Description: Creates a DragListBox (using ComCtl32.dll),  #
'#     and processes the windows messages sent.              #
'#     It will raise events allowing you to process them     #
'#     easily or it will do everything automatically.        #
'# --------------------------------------------------------- #
'#    Needs SSubTmr6.dll from vbAccelerator                  #
'#    (www.vbAccelerator.com) to implement subclassing !     #
'#############################################################

'Subclassing:
Implements ISubclass

'API-Declarations:
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Translated from C-Header files:
Private Type DRAGLISTINFO
    uNotification As Long
    hWndLB As Long
    ptCursor As POINTAPI
End Type

Private Declare Function MakeDragList Lib "Comctl32.dll" (ByVal hWndLB As Long) As Long
Private Declare Function LBItemFromPt Lib "Comctl32.dll" (ByVal hWndLB As Long, ByVal X As Long, ByVal Y As Long, ByVal bAutoScroll As Long) As Long
Private Declare Sub DrawInsert Lib "Comctl32.dll" (ByVal hWndLBParent As Long, ByVal hWndLB As Long, ByVal nItem As Long)
Private Const DL_BEGINDRAG = 1157
Private Const DL_CANCELDRAG = 1160
Private Const DL_DRAGGING = 1158
Private Const DL_DROPPED = 1159
Private Const DRAGLISTMSGSTRING = "commctrl_DragListMsg"

'From API-Viewer:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal cBytes As Long)
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

'For Auto Mode:
Private Const LB_INSERTSTRING = &H181
Private Const LB_DELETESTRING = &H182
Private Const LB_ERR = (-1)
Private Const LB_GETITEMDATA = &H199
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_SETITEMDATA = &H19A
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


'Private variables:
Private m_hWndLB As Long, m_hWndLBParent As Long
Private m_DragListMsg As Long
Private m_LastIdx As Long
Private m_OldMousePointer As Long, m_OldMouseIcon As IPictureDisp

'Properties:
Public DrawInsertIcon As Boolean
Attribute DrawInsertIcon.VB_VarDescription = "Gibt an, ob das Icon, das den Einfügepunkt markiert, gezeichnet werden soll."
Public AutoScroll As Boolean
Attribute AutoScroll.VB_VarDescription = "Legt fest, ob die ListBox beim  ""Draggen"" automatisch scrollen soll."
'Sets the cursor to MouseIcon (or to vbArrow if is Nothing)
' if the Mouse is over the ListBox, and to vbNoDrop otherwise
Public AutoCursor As Boolean
'Exchanges the items and the ItemData automatically when finished
Public AutoItems As Boolean

'Events:

'Cancel can be set to True to cancel the Drag&Drop
Event BeginDrag(ByVal nItemIndex As Long, ByRef Cancel As Boolean)
Attribute BeginDrag.VB_Description = "Wird ausgelöst, wenn der Benutzer Drag&Drop beginnt."

'Change the MousePointer, if needed + draw a user defined insert icon in this event
' This event is raised very often during Drag&Drop, even if the list item hasn't changed,
'  so change the Cursor only if needed, not every time when the event is raised
Event Dragging(ByVal nItemDragging As Long, ByVal nTargetItem As Long, ByRef ChangeCursor As dlChangeCursor, ByVal PixelX As Long, ByVal PixelY As Long)
Attribute Dragging.VB_Description = "Wird während des Drag&Drop in regelmäßigen Abständen ausgelöst."

Event CancelDrag()
Attribute CancelDrag.VB_Description = "Wird ausgelöst wenn das Drag&Drop abgebrochen wird."

'IMPORTANT: You'll have to put code to exchange the list items
'    into this event if AutoItems = False !
Event Dropped(ByVal nOldIdx As Long, ByVal nTargetIdx As Long)
Attribute Dropped.VB_Description = "Wird ausgelöst, wenn das Drag&Drop erfolgreich beendet wird."

'Constants for the cursor:
Enum dlChangeCursor
    NoChange = 0
    DL_MOVECURSOR = 3
    DL_COPYCURSOR = 2
    DL_STOPCURSOR = 1
End Enum

Public Sub ClearInsertIcon()
Attribute ClearInsertIcon.VB_Description = """Löscht"" das Icon, das den Einfügepunkt markiert."
    If m_hWndLBParent <> 0 And m_hWndLB <> 0 Then
        'Clear icon:
        DrawInsert m_hWndLBParent, m_hWndLB, -1
    End If
End Sub


Public Property Get hWndListBox() As Long
Attribute hWndListBox.VB_Description = "Legt das Handle der ListBox fest, die in eine DragListBox umgewandelt werden soll, fest.\r\nKann nur einmal gesetzt werden !"
    hWndListBox = m_hWndLB
End Property
Public Property Let hWndListBox(ByVal NewVal As Long)
    Static NotFirstTime As Boolean
    'Can only be set once, because MakeDragListBox()
    ' cannot be "undone". So you could encounter
    ' problems with the messages if you did allow
    ' setting this property multiple times
    If NotFirstTime Then
        Err.Raise vbObjectError + 100, "DragList", "hWndListBox can only be set one time !"
    End If
    If IsWindow(NewVal) Then 'Valid window ?
        m_hWndLB = NewVal
        If Not pCreate Then Err.Raise vbObjectError + 101, "DragList", "ListBox couldn't be changed into a DragListBox !"
        NotFirstTime = True
    Else
        Err.Raise 380
    End If
End Property
Public Property Get LastDraggedItemIndex() As Long
Attribute LastDraggedItemIndex.VB_Description = "Gibt den ListIndex des zuletzt gezogenen Listenpunkts zurück."
    LastDraggedItemIndex = m_LastIdx
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = m_OldMouseIcon
End Property

Public Property Set MouseIcon(NewMI As IPictureDisp)
    If Not (m_OldMouseIcon Is NewMI) Then
        Set m_OldMouseIcon = NewMI
        Set Screen.MouseIcon = NewMI
    End If
End Property

Private Function pCreate() As Boolean
    pCreate = False
    'Create DragList
    If (MakeDragList(m_hWndLB) = 0) Then Exit Function
    'Register DragList-Message and trap it
    m_DragListMsg = RegisterWindowMessage(DRAGLISTMSGSTRING)
    m_hWndLBParent = GetParent(m_hWndLB)
    AttachMessage Me, m_hWndLBParent, m_DragListMsg
    pCreate = True
End Function


Private Sub pDetach()
    'End subclassing:
    DetachMessage Me, m_hWndLBParent, m_DragListMsg
    'Remove property:
    RemoveProp m_hWndLB, "clsDL_WPResponse"
End Sub

Public Property Let MousePointer(ByVal NewMP As VBRUN.MousePointerConstants)
    If m_OldMousePointer <> NewMP Then
        m_OldMousePointer = NewMP
        Screen.MousePointer = NewMP
    End If
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
    MousePointer = m_OldMousePointer
End Property
Private Sub pDoAutoItems(ByVal nOldIdx As Long, ByVal nTargetIdx As Long)
    Dim strOld As String, lOldData As Long
    Dim lNewIdx As Long
    If Not AutoItems Then Exit Sub
    If nTargetIdx < 0 Then Exit Sub
    If nTargetIdx = nOldIdx Then Exit Sub
    
    If nOldIdx > nTargetIdx Then
        lNewIdx = nTargetIdx
    Else
        'If you remove a item, the indexes will change:
        lNewIdx = nTargetIdx - 1
    End If
    'Allocate Buffer:
    strOld = Space$(SendMessage(m_hWndLB, LB_GETTEXTLEN, nOldIdx, ByVal 0&) + 1)
    'Get the Text:
    SendMessage m_hWndLB, LB_GETTEXT, nOldIdx, ByVal strOld
    'Get the old ItemData:
    lOldData = SendMessage(m_hWndLB, LB_GETITEMDATA, nOldIdx, ByVal 0&)
    'Remove the old item:
    SendMessage m_hWndLB, LB_DELETESTRING, nOldIdx, ByVal 0&
    'Insert the item at the new position:
    SendMessage m_hWndLB, LB_INSERTSTRING, lNewIdx, ByVal strOld
    'Set the ItemData:
    SendMessage m_hWndLB, LB_SETITEMDATA, lNewIdx, ByVal lOldData
End Sub

Private Sub Class_Initialize()
    DrawInsertIcon = True
    AutoScroll = True
    AutoCursor = True
    AutoItems = True
    m_OldMousePointer = vbDefault
End Sub


Private Sub Class_Terminate()
    pDetach
    Set m_OldMouseIcon = Nothing
End Sub


Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    'Not used
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property


Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If iMsg = m_DragListMsg Then
        Dim DLI As DRAGLISTINFO
        Dim MyCancel As Boolean, lngLI As Long
        Dim Cursor As dlChangeCursor
        'Copy the struct pointed to by lParam into a local one to read its data
        CopyMemory DLI, ByVal lParam, Len(DLI)
        If DLI.hWndLB = m_hWndLB Then 'Is it our ListBox ?
            Select Case DLI.uNotification
            Case DL_BEGINDRAG 'Drag&Drop started
                MyCancel = False
                'Get the Index:
                lngLI = LBItemFromPt(DLI.hWndLB, DLI.ptCursor.X, DLI.ptCursor.Y, 0)
                RaiseEvent BeginDrag(lngLI, MyCancel)
                If MyCancel Then 'The class's user has selected to cancel Drag&Drop
                    ISubclass_WindowProc = 0&
                Else
                    m_LastIdx = lngLI
                    ISubclass_WindowProc = 1&
                End If
                If AutoCursor Then
                    If m_OldMouseIcon Is Nothing Then
                        MousePointer = vbArrow
                    Else
                        MousePointer = vbCustom
                    End If
                End If
            Case DL_DRAGGING 'Drag&Drop is in progress
                lngLI = LBItemFromPt(m_hWndLB, DLI.ptCursor.X, DLI.ptCursor.Y, CLng(AutoScroll))
                Cursor = NoChange
                RaiseEvent Dragging(m_LastIdx, lngLI, Cursor, DLI.ptCursor.X, DLI.ptCursor.Y)
                ISubclass_WindowProc = Cursor
                If DrawInsertIcon Then 'Draw icon if wanted
                    DrawInsert m_hWndLBParent, m_hWndLB, lngLI
                End If
                If AutoCursor Then
                    If lngLI < 0 Then
                        MousePointer = vbNoDrop
                    Else
                        If m_OldMouseIcon Is Nothing Then
                            MousePointer = vbArrow
                        Else
                            MousePointer = vbCustom
                        End If
                    End If
                End If
            Case DL_CANCELDRAG 'Drag&Drop cancelled
                RaiseEvent CancelDrag
                'Clear icon
                DrawInsert m_hWndLBParent, m_hWndLB, -1
                If AutoCursor Then
                    MousePointer = vbDefault
                End If
            Case DL_DROPPED 'Drag&Drop ended successfully
                lngLI = LBItemFromPt(m_hWndLB, DLI.ptCursor.X, DLI.ptCursor.Y, 0&)
                RaiseEvent Dropped(m_LastIdx, lngLI)
                'Clear icon
                DrawInsert m_hWndLBParent, m_hWndLB, -1
                If AutoCursor Then
                    MousePointer = vbDefault
                End If
                If AutoItems Then pDoAutoItems m_LastIdx, lngLI
            End Select
            'For several classes:
            SetProp hwnd, "clsDL_WPResponse", ISubclass_WindowProc
        Else
            'For several classes:
            ISubclass_WindowProc = GetProp(hwnd, "clsDL_WPResponse")
        End If
    End If
End Function


