VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DragListBox-sample"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkNoDrag 
      Caption         =   "Don't allow Drag&&Drop (for non-auto mode only)"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ListBox lstNoAuto 
      Height          =   3375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstAuto 
      Height          =   3375
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Look into Form_Load to find out how to initialize the classes."
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Auto ListBox:"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   """Manual"" ListBox:"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Drag some items from the ListBoxes at the right !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "DragList Box- Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' Sample for clsDragList
' -------------------------------------------------------------
' This sample shows how to use and initialize clsDragList
' both in auto mode and in non-auto mode (you usually won't need
' the latter mode !)
' ==============================================================

'Auto DragList, you'll only need to initialize
'it in Form_Load
Private DLAuto As clsDragList

'Non-auto mode class: Declared with WithEvents because
' you'll need to trap the events raised when using this mode
Private WithEvents DLNoAuto As clsDragList
Attribute DLNoAuto.VB_VarHelpID = -1



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DLNoAuto_BeginDrag(ByVal nItemIndex As Long, Cancel As Boolean)
    'Deactivate Drag&Drop if selected: (code in the Dropped event needed, too)
    If chkNoDrag.Value = vbChecked Then
        Cancel = True
        Exit Sub
    End If
    'Use the class's MousePointer property:
    DLNoAuto.MousePointer = vbCustom
End Sub

Private Sub DLNoAuto_CancelDrag()
    DLNoAuto.MousePointer = vbDefault
End Sub

Private Sub DLNoAuto_Dragging(ByVal nItemDragging As Long, ByVal nTargetItem As Long, ChangeCursor As dlChangeCursor, ByVal PixelX As Long, ByVal PixelY As Long)
    'Use the DragBox's MousePointer property here !!
    'This property ensures that the MousePointer is only
    'set when needed (avoids flickering of the pointer)
    If nTargetItem < 0 Then
        DLNoAuto.MousePointer = vbNoDrop
    End If
    If nTargetItem >= 0 Then
        DLNoAuto.MousePointer = vbCustom
    End If
End Sub

Private Sub DLNoAuto_Dropped(ByVal nOldIdx As Long, ByVal nTargetIdx As Long)
    Dim strText As String, lDelta As Long
    'Reset the MousePointer first
    DLNoAuto.MousePointer = vbDefault
    'Dropped event sometimes is also raised when Drag&Drop was cancelled !?
    If chkNoDrag.Value = vbChecked Then Exit Sub
    If nTargetIdx < 0 Or nTargetIdx = nOldIdx Then Exit Sub '-> Cancelled or at the same position as before
    '------------------------------------------------------------------------------------------
    'Move the item:
    '------------------------------------------------------------------------------------------
    ' Insert code to move ItemData, etc., too, if needed
    If nTargetIdx < nOldIdx Then 'Because the Index changes when you use RemoveItem
        lDelta = 0
    Else
        lDelta = 1
    End If
    strText = lstNoAuto.List(DLNoAuto.LastDraggedItemIndex) 'Save old text
    lstNoAuto.RemoveItem nOldIdx 'Remove old item
    'Insert item in its new place.
    '  Pay attention on the change of the indexes caused by RemoveItem (->lDelta)
    lstNoAuto.AddItem strText, nTargetIdx - lDelta

End Sub

Private Sub Form_Load()
    Dim L As Long
    'Add some sample items to the ListBoxes:
    For L = 1 To 20
        lstNoAuto.AddItem "Item " & CStr(L)
        lstAuto.AddItem "Item " & CStr(L)
    Next
    
    'Init the auto class:
    'Create it:
    Set DLAuto = New clsDragList
    'For Auto Mode, you'll just have to set
    ' the hWnd of the ListBox (it's usually good to
    ' set the MouseIcon property to a nice icon, too)
    DLAuto.hWndListBox = lstAuto.hwnd
    'Set the drag icon:
    Set DLAuto.MouseIcon = LoadResPicture("Drag", vbResCursor)
    ' Done for auto mode !
    
    'Init the "non-Auto" class:
    Set DLNoAuto = New clsDragList
    DLNoAuto.hWndListBox = lstNoAuto.hwnd
    DLNoAuto.AutoCursor = False
    DLNoAuto.AutoItems = False
    'Load the icon:
    Set DLNoAuto.MouseIcon = LoadResPicture("Drag", vbResCursor)
    
End Sub


