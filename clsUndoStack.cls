VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MAX_UNDO_STEPS As Integer = 24
Private mCur As Integer
Private mStackPicture(MAX_UNDO_STEPS) As StdPicture
Private mStackLgndChar(MAX_UNDO_STEPS) As String

Public Function CanRedo() As Boolean
    If mCur > UBound(mStackPicture) Then Exit Function
    CanRedo = Not mStackPicture(mCur) Is Nothing
End Function

Public Function CanUndo() As Boolean
    CanUndo = (mCur > 1)
End Function

Public Sub CreateUndoStep(p As StdPicture, Optional LgndChar As String)
Dim i As Integer
    If mCur > UBound(mStackPicture) Then
        For i = 0 To MAX_UNDO_STEPS - 1
            Set mStackPicture(i) = mStackPicture(i + 1)
            mStackLgndChar(i) = mStackLgndChar(i + 1)
        Next i
        Set mStackPicture(mCur - 1) = p
        mStackLgndChar(mCur - 1) = LgndChar
    Else
        Set mStackPicture(mCur) = p
        mStackLgndChar(mCur) = LgndChar
        mCur = mCur + 1
        For i = mCur To MAX_UNDO_STEPS
            Set mStackPicture(i) = Nothing
            mStackLgndChar(mCur) = ""
        Next i
    End If
End Sub

Public Function GetRedo(ByRef Picture As StdPicture, ByRef LgndChar As String) As Boolean
    If mCur <= UBound(mStackPicture) Then
        Set Picture = mStackPicture(mCur)
        LgndChar = mStackLgndChar(mCur)
        mCur = mCur + 1
        GetRedo = True
    End If
End Function

Public Function GetUndo(ByRef Picture As StdPicture, ByRef LgndChar As String) As Boolean
    If mCur > 1 Then
        mCur = mCur - 1
        Set Picture = mStackPicture(mCur - 1)
        LgndChar = mStackLgndChar(mCur)
        GetUndo = True
    End If
End Function





