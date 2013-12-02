Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class mTextBox
    Inherits TextBox
    Public Const ECM_FIRST = &H1500
    Public Const EM_SETCUEBANNER = ECM_FIRST + 1
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As String) As IntPtr
    End Function
    Private _CueText As String
    Private _CueFocus As Boolean
    Public Property CueText() As String
        Get
            Return _CueText
        End Get
        Set(ByVal value As String)
            _CueText = value
            If Me.Handle <> IntPtr.Zero Then SendMessage(Me.Handle, EM_SETCUEBANNER, New IntPtr(CInt(_CueFocus)), _CueText)
        End Set
    End Property
    Property CueTextFocus() As Boolean
        Get
            Return _CueFocus
        End Get
        Set(ByVal value As Boolean)
            _CueFocus = value
            CueText = CueText
        End Set
    End Property
End Class
