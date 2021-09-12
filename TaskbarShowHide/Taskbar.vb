Option Explicit On
Imports System.Runtime.InteropServices

Public Class Taskbar
    Private Declare Auto Function FindWindow Lib "user32.dll" (ByVal strClassName As String, ByVal strWindowName As String) As IntPtr

    Private Declare Auto Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As UInt32, ByRef pData As APPBARDATA) As UInt32

    Public Enum AppBarMessages
        [New] = &H0
        Remove = &H1
        QueryPos = &H2
        SetPos = &H3
        GetState = &H4
        GetTaskBarPos = &H5
        Activate = &H6
        GetAutoHideBar = &H7
        SetAutoHideBar = &H8
        WindowPosChanged = &H9
        SetState = &HA
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure APPBARDATA
        Public cbSize As UInt32
        Public hWnd As IntPtr
        Public uCallbackMessage As UInt32
        Public uEdge As UInt32
        Public rc As Rectangle
        Public lParam As Int32
    End Structure

    Public Enum AppBarStates
        AutoHide = &H1
        AlwaysOnTop = &H2
    End Enum

    Public Sub SetTaskbarState(ByVal [option] As AppBarStates)
        Dim msgData As APPBARDATA = New APPBARDATA()
        msgData.cbSize = CType(Marshal.SizeOf(msgData), UInt32)
        msgData.hWnd = FindWindow("System_TrayWnd", Nothing)
        msgData.lParam = CType(([option]), Int32)
        SHAppBarMessage(CType(AppBarMessages.SetState, UInt32), msgData)
    End Sub

    Public Function GetTaskbarState() As AppBarStates
        Dim msgData As APPBARDATA = New APPBARDATA()
        msgData.cbSize = CType(Marshal.SizeOf(msgData), UInt32)
        msgData.hWnd = FindWindow("System_TrayWnd", Nothing)
        Return CType(SHAppBarMessage(CType(AppBarMessages.GetState, UInt32), msgData), AppBarStates)
    End Function
End Class
