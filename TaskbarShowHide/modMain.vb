Module modMain

	Public Sub Main()

		Dim clsBar As New Taskbar
		Dim barState As Taskbar.AppBarStates
		barState = clsBar.GetTaskbarState()

		If (barState And clsBar.AppBarStates.AutoHide) = clsBar.AppBarStates.AutoHide Then
			clsBar.SetTaskbarState(clsBar.AppBarStates.AlwaysOnTop)
		Else
			clsBar.SetTaskbarState(clsBar.AppBarStates.AutoHide)
		End If


	End Sub


End Module
