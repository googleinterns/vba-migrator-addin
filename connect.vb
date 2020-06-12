 Private Sub InitializeAddIn()
        'On initialization of add-in this button will create
        'Private WithEvents _Button As CommandBarButton declare  as a member variable of class
        Dim standardCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl

        Try

            standardCommandBar = _VBE.CommandBars.Item("Menu Bar")      'There are many command bar like standard, Menu Bar etc. So to decide in which we want to make it. 

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _Button = DirectCast(commandBarControl, CommandBarButton)
            _Button.Caption = "Sheets_Compatibility"             	 'Name of button
            _Button.Style = MsoButtonStyle.msoButtonIconAndCaption
            _Button.BeginGroup = True

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub