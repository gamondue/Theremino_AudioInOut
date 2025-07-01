Module Module_Utils

    Public ReportString As String

    Friend Sub InitPictureboxImage(ByVal pbox As PictureBox)
        With pbox
		 If .ClientSize.Height < 1 Then Return
            If .Image Is Nothing OrElse .Image.Width <> .ClientSize.Width OrElse .Image.Height <> .ClientSize.Height Then
                .Image = New Bitmap(.ClientRectangle.Width, .ClientRectangle.Height, Imaging.PixelFormat.Format24bppRgb)
            End If
        End With
    End Sub

    ' ----------------------------------------------------------------------------
    '  COMBO
    ' ----------------------------------------------------------------------------
    Friend Sub Combo_InitFromEnum(ByVal combo As ComboBox, ByVal enumType As Type) ' use GeType(enum)
        For Each s As String In [Enum].GetNames(enumType)
            combo.Items.Add(s)
        Next
    End Sub
    Friend Sub Combo_SetIndex(ByVal combo As ComboBox, ByRef index As Int32)
        Dim n As Int32 = combo.Items.Count
        If n = 0 Then Return
        If index < 0 Then index = 0
        If index > n - 1 Then index = n - 1
        combo.SelectedIndex = index
    End Sub

    ' =======================================================================================================
    '   CULTURE INFO FOR STRING CONVERSIONS
    ' =======================================================================================================
    Friend GCI As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

    ' ----------------------------------------------------------------------------
    '  OPEN AUDIO CONTROLS
    ' ----------------------------------------------------------------------------
    Friend Sub Open_AudioOutputs()
        Shell("control mmsys.cpl,,0", AppWinStyle.NormalFocus)
    End Sub
    Friend Sub Open_AudioInputs()
        Shell("control mmsys.cpl,,1", AppWinStyle.NormalFocus)
    End Sub
    'Friend Sub Open_AudioMixer()
    '    Dim mixerpath As String
    '    mixerpath = "C:\WINDOWS\system32\sndvol32.exe"
    '    If Not FileExists(mixerpath) Then
    '        mixerpath = "C:\WINDOWS\system32\sndvol.exe"
    '    End If
    '    If FileExists(mixerpath) Then
    '        Shell(mixerpath, AppWinStyle.NormalFocus)
    '    End If
    '    'Shell("C:\WINDOWS\system32\sndvol32.exe /rec", AppWinStyle.NormalFocus)
    '    'Shell("C:\WINDOWS\system32\sndvol.exe /rec", AppWinStyle.NormalFocus)
    'End Sub

End Module
