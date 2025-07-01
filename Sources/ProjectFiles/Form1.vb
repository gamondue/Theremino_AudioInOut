
Public Class Form1

    Private Slot_Spectrum As Int32

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        EventsAreEnabled = False
        Text = AppTitleAndVersion("")
        ' --------------------------------------------------------
        Combo_InitFromEnum(cmb_Gen1_Wave, GetType(WavePlayer.WaveTypes))
        Combo_InitFromEnum(cmb_Gen2_Wave, GetType(WavePlayer.WaveTypes))
        Combo_InitFromEnum(cmb_Gen3_Wave, GetType(WavePlayer.WaveTypes))
        Combo_InitFromEnum(cmb_Spec_Window, GetType(FHT.WindowTypes))
        cmb_Spec_Window.SelectedIndex = 15
        ' --------------------------------------------------------
        cmb_Gen1_Wave.SelectedIndex = 0
        cmb_Gen2_Wave.SelectedIndex = 0
        cmb_Gen3_Wave.SelectedIndex = 0
        Load_INI()
        ' -------------------------------------------------------- Init Audio InOut Combo
        FillAudioInDevicesCombo()
        FillAudioOutDevicesCombo()
        ' --------------------------------------------------------
        WaveGenTrim()
        WaveGenSetWaveTypes()
        WaveGenerators.StartStop(SelectedAudioOut, _
                                 chk_Generator1.Checked, _
                                 chk_Generator2.Checked, _
                                 chk_Generator3.Checked)
        ' --------------------------------------------------------
        Scope.Initialize(pbox_Scope, SelectedAudioIn)
        ScopeTrim()
        Spectrum.Initialize(pbox_Spectrum, SelectedAudioIn)
        SpectrumBands.Initialize(pbox_SpectrumBars, SelectedAudioIn)
        SetSpectrumWindowType()
        ' -------------------------------------------------------- EXTREME test Rec and Play 
        'RecPlayer.RecPlayStart(SelectedAudioOut, SelectedAudioIn)
        ' -------------------------------------------------------- EXTREME test Multigenerator 
        'MultigenTest.Initialize(SelectedAudioOut)
        ' -------------------------------------------------------- TIMER START
        Timer1.Interval = 20
        Timer1.Start()
        ' -------------------------------------------------------- CPU Affinity test
        'Dim Proc As Process = Process.GetCurrentProcess()
        'For i As Int32 = 0 To Proc.Threads.Count
        '    Dim Thread As ProcessThread = Proc.Threads(0)
        '    Dim AffinityMask As Int32 = 2  ' use only flagged processors, despite availability
        '    Thread.ProcessorAffinity = CType(AffinityMask, IntPtr)
        'Next
        ' -------------------------------------------------------- 
        EventsAreEnabled = True
        Refresh()
        Opacity = 1
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MultigenTest.CloseAll()
        RecPlayer.RecPlayStop()
        WaveGenerators.CloseAll()
        Scope.CloseAll()
        SpectrumBands.CloseAll()
        Spectrum.CloseAll()
        Save_INI()
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Not EventsAreEnabled Then Return
        UpdateAll()
    End Sub

    Private Sub Form_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        If Not EventsAreEnabled Then Exit Sub
        LimitFormPosition(Me)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Not EventsAreEnabled Then Return
        UpdateAll()
    End Sub

    Private SpectrumWindowType As FHT.WindowTypes
    Private Sub UpdateAll()
        'ReportString = ""
        ' ----------------------------------------------------- EXTREME test for repeated stop/start
        'WaveGenerators.StartStop(False, False, False)
        'WaveGenerators.StartStop(chk_Generator1.Checked, chk_Generator2.Checked, chk_Generator3.Checked)

        ' ----------------------------------------------------- 
        'WaveGenerators.Sweep()
        Scope.Update(pbox_Scope, chk_ScopeNeg.Checked, chk_ScopeTrig.Checked)
        SpectrumBands.Update(pbox_SpectrumBars, _
                             CSng(txt_BandsMinFreq.NumericValue), _
                             CSng(txt_BandsMaxFreq.NumericValue), _
                             CSng(txt_BandsMinDb.NumericValue), _
                             CSng(txt_BandsMaxDb.NumericValue), _
                             chk_BandsLogX.Checked, _
                             chk_BandsLogY.Checked, _
                             txt_BandsSpeed.NumericValueInteger / 100.0F, _
                             SpectrumWindowType, _
                             txt_BandsCount.NumericValueInteger, _
                             chk_BandsAGC.Checked)
        Spectrum.Update(pbox_Spectrum, _
                        CSng(txt_SpecMinFreq.NumericValue), _
                        CSng(txt_SpecMaxFreq.NumericValue), _
                        CSng(txt_SpecMinDb.NumericValue), _
                        CSng(txt_SpecMaxDb.NumericValue), _
                        chk_SpecLogX.Checked, _
                        chk_SpecLogY.Checked, _
                        txt_SpecSpeed.NumericValueInteger / 100.0F, _
                        SpectrumWindowType)
        ' -----------------------------------------------------
        'If ReportString <> "" Then Text = ReportString


        ' ----------------------------------------------------- IN SLOTS
        If txt_FirstInSlot.NumericValueInteger >= 0 Then
            Dim n As Int32 = txt_FirstInSlot.NumericValueInteger
            If chk_Generator1.Checked Then
                Combo_SetIndex(cmb_Gen1_Wave, CInt(Slots.ReadSlot(n)))
                txt_Gen1_Freq.NumericValueInteger = CInt(Slots.ReadSlot(n + 1))
                txt_Gen1_dB.NumericValueInteger = CInt(Slots.ReadSlot(n + 2))
            End If
            If chk_Generator2.Checked Then
                Combo_SetIndex(cmb_Gen2_Wave, CInt(Slots.ReadSlot(n + 3)))
                txt_Gen2_Freq.NumericValueInteger = CInt(Slots.ReadSlot(n + 4))
                txt_Gen2_dB.NumericValueInteger = CInt(Slots.ReadSlot(n + 5))
            End If
            If chk_Generator3.Checked Then
                Combo_SetIndex(cmb_Gen3_Wave, CInt(Slots.ReadSlot(n + 6)))
                txt_Gen3_Freq.NumericValueInteger = CInt(Slots.ReadSlot(n + 7))
                txt_Gen3_dB.NumericValueInteger = CInt(Slots.ReadSlot(n + 8))
            End If
        End If
        ' ----------------------------------------------------- OUT SLOTS
        If txt_FirstOutSlot.NumericValueInteger >= 0 Then
            For i As Int32 = 0 To txt_BandsCount.NumericValueInteger - 1
                Slots.WriteSlot(txt_FirstOutSlot.NumericValueInteger + i, 1000.0F * SpectrumBands.m_fht.BandsBuffer(i))
            Next
        End If
    End Sub

    ' ===================================================================================================
    '  SELECT AUDIO INPUT 
    ' ===================================================================================================
    Friend SelectedAudioIn As Int32 = 0
    Private Sub cmb_AudioInDevices_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_AudioInDevices.DropDown
        cmb_AudioInDevices.ItemHeight = 16
        FillAudioInDevicesCombo()
    End Sub
    Private Sub cmb_AudioInDevices_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_AudioInDevices.DropDownClosed
        cmb_AudioInDevices.ItemHeight = 11
    End Sub
    Private Sub cmb_AudioInDevices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_AudioInDevices.SelectedIndexChanged
        If Not EventsAreEnabled Then Return
        SelectedAudioIn = cmb_AudioInDevices.SelectedIndex
        Scope.Initialize(pbox_Scope, SelectedAudioIn)
        Spectrum.Initialize(pbox_Spectrum, SelectedAudioIn)
        SpectrumBands.Initialize(pbox_SpectrumBars, SelectedAudioIn)
    End Sub
    Private Sub FillAudioInDevicesCombo()
        Dim sa As String() = WaveNative.GetInputDevicesNames
        If sa.Length = 0 Then ReDim sa(0)
        cmb_AudioInDevices.Items.Clear()
        For i As Int32 = 0 To sa.Length - 1
            cmb_AudioInDevices.Items.Add(ExtractDeviceName(sa(i)))
        Next
        Combo_SetIndex(cmb_AudioInDevices, SelectedAudioIn)
    End Sub

    ' ===================================================================================================
    '  SELECT AUDIO OUTPUT 
    ' ===================================================================================================
    Friend SelectedAudioOut As Int32 = 0
    Private Sub cmb_AudioOutDevices_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_AudioOutDevices.DropDown
        cmb_AudioOutDevices.ItemHeight = 16
        FillAudioOutDevicesCombo()
    End Sub
    Private Sub cmb_AudioOutDevices_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_AudioOutDevices.DropDownClosed
        cmb_AudioOutDevices.ItemHeight = 11
    End Sub
    Private Sub ccmb_AudioOutDevices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_AudioOutDevices.SelectedIndexChanged
        If Not EventsAreEnabled Then Return
        SelectedAudioOut = cmb_AudioOutDevices.SelectedIndex
        WaveGenerators.StartStop(SelectedAudioOut, _
                                 chk_Generator1.Checked, _
                                 chk_Generator2.Checked, _
                                 chk_Generator3.Checked)
    End Sub
    Private Sub FillAudioOutDevicesCombo()
        Dim sa As String() = WaveNative.GetOutputDevicesNames
        If sa.Length = 0 Then ReDim sa(0) : sa(0) = ""
        cmb_AudioOutDevices.Items.Clear()
        For i As Int32 = 0 To sa.Length - 1
            cmb_AudioOutDevices.Items.Add(ExtractDeviceName(sa(i)))
        Next
        Combo_SetIndex(cmb_AudioOutDevices, SelectedAudioOut)
    End Sub

    ' ===================================================================================================
    '  AUDIO IN OUT HELPERS
    ' ===================================================================================================
    Private Sub btn_AudioInputs_ClickButtonArea(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles btn_AudioInputs.ClickButtonArea
        Open_AudioInputs()
    End Sub
    Private Sub btn_AudioOutputs_ClickButtonArea(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles btn_AudioOutputs.ClickButtonArea
        Open_AudioOutputs()
    End Sub
    Private Function ExtractDeviceName(ByVal s As String) As String
        If s = Nothing Then Return ""
        Dim i As Int32 = InStr(s, "(") - 1
        Dim s2 As String = ""
        If i > 1 Then s2 = s.Remove(i)
        s2 = s2.Trim
        If s.ToLower.Contains("usb") Then s2 = s2 + " USB"
        Return s2
    End Function


    ' ===================================================================================================
    '  USER PARAMS 
    ' ===================================================================================================

    Private Sub SaveParamsOnLostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles chk_Generator1.LostFocus, _
                                                                        chk_Generator2.LostFocus, _
                                                                        chk_Generator3.LostFocus, _
                                                                        cmb_Gen1_Wave.LostFocus, _
                                                                        cmb_Gen2_Wave.LostFocus, _
                                                                        cmb_Gen3_Wave.LostFocus, _
                                                                        txt_Gen1_Freq.LostFocus, _
                                                                        txt_Gen2_Freq.LostFocus, _
                                                                        txt_Gen3_Freq.LostFocus, _
                                                                        txt_Gen1_dB.LostFocus, _
                                                                        txt_Gen2_dB.LostFocus, _
                                                                        txt_Gen3_dB.LostFocus, _
                                                                        tbar_ScopeVoltage.LostFocus, _
                                                                        tbar_ScopeTime.LostFocus, _
                                                                        chk_ScopeNeg.LostFocus, _
                                                                        chk_ScopeTrig.LostFocus, _
                                                                        txt_BandsCount.LostFocus, _
                                                                        chk_BandsAGC.LostFocus, _
                                                                        txt_FirstOutSlot.LostFocus, _
                                                                        txt_BandsMaxDb.LostFocus, _
                                                                        txt_BandsMinDb.LostFocus, _
                                                                        txt_BandsMaxFreq.LostFocus, _
                                                                        txt_BandsMinFreq.LostFocus, _
                                                                        txt_BandsSpeed.LostFocus, _
                                                                        chk_BandsLogX.LostFocus, _
                                                                        chk_BandsLogY.LostFocus, _
                                                                        txt_SpecMaxDb.LostFocus, _
                                                                        txt_SpecMinDb.LostFocus, _
                                                                        txt_SpecMaxFreq.LostFocus, _
                                                                        txt_SpecMinFreq.LostFocus, _
                                                                        txt_SpecSpeed.LostFocus, _
                                                                        chk_SpecLogX.LostFocus, _
                                                                        chk_SpecLogY.LostFocus, _
                                                                        cmb_Spec_Window.LostFocus
        If Not EventsAreEnabled Then Return
        Save_INI()
    End Sub

    Private Sub Gen_TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Gen1_Freq.TextChanged, _
                                                                                                            txt_Gen2_Freq.TextChanged, _
                                                                                                            txt_Gen3_Freq.TextChanged, _
                                                                                                            txt_SpecMaxFreq.TextChanged, _
                                                                                                            txt_SpecMinFreq.TextChanged, _
                                                                                                            txt_BandsMaxFreq.TextChanged, _
                                                                                                            txt_BandsMinFreq.TextChanged
        If Not EventsAreEnabled Then Return
        Dim txt As MyTextBox
        txt = DirectCast(sender, MyTextBox)
        If txt.NumericValue >= 1000 Then
            txt.Decimals = 0
            txt.ArrowsIncrement = 100
            txt.Increment = 1
            txt.RoundingStep = 100
        ElseIf txt.NumericValue >= 100 Then
            txt.Decimals = 0
            txt.ArrowsIncrement = 10
            txt.Increment = 1
            txt.RoundingStep = 10
        Else
            txt.Decimals = 0
            txt.ArrowsIncrement = 1
            txt.Increment = 1
            txt.RoundingStep = 0
        End If
        WaveGenTrim()
    End Sub

    Private Sub Gen_dB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Gen1_dB.TextChanged, _
                                                                                                       txt_Gen2_dB.TextChanged, _
                                                                                                       txt_Gen3_dB.TextChanged
        If Not EventsAreEnabled Then Return
        WaveGenTrim()
    End Sub

    Private Sub chk_AllGenerators_CheckedChanged(ByVal Sender As Object, ByVal e As System.EventArgs) Handles chk_Generator1.CheckedChanged, _
                                                                                                              chk_Generator2.CheckedChanged, _
                                                                                                              chk_Generator3.CheckedChanged
        If Not EventsAreEnabled Then Return
        WaveGenerators.StartStop(SelectedAudioOut, chk_Generator1.Checked, chk_Generator2.Checked, chk_Generator3.Checked)
    End Sub

    Private Sub AllCombo_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_Gen1_Wave.DropDown, _
                                                                                               cmb_Gen2_Wave.DropDown, _
                                                                                               cmb_Gen3_Wave.DropDown, _
                                                                                               cmb_Spec_Window.DropDown
        CType(sender, ComboBox).ItemHeight = 16
    End Sub

    Private Sub AllCombo_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_Gen1_Wave.DropDownClosed, _
                                                                                                     cmb_Gen2_Wave.DropDownClosed, _
                                                                                                     cmb_Gen3_Wave.DropDownClosed, _
                                                                                                     cmb_Spec_Window.DropDownClosed
        CType(sender, ComboBox).ItemHeight = 12
    End Sub

    Private Sub cmb_Gen1_Wave_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) _
                                                                    Handles cmb_Gen1_Wave.SelectionChangeCommitted, _
                                                                            cmb_Gen2_Wave.SelectionChangeCommitted, _
                                                                            cmb_Gen3_Wave.SelectionChangeCommitted
        If Not EventsAreEnabled Then Return
        WaveGenSetWaveTypes()
        GroupBox_Scope.Focus()
    End Sub

    Private Sub cmb_Spec_Window_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) _
                                                                    Handles cmb_Spec_Window.SelectionChangeCommitted
        If Not EventsAreEnabled Then Return
        SetSpectrumWindowType()
        GroupBox_Scope.Focus()
    End Sub

    Private Sub WaveGenSetWaveTypes()
        WaveGenerators.SetWavetypes(cmb_Gen1_Wave.SelectedIndex, _
                                    cmb_Gen2_Wave.SelectedIndex, _
                                    cmb_Gen3_Wave.SelectedIndex)
    End Sub

    Private Sub WaveGenTrim()
        WaveGenerators.SetFreq(CSng(txt_Gen1_Freq.NumericValue), _
                               CSng(txt_Gen2_Freq.NumericValue), _
                               CSng(txt_Gen3_Freq.NumericValue))
        WaveGenerators.SetOutLevel(txt_Gen1_dB.NumericValueInteger, _
                                   txt_Gen2_dB.NumericValueInteger, _
                                   txt_Gen3_dB.NumericValueInteger)
    End Sub

    Private Sub SetSpectrumWindowType()
        SpectrumWindowType = CType(cmb_Spec_Window.SelectedIndex, FHT.WindowTypes)
    End Sub

    Private Sub tbar_ScopeParams_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                        Handles tbar_ScopeVoltage.Scroll, _
                                                                                tbar_ScopeTime.Scroll
        If Not EventsAreEnabled Then Return
        ScopeTrim()
    End Sub

    Private Sub ScopeTrim()
        Scope.SetVoltage(VoltageFromTrackBar(tbar_ScopeVoltage.Value))
        Scope.SetTime(TimeFromTrackBar(tbar_ScopeTime.Value))
    End Sub

    Private Function VoltageFromTrackBar(ByVal n As Int32) As Single
        Select Case n
            Case 0 : Return 500
            Case 1 : Return 250
            Case 2 : Return 200
            Case 3 : Return 100
            Case 4 : Return 50
            Case 5 : Return 25
            Case 6 : Return 20
            Case 7 : Return 10
            Case 8 : Return 5
            Case 9 : Return 2.5
            Case 10 : Return 2
            Case 11 : Return 1
            Case 12 : Return 0.5
            Case 13 : Return 0.2
            Case 14 : Return 0.1
        End Select
    End Function

    Private Function TimeFromTrackBar(ByVal n As Int32) As Single
        Select Case n
            Case 0 : Return 100
            Case 1 : Return 50
            Case 2 : Return 20
            Case 3 : Return 10
            Case 4 : Return 5
            Case 5 : Return 2
            Case 6 : Return 1
            Case 7 : Return 0.5
            Case 8 : Return 0.2
            Case 9 : Return 0.1
        End Select
    End Function


    Private Sub txt_BandsCount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
