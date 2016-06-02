Public Class MacroAssistant
    ' Constructeur
    Public Function getContentAssistNew() As String
        Dim contentTest As String
        contentTest = "Sub AssistNew(StrTitre As String, StrLibelle As String)" & _
        vbNewLine & vbTab & "' Correction Compatibilite Office 2010" & _
        vbNewLine & vbTab & "' Assistant non supporté pour version Office > 2007 (12.0)" & _
        vbNewLine & vbTab & "' ------------------------------------------------------" & _
        vbNewLine & _
        vbNewLine & vbTab & "' Test pour afficage de l'assistant selon version Office" & _
        vbNewLine & vbTab & "Dim versionOffice As Integer" & _
        vbNewLine & vbTab & "versionOffice = Left(Application.Version, 2)" & _
        vbNewLine & _
        vbNewLine & vbTab & "If versionOffice < 13 Then" & _
        vbNewLine & vbTab & vbTab & "Dim StrAssistNom As String" & _
        vbNewLine & vbTab & vbTab & "StrAssistNom = Application.Assistant.Name" & _
        vbNewLine & vbTab & vbTab & "With Assistant" & _
        vbNewLine & vbTab & vbTab & vbTab & ".On = True" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Visible = True" & _
        vbNewLine & vbTab & vbTab & vbTab & ".MoveWhenInTheWay = False" & _
        vbNewLine & vbTab & vbTab & "End With" & _
        vbNewLine & vbTab & vbTab & "Application.Assistant.On = True" & _
        vbNewLine & vbTab & vbTab & "With Application.Assistant.NewBalloon" & _
        vbNewLine & vbTab & vbTab & vbTab & ".BalloonType = msoBalloonTypeBullets" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Icon = msoIconAlertWarning" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Button = msoButtonSetOK" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Heading = ""{cf 252}"" & StrTitre & ""{cf 0}""" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Text = ""{cf 252}"" & StrLibelle & ""{cf 0}""" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Button = msoButtonSetNone" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Mode = msoModeModeless" & _
        vbNewLine & vbTab & vbTab & vbTab & "Select Case StrAssistNom" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case ""MécanOffice""" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationSearching" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case ""Merlin""" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationWritingNotingSomething" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case ""Tifauve""" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationWorkingAtSomething" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case ""Toufou""" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationWritingNotingSomething" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case ""Trombine""" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationWritingNotingSomething" & _
        vbNewLine & vbTab & vbTab & vbTab & "Case Else" & _
        vbNewLine & vbTab & vbTab & vbTab & vbTab & ".Animation = msoAnimationCheckingSomething" & _
        vbNewLine & vbTab & vbTab & vbTab & "End Select" & _
        vbNewLine & vbTab & vbTab & vbTab & ".Show" & _
        vbNewLine & vbTab & vbTab & "End With" & _
        vbNewLine & vbTab & "End If" & _
        vbNewLine & "End Sub"

        Return contentTest
    End Function

    Public Function getContentAssistFin() As String
        Dim content As String
        content = "Sub AssistFin()" & _
        vbNewLine & vbTab & "' Correction Compatibilite Office 2010" & _
        vbNewLine & vbTab & "' Assistant non supporté pour version Office > 2007 (12.0)" & _
        vbNewLine & vbTab & "' --------------------------------------------------------" & _
        vbNewLine & vbTab & _
        vbNewLine & vbTab & "' Test pour affichage de l'assistant selon version Office" & _
        vbNewLine & vbTab & "Dim versionOffice As Integer" & _
        vbNewLine & vbTab & "versionOffice = Left(Application.Version, 2)" & _
        vbNewLine & vbTab & "If versionOffice < 13 Then" & _
        vbNewLine & vbTab & vbTab & "Application.Assistant.NewBalloon.Close" & _
        vbNewLine & vbTab & vbTab & "Application.Assistant.On = False" & _
        vbNewLine & vbTab & vbTab & "If BooAssistOn = True Then" & _
        vbNewLine & vbTab & vbTab & vbTab & "Application.Assistant.On = True" & _
        vbNewLine & vbTab & vbTab & "End If" & _
        vbNewLine & vbTab & vbTab & "Application.Assistant.MoveWhenInTheWay = BooAssistMove" & _
        vbNewLine & vbTab & vbTab & "Application.Assistant.Visible = BooAssistVisible" & _
        vbNewLine & vbTab & "End If" & _
        vbNewLine & vbTab & _
        vbNewLine & vbTab & "If ActiveWindow.View.SplitSpecial = wdPaneNone Then" & _
        vbNewLine & vbTab & vbTab & "ActiveWindow.ActivePane.View.Type = wdPrintView" & _
        vbNewLine & vbTab & "Else" & _
        vbNewLine & vbTab & vbTab & "ActiveWindow.View.Type = wdPrintView" & _
        vbNewLine & vbTab & "End If" & _
        vbNewLine & "End Sub"

        Return content
    End Function
End Class
