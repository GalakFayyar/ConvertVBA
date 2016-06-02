Imports VBIDE
Imports System.IO
Imports System.IO.File
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports Microsoft.Win32

Public Class Form1
    Inherits System.Windows.Forms.Form
#Region "Déclaration des varibles"
    Public cheminRapport As String              ' Chemin vers les rapports ANADEFi
    Public defCheminRapport As String           ' Chemin vers les rapports locaux ANADEFi (disque dur pc)
    Public nbThreadMax As Integer = 4           ' Nombre de threads pour le traitement parallèle
    Public nbFichiers As Integer = 0            ' Nombre total de fichiers à traiter (progressbar)
    Public objWordApp As Microsoft.Office.Interop.Word.Application
    Public valeurProgressBar As Integer
    Public threadOperation As Thread
#End Region

#Region "Traitement de Modification du VBA"
    Private Sub TraitementModifVBA(ByVal unFichier As String)
        cheminRapport = txtChemin.Text
        Dim objWordDocument As Microsoft.Office.Interop.Word.Document

        ' Declare variables to access the macros in the workbook.
        Dim objProject As Microsoft.Vbe.Interop.VBProject
        Dim objComponent As VBIDE.VBComponent
        Dim objCode As VBIDE.CodeModule

        ' Declare other miscellaneous variables.
        Dim iLine As Integer
        Dim sProcName As String
        Dim pk As vbext_ProcKind

        Try
            'objWordApp = New Microsoft.Office.Interop.Word.Application
            'objWordApp.Visible = False

            objWordDocument = objWordApp.Documents.Open(cheminRapport)

            ' Get the project details in the workbook.
            objProject = objWordDocument.VBProject

            ' Iterate through each component in the project.
            For Each objComponent In objProject.VBComponents

                ' Find the code module for the project.
                objCode = objComponent.CodeModule

                ' Recherche Module AssistantBulle
                If objCode.Name = "AssistantBulle" Then
ResetProcLine:
                    ' Correction Anomalie Assistant Office
                    iLine = 1
                    Do While iLine < objCode.CountOfLines
                        sProcName = objCode.ProcOfLine(iLine, pk)

                        If sProcName <> Nothing Then
                            Dim tailleProc As Integer
                            Select Case sProcName
                                Case "AssistNew"
                                    tailleProc = objCode.ProcCountLines(sProcName, pk)
                                    If Not objCode.Find("Correction Compatibilite Office 2010", iLine, 1, tailleProc + iLine, 450) Then
                                        objCode.DeleteLines(iLine, tailleProc)
                                        Dim contentAssitNew As String
                                        Dim MacroAssistant As New MacroAssistant
                                        contentAssitNew = MacroAssistant.getContentAssistNew()
                                        objCode.InsertLines(iLine, contentAssitNew)
                                        GoTo ResetProcLine
                                    End If
                                Case "AssistFin"
                                    tailleProc = objCode.ProcCountLines(sProcName, pk)
                                    If Not objCode.Find("Correction Compatibilite Office 2010", iLine, 1, tailleProc + iLine, 450) Then
                                        objCode.DeleteLines(iLine - 1, tailleProc)
                                        Dim contentAssitFin As String
                                        Dim MacroAssistant As New MacroAssistant
                                        contentAssitFin = MacroAssistant.getContentAssistFin()
                                        objCode.InsertLines(iLine, contentAssitFin)
                                        GoTo ResetProcLine
                                    End If
                            End Select
                            tailleProc = objCode.ProcCountLines(sProcName, pk)
                            iLine = iLine + tailleProc
                        Else
                            iLine += 1
                        End If
                    Loop
                End If

                ' Correction Anomalie Comité CR
                If objCode.Name = "Rapport" Then
                    Dim jLine As Integer
                    For jLine = 1 To objCode.CountOfLines
                        If objCode.Lines(jLine, 1).Contains("Windows(StrFicComite).Activate") Then
                            objCode.ReplaceLine(jLine, vbTab & "Documents(StrFicComite).Activate")
                        End If
                        If objCode.Lines(jLine, 1).Contains("Windows(StrFicComite).Close") Then
                            objCode.ReplaceLine(jLine, vbTab & "Documents(StrFicComite).Close")
                        End If
                    Next jLine

                End If

                objCode = Nothing
                objComponent = Nothing
            Next

            objProject = Nothing

            ' Clean up and exit.
            objWordDocument.Save()
            objWordDocument.Close()
            objWordDocument = Nothing

            Debug.Print("Fin du traitement sur : " & cheminRapport)
            Debug.Print("-----------------------------------------------")

        Catch ex As Exception
            MsgBox("Erreur d'ouverture du fichier : " & ex.Message)
        End Try

    End Sub
#End Region

#Region "Chargement du formulaire"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        defCheminRapport = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\O.R.System\Anadefi\Client\Session Data\Rapport.doc"
        'defCheminRapport = "C:\ANADEFI\Rapport2.doc"
        btnTraitementAnadefi_Click(sender, e)
    End Sub
#End Region

#Region "Bouton Fermer"
    Private Sub btnFermer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFermer.Click
        'killProcessWord()
        Me.Dispose()
    End Sub
#End Region

#Region "Bouton Parcourir"
    Private Sub btnParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParcourir.Click
        Dim ofd As New OpenFileDialog()
        With ofd
            .Filter = "Fichiers Word (*.doc)|*.DOC"
            .InitialDirectory = Environment.SpecialFolder.Desktop
            .Multiselect = False
            .CheckFileExists = True
        End With
        If (ofd.ShowDialog = System.Windows.Forms.DialogResult.OK) Then
            'MsgBox(ofd.FileName)
            cheminRapport = ofd.FileName
            btnExec.Enabled = True
            txtChemin.Text = cheminRapport
            lblProgress.Visible = False
        End If
        ofd.Dispose()
    End Sub
#End Region

#Region "Bouton Executer"
    Private Sub btnExec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExec.Click

        If Exists(cheminRapport) Then
            Dim answer As DialogResult
            answer = MessageBox.Show("Attention, tous les documents word ouverts non enregistrés vont être perdus. Souhaitez-vous continuer ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If answer = vbYes Then
                btnExec.Enabled = False
                Debug.Flush()

                ' Reset Progress bar
                ProgressBar.Value = 0
                lblProgress.Visible = True
                lblProgress.Text = "Traitement en cours ..."

                threadOperation = New Thread(AddressOf ExecutionOperation)
                threadOperation.Start()
            End If
        Else
            MsgBox("Le fichier spécifié  à l'adresse suivante : " & vbNewLine & cheminRapport & vbNewLine & "n'existe pas.")
        End If

    End Sub
#End Region

#Region "Barre de menu"
    Private Sub ParamètresToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ParamètresToolStripMenuItem1.Click
        'defCheminRapport = InputBox("Entrer le chemin du répertoire contenant les rapports Anadefi locaux : ", "Paramétrage chemin Aandefi", defCheminRapport)
        formParam.Show()
    End Sub
#End Region

#Region "Bouton Sélectionner Document courant"
    Private Sub btnDocCourant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            objWordApp = GetObject(, "Word.Application")
            If Not objWordApp Is Nothing Then
                ' Si condition réalisée : word ouvert
                'Dim objWordDocument As Microsoft.Office.Interop.Word.Document
                'objWordDocument = objWord.Documents.Item(1).Name

                ' Récupération du nom du fichier
                If objWordApp.Documents.Count > 1 Then
                    MsgBox("Plusieurs documents Word sont ouverts ou en tâche de fond." & vbNewLine & "Fermez ces documents avant de poursuivre l'opération.")
                Else
                    cheminRapport = objWordApp.Documents.Item(1).Path & "\" & objWordApp.Documents.Item(1).Name
                    txtChemin.Text = cheminRapport
                End If
                lblProgress.Text = ""
            End If
        Catch ex As Exception
            MsgBox("Erreur d'ouverture du fichier : " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Bouton Sélectionner Document Anadefi"
    Private Sub btnTraitementAnadefi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTraitementAnadefi.Click
        lblProgress.Visible = False
        cheminRapport = defCheminRapport
        txtChemin.Text = cheminRapport
        btnExec.Enabled = True
    End Sub
#End Region

#Region "Test et correction des références du document"
    Private Sub testReferences()
        Dim ref As Microsoft.Vbe.Interop.Reference
        Dim objWordDocument As Microsoft.Office.Interop.Word.Document
        Dim objProject As Microsoft.Vbe.Interop.VBProject

        Try
            'objWordApp.Visible = False
            objWordApp.WordBasic.DisableAutoMacros()
            objWordDocument = objWordApp.Documents.Open(cheminRapport)
            objProject = objWordDocument.VBProject

            For Each ref In objProject.References

                ' If the reference is broken, send the name to the Immediate Window.
                If ref.IsBroken Then
                    objProject.References.Remove(ref)
                End If

                ' Suppression de la référence pour le monitoring sysmon.ocx (bug 2010 -> 2003)
                If ref.Name = "SystemMonitor" Then
                    objProject.References.Remove(ref)
                End If

            Next

        Catch ex As Exception
            MsgBox("Erreur lors du nettoyage des références du document : " & ex.Message)
        End Try

    End Sub
#End Region

#Region "Suppression des processus word"
    Private Sub killProcessWord()
        Dim p As Process
        For Each p In System.Diagnostics.Process.GetProcessesByName("winword")
            Try
                p.Kill()
                p.WaitForExit()
            Catch ex As Win32Exception
                MsgBox("Erreur lors de la fermeture de l'application Word : " & ex.Message)
            End Try
        Next
    End Sub
#End Region

#Region "Element de mise à jour de l'interface graphique"
    Public Delegate Sub RefreshUI()

    Public Sub RefreshProgressBar()
        ProgressBar.Value = valeurProgressBar
    End Sub

    Public Sub majPB(ByVal val As Integer)
        valeurProgressBar = val
        Try
            Invoke(New RefreshUI(AddressOf RefreshProgressBar))
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub

    Public Sub RefreshEltFinOperation()
        lblProgress.Text = "Traitement terminé"
        btnExec.Enabled = True
    End Sub

    Public Sub majFinOperation()
        Try
            Invoke(New RefreshUI(AddressOf RefreshEltFinOperation))
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub
#End Region

    Sub ExecutionOperation()
        Try
            ' Suppression des processus word en tâche de fond
            killProcessWord()

            majPB(10)

            ' Création de l'object applicatif Word
            objWordApp = New Microsoft.Office.Interop.Word.Application
            objWordApp.Visible = False
            objWordApp.DisplayAlerts = False

            ' Ajout de l'autorisation de modification du code VBA dans le registre
            CheckAccessVBAON(objWordApp)

            ' Nettoyage des références
            testReferences()

            majPB(30)

            TraitementModifVBA(cheminRapport)

            ' Mise à jour de la barre de progression
            majPB(70)
            objWordApp.Quit()
            objWordApp = Nothing
            majPB(80)
            ' Suppression des processus word en tâche de fond
            killProcessWord()
            majPB(90)
            majFinOperation()
            majPB(100)
            Dim answer As DialogResult
            answer = MessageBox.Show("Souhaitez-vous ouvrir le document mis à jour (cette opération va fermer l'outil) ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If answer = vbYes Then
                objWordApp = New Microsoft.Office.Interop.Word.Application
                objWordApp.Documents.Open(cheminRapport)
                objWordApp.Visible = True
                'Me.Dispose()
                CloseMe()
            End If
        Catch ex As Exception
            MsgBox("Une erreur s'est produite durant le traitement. " & vbNewLine & _
                   "Vérifiez que le fichier n'est pas protégé en écriture sur le disque (Lecture seule) " & _
                   "ou bien en cours d'utilisation." & vbNewLine & _
                   "Erreur : " & ex.Message)
        End Try
        
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitterToolStripMenuItem.Click
        btnFermer_Click(sender, e)
    End Sub

    Private Sub AProposToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AProposToolStripMenuItem.Click
        formAide.Show()
    End Sub

    Private Sub CheckAccessVBAON(ByVal application As Object)
        Dim strRegPath As String
        strRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & application.Version & "\Word\Security\AccessVBOM"
        Dim WshShell
        ' Si la clé de registre existe, on modifie la valeur pour autoriser la programmatique VBA
        If TestIfKeyExists(strRegPath) Then
            On Error Resume Next

            WshShell = CreateObject("WScript.Shell")
            WshShell.RegWrite(strRegPath, 1, "REG_DWORD")

            If Err.Number <> 0 Then
                MsgBox("Erreur " & Chr(13) & Chr(10) & Err.Source & Chr(13) & Chr(10) & Err.Description)
            End If
        Else
            ' Si la clé de registre n'existe pas, on la créé
            Registry.CurrentUser.CreateSubKey("Software\Microsoft\Office\" & application.Version & "\Word\Security\AccessVBOM")
            WshShell = CreateObject("WScript.Shell")
            WshShell.RegWrite(strRegPath, 1, "REG_DWORD")
        End If
    End Sub

    Public Function TestIfKeyExists(ByVal path As String) As Boolean
        Dim WshShell As Object
        WshShell = CreateObject("Wscript.Shell")
        On Error Resume Next
        WshShell.RegRead(path)
        If Err.Number <> 0 Then
            Err.Clear()
            TestIfKeyExists = False
        Else
            TestIfKeyExists = True
        End If
        On Error GoTo 0
    End Function

    Private Sub CloseMe()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf CloseMe))
            Exit Sub
        End If
        Me.Dispose()
    End Sub
End Class
