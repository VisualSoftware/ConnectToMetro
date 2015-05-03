'============================================================================
'
'    ConnectToMetro
'    Copyright (C) 2013 - 2015 Visual Software Corporation
'
'    Author: ASV93
'    File: Form1.vb
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'============================================================================


Imports NativeWifi
Imports System.Net.NetworkInformation
Public Class Form1

    Dim client As New WlanClient()
    Dim currentwlan As String
    Dim oldlabel As String = ""
    Dim getcurrentIP As String
    Dim currentIP As String
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Dim VSTools As VSSharedSource = New VSSharedSource
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListViewEx1.Nodes.Clear()
        Label7.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        ButtonX6_Click(sender, e)
        Try
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\Setup.upd") = True Then
                If IO.File.Exists(My.Application.Info.DirectoryPath & "\Setup.exe") = True Then
                    IO.File.Delete(My.Application.Info.DirectoryPath & "\Setup.exe")
                Else

                End If
                My.Computer.FileSystem.RenameFile(My.Application.Info.DirectoryPath & "\Setup.upd", "Setup.exe")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        Try
            Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
            wlanIface.Scan()
            ButtonX4_Click(sender, e)
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonX4_Click(sender As Object, e As EventArgs) Handles ButtonX4.Click
        Dim networkBssList As IEnumerable(Of Wlan.WlanBssEntry) = Nothing
        Dim availableNetworkList As IEnumerable(Of Wlan.WlanAvailableNetwork) = Nothing
        If ListViewEx1.Nodes.Count > 0 Then
            ListViewEx1.Nodes.Clear()
        End If
        Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
        'GROUP/NODE
        Try
            networkBssList = wlanIface.GetNetworkBssList()
            availableNetworkList = wlanIface.GetAvailableNetworkList(Wlan.WlanGetAvailableNetworkFlags.IncludeAllManualHiddenProfiles)
        Catch generatedExceptionName As NullReferenceException
            MessageBox.Show("CRITICAL ERROR", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Dim foundNetwork As New Wlan.WlanAvailableNetwork()
        For Each entry As Wlan.WlanBssEntry In networkBssList
            Dim ssid As String = System.Text.Encoding.ASCII.GetString(entry.dot11Ssid.SSID, 0, CInt(entry.dot11Ssid.SSIDLength))
            If FindNetwork(ssid, availableNetworkList, foundNetwork) Then
                Dim macAddr As Byte() = entry.dot11Bssid
                Dim tMac As String = ""
                For i As Integer = 0 To macAddr.Length - 1
                    tMac += macAddr(i).ToString("X2") + ":"c
                Next
                tMac = tMac.Remove(tMac.LastIndexOf(":"))
                Dim newitem As New DevComponents.AdvTree.Node
                If entry.linkQuality >= 80 Then
                    newitem.Image = My.Resources.metro5
                ElseIf entry.linkQuality >= 60 Then
                    newitem.Image = My.Resources.metro4
                ElseIf entry.linkQuality >= 40 Then
                    newitem.Image = My.Resources.metro3
                ElseIf entry.linkQuality >= 20 Then
                    newitem.Image = My.Resources.metro2
                ElseIf entry.linkQuality >= 10 Then
                    newitem.Image = My.Resources.metro1
                Else
                    newitem.Image = My.Resources.metro0
                End If
                Dim NetworkSSID As String
                NetworkSSID = GetStringForSSID(entry.dot11Ssid)
                If NetworkSSID = currentwlan Then
                    Dim newcell0 As New DevComponents.AdvTree.Cell
                    newcell0.Text = NetworkSSID ' & " [Connected]"
                    newcell0.Images.Image = My.Resources.NewConnected
                    newitem.Cells.Add(newcell0)
                Else
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell(NetworkSSID))
                End If
                newitem.Cells.Add(New DevComponents.AdvTree.Cell(tMac))
                'ENCRYPTION
                Dim newcell As New DevComponents.AdvTree.Cell
                If foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.WEP Then
                    newcell.Text = "WEP"
                    newcell.Images.Image = My.Resources.NewLowSec
                    newitem.Cells.Add(newcell)
                ElseIf foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.WEP104 Then
                    newcell.Text = "WEP (128-Bits)"
                    newcell.Images.Image = My.Resources.NewLowSec
                    newitem.Cells.Add(newcell)
                ElseIf foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.WEP40 Then
                    newcell.Text = "WEP (64-Bits)"
                    newcell.Images.Image = My.Resources.NewLowSec
                    newitem.Cells.Add(newcell)
                ElseIf foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.CCMP Then
                    newcell.Text = "CCMP"
                    newcell.Images.Image = My.Resources.NewHighSec
                    newitem.Cells.Add(newcell)
                ElseIf foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.TKIP Then
                    newcell.Text = "TKIP"
                    newcell.Images.Image = My.Resources.NewHighSec
                    newitem.Cells.Add(newcell)
                ElseIf foundNetwork.dot11DefaultCipherAlgorithm = Wlan.Dot11CipherAlgorithm.None Then
                    newcell.Text = "None"
                    newcell.Images.Image = My.Resources.NewNoSec
                    newitem.Cells.Add(newcell)
                Else
                    newcell.Text = "Unknown (" & foundNetwork.dot11DefaultCipherAlgorithm.ToString & ")"
                    newcell.Images.Image = My.Resources.Unknown_Security
                    newitem.Cells.Add(newcell)
                End If
                'AUTHENTICATION
                If foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.IEEE80211_Open Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("Open"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.IEEE80211_SharedKey Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("Shared Key"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.RSNA Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("WPA2-Enterprise"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.RSNA_PSK Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("WPA2-Personal"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.WPA Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("WPA-Enterprise"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.WPA_PSK Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("WPA-Personal"))
                ElseIf foundNetwork.dot11DefaultAuthAlgorithm = Wlan.Dot11AuthAlgorithm.WPA_None Then
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("WPA-None"))
                Else
                    newitem.Cells.Add(New DevComponents.AdvTree.Cell("Unknown (" & foundNetwork.dot11DefaultAuthAlgorithm.ToString & ")"))
                End If
                Dim newcell2 As New DevComponents.AdvTree.Cell
                If entry.dot11BssType = 1 Then
                    newcell2.Text = "Infrastructure"
                    newcell2.Images.Image = My.Resources.NewInfr
                    newitem.Cells.Add(newcell2)
                Else
                    newcell2.Text = "Ad-Hoc"
                    newcell2.Images.Image = My.Resources.NewAdHoc
                    newitem.Cells.Add(newcell2)
                End If
                Dim channelfrec As UInteger = entry.chCenterFrequency
                newitem.Cells.Add(New DevComponents.AdvTree.Cell(returnChannelStr(channelfrec)))
                Dim newcell3 As New DevComponents.AdvTree.Cell
                If entry.dot11BssPhyType.ToString = "7" Then
                    newcell3.Images.Image = My.Resources.NewN
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "HT" Then
                    newcell3.Images.Image = My.Resources.NewN
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "ERP" Then
                    newcell3.Images.Image = My.Resources.G
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "6" Then
                    newcell3.Images.Image = My.Resources.G
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "HRDSSS" Then
                    newcell3.Images.Image = My.Resources.NewB
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "5" Then
                    newcell3.Images.Image = My.Resources.NewB
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "OFDM" Then
                    newcell3.Images.Image = My.Resources.NewA
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "4" Then
                    newcell3.Images.Image = My.Resources.NewA
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "VT" Then
                    newcell3.Images.Image = My.Resources.NewAC
                    newitem.Cells.Add(newcell3)
                ElseIf entry.dot11BssPhyType.ToString = "8" Then
                    newcell3.Images.Image = My.Resources.NewAC
                    newitem.Cells.Add(newcell3)
                Else
                    newcell3.Text = "Unknown (" & entry.dot11BssPhyType.ToString & ")"
                    newcell3.Images.Image = My.Resources.UN
                    newitem.Cells.Add(newcell3)
                End If
                ListViewEx1.Nodes.Add(newitem)
            End If
        Next
    End Sub

    Private Shared Function FindNetwork(ssid As String, networks As IEnumerable(Of Wlan.WlanAvailableNetwork), ByRef foundNetwork As Wlan.WlanAvailableNetwork) As Boolean
        If networks IsNot Nothing Then
            For Each network As Wlan.WlanAvailableNetwork In networks
                Dim str As String = System.Text.Encoding.ASCII.GetString(network.dot11Ssid.SSID, 0, CInt(network.dot11Ssid.SSIDLength))
                If Not String.IsNullOrEmpty(str) AndAlso str.Equals(ssid) Then
                    foundNetwork = network
                    Return True
                End If
            Next
        End If
        Return False
    End Function
    Public Shared Function IndexOf(arrayToSearchThrough As Byte(), patternToFind As Byte()) As Integer
        If patternToFind.Length > arrayToSearchThrough.Length Then
            Return -2
        End If
        For i As Integer = 0 To arrayToSearchThrough.Length - patternToFind.Length - 1
            Dim found As Boolean = True
            For j As Integer = 0 To patternToFind.Length - 1
                If arrayToSearchThrough(i + j) <> patternToFind(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then
                Return i
            End If
        Next
        Return -1
    End Function
    Public Shared Function StrToByteArray(str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function
    Private Shared Function returnChannelStr(frequency As UInteger) As String

        Dim value As String = "0"
        Dim channelsFR As UInteger() = {2412000, 2417000, 2422000, 2427000, 2432000, 2437000, _
            2442000, 2447000, 2452000, 2457000, 2462000, 2467000, _
            2472000, 2484000}

        'Input the channel frequenzy
        For i As Integer = 0 To channelsFR.Length - 1
            If frequency = channelsFR(i) Then
                Return (i + 1).ToString()
            End If
        Next
        Return value
    End Function

    Private Shared Function GetStringForSSID(ssid As Wlan.Dot11Ssid) As String
        Return System.Text.Encoding.ASCII.GetString(ssid.SSID, 0, CInt(ssid.SSIDLength))
    End Function

    Private Sub ButtonX2_Click(sender As Object, e As EventArgs) Handles ButtonX2.Click
        End
    End Sub

    Private Sub ButtonX5_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Created by ASV93. Visual Software (c) 2004-2015", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ButtonX3_Click(sender As Object, e As EventArgs) Handles ButtonX3.Click
        Try
            For Each item As DevComponents.AdvTree.Node In ListViewEx1.SelectedNodes
                If item.IsSelected = True Then
                    NewConnection.LabelX6.Text = item.Cells(1).Text
                    NewConnection.LabelX7.Text = item.Cells(2).Text
                    NewConnection.LabelX8.Text = item.Cells(4).Text & " (" & item.Cells(3).Text & ")"
                    NewConnection.encryptiontype = item.Cells(3).Text
                    NewConnection.authtype = item.Cells(4).Text
                    NewConnection.Text = "New Connection"
                    NewConnection.createadhoc = 0
                    If currentwlan = item.Cells(1).Text Then

                    Else
                        NewConnection.ShowDialog()
                    End If
                End If
                Exit For
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ButtonX6_Click(sender As Object, e As EventArgs) Handles ButtonX6.Click
        ComboBoxEx1.Items.Clear()
        For Each wlanIface As WlanClient.WlanInterface In client.Interfaces
            Dim MAC = wlanIface.NetworkInterface.GetPhysicalAddress.ToString
            Dim adMAC = ""
            For i = 1 To MAC.Length Step 2
                adMAC &= MAC.Substring(i - 1, 1) & MAC.Substring(i, 1) & ":"
            Next
            ComboBoxEx1.Items.Add(wlanIface.InterfaceDescription & " (" & adMAC.Remove(adMAC.Length - 1) & ")")
        Next
        If ComboBoxEx1.Items.Count > 0 Then
            ComboBoxEx1.SelectedIndex = 0
        Else

        End If
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        MessageBox.Show("Developed by ASV93" & vbCrLf & "UI by Winor" & vbCrLf & vbCrLf & "Visual Software" & VSTools.GetCopyrightDate() & vbCrLf & "http://visualsoftware.wordpress.com", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub CopyMACAddressToClipboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyMACAddressToClipboardToolStripMenuItem.Click
        Try
            For Each item As DevComponents.AdvTree.Node In ListViewEx1.SelectedNodes
                If item.IsSelected = True Then
                    Clipboard.SetText(item.Cells(2).Text)
                    Exit For
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Please select a network before using this option", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Button3_Click(sender, e)
    End Sub

    Private Sub DisconnectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisconnectToolStripMenuItem.Click
        Try
            For Each item As DevComponents.AdvTree.Node In ListViewEx1.SelectedNodes
                If item.IsSelected = True Then
                    If currentwlan = item.Cells(1).Text Then
                        Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
                        wlanIface.Disconnect()
                    Else
                        MessageBox.Show("You're not currently connected to the selected network", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Exit For
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Please select a network before using this option", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub ConnectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConnectToolStripMenuItem.Click
        ButtonX3_Click(sender, e)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
        wlanIface.Disconnect()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
        For Each profileInfo As Wlan.WlanProfileInfo In wlanIface.GetProfiles()
            Dim name As String = profileInfo.profileName
            Dim xml As String = wlanIface.GetProfileXml(profileInfo.profileName)
            TextBox1.Text = TextBox1.Text & name & vbCrLf & xml & vbCrLf & vbCrLf
        Next
        IO.File.WriteAllText("Profiles.txt", TextBox1.Text)
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            For Each item As DevComponents.AdvTree.Node In ListViewEx1.SelectedNodes
                If item.IsSelected = True Then
                    If item.Cells(1).Text = currentwlan Then
                        ButtonX3.Enabled = False
                        ConnectToolStripMenuItem.Enabled = False
                    Else
                        If item.Cells(2).Text = "" Then
                            ButtonX3.Enabled = False
                            ConnectToolStripMenuItem.Enabled = False
                        Else
                            ButtonX3.Enabled = True
                            ConnectToolStripMenuItem.Enabled = True
                        End If

                    End If

                Else
                    ButtonX3.Enabled = False
                    ConnectToolStripMenuItem.Enabled = False
                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub LabelX1_Click(sender As Object, e As EventArgs) Handles LabelX1.Click

    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox2.MouseEnter
        PictureBox2.BackColor = Color.FromArgb(255, 7, 198, 217)
    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        PictureBox2.BackColor = Color.FromArgb(255, 7, 178, 217)
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        PictureBox2.BackColor = Color.FromArgb(255, 7, 158, 217)
        ButtonX1_Click(sender, e)
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        PictureBox4.BackColor = Color.FromArgb(255, 7, 158, 217)
        Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
        Dim MAC = wlanIface.NetworkInterface.GetPhysicalAddress.ToString
        Dim adMAC = ""
        For i = 1 To MAC.Length Step 2
            adMAC &= MAC.Substring(i - 1, 1) & MAC.Substring(i, 1) & ":"
        Next
        NewConnection.LabelX6.Text = ""
        NewConnection.LabelX7.Text = adMAC.Remove(adMAC.Length - 1)
        NewConnection.LabelX8.Text = "WPA2-Personal" & " (" & "CCMP" & ")"
        NewConnection.encryptiontype = "CCMP"
        NewConnection.authtype = "WPA2-Personal"
        NewConnection.createadhoc = 1
        NewConnection.Text = "Create an Ad-Hoc Network"
        NewConnection.ShowDialog()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        PictureBox5.BackColor = Color.FromArgb(255, 7, 158, 217)
        ButtonX3_Click(sender, e)
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        PictureBox8.BackColor = Color.FromArgb(255, 7, 158, 217)
        PictureBox1_Click(sender, e)
    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox4.MouseEnter
        PictureBox4.BackColor = Color.FromArgb(255, 7, 198, 217)
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave
        PictureBox4.BackColor = Color.FromArgb(255, 7, 178, 217)
    End Sub

    Private Sub PictureBox5_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox5.MouseEnter
        PictureBox5.BackColor = Color.FromArgb(255, 7, 198, 217)
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave
        PictureBox5.BackColor = Color.FromArgb(255, 7, 178, 217)
    End Sub

    Private Sub PictureBox8_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox8.MouseEnter
        PictureBox8.BackColor = Color.FromArgb(255, 7, 198, 217)
    End Sub

    Private Sub PictureBox8_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox8.MouseLeave
        PictureBox8.BackColor = Color.FromArgb(255, 7, 178, 217)
    End Sub

    Private Sub PictureBox6_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox6.MouseEnter
        PictureBox6.BackColor = Color.FromArgb(255, 238, 238, 238)
    End Sub

    Private Sub PictureBox6_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox6.MouseLeave
        PictureBox6.BackColor = Color.FromArgb(255, 255, 255, 255)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If oldlabel = LabelX3.Text Then

            Else
                oldlabel = LabelX3.Text
                ButtonX4_Click(sender, e)
            End If
            Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(ComboBoxEx1.SelectedIndex)
            If wlanIface.InterfaceState = Wlan.WlanInterfaceState.Connected Or wlanIface.InterfaceState = Wlan.WlanInterfaceState.AdHocNetworkFormed Then
                currentwlan = wlanIface.CurrentConnection.profileName
                getcurrentIP = "0"
                Select Case wlanIface.CurrentConnection.isState
                    Case Wlan.WlanInterfaceState.AdHocNetworkFormed
                        LabelX3.Text = "Connected to Ad-Hoc Network: " & wlanIface.CurrentConnection.profileName
                        getcurrentIP = "1"
                    Case Wlan.WlanInterfaceState.Associating
                        LabelX3.Text = "Associating..."
                    Case Wlan.WlanInterfaceState.Authenticating
                        LabelX3.Text = "Authenticating..."
                    Case Wlan.WlanInterfaceState.Connected
                        LabelX3.Text = "Connected to " & wlanIface.CurrentConnection.profileName
                        getcurrentIP = "1"
                    Case Wlan.WlanInterfaceState.Disconnected
                        LabelX3.Text = "Disconnected"
                    Case Wlan.WlanInterfaceState.Disconnecting
                        LabelX3.Text = "Disconnecting..."
                    Case Wlan.WlanInterfaceState.Discovering
                        LabelX3.Text = "Refreshing networks list..."
                    Case Wlan.WlanInterfaceState.NotReady
                        LabelX3.Text = "Not Ready"
                End Select
            Else
                LabelX3.Text = "Disconnected"
                getcurrentIP = "0"
            End If
            If getcurrentIP = "1" Then
                For Each ni As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                    If ni.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                        For Each ip As UnicastIPAddressInformation In ni.GetIPProperties().UnicastAddresses
                            If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                                If ni.Name = wlanIface.NetworkInterface.Name Then
                                    If currentIP = ip.Address.ToString() Then

                                    Else
                                        currentIP = ip.Address.ToString()
                                        Label2.Text = "    Connect To Metro - " & currentIP
                                    End If
                                End If
                            End If
                        Next
                    End If

                Next
            Else
                currentIP = ""
                If Label2.Text = "    Connect To Metro" Then

                Else
                    Label2.Text = "    Connect To Metro"
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Label2_MouseDown(sender As Object, e As MouseEventArgs) Handles Label2.MouseDown
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    Private Sub Label2_MouseMove(sender As Object, e As MouseEventArgs) Handles Label2.MouseMove
        If IsFormBeingDragged Then
            Dim temp As Point = New Point()
            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    Private Sub Label2_MouseUp(sender As Object, e As MouseEventArgs) Handles Label2.MouseUp
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        VSTools.OpenDonationPage()
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If BackgroundWorker2.IsBusy = True Then

        Else
            BackgroundWorker2.RunWorkerAsync()

        End If
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        Process.Start("https://www.twitter.com/VisualSoftCorp")
    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        PictureBox11.BackColor = Color.FromArgb(255, 7, 158, 217)
    End Sub

    Private Sub PictureBox11_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox11.MouseEnter
        PictureBox11.BackColor = Color.FromArgb(255, 7, 198, 217)
    End Sub

    Private Sub PictureBox11_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox11.MouseLeave
        PictureBox11.BackColor = Color.FromArgb(255, 7, 178, 217)
    End Sub
End Class
