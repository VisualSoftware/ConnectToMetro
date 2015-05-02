Imports NativeWifi
Public Class NewConnection

    Dim client As New WlanClient()
    Public encryptiontype As String
    Public authtype As String
    Dim hiddenssid As Integer
    Public createadhoc As Integer
    Public enterprisemode As Integer
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Private Sub ButtonX4_Click(sender As Object, e As EventArgs) Handles ButtonX4.Click
        Me.Close()
    End Sub

    Private Sub NewConnection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoxX1.Enabled = True
        Label2.Text = "    " & Me.Text
        If Me.Text = "New Connection" Then
            Me.Icon = My.Resources.NC
            Label2.Image = My.Resources.NC_X16
        Else
            Me.Icon = My.Resources.CA
            Label2.Image = My.Resources.CA_X16
        End If
        If LabelX6.Text.Length = 0 Then
            TextBoxX2.Visible = True
            hiddenssid = 1
            TextBoxX2.Text = ""
        Else
            TextBoxX2.Visible = False
            hiddenssid = 0
        End If
        If createadhoc = 1 Then
            ComboBoxEx1.Visible = True
            ComboBoxEx1.SelectedIndex = 0
            TextBoxX2.Text = "NewAdHocNetwork"
        Else
            ComboBoxEx1.Visible = False
            If encryptiontype = "None" Then
                TextBoxX1.Enabled = False
            End If
        End If

        ComboBoxEx2.SelectedIndex = 0
        TextBoxX1.Text = ""
        If authtype = "WPA" OrElse authtype = "WPA2" Then
            LabelX5.Visible = False
            TextBoxX1.Visible = False
            CheckBox1.Visible = False
            LabelX1.Visible = True
            ComboBoxEx2.Visible = True
        Else
            LabelX5.Visible = True
            TextBoxX1.Visible = True
            CheckBox1.Visible = True
            LabelX1.Visible = False
            ComboBoxEx2.Visible = False
        End If
    End Sub

    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        Try
            'TODO: ADD SUPPORT FOR ENTERPRISE
            If hiddenssid = 1 Then
                LabelX6.Text = TextBoxX2.Text
            Else
            End If
            Dim wlanIface As WlanClient.WlanInterface = client.Interfaces(My.Forms.Form1.ComboBoxEx1.SelectedIndex)
            Dim mac As String = HexIt(LabelX6.Text).Replace(" ", "")
            Dim profileXml As String
            If createadhoc = 1 Then
                'AD-HOC NETWORK
                If ComboBoxEx1.SelectedIndex = 1 Then
                    'wep
                    If TextBoxX1.Text.Count = 5 Then
                        'OK
                    ElseIf TextBoxX1.Text.Count = 13 Then
                        'OK
                    Else
                        MessageBox.Show("A WEP Password should contain 5 or 13 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                ElseIf ComboBoxEx1.SelectedIndex = 2 Then
                    'wpa2
                    If TextBoxX1.Text.Count > 9 Then
                        'OK
                    Else
                        MessageBox.Show("A WPA2-Personal Password should contain at least 10 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                End If
                If ComboBoxEx1.SelectedIndex = 0 Then
                    'Open
                    profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>IBSS</connectionType><MSM><security><authEncryption><authentication>open</authentication><encryption>none</encryption><useOneX>false</useOneX></authEncryption></security></MSM></WLANProfile>", LabelX6.Text, mac)
                ElseIf ComboBoxEx1.SelectedIndex = 1 Then
                    'WEP
                    profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>IBSS</connectionType><MSM><security><authEncryption><authentication>open</authentication><encryption>WEP</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>networkKey</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                Else
                    'WPA2-Personal CCMP
                    profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>IBSS</connectionType><MSM><security><authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                End If
            Else
                'INFRASTRUCTURE
                If encryptiontype = "WEP" Then
                    'WEP
                    If TextBoxX1.Text.Count = 5 Then
                        'OK
                    ElseIf TextBoxX1.Text.Count = 13 Then
                        'OK
                    Else
                        MessageBox.Show("A WEP Password should contain 5 or 13 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>open</authentication><encryption>WEP</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>networkKey</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                ElseIf encryptiontype = "CCMP" Then
                    If authtype = "WPA-Personal" Then
                        If TextBoxX1.Text.Count > 9 Then
                            'OK
                        Else
                            MessageBox.Show("A WPA-PSK/WPA2-Personal Password should contain at least 10 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        'WPA-PSK CCMP
                        profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>WPAPSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                    ElseIf authtype = "WPA2-Personal" Then
                        If TextBoxX1.Text.Count > 9 Then
                            'OK
                        Else
                            MessageBox.Show("A WPA-PSK/WPA2-Personal Password should contain at least 10 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        'WPA2-PSK CCMP
                        profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                    ElseIf authtype = "WPA" Then
                        If ComboBoxEx2.SelectedIndex = 0 Then
                            'WPA CCMP PEAP-MSCHAPv2
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA</authentication><encryption>AES</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>25</eapCommon:Type> <eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:msPeap=""http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV1"" xmlns:msChapV2=""http://www.microsoft.com/provisioning/MsChapV2ConnectionPropertiesV1""><baseEap:Eap><baseEap:Type>25</baseEap:Type> <msPeap:EapType><msPeap:ServerValidation><msPeap:DisableUserPromptForServerValidation>false</msPeap:DisableUserPromptForServerValidation><msPeap:TrustedRootCA /></msPeap:ServerValidation><msPeap:FastReconnect>true</msPeap:FastReconnect><msPeap:InnerEapOptional>0</msPeap:InnerEapOptional><baseEap:Eap><baseEap:Type>26</baseEap:Type><msChapV2:EapType><msChapV2:UseWinLogonCredentials>false</msChapV2:UseWinLogonCredentials></msChapV2:EapType></baseEap:Eap><msPeap:EnableQuarantineChecks>false</msPeap:EnableQuarantineChecks><msPeap:RequireCryptoBinding>false</msPeap:RequireCryptoBinding><msPeap:PeapExtensions /></msPeap:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        Else
                            'WPA CCMP TLS
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID><nonBroadcast>false</nonBroadcast></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><autoSwitch>false</autoSwitch><MSM><security><authEncryption><authentication>WPA</authentication><encryption>AES</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>13</eapCommon:Type><eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:eapTls=""http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV1"">baseEap:Eap><baseEap:Type>13</baseEap:Type><eapTls:EapType><eapTls:CredentialsSource><eapTls:CertificateStore /></eapTls:CredentialsSource><eapTls:ServerValidation><eapTls:DisableUserPromptForServerValidation>false</eapTls:DisableUserPromptForServerValidation><eapTls:ServerNames /></eapTls:ServerValidation><eapTls:DifferentUsername>false</eapTls:DifferentUsername></eapTls:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        End If
                    ElseIf authtype = "WPA2" Then
                        If ComboBoxEx2.SelectedIndex = 0 Then
                            'WPA2 CCMP PEAP-MSCHAPv2
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA2</authentication><encryption>AES</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>25</eapCommon:Type> <eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:msPeap=""http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV1"" xmlns:msChapV2=""http://www.microsoft.com/provisioning/MsChapV2ConnectionPropertiesV1""><baseEap:Eap><baseEap:Type>25</baseEap:Type> <msPeap:EapType><msPeap:ServerValidation><msPeap:DisableUserPromptForServerValidation>false</msPeap:DisableUserPromptForServerValidation><msPeap:TrustedRootCA /></msPeap:ServerValidation><msPeap:FastReconnect>true</msPeap:FastReconnect><msPeap:InnerEapOptional>0</msPeap:InnerEapOptional><baseEap:Eap><baseEap:Type>26</baseEap:Type><msChapV2:EapType><msChapV2:UseWinLogonCredentials>false</msChapV2:UseWinLogonCredentials></msChapV2:EapType></baseEap:Eap><msPeap:EnableQuarantineChecks>false</msPeap:EnableQuarantineChecks><msPeap:RequireCryptoBinding>false</msPeap:RequireCryptoBinding><msPeap:PeapExtensions /></msPeap:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        Else
                            'WPA2 CCMP TLS
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID><nonBroadcast>false</nonBroadcast></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><autoSwitch>false</autoSwitch><MSM><security><authEncryption><authentication>WPA2</authentication><encryption>AES</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>13</eapCommon:Type><eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:eapTls=""http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV1"">baseEap:Eap><baseEap:Type>13</baseEap:Type><eapTls:EapType><eapTls:CredentialsSource><eapTls:CertificateStore /></eapTls:CredentialsSource><eapTls:ServerValidation><eapTls:DisableUserPromptForServerValidation>false</eapTls:DisableUserPromptForServerValidation><eapTls:ServerNames /></eapTls:ServerValidation><eapTls:DifferentUsername>false</eapTls:DifferentUsername></eapTls:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        End If
                    End If
                ElseIf encryptiontype = "TKIP" Then
                    If authtype = "WPA-Personal" Then
                        If TextBoxX1.Text.Count > 9 Then
                            'OK
                        Else
                            MessageBox.Show("A WPA-PSK/WPA2-Personal Password should contain at least 10 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        'WPA-PSK TKIP
                        profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>WPAPSK</authentication><encryption>TKIP</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                    ElseIf authtype = "WPA2-Personal" Then
                        If TextBoxX1.Text.Count > 9 Then
                            'OK
                        Else
                            MessageBox.Show("A WPA-PSK/WPA2-Personal Password should contain at least 10 ASCII chars", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        'WPA2-PSK TKIP
                        profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>WPA2PSK</authentication><encryption>TKIP</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{2}</keyMaterial></sharedKey><keyIndex>0</keyIndex></security></MSM></WLANProfile>", LabelX6.Text, mac, TextBoxX1.Text)
                    ElseIf authtype = "WPA" Then
                        If ComboBoxEx2.SelectedIndex = 0 Then
                            'WPA TKIP PEAP-MSCHAPv2
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA</authentication><encryption>TKIP</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>25</eapCommon:Type> <eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:msPeap=""http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV1"" xmlns:msChapV2=""http://www.microsoft.com/provisioning/MsChapV2ConnectionPropertiesV1""><baseEap:Eap><baseEap:Type>25</baseEap:Type> <msPeap:EapType><msPeap:ServerValidation><msPeap:DisableUserPromptForServerValidation>false</msPeap:DisableUserPromptForServerValidation><msPeap:TrustedRootCA /></msPeap:ServerValidation><msPeap:FastReconnect>true</msPeap:FastReconnect><msPeap:InnerEapOptional>0</msPeap:InnerEapOptional><baseEap:Eap><baseEap:Type>26</baseEap:Type><msChapV2:EapType><msChapV2:UseWinLogonCredentials>false</msChapV2:UseWinLogonCredentials></msChapV2:EapType></baseEap:Eap><msPeap:EnableQuarantineChecks>false</msPeap:EnableQuarantineChecks><msPeap:RequireCryptoBinding>false</msPeap:RequireCryptoBinding><msPeap:PeapExtensions /></msPeap:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        Else
                            'WPA TKIP TLS
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID><nonBroadcast>false</nonBroadcast></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><autoSwitch>false</autoSwitch><MSM><security><authEncryption><authentication>WPA</authentication><encryption>TKIP</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>13</eapCommon:Type><eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:eapTls=""http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV1"">baseEap:Eap><baseEap:Type>13</baseEap:Type><eapTls:EapType><eapTls:CredentialsSource><eapTls:CertificateStore /></eapTls:CredentialsSource><eapTls:ServerValidation><eapTls:DisableUserPromptForServerValidation>false</eapTls:DisableUserPromptForServerValidation><eapTls:ServerNames /></eapTls:ServerValidation><eapTls:DifferentUsername>false</eapTls:DifferentUsername></eapTls:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        End If
                    ElseIf authtype = "WPA2" Then
                        If ComboBoxEx2.SelectedIndex = 0 Then
                            'WPA2 TKIP PEAP-MSCHAPv2
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA2</authentication><encryption>TKIP</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>25</eapCommon:Type> <eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:msPeap=""http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV1"" xmlns:msChapV2=""http://www.microsoft.com/provisioning/MsChapV2ConnectionPropertiesV1""><baseEap:Eap><baseEap:Type>25</baseEap:Type> <msPeap:EapType><msPeap:ServerValidation><msPeap:DisableUserPromptForServerValidation>false</msPeap:DisableUserPromptForServerValidation><msPeap:TrustedRootCA /></msPeap:ServerValidation><msPeap:FastReconnect>true</msPeap:FastReconnect><msPeap:InnerEapOptional>0</msPeap:InnerEapOptional><baseEap:Eap><baseEap:Type>26</baseEap:Type><msChapV2:EapType><msChapV2:UseWinLogonCredentials>false</msChapV2:UseWinLogonCredentials></msChapV2:EapType></baseEap:Eap><msPeap:EnableQuarantineChecks>false</msPeap:EnableQuarantineChecks><msPeap:RequireCryptoBinding>false</msPeap:RequireCryptoBinding><msPeap:PeapExtensions /></msPeap:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        Else
                            'WPA2 TKIP TLS
                            profileXml = String.Format("<?xml version=""1.0"" encoding=""US-ASCII""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID><nonBroadcast>false</nonBroadcast></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><autoSwitch>false</autoSwitch><MSM><security><authEncryption><authentication>WPA2</authentication><encryption>TKIP</encryption><useOneX>true</useOneX></authEncryption><OneX xmlns=""http://www.microsoft.com/networking/OneX/v1""><EAPConfig><EapHostConfig xmlns=""http://www.microsoft.com/provisioning/EapHostConfig"" xmlns:eapCommon=""http://www.microsoft.com/provisioning/EapCommon"" xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapMethodConfig""><EapMethod><eapCommon:Type>13</eapCommon:Type><eapCommon:AuthorId>0</eapCommon:AuthorId></EapMethod><Config xmlns:baseEap=""http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"" xmlns:eapTls=""http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV1"">baseEap:Eap><baseEap:Type>13</baseEap:Type><eapTls:EapType><eapTls:CredentialsSource><eapTls:CertificateStore /></eapTls:CredentialsSource><eapTls:ServerValidation><eapTls:DisableUserPromptForServerValidation>false</eapTls:DisableUserPromptForServerValidation><eapTls:ServerNames /></eapTls:ServerValidation><eapTls:DifferentUsername>false</eapTls:DifferentUsername></eapTls:EapType></baseEap:Eap></Config></EapHostConfig></EAPConfig></OneX></security></MSM></WLANProfile>", LabelX6.Text, mac)
                        End If
                    End If
                ElseIf encryptiontype = "None" Then
                    profileXml = String.Format("<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>{0}</name><SSIDConfig><SSID><hex>{1}</hex><name>{0}</name></SSID></SSIDConfig><connectionType>ESS</connectionType><MSM><security><authEncryption><authentication>open</authentication><encryption>none</encryption><useOneX>false</useOneX></authEncryption></security></MSM></WLANProfile>", LabelX6.Text, mac)
                End If
            End If
            wlanIface.SetProfile(Wlan.WlanProfileFlags.AllUser, profileXml, True)
            wlanIface.Connect(Wlan.WlanConnectionMode.Profile, Wlan.Dot11BssType.Any, LabelX6.Text)
            Me.Hide()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBoxX1.PasswordChar = "●"
        Else
            TextBoxX1.PasswordChar = ""
        End If
    End Sub

    Public Function HexIt(sText As String) As String
        Dim A As Long
        For A = 1 To Len(sText)
            HexIt = HexIt & Hex$(Asc(Mid(sText, A, 1))) & Space$(1)
            On Error Resume Next
        Next A
        HexIt = Mid$(HexIt, 1, Len(HexIt) - 1)
    End Function

    Private Sub Label2_MouseUp(sender As Object, e As MouseEventArgs) Handles Label2.MouseUp
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
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

    Private Sub ComboBoxEx1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        If ComboBoxEx1.SelectedIndex = 0 Then
            TextBoxX1.Enabled = False
        Else
            TextBoxX1.Enabled = True
        End If
    End Sub
End Class