Imports System.Globalization
Imports System.Xml
Imports System.IO
Imports mRemoteNG.App.Runtime

Namespace Config
    Namespace Connections
        Public Class XML
            Inherits Base
#Region "Public Functions"
            Public Overrides Function Load() As Boolean
                Dim strCons As String = DecryptCompleteFile()
                LoadFromXML(strCons)

                If Import = False Then
                    SetMainFormText(ConnectionFileName)
                End If
            End Function

            Public Overrides Function Save() As Boolean
                SaveToXML()
                If My.Settings.EncryptCompleteConnectionsFile Then
                    EncryptCompleteFile()
                End If
                SetMainFormText(ConnectionFileName)
            End Function
#End Region

#Region "Private Variables"
            Private confVersion As Double
            Private xDom As XmlDocument
            Private xW As XmlTextWriter
#End Region

#Region "Private Functions"
#Region "Load"
            Private Function DecryptCompleteFile() As String
                Dim sRd As New StreamReader(ConnectionFileName)

                Dim strCons As String
                strCons = sRd.ReadToEnd
                sRd.Close()

                If strCons <> "" Then
                    Dim strDecr As String = ""
                    Dim notDecr As Boolean = True

                    If strCons.Contains("<?xml version=""1.0"" encoding=""utf-8""?>") Then
                        strDecr = strCons
                        Return strDecr
                    End If

                    Try
                        strDecr = Security.Crypt.Decrypt(strCons, pW)

                        If strDecr <> strCons Then
                            notDecr = False
                        Else
                            notDecr = True
                        End If
                    Catch ex As Exception
                        notDecr = True
                    End Try

                    If notDecr Then
                        If Authenticate(strCons, True) = True Then
                            strDecr = Security.Crypt.Decrypt(strCons, pW)
                            notDecr = False
                        Else
                            notDecr = True
                        End If

                        If notDecr = False Then
                            Return strDecr
                        End If
                    Else
                        Return strDecr
                    End If
                End If

                Return ""
            End Function

            Private Sub LoadFromXML(Optional ByVal cons As String = "")
                Try
                    App.Runtime.IsConnectionsFileLoaded = False

                    ' SECTION 1. Create a DOM Document and load the XML data into it.
                    Me.xDom = New XmlDocument()
                    If cons <> "" Then
                        xDom.LoadXml(cons)
                    Else
                        xDom.Load(ConnectionFileName)
                    End If

                    If xDom.DocumentElement.HasAttribute("ConfVersion") Then
                        Dim enCulture As New CultureInfo("en-US")
                        Me.confVersion = Convert.ToDouble(xDom.DocumentElement.Attributes("ConfVersion").Value, enCulture)
                    Else
                        MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, My.Resources.strOldConffile)
                    End If

                    ' SECTION 2. Initialize the treeview control.
                    Dim rootNode As TreeNode

                    Try
                        rootNode = New TreeNode(xDom.DocumentElement.Attributes("Name").Value)
                    Catch ex As Exception
                        rootNode = New TreeNode(xDom.DocumentElement.Name)
                    End Try

                    Dim rInfo As New Root.Info(Root.Info.RootType.Connection)
                    rInfo.Name = rootNode.Text
                    rInfo.TreeNode = rootNode

                    rootNode.Tag = rInfo

                    If Me.confVersion > 1.3 Then '1.4
                        If Security.Crypt.Decrypt(xDom.DocumentElement.Attributes("Protected").Value, pW) <> "ThisIsNotProtected" Then
                            If Authenticate(xDom.DocumentElement.Attributes("Protected").Value, False, rInfo) = False Then
                                My.Settings.LoadConsFromCustomLocation = False
                                My.Settings.CustomConsPath = ""
                                rootNode.Remove()
                                Exit Sub
                            End If
                        End If
                    End If

                    Dim imp As Boolean = False

                    If Me.confVersion > 0.9 Then '1.0
                        If xDom.DocumentElement.Attributes("Export").Value = True Then
                            imp = True
                        End If
                    End If

                    If Import = True And imp = False Then
                        MessageCollector.AddMessage(Messages.MessageClass.InformationMsg, My.Resources.strCannotImportNormalSessionFile)

                        Exit Sub
                    End If

                    If imp = False Then
                        RootTreeNode.Text = rootNode.Text
                        RootTreeNode.Tag = rootNode.Tag
                        RootTreeNode.ImageIndex = Images.Enums.TreeImage.Root
                        RootTreeNode.SelectedImageIndex = Images.Enums.TreeImage.Root
                    End If

                    ' SECTION 3. Populate the TreeView with the DOM nodes.
                    AddNodeFromXML(xDom.DocumentElement, RootTreeNode)

                    RootTreeNode.Expand()

                    'expand containers
                    For Each contI As Container.Info In ContainerList
                        If contI.IsExpanded = True Then
                            contI.TreeNode.Expand()
                        End If
                    Next

                    'open connections from last mremote session
                    If My.Settings.OpenConsFromLastSession = True And My.Settings.NoReconnect = False Then
                        For Each conI As Connection.Info In ConnectionList
                            If conI.PleaseConnect = True Then
                                App.Runtime.OpenConnection(conI)
                            End If
                        Next
                    End If

                    RootTreeNode.EnsureVisible()

                    App.Runtime.IsConnectionsFileLoaded = True
                    App.Runtime.Windows.treeForm.InitialRefresh()
                Catch ex As Exception
                    App.Runtime.IsConnectionsFileLoaded = False
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strLoadFromXmlFailed & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private prevCont As Container.Info
            Private Sub AddNodeFromXML(ByRef inXmlNode As XmlNode, ByRef inTreeNode As TreeNode)
                Try
                    Dim i As Integer

                    Dim xNode As XmlNode
                    Dim xNodeList As XmlNodeList
                    Dim tNode As TreeNode

                    ' Loop through the XML nodes until the leaf is reached.
                    ' Add the nodes to the TreeView during the looping process.
                    If inXmlNode.HasChildNodes() Then
                        xNodeList = inXmlNode.ChildNodes
                        For i = 0 To xNodeList.Count - 1
                            xNode = xNodeList(i)
                            inTreeNode.Nodes.Add(New TreeNode(xNode.Attributes("Name").Value))
                            tNode = inTreeNode.Nodes(i)

                            If Tree.Node.GetNodeTypeFromString(xNode.Attributes("Type").Value) = Tree.Node.Type.Connection Then 'connection info
                                Dim conI As Connection.Info = GetConnectionInfoFromXml(xNode)
                                conI.TreeNode = tNode
                                conI.Parent = prevCont 'NEW

                                ConnectionList.Add(conI)

                                tNode.Tag = conI
                                tNode.ImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                tNode.SelectedImageIndex = Images.Enums.TreeImage.ConnectionClosed
                            ElseIf Tree.Node.GetNodeTypeFromString(xNode.Attributes("Type").Value) = Tree.Node.Type.Container Then  'container info
                                Dim contI As New Container.Info
                                If tNode.Parent IsNot Nothing Then
                                    If Tree.Node.GetNodeType(tNode.Parent) = Tree.Node.Type.Container Then
                                        contI.Parent = tNode.Parent.Tag
                                    End If
                                End If
                                prevCont = contI 'NEW
                                contI.TreeNode = tNode

                                contI.Name = xNode.Attributes("Name").Value

                                If Me.confVersion > 0.7 Then '0.8
                                    If xNode.Attributes("Expanded").Value = "True" Then
                                        contI.IsExpanded = True
                                    Else
                                        contI.IsExpanded = False
                                    End If
                                End If

                                Dim conI As Connection.Info
                                If Me.confVersion > 0.8 Then '0.9
                                    conI = GetConnectionInfoFromXml(xNode)
                                Else
                                    conI = New Connection.Info
                                End If

                                conI.Parent = contI
                                conI.IsContainer = True
                                contI.ConnectionInfo = conI

                                ContainerList.Add(contI)

                                tNode.Tag = contI
                                tNode.ImageIndex = Images.Enums.TreeImage.Container
                                tNode.SelectedImageIndex = Images.Enums.TreeImage.Container
                            End If

                            AddNodeFromXML(xNode, tNode)
                        Next
                    Else
                        inTreeNode.Text = inXmlNode.Attributes("Name").Value.Trim
                    End If
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strAddNodeFromXmlFailed & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Function GetConnectionInfoFromXml(ByVal xxNode As XmlNode) As Connection.Info
                Dim conI As New Connection.Info

                Try
                    With xxNode
                        If Me.confVersion > 0.1 Then '0.2
                            conI.Name = .Attributes("Name").Value
                            conI.Description = .Attributes("Descr").Value
                            conI.Hostname = .Attributes("Hostname").Value
                            conI.Username = .Attributes("Username").Value
                            conI.Password = Security.Crypt.Decrypt(.Attributes("Password").Value, pW)
                            conI.Domain = .Attributes("Domain").Value
                            conI.DisplayWallpaper = .Attributes("DisplayWallpaper").Value
                            conI.DisplayThemes = .Attributes("DisplayThemes").Value
                            conI.CacheBitmaps = .Attributes("CacheBitmaps").Value

                            If Me.confVersion < 1.1 Then '1.0 - 0.1
                                If .Attributes("Fullscreen").Value = True Then
                                    conI.Resolution = Connection.Protocol.RDP.RDPResolutions.Fullscreen
                                Else
                                    conI.Resolution = Connection.Protocol.RDP.RDPResolutions.FitToWindow
                                End If
                            End If
                        End If

                        If Me.confVersion > 0.2 Then '0.3
                            If Me.confVersion < 0.7 Then
                                If CType(.Attributes("UseVNC").Value, Boolean) = True Then
                                    conI.Protocol = Connection.Protocol.Protocols.VNC
                                    conI.Port = .Attributes("VNCPort").Value
                                Else
                                    conI.Protocol = Connection.Protocol.Protocols.RDP
                                End If
                            End If
                        Else
                            conI.Port = Connection.Protocol.RDP.Defaults.Port
                            conI.Protocol = Connection.Protocol.Protocols.RDP
                        End If

                        If Me.confVersion > 0.3 Then '0.4
                            If Me.confVersion < 0.7 Then
                                If CType(.Attributes("UseVNC").Value, Boolean) = True Then
                                    conI.Port = .Attributes("VNCPort").Value
                                Else
                                    conI.Port = .Attributes("RDPPort").Value
                                End If
                            End If

                            conI.UseConsoleSession = .Attributes("ConnectToConsole").Value
                        Else
                            If Me.confVersion < 0.7 Then
                                If CType(.Attributes("UseVNC").Value, Boolean) = True Then
                                    conI.Port = Connection.Protocol.VNC.Defaults.Port
                                Else
                                    conI.Port = Connection.Protocol.RDP.Defaults.Port
                                End If
                            End If
                            conI.UseConsoleSession = False
                        End If

                        If Me.confVersion > 0.4 Then '0.5 and 0.6
                            conI.RedirectDiskDrives = .Attributes("RedirectDiskDrives").Value
                            conI.RedirectPrinters = .Attributes("RedirectPrinters").Value
                            conI.RedirectPorts = .Attributes("RedirectPorts").Value
                            conI.RedirectSmartCards = .Attributes("RedirectSmartCards").Value
                        Else
                            conI.RedirectDiskDrives = False
                            conI.RedirectPrinters = False
                            conI.RedirectPorts = False
                            conI.RedirectSmartCards = False
                        End If

                        If Me.confVersion > 0.6 Then '0.7
                            conI.Protocol = Tools.Misc.StringToEnum(GetType(Connection.Protocol.Protocols), .Attributes("Protocol").Value)
                            conI.Port = .Attributes("Port").Value
                        End If

                        If Me.confVersion > 0.9 Then '1.0
                            conI.RedirectKeys = .Attributes("RedirectKeys").Value
                        End If

                        If Me.confVersion > 1.1 Then '1.2
                            conI.PuttySession = .Attributes("PuttySession").Value
                        End If

                        If Me.confVersion > 1.2 Then '1.3
                            conI.Colors = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPColors), .Attributes("Colors").Value)
                            conI.Resolution = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPResolutions), .Attributes("Resolution").Value)
                            conI.RedirectSound = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPSounds), .Attributes("RedirectSound").Value)
                        Else
                            Select Case .Attributes("Colors").Value
                                Case 0
                                    conI.Colors = Connection.Protocol.RDP.RDPColors.Colors256
                                Case 1
                                    conI.Colors = Connection.Protocol.RDP.RDPColors.Colors16Bit
                                Case 2
                                    conI.Colors = Connection.Protocol.RDP.RDPColors.Colors24Bit
                                Case 3
                                    conI.Colors = Connection.Protocol.RDP.RDPColors.Colors32Bit
                                Case 4
                                    conI.Colors = Connection.Protocol.RDP.RDPColors.Colors15Bit
                            End Select

                            conI.RedirectSound = .Attributes("RedirectSound").Value
                        End If

                        If Me.confVersion > 1.2 Then '1.3
                            conI.Inherit = New Connection.Info.Inheritance(conI)
                            conI.Inherit.CacheBitmaps = .Attributes("InheritCacheBitmaps").Value
                            conI.Inherit.Colors = .Attributes("InheritColors").Value
                            conI.Inherit.Description = .Attributes("InheritDescription").Value
                            conI.Inherit.DisplayThemes = .Attributes("InheritDisplayThemes").Value
                            conI.Inherit.DisplayWallpaper = .Attributes("InheritDisplayWallpaper").Value
                            conI.Inherit.Domain = .Attributes("InheritDomain").Value
                            conI.Inherit.Icon = .Attributes("InheritIcon").Value
                            conI.Inherit.Panel = .Attributes("InheritPanel").Value
                            conI.Inherit.Password = .Attributes("InheritPassword").Value
                            conI.Inherit.Port = .Attributes("InheritPort").Value
                            conI.Inherit.Protocol = .Attributes("InheritProtocol").Value
                            conI.Inherit.PuttySession = .Attributes("InheritPuttySession").Value
                            conI.Inherit.RedirectDiskDrives = .Attributes("InheritRedirectDiskDrives").Value
                            conI.Inherit.RedirectKeys = .Attributes("InheritRedirectKeys").Value
                            conI.Inherit.RedirectPorts = .Attributes("InheritRedirectPorts").Value
                            conI.Inherit.RedirectPrinters = .Attributes("InheritRedirectPrinters").Value
                            conI.Inherit.RedirectSmartCards = .Attributes("InheritRedirectSmartCards").Value
                            conI.Inherit.RedirectSound = .Attributes("InheritRedirectSound").Value
                            conI.Inherit.Resolution = .Attributes("InheritResolution").Value
                            conI.Inherit.UseConsoleSession = .Attributes("InheritUseConsoleSession").Value
                            conI.Inherit.Username = .Attributes("InheritUsername").Value

                            conI.Icon = .Attributes("Icon").Value
                            conI.Panel = .Attributes("Panel").Value
                        Else
                            conI.Inherit = New Connection.Info.Inheritance(conI, .Attributes("Inherit").Value)

                            conI.Icon = .Attributes("Icon").Value.Replace(".ico", "")
                            conI.Panel = My.Resources.strGeneral
                        End If

                        If Me.confVersion > 1.4 Then '1.5
                            conI.PleaseConnect = .Attributes("Connected").Value
                        End If

                        If Me.confVersion > 1.5 Then '1.6
                            conI.ICAEncryption = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.ICA.EncryptionStrength), .Attributes("ICAEncryptionStrength").Value)
                            conI.Inherit.ICAEncryption = .Attributes("InheritICAEncryptionStrength").Value

                            conI.PreExtApp = .Attributes("PreExtApp").Value
                            conI.PostExtApp = .Attributes("PostExtApp").Value
                            conI.Inherit.PreExtApp = .Attributes("InheritPreExtApp").Value
                            conI.Inherit.PostExtApp = .Attributes("InheritPostExtApp").Value
                        End If

                        If Me.confVersion > 1.6 Then '1.7
                            conI.VNCCompression = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.Compression), .Attributes("VNCCompression").Value)
                            conI.VNCEncoding = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.Encoding), .Attributes("VNCEncoding").Value)
                            conI.VNCAuthMode = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.AuthMode), .Attributes("VNCAuthMode").Value)
                            conI.VNCProxyType = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.ProxyType), .Attributes("VNCProxyType").Value)
                            conI.VNCProxyIP = .Attributes("VNCProxyIP").Value
                            conI.VNCProxyPort = .Attributes("VNCProxyPort").Value
                            conI.VNCProxyUsername = .Attributes("VNCProxyUsername").Value
                            conI.VNCProxyPassword = Security.Crypt.Decrypt(.Attributes("VNCProxyPassword").Value, pW)
                            conI.VNCColors = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.Colors), .Attributes("VNCColors").Value)
                            conI.VNCSmartSizeMode = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.VNC.SmartSizeMode), .Attributes("VNCSmartSizeMode").Value)
                            conI.VNCViewOnly = .Attributes("VNCViewOnly").Value

                            conI.Inherit.VNCCompression = .Attributes("InheritVNCCompression").Value
                            conI.Inherit.VNCEncoding = .Attributes("InheritVNCEncoding").Value
                            conI.Inherit.VNCAuthMode = .Attributes("InheritVNCAuthMode").Value
                            conI.Inherit.VNCProxyType = .Attributes("InheritVNCProxyType").Value
                            conI.Inherit.VNCProxyIP = .Attributes("InheritVNCProxyIP").Value
                            conI.Inherit.VNCProxyPort = .Attributes("InheritVNCProxyPort").Value
                            conI.Inherit.VNCProxyUsername = .Attributes("InheritVNCProxyUsername").Value
                            conI.Inherit.VNCProxyPassword = .Attributes("InheritVNCProxyPassword").Value
                            conI.Inherit.VNCColors = .Attributes("InheritVNCColors").Value
                            conI.Inherit.VNCSmartSizeMode = .Attributes("InheritVNCSmartSizeMode").Value
                            conI.Inherit.VNCViewOnly = .Attributes("InheritVNCViewOnly").Value
                        End If

                        If Me.confVersion > 1.7 Then '1.8
                            conI.RDPAuthenticationLevel = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.RDP.AuthenticationLevel), .Attributes("RDPAuthenticationLevel").Value)

                            conI.Inherit.RDPAuthenticationLevel = .Attributes("InheritRDPAuthenticationLevel").Value
                        End If

                        If Me.confVersion > 1.8 Then '1.9
                            conI.RenderingEngine = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.HTTPBase.RenderingEngine), .Attributes("RenderingEngine").Value)
                            conI.MacAddress = .Attributes("MacAddress").Value

                            conI.Inherit.RenderingEngine = .Attributes("InheritRenderingEngine").Value
                            conI.Inherit.MacAddress = .Attributes("InheritMacAddress").Value
                        End If

                        If Me.confVersion > 1.9 Then '2.0
                            conI.UserField = .Attributes("UserField").Value
                            conI.Inherit.UserField = .Attributes("InheritUserField").Value
                        End If

                        If Me.confVersion > 2.0 Then '2.1
                            conI.ExtApp = .Attributes("ExtApp").Value
                            conI.Inherit.ExtApp = .Attributes("InheritExtApp").Value
                        End If

                        If Me.confVersion > 2.1 Then '2.2
                            ' Get settings
                            conI.RDGatewayUsageMethod = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.RDP.RDGatewayUsageMethod), .Attributes("RDGatewayUsageMethod").Value)
                            conI.RDGatewayHostname = .Attributes("RDGatewayHostname").Value
                            conI.RDGatewayUseConnectionCredentials = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.RDP.RDGatewayUseConnectionCredentials), .Attributes("RDGatewayUseConnectionCredentials").Value)
                            conI.RDGatewayUsername = .Attributes("RDGatewayUsername").Value
                            conI.RDGatewayPassword = Security.Crypt.Decrypt(.Attributes("RDGatewayPassword").Value, pW)
                            conI.RDGatewayDomain = .Attributes("RDGatewayDomain").Value

                            ' Get inheritance settings
                            conI.Inherit.RDGatewayUsageMethod = .Attributes("InheritRDGatewayUsageMethod").Value
                            conI.Inherit.RDGatewayHostname = .Attributes("InheritRDGatewayHostname").Value
                            conI.Inherit.RDGatewayUseConnectionCredentials = .Attributes("InheritRDGatewayUseConnectionCredentials").Value
                            conI.Inherit.RDGatewayUsername = .Attributes("InheritRDGatewayUsername").Value
                            conI.Inherit.RDGatewayPassword = .Attributes("InheritRDGatewayPassword").Value
                            conI.Inherit.RDGatewayDomain = .Attributes("InheritRDGatewayDomain").Value
                        End If

                        If Me.confVersion > 2.2 Then '2.3
                            ' Get settings
                            conI.EnableFontSmoothing = .Attributes("EnableFontSmoothing").Value
                            conI.EnableDesktopComposition = .Attributes("EnableDesktopComposition").Value

                            ' Get inheritance settings
                            conI.Inherit.EnableFontSmoothing = .Attributes("InheritEnableFontSmoothing").Value
                            conI.Inherit.EnableDesktopComposition = .Attributes("InheritEnableDesktopComposition").Value
                        End If
                    End With
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, String.Format(My.Resources.strGetConnectionInfoFromXmlFailed, conI.Name, Me.ConnectionFileName, ex.Message), False)
                End Try
                Return conI
            End Function
#End Region

#Region "Save"
            Private Sub EncryptCompleteFile()
                Dim sRd As New StreamReader(ConnectionFileName)

                Dim strCons As String
                strCons = sRd.ReadToEnd
                sRd.Close()

                If strCons <> "" Then
                    Dim sWr As New StreamWriter(ConnectionFileName)
                    sWr.Write(Security.Crypt.Encrypt(strCons, pW))
                    sWr.Close()
                End If
            End Sub

            Private Sub SaveToXML()
                Try
                    If App.Runtime.IsConnectionsFileLoaded = False Then
                        Exit Sub
                    End If

                    Dim tN As TreeNode
                    Dim exp As Boolean = False

                    If Tree.Node.GetNodeType(RootTreeNode) = Tree.Node.Type.Root Then
                        tN = RootTreeNode.Clone
                    Else
                        tN = New TreeNode("mR|Export (" + Tools.Misc.DBDate(Now) + ")")
                        tN.Nodes.Add(RootTreeNode.Clone)
                        exp = True
                    End If

                    xW = New XmlTextWriter(ConnectionFileName, System.Text.Encoding.UTF8)

                    xW.Formatting = Formatting.Indented
                    xW.Indentation = 4

                    xW.WriteStartDocument()

                    xW.WriteStartElement("Connections")
                    xW.WriteAttributeString("Name", "", tN.Text)
                    xW.WriteAttributeString("Export", "", exp)

                    If exp Then
                        xW.WriteAttributeString("Protected", "", Security.Crypt.Encrypt("ThisIsNotProtected", pW))
                    Else
                        If TryCast(tN.Tag, mRemoteNG.Root.Info).Password = True Then
                            pW = TryCast(tN.Tag, mRemoteNG.Root.Info).PasswordString
                            xW.WriteAttributeString("Protected", "", Security.Crypt.Encrypt("ThisIsProtected", pW))
                        Else
                            xW.WriteAttributeString("Protected", "", Security.Crypt.Encrypt("ThisIsNotProtected", pW))
                        End If
                    End If

                    Dim enCulture As New CultureInfo("en-US")
                    xW.WriteAttributeString("ConfVersion", "", App.Info.Connections.ConnectionFileVersion.ToString(enCulture))

                    Dim tNC As TreeNodeCollection
                    tNC = tN.Nodes

                    saveNode(tNC)

                    xW.WriteEndElement()
                    xW.Close()
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "SaveToXML failed" & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Sub saveNode(ByVal tNC As TreeNodeCollection)
                Try
                    For Each node As TreeNode In tNC
                        Dim curConI As Connection.Info

                        If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Or Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then
                            xW.WriteStartElement("Node")
                            xW.WriteAttributeString("Name", "", node.Text)
                            xW.WriteAttributeString("Type", "", Tree.Node.GetNodeType(node).ToString)
                        End If

                        If Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then 'container
                            xW.WriteAttributeString("Expanded", "", ContainerList(node.Tag).TreeNode.IsExpanded)
                            curConI = ContainerList(node.Tag).ConnectionInfo
                            SaveConnectionFields(curConI)
                            saveNode(node.Nodes)
                            xW.WriteEndElement()
                        End If

                        If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Then
                            curConI = ConnectionList(node.Tag)
                            SaveConnectionFields(curConI)
                            xW.WriteEndElement()
                        End If
                    Next
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "saveNode failed" & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Sub SaveConnectionFields(ByVal curConI As Connection.Info)
                Try
                    xW.WriteAttributeString("Descr", "", curConI.Description)

                    xW.WriteAttributeString("Icon", "", curConI.Icon)

                    xW.WriteAttributeString("Panel", "", curConI.Panel)

                    If SaveSecurity.Username = True Then
                        xW.WriteAttributeString("Username", "", curConI.Username)
                    Else
                        xW.WriteAttributeString("Username", "", "")
                    End If

                    If SaveSecurity.Domain = True Then
                        xW.WriteAttributeString("Domain", "", curConI.Domain)
                    Else
                        xW.WriteAttributeString("Domain", "", "")
                    End If

                    If SaveSecurity.Password = True Then
                        xW.WriteAttributeString("Password", "", Security.Crypt.Encrypt(curConI.Password, pW))
                    Else
                        xW.WriteAttributeString("Password", "", "")
                    End If

                    xW.WriteAttributeString("Hostname", "", curConI.Hostname)

                    xW.WriteAttributeString("Protocol", "", curConI.Protocol.ToString)

                    xW.WriteAttributeString("PuttySession", "", curConI.PuttySession)

                    xW.WriteAttributeString("Port", "", curConI.Port)

                    xW.WriteAttributeString("ConnectToConsole", "", curConI.UseConsoleSession)

                    xW.WriteAttributeString("RenderingEngine", "", curConI.RenderingEngine.ToString)

                    xW.WriteAttributeString("ICAEncryptionStrength", "", curConI.ICAEncryption.ToString)

                    xW.WriteAttributeString("RDPAuthenticationLevel", "", curConI.RDPAuthenticationLevel.ToString)

                    xW.WriteAttributeString("Colors", "", curConI.Colors.ToString)

                    xW.WriteAttributeString("Resolution", "", curConI.Resolution.ToString)

                    xW.WriteAttributeString("DisplayWallpaper", "", curConI.DisplayWallpaper)

                    xW.WriteAttributeString("DisplayThemes", "", curConI.DisplayThemes)

                    xW.WriteAttributeString("EnableFontSmoothing", "", curConI.EnableFontSmoothing)

                    xW.WriteAttributeString("EnableDesktopComposition", "", curConI.EnableDesktopComposition)

                    xW.WriteAttributeString("CacheBitmaps", "", curConI.CacheBitmaps)

                    xW.WriteAttributeString("RedirectDiskDrives", "", curConI.RedirectDiskDrives)

                    xW.WriteAttributeString("RedirectPorts", "", curConI.RedirectPorts)

                    xW.WriteAttributeString("RedirectPrinters", "", curConI.RedirectPrinters)

                    xW.WriteAttributeString("RedirectSmartCards", "", curConI.RedirectSmartCards)

                    xW.WriteAttributeString("RedirectSound", "", curConI.RedirectSound.ToString)

                    xW.WriteAttributeString("RedirectKeys", "", curConI.RedirectKeys)

                    If curConI.OpenConnections.Count > 0 Then
                        xW.WriteAttributeString("Connected", "", True)
                    Else
                        xW.WriteAttributeString("Connected", "", False)
                    End If

                    xW.WriteAttributeString("PreExtApp", "", curConI.PreExtApp)
                    xW.WriteAttributeString("PostExtApp", "", curConI.PostExtApp)
                    xW.WriteAttributeString("MacAddress", "", curConI.MacAddress)
                    xW.WriteAttributeString("UserField", "", curConI.UserField)
                    xW.WriteAttributeString("ExtApp", "", curConI.ExtApp)

                    xW.WriteAttributeString("VNCCompression", "", curConI.VNCCompression.ToString)
                    xW.WriteAttributeString("VNCEncoding", "", curConI.VNCEncoding.ToString)
                    xW.WriteAttributeString("VNCAuthMode", "", curConI.VNCAuthMode.ToString)
                    xW.WriteAttributeString("VNCProxyType", "", curConI.VNCProxyType.ToString)
                    xW.WriteAttributeString("VNCProxyIP", "", curConI.VNCProxyIP)
                    xW.WriteAttributeString("VNCProxyPort", "", curConI.VNCProxyPort)
                    xW.WriteAttributeString("VNCProxyUsername", "", curConI.VNCProxyUsername)
                    xW.WriteAttributeString("VNCProxyPassword", "", Security.Crypt.Encrypt(curConI.VNCProxyPassword, pW))
                    xW.WriteAttributeString("VNCColors", "", curConI.VNCColors.ToString)
                    xW.WriteAttributeString("VNCSmartSizeMode", "", curConI.VNCSmartSizeMode.ToString)
                    xW.WriteAttributeString("VNCViewOnly", "", curConI.VNCViewOnly)

                    xW.WriteAttributeString("RDGatewayUsageMethod", "", curConI.RDGatewayUsageMethod.ToString)
                    xW.WriteAttributeString("RDGatewayHostname", "", curConI.RDGatewayHostname)

                    xW.WriteAttributeString("RDGatewayUseConnectionCredentials", "", curConI.RDGatewayUseConnectionCredentials.ToString)

                    If SaveSecurity.Username = True Then
                        xW.WriteAttributeString("RDGatewayUsername", "", curConI.RDGatewayUsername)
                    Else
                        xW.WriteAttributeString("RDGatewayUsername", "", "")
                    End If

                    If SaveSecurity.Password = True Then
                        xW.WriteAttributeString("RDGatewayPassword", "", curConI.RDGatewayPassword)
                    Else
                        xW.WriteAttributeString("RDGatewayPassword", "", "")
                    End If

                    If SaveSecurity.Domain = True Then
                        xW.WriteAttributeString("RDGatewayDomain", "", curConI.RDGatewayDomain)
                    Else
                        xW.WriteAttributeString("RDGatewayDomain", "", "")
                    End If

                    If SaveSecurity.Inheritance = True Then
                        xW.WriteAttributeString("InheritCacheBitmaps", "", curConI.Inherit.CacheBitmaps)
                        xW.WriteAttributeString("InheritColors", "", curConI.Inherit.Colors)
                        xW.WriteAttributeString("InheritDescription", "", curConI.Inherit.Description)
                        xW.WriteAttributeString("InheritDisplayThemes", "", curConI.Inherit.DisplayThemes)
                        xW.WriteAttributeString("InheritDisplayWallpaper", "", curConI.Inherit.DisplayWallpaper)
                        xW.WriteAttributeString("InheritEnableFontSmoothing", "", curConI.Inherit.EnableFontSmoothing)
                        xW.WriteAttributeString("InheritEnableDesktopComposition", "", curConI.Inherit.EnableDesktopComposition)
                        xW.WriteAttributeString("InheritDomain", "", curConI.Inherit.Domain)
                        xW.WriteAttributeString("InheritIcon", "", curConI.Inherit.Icon)
                        xW.WriteAttributeString("InheritPanel", "", curConI.Inherit.Panel)
                        xW.WriteAttributeString("InheritPassword", "", curConI.Inherit.Password)
                        xW.WriteAttributeString("InheritPort", "", curConI.Inherit.Port)
                        xW.WriteAttributeString("InheritProtocol", "", curConI.Inherit.Protocol)
                        xW.WriteAttributeString("InheritPuttySession", "", curConI.Inherit.PuttySession)
                        xW.WriteAttributeString("InheritRedirectDiskDrives", "", curConI.Inherit.RedirectDiskDrives)
                        xW.WriteAttributeString("InheritRedirectKeys", "", curConI.Inherit.RedirectKeys)
                        xW.WriteAttributeString("InheritRedirectPorts", "", curConI.Inherit.RedirectPorts)
                        xW.WriteAttributeString("InheritRedirectPrinters", "", curConI.Inherit.RedirectPrinters)
                        xW.WriteAttributeString("InheritRedirectSmartCards", "", curConI.Inherit.RedirectSmartCards)
                        xW.WriteAttributeString("InheritRedirectSound", "", curConI.Inherit.RedirectSound)
                        xW.WriteAttributeString("InheritResolution", "", curConI.Inherit.Resolution)
                        xW.WriteAttributeString("InheritUseConsoleSession", "", curConI.Inherit.UseConsoleSession)
                        xW.WriteAttributeString("InheritRenderingEngine", "", curConI.Inherit.RenderingEngine)
                        xW.WriteAttributeString("InheritUsername", "", curConI.Inherit.Username)
                        xW.WriteAttributeString("InheritICAEncryptionStrength", "", curConI.Inherit.ICAEncryption)
                        xW.WriteAttributeString("InheritRDPAuthenticationLevel", "", curConI.Inherit.RDPAuthenticationLevel)
                        xW.WriteAttributeString("InheritPreExtApp", "", curConI.Inherit.PreExtApp)
                        xW.WriteAttributeString("InheritPostExtApp", "", curConI.Inherit.PostExtApp)
                        xW.WriteAttributeString("InheritMacAddress", "", curConI.Inherit.MacAddress)
                        xW.WriteAttributeString("InheritUserField", "", curConI.Inherit.UserField)
                        xW.WriteAttributeString("InheritExtApp", "", curConI.Inherit.ExtApp)
                        xW.WriteAttributeString("InheritVNCCompression", "", curConI.Inherit.VNCCompression)
                        xW.WriteAttributeString("InheritVNCEncoding", "", curConI.Inherit.VNCEncoding)
                        xW.WriteAttributeString("InheritVNCAuthMode", "", curConI.Inherit.VNCAuthMode)
                        xW.WriteAttributeString("InheritVNCProxyType", "", curConI.Inherit.VNCProxyType)
                        xW.WriteAttributeString("InheritVNCProxyIP", "", curConI.Inherit.VNCProxyIP)
                        xW.WriteAttributeString("InheritVNCProxyPort", "", curConI.Inherit.VNCProxyPort)
                        xW.WriteAttributeString("InheritVNCProxyUsername", "", curConI.Inherit.VNCProxyUsername)
                        xW.WriteAttributeString("InheritVNCProxyPassword", "", curConI.Inherit.VNCProxyPassword)
                        xW.WriteAttributeString("InheritVNCColors", "", curConI.Inherit.VNCColors)
                        xW.WriteAttributeString("InheritVNCSmartSizeMode", "", curConI.Inherit.VNCSmartSizeMode)
                        xW.WriteAttributeString("InheritVNCViewOnly", "", curConI.Inherit.VNCViewOnly)
                        xW.WriteAttributeString("InheritRDGatewayUsageMethod", "", curConI.Inherit.RDGatewayUsageMethod)
                        xW.WriteAttributeString("InheritRDGatewayHostname", "", curConI.Inherit.RDGatewayHostname)
                        xW.WriteAttributeString("InheritRDGatewayUseConnectionCredentials", "", curConI.Inherit.RDGatewayUseConnectionCredentials)
                        xW.WriteAttributeString("InheritRDGatewayUsername", "", curConI.Inherit.RDGatewayUsername)
                        xW.WriteAttributeString("InheritRDGatewayPassword", "", curConI.Inherit.RDGatewayPassword)
                        xW.WriteAttributeString("InheritRDGatewayDomain", "", curConI.Inherit.RDGatewayDomain)
                    Else
                        xW.WriteAttributeString("InheritCacheBitmaps", "", False)
                        xW.WriteAttributeString("InheritColors", "", False)
                        xW.WriteAttributeString("InheritDescription", "", False)
                        xW.WriteAttributeString("InheritDisplayThemes", "", False)
                        xW.WriteAttributeString("InheritDisplayWallpaper", "", False)
                        xW.WriteAttributeString("InheritEnableFontSmoothing", "", False)
                        xW.WriteAttributeString("InheritEnableDesktopComposition", "", False)
                        xW.WriteAttributeString("InheritDomain", "", False)
                        xW.WriteAttributeString("InheritIcon", "", False)
                        xW.WriteAttributeString("InheritPanel", "", False)
                        xW.WriteAttributeString("InheritPassword", "", False)
                        xW.WriteAttributeString("InheritPort", "", False)
                        xW.WriteAttributeString("InheritProtocol", "", False)
                        xW.WriteAttributeString("InheritPuttySession", "", False)
                        xW.WriteAttributeString("InheritRedirectDiskDrives", "", False)
                        xW.WriteAttributeString("InheritRedirectKeys", "", False)
                        xW.WriteAttributeString("InheritRedirectPorts", "", False)
                        xW.WriteAttributeString("InheritRedirectPrinters", "", False)
                        xW.WriteAttributeString("InheritRedirectSmartCards", "", False)
                        xW.WriteAttributeString("InheritRedirectSound", "", False)
                        xW.WriteAttributeString("InheritResolution", "", False)
                        xW.WriteAttributeString("InheritUseConsoleSession", "", False)
                        xW.WriteAttributeString("InheritRenderingEngine", "", False)
                        xW.WriteAttributeString("InheritUsername", "", False)
                        xW.WriteAttributeString("InheritICAEncryptionStrength", "", False)
                        xW.WriteAttributeString("InheritRDPAuthenticationLevel", "", False)
                        xW.WriteAttributeString("InheritPreExtApp", "", False)
                        xW.WriteAttributeString("InheritPostExtApp", "", False)
                        xW.WriteAttributeString("InheritMacAddress", "", False)
                        xW.WriteAttributeString("InheritUserField", "", False)
                        xW.WriteAttributeString("InheritExtApp", "", False)
                        xW.WriteAttributeString("InheritVNCCompression", "", False)
                        xW.WriteAttributeString("InheritVNCEncoding", "", False)
                        xW.WriteAttributeString("InheritVNCAuthMode", "", False)
                        xW.WriteAttributeString("InheritVNCProxyType", "", False)
                        xW.WriteAttributeString("InheritVNCProxyIP", "", False)
                        xW.WriteAttributeString("InheritVNCProxyPort", "", False)
                        xW.WriteAttributeString("InheritVNCProxyUsername", "", False)
                        xW.WriteAttributeString("InheritVNCProxyPassword", "", False)
                        xW.WriteAttributeString("InheritVNCColors", "", False)
                        xW.WriteAttributeString("InheritVNCSmartSizeMode", "", False)
                        xW.WriteAttributeString("InheritVNCViewOnly", "", False)
                        xW.WriteAttributeString("InheritRDGatewayHostname", "", False)
                        xW.WriteAttributeString("InheritRDGatewayUseConnectionCredentials", "", False)
                        xW.WriteAttributeString("InheritRDGatewayUsername", "", False)
                        xW.WriteAttributeString("InheritRDGatewayPassword", "", False)
                        xW.WriteAttributeString("InheritRDGatewayDomain", "", False)
                    End If
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "SaveConnectionFields failed" & vbNewLine & ex.Message, True)
                End Try
            End Sub
#End Region
#End Region
        End Class
    End Namespace
End Namespace

