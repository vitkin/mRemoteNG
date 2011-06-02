Imports System.Globalization
Imports System.Data.SqlClient
Imports mRemoteNG.App.Runtime

Namespace Config
    Namespace Connections
        Public Class MSSQL
            Inherits Base
#Region "Public Properties"
            Private _UseSQL As Boolean
            Public Property UseSQL() As Boolean
                Get
                    Return _UseSQL
                End Get
                Set(ByVal value As Boolean)
                    _UseSQL = value
                End Set
            End Property

            Private _SQLHost As String
            Public Property SQLHost() As String
                Get
                    Return _SQLHost
                End Get
                Set(ByVal value As String)
                    _SQLHost = value
                End Set
            End Property

            Private _SQLDatabaseName As String
            Public Property SQLDatabaseName() As String
                Get
                    Return _SQLDatabaseName
                End Get
                Set(ByVal value As String)
                    _SQLDatabaseName = value
                End Set
            End Property

            Private _SQLUsername As String
            Public Property SQLUsername() As String
                Get
                    Return _SQLUsername
                End Get
                Set(ByVal value As String)
                    _SQLUsername = value
                End Set
            End Property

            Private _SQLPassword As String
            Public Property SQLPassword() As String
                Get
                    Return _SQLPassword
                End Get
                Set(ByVal value As String)
                    _SQLPassword = value
                End Set
            End Property

            Private _SQLUpdate As Boolean
            Public Property SQLUpdate() As Boolean
                Get
                    Return _SQLUpdate
                End Get
                Set(ByVal value As Boolean)
                    _SQLUpdate = value
                End Set
            End Property
#End Region

#Region "Public Members"
            Public Overrides Function Load() As Boolean
                LoadFromSQL()
                SetMainFormText("SQL Server")
            End Function

            Public Overrides Function Save() As Boolean
                SaveToSQL()
                SetMainFormText("SQL Server")
            End Function

            Public Sub TestConnection()
                Dim sqlRd As SqlDataReader = Nothing

                Try
                    SqlConnection.Open()

                    Dim sqlQuery As SqlCommand = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                    sqlRd = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                    sqlRd.Read()

                    sqlRd.Close()
                    SqlConnection.Close()

                    MessageBox.Show(frmMain, "The database connection was successful.", "SQL Server Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(frmMain, ex.Message, "SQL Server Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    If sqlRd IsNot Nothing Then
                        If Not sqlRd.IsClosed Then sqlRd.Close()
                    End If
                    If SqlConnection IsNot Nothing Then
                        If Not SqlConnection.State = ConnectionState.Closed Then SqlConnection.Close()
                    End If
                End Try
            End Sub

            Public Sub CreateDatabase()
                ' TODO: Do stuff
            End Sub

            Public Sub MigrateDatabase()
                If Not OpenConnection() Then

                End If

                CloseConnection()
            End Sub
#End Region

#Region "Private Properties"
            Private _SqlConnection As SqlConnection = Nothing
            Private ReadOnly Property SqlConnection() As SqlConnection
                Get
                    If _SqlConnection IsNot Nothing Then Return _SqlConnection

                    Dim stringBuilder As New SqlConnectionStringBuilder
                    stringBuilder.DataSource = SQLHost
                    stringBuilder.InitialCatalog = SQLDatabaseName
                    If String.IsNullOrEmpty(SQLUsername) Then
                        stringBuilder.IntegratedSecurity = True
                    Else
                        stringBuilder.UserID = SQLUsername
                        stringBuilder.Password = SQLPassword
                    End If

                    _SqlConnection = New SqlConnection(stringBuilder.ConnectionString)
                    Return _SqlConnection
                End Get
            End Property

            Private Property DatabaseVersion() As System.Version
                Get
                    OpenConnection()

                    sqlQuery = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                    sqlRd = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)
                    sqlRd.Read()

                    Dim version As System.Version = sqlRd.Item("confVersion")

                    CloseConnection()

                    Return version
                End Get
                Set(ByVal value)

                End Set
            End Property
#End Region

#Region "Private Variables"
            Private confVersion As Double

            Private sqlQuery As SqlCommand
            Private sqlRd As SqlDataReader

            Private gIndex As Integer = 0
            Private parentID As String = 0
#End Region

#Region "Private Methods"

#Region "SQL Server Connection"
            Private Function OpenConnection() As Boolean
                If SqlConnection Is Nothing Then Return False

                ' Already open?
                If Not (SqlConnection.State = ConnectionState.Closed Or SqlConnection.State = ConnectionState.Broken) Then Return True

                Try
                    SqlConnection.Open()
                    Return True
                Catch ex As System.Exception
                    Return False
                End Try
            End Function

            Private Sub CloseConnection()
                If SqlConnection Is Nothing Then Return
                If SqlConnection.State = ConnectionState.Closed Then Return
                SqlConnection.Close()
            End Sub
#End Region

#Region "Load"
            Private Sub LoadFromSQL()
                Try
                    App.Runtime.ConnectionsFileLoaded = False

                    SqlConnection.Open()

                    sqlQuery = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                    sqlRd = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                    sqlRd.Read()

                    If sqlRd.HasRows = False Then
                        App.Runtime.SaveConnections()

                        sqlQuery = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                        sqlRd = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                        sqlRd.Read()
                    End If

                    Dim enCulture As CultureInfo = New CultureInfo("en-US")
                    Me.confVersion = Convert.ToDouble(sqlRd.Item("confVersion"), enCulture)

                    Dim rootNode As TreeNode
                    rootNode = New TreeNode(sqlRd.Item("Name"))

                    Dim rInfo As New Root.Info(Root.Info.RootType.Connection)
                    rInfo.Name = rootNode.Text
                    rInfo.TreeNode = rootNode

                    rootNode.Tag = rInfo
                    rootNode.ImageIndex = Images.Enums.TreeImage.Root
                    rootNode.SelectedImageIndex = Images.Enums.TreeImage.Root

                    If Security.Crypt.Decrypt(sqlRd.Item("Protected"), pW) <> "ThisIsNotProtected" Then
                        If Authenticate(sqlRd.Item("Protected"), False, rInfo) = False Then
                            My.Settings.LoadConsFromCustomLocation = False
                            My.Settings.CustomConsPath = ""
                            rootNode.Remove()
                            Exit Sub
                        End If
                    End If

                    'Me._RootTreeNode.Text = rootNode.Text
                    'Me._RootTreeNode.Tag = rootNode.Tag
                    'Me._RootTreeNode.ImageIndex = Images.Enums.TreeImage.Root
                    'Me._RootTreeNode.SelectedImageIndex = Images.Enums.TreeImage.Root

                    sqlRd.Close()

                    ' SECTION 3. Populate the TreeView with the DOM nodes.
                    AddNodesFromSQL(rootNode)
                    'AddNodeFromXML(xDom.DocumentElement, Me._RootTreeNode)

                    rootNode.Expand()

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

                    'Tree.Node.TreeView.Nodes.Clear()
                    'Tree.Node.TreeView.Nodes.Add(rootNode)

                    AddNodeToTree(rootNode)
                    SetSelectedNode(selNode)

                    App.Runtime.ConnectionsFileLoaded = True
                    'App.Runtime.Windows.treeForm.InitialRefresh()

                    SqlConnection.Close()
                Catch ex As Exception
                    mC.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strLoadFromSqlFailed & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Delegate Sub AddNodeToTreeCB(ByVal TreeNode As TreeNode)
            Private Sub AddNodeToTree(ByVal TreeNode As TreeNode)
                If Tree.Node.TreeView.InvokeRequired Then
                    Dim d As New AddNodeToTreeCB(AddressOf AddNodeToTree)
                    App.Runtime.Windows.treeForm.Invoke(d, New Object() {TreeNode})
                Else
                    App.Runtime.Windows.treeForm.tvConnections.Nodes.Clear()
                    App.Runtime.Windows.treeForm.tvConnections.Nodes.Add(TreeNode)
                    App.Runtime.Windows.treeForm.InitialRefresh()
                End If
            End Sub

            Private Delegate Sub SetSelectedNodeCB(ByVal TreeNode As TreeNode)
            Private Sub SetSelectedNode(ByVal TreeNode As TreeNode)
                If Tree.Node.TreeView.InvokeRequired Then
                    Dim d As New SetSelectedNodeCB(AddressOf SetSelectedNode)
                    App.Runtime.Windows.treeForm.Invoke(d, New Object() {TreeNode})
                Else
                    App.Runtime.Windows.treeForm.tvConnections.SelectedNode = TreeNode
                End If
            End Sub

            Private Sub AddNodesFromSQL(ByVal baseNode As TreeNode)
                Try
                    SqlConnection.Open()
                    sqlQuery = New SqlCommand("SELECT * FROM tblCons ORDER BY PositionID ASC", SqlConnection)
                    sqlRd = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                    If sqlRd.HasRows = False Then
                        Exit Sub
                    End If

                    Dim tNode As TreeNode

                    While sqlRd.Read
                        tNode = New TreeNode(sqlRd.Item("Name"))
                        'baseNode.Nodes.Add(tNode)

                        If Tree.Node.GetNodeTypeFromString(sqlRd.Item("Type")) = Tree.Node.Type.Connection Then
                            Dim conI As Connection.Info = GetConnectionInfoFromSQL()
                            conI.TreeNode = tNode
                            'conI.Parent = prevCont 'NEW

                            ConnectionList.Add(conI)

                            tNode.Tag = conI

                            If SQLUpdate = True Then
                                Dim prevCon As Connection.Info = PreviousConnectionList.FindByConstantID(conI.ConstantID)

                                If prevCon IsNot Nothing Then
                                    For Each prot As Connection.Protocol.Base In prevCon.OpenConnections
                                        prot.InterfaceControl.Info = conI
                                        conI.OpenConnections.Add(prot)
                                    Next

                                    If conI.OpenConnections.Count > 0 Then
                                        tNode.ImageIndex = Images.Enums.TreeImage.ConnectionOpen
                                        tNode.SelectedImageIndex = Images.Enums.TreeImage.ConnectionOpen
                                    Else
                                        tNode.ImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                        tNode.SelectedImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                    End If
                                Else
                                    tNode.ImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                    tNode.SelectedImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                End If

                                If conI.ConstantID = PreviousSelected Then
                                    selNode = tNode
                                End If
                            Else
                                tNode.ImageIndex = Images.Enums.TreeImage.ConnectionClosed
                                tNode.SelectedImageIndex = Images.Enums.TreeImage.ConnectionClosed
                            End If
                        ElseIf Tree.Node.GetNodeTypeFromString(sqlRd.Item("Type")) = Tree.Node.Type.Container Then
                            Dim contI As New Container.Info
                            'If tNode.Parent IsNot Nothing Then
                            '    If Tree.Node.GetNodeType(tNode.Parent) = Tree.Node.Type.Container Then
                            '        contI.Parent = tNode.Parent.Tag
                            '    End If
                            'End If
                            'prevCont = contI 'NEW
                            contI.TreeNode = tNode

                            contI.Name = sqlRd.Item("Name")

                            Dim conI As Connection.Info

                            conI = GetConnectionInfoFromSQL()

                            conI.Parent = contI
                            conI.IsContainer = True
                            contI.ConnectionInfo = conI

                            If SQLUpdate = True Then
                                Dim prevCont As Container.Info = PreviousContainerList.FindByConstantID(conI.ConstantID)
                                If prevCont IsNot Nothing Then
                                    contI.IsExpanded = prevCont.IsExpanded
                                End If

                                If conI.ConstantID = PreviousSelected Then
                                    selNode = tNode
                                End If
                            Else
                                If sqlRd.Item("Expanded") = True Then
                                    contI.IsExpanded = True
                                Else
                                    contI.IsExpanded = False
                                End If
                            End If

                            ContainerList.Add(contI)
                            ConnectionList.Add(conI)

                            tNode.Tag = contI
                            tNode.ImageIndex = Images.Enums.TreeImage.Container
                            tNode.SelectedImageIndex = Images.Enums.TreeImage.Container
                        End If

                        If sqlRd.Item("ParentID") <> 0 Then
                            Dim pNode As TreeNode = Tree.Node.GetNodeFromConstantID(sqlRd.Item("ParentID"))

                            If pNode IsNot Nothing Then
                                pNode.Nodes.Add(tNode)

                                If Tree.Node.GetNodeType(tNode) = Tree.Node.Type.Connection Then
                                    TryCast(tNode.Tag, Connection.Info).Parent = pNode.Tag
                                ElseIf Tree.Node.GetNodeType(tNode) = Tree.Node.Type.Container Then
                                    TryCast(tNode.Tag, Container.Info).Parent = pNode.Tag
                                End If
                            Else
                                baseNode.Nodes.Add(tNode)
                            End If
                        Else
                            baseNode.Nodes.Add(tNode)
                        End If

                        'AddNodesFromSQL(tNode)
                    End While
                Catch ex As Exception
                    mC.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strAddNodesFromSqlFailed & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Function GetConnectionInfoFromSQL() As Connection.Info
                Try
                    Dim conI As New Connection.Info

                    With sqlRd
                        conI.PositionID = .Item("PositionID")
                        conI.ConstantID = .Item("ConstantID")
                        conI.Name = .Item("Name")
                        conI.Description = .Item("Description")
                        conI.Hostname = .Item("Hostname")
                        conI.Username = .Item("Username")
                        conI.Password = Security.Crypt.Decrypt(.Item("Password"), pW)
                        conI.Domain = .Item("DomainName")
                        conI.DisplayWallpaper = .Item("DisplayWallpaper")
                        conI.DisplayThemes = .Item("DisplayThemes")
                        conI.CacheBitmaps = .Item("CacheBitmaps")
                        conI.UseConsoleSession = .Item("ConnectToConsole")

                        conI.RedirectDiskDrives = .Item("RedirectDiskDrives")
                        conI.RedirectPrinters = .Item("RedirectPrinters")
                        conI.RedirectPorts = .Item("RedirectPorts")
                        conI.RedirectSmartCards = .Item("RedirectSmartCards")
                        conI.RedirectKeys = .Item("RedirectKeys")
                        conI.RedirectSound = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPSounds), .Item("RedirectSound"))

                        conI.Protocol = Tools.Misc.StringToEnum(GetType(Connection.Protocol.Protocols), .Item("Protocol"))
                        conI.Port = .Item("Port")
                        conI.PuttySession = .Item("PuttySession")

                        conI.Colors = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPColors), .Item("Colors"))
                        conI.Resolution = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.RDPResolutions), .Item("Resolution"))

                        conI.Inherit = New Connection.Info.Inheritance(conI)
                        conI.Inherit.CacheBitmaps = .Item("InheritCacheBitmaps")
                        conI.Inherit.Colors = .Item("InheritColors")
                        conI.Inherit.Description = .Item("InheritDescription")
                        conI.Inherit.DisplayThemes = .Item("InheritDisplayThemes")
                        conI.Inherit.DisplayWallpaper = .Item("InheritDisplayWallpaper")
                        conI.Inherit.Domain = .Item("InheritDomain")
                        conI.Inherit.Icon = .Item("InheritIcon")
                        conI.Inherit.Panel = .Item("InheritPanel")
                        conI.Inherit.Password = .Item("InheritPassword")
                        conI.Inherit.Port = .Item("InheritPort")
                        conI.Inherit.Protocol = .Item("InheritProtocol")
                        conI.Inherit.PuttySession = .Item("InheritPuttySession")
                        conI.Inherit.RedirectDiskDrives = .Item("InheritRedirectDiskDrives")
                        conI.Inherit.RedirectKeys = .Item("InheritRedirectKeys")
                        conI.Inherit.RedirectPorts = .Item("InheritRedirectPorts")
                        conI.Inherit.RedirectPrinters = .Item("InheritRedirectPrinters")
                        conI.Inherit.RedirectSmartCards = .Item("InheritRedirectSmartCards")
                        conI.Inherit.RedirectSound = .Item("InheritRedirectSound")
                        conI.Inherit.Resolution = .Item("InheritResolution")
                        conI.Inherit.UseConsoleSession = .Item("InheritUseConsoleSession")
                        conI.Inherit.Username = .Item("InheritUsername")

                        conI.Icon = .Item("Icon")
                        conI.Panel = .Item("Panel")

                        If Me.confVersion > 1.5 Then '1.6
                            conI.ICAEncryption = Tools.Misc.StringToEnum(GetType(Connection.Protocol.ICA.EncryptionStrength), .Item("ICAEncryptionStrength"))
                            conI.Inherit.ICAEncryption = .Item("InheritICAEncryptionStrength")

                            conI.PreExtApp = .Item("PreExtApp")
                            conI.PostExtApp = .Item("PostExtApp")
                            conI.Inherit.PreExtApp = .Item("InheritPreExtApp")
                            conI.Inherit.PostExtApp = .Item("InheritPostExtApp")
                        End If

                        If Me.confVersion > 1.6 Then '1.7
                            conI.VNCCompression = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.Compression), .Item("VNCCompression"))
                            conI.VNCEncoding = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.Encoding), .Item("VNCEncoding"))
                            conI.VNCAuthMode = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.AuthMode), .Item("VNCAuthMode"))
                            conI.VNCProxyType = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.ProxyType), .Item("VNCProxyType"))
                            conI.VNCProxyIP = .Item("VNCProxyIP")
                            conI.VNCProxyPort = .Item("VNCProxyPort")
                            conI.VNCProxyUsername = .Item("VNCProxyUsername")
                            conI.VNCProxyPassword = Security.Crypt.Decrypt(.Item("VNCProxyPassword"), pW)
                            conI.VNCColors = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.Colors), .Item("VNCColors"))
                            conI.VNCSmartSizeMode = Tools.Misc.StringToEnum(GetType(Connection.Protocol.VNC.SmartSizeMode), .Item("VNCSmartSizeMode"))
                            conI.VNCViewOnly = .Item("VNCViewOnly")

                            conI.Inherit.VNCCompression = .Item("InheritVNCCompression")
                            conI.Inherit.VNCEncoding = .Item("InheritVNCEncoding")
                            conI.Inherit.VNCAuthMode = .Item("InheritVNCAuthMode")
                            conI.Inherit.VNCProxyType = .Item("InheritVNCProxyType")
                            conI.Inherit.VNCProxyIP = .Item("InheritVNCProxyIP")
                            conI.Inherit.VNCProxyPort = .Item("InheritVNCProxyPort")
                            conI.Inherit.VNCProxyUsername = .Item("InheritVNCProxyUsername")
                            conI.Inherit.VNCProxyPassword = .Item("InheritVNCProxyPassword")
                            conI.Inherit.VNCColors = .Item("InheritVNCColors")
                            conI.Inherit.VNCSmartSizeMode = .Item("InheritVNCSmartSizeMode")
                            conI.Inherit.VNCViewOnly = .Item("InheritVNCViewOnly")
                        End If

                        If Me.confVersion > 1.7 Then '1.8
                            conI.RDPAuthenticationLevel = Tools.Misc.StringToEnum(GetType(Connection.Protocol.RDP.AuthenticationLevel), .Item("RDPAuthenticationLevel"))

                            conI.Inherit.RDPAuthenticationLevel = .Item("InheritRDPAuthenticationLevel")
                        End If

                        If Me.confVersion > 1.8 Then '1.9
                            conI.RenderingEngine = Tools.Misc.StringToEnum(GetType(Connection.Protocol.HTTPBase.RenderingEngine), .Item("RenderingEngine"))
                            conI.MacAddress = .Item("MacAddress")

                            conI.Inherit.RenderingEngine = .Item("InheritRenderingEngine")
                            conI.Inherit.MacAddress = .Item("InheritMacAddress")
                        End If

                        If Me.confVersion > 1.9 Then '2.0
                            conI.UserField = .Item("UserField")

                            conI.Inherit.UserField = .Item("InheritUserField")
                        End If

                        If Me.confVersion > 2.0 Then '2.1
                            conI.ExtApp = .Item("ExtApp")

                            conI.Inherit.ExtApp = .Item("InheritExtApp")
                        End If

                        If Me.confVersion >= 2.2 Then
                            conI.RDGatewayUsageMethod = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.RDP.RDGatewayUsageMethod), .Item("RDGatewayUsageMethod"))
                            conI.RDGatewayHostname = .Item("RDGatewayHostname")
                            conI.RDGatewayUseConnectionCredentials = Tools.Misc.StringToEnum(GetType(mRemoteNG.Connection.Protocol.RDP.RDGatewayUseConnectionCredentials), .Item("RDGatewayUseConnectionCredentials"))
                            conI.RDGatewayUsername = .Item("RDGatewayUsername")
                            conI.RDGatewayPassword = Security.Crypt.Decrypt(.Item("RDGatewayPassword"), pW)
                            conI.RDGatewayDomain = .Item("RDGatewayDomain")
                            conI.Inherit.RDGatewayUsageMethod = .Item("InheritRDGatewayUsageMethod")
                            conI.Inherit.RDGatewayHostname = .Item("InheritRDGatewayHostname")
                            conI.Inherit.RDGatewayUsername = .Item("InheritRDGatewayUsername")
                            conI.Inherit.RDGatewayPassword = .Item("InheritRDGatewayPassword")
                            conI.Inherit.RDGatewayDomain = .Item("InheritRDGatewayDomain")
                        End If

                        If Me.confVersion >= 2.3 Then
                            conI.EnableFontSmoothing = .Item("EnableFontSmoothing")
                            conI.EnableDesktopComposition = .Item("EnableDesktopComposition")
                            conI.Inherit.EnableFontSmoothing = .Item("InheritEnableFontSmoothing")
                            conI.Inherit.EnableDesktopComposition = .Item("InheritEnableDesktopComposition")
                        End If

                        If SQLUpdate = True Then
                            conI.PleaseConnect = .Item("Connected")
                        End If
                    End With

                    Return conI
                Catch ex As Exception
                    mC.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strGetConnectionInfoFromSqlFailed & vbNewLine & ex.Message, True)
                End Try

                Return Nothing
            End Function
#End Region

#Region "Save"
            Private Sub SaveToSQL()
                SqlConnection.Open()

                Dim tN As TreeNode
                tN = RootTreeNode.Clone

                Dim strProtected As String
                If tN.Tag IsNot Nothing Then
                    If TryCast(tN.Tag, mRemoteNG.Root.Info).Password = True Then
                        pW = TryCast(tN.Tag, mRemoteNG.Root.Info).PasswordString
                        strProtected = Security.Crypt.Encrypt("ThisIsProtected", pW)
                    Else
                        strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", pW)
                    End If
                Else
                    strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", pW)
                End If

                sqlQuery = New SqlCommand("DELETE FROM tblRoot", SqlConnection)
                Dim sqlWr As Integer = sqlQuery.ExecuteNonQuery

                Dim enCulture As CultureInfo = New CultureInfo("en-US")
                sqlQuery = New SqlCommand("INSERT INTO tblRoot (Name, Export, Protected, ConfVersion) VALUES('" & PrepareValueForDB(tN.Text) & "', 0, '" & strProtected & "'," & App.Info.Connections.ConnectionFileVersion.ToString(enCulture) & ")", SqlConnection)
                sqlWr = sqlQuery.ExecuteNonQuery

                sqlQuery = New SqlCommand("DELETE FROM tblCons", SqlConnection)
                sqlWr = sqlQuery.ExecuteNonQuery

                Dim tNC As TreeNodeCollection
                tNC = tN.Nodes

                SaveNodesSQL(tNC)

                sqlQuery = New SqlCommand("DELETE FROM tblUpdate", SqlConnection)
                sqlWr = sqlQuery.ExecuteNonQuery
                sqlQuery = New SqlCommand("INSERT INTO tblUpdate (LastUpdate) VALUES('" & Tools.Misc.DBDate(Now) & "')", SqlConnection)
                sqlWr = sqlQuery.ExecuteNonQuery

                SqlConnection.Close()
            End Sub

            Private Sub SaveNodesSQL(ByVal tnc As TreeNodeCollection)
                For Each node As TreeNode In tnc
                    gIndex += 1

                    Dim curConI As Connection.Info
                    sqlQuery = New SqlCommand("INSERT INTO tblCons (Name, Type, Expanded, Description, Icon, Panel, Username, " & _
                                               "DomainName, Password, Hostname, Protocol, PuttySession, " & _
                                               "Port, ConnectToConsole, RenderingEngine, ICAEncryptionStrength, RDPAuthenticationLevel, Colors, Resolution, DisplayWallpaper, " & _
                                               "DisplayThemes, EnableFontSmoothing, EnableDesktopComposition, CacheBitmaps, RedirectDiskDrives, RedirectPorts, " & _
                                               "RedirectPrinters, RedirectSmartCards, RedirectSound, RedirectKeys, " & _
                                               "Connected, PreExtApp, PostExtApp, MacAddress, UserField, ExtApp, VNCCompression, VNCEncoding, VNCAuthMode, " & _
                                               "VNCProxyType, VNCProxyIP, VNCProxyPort, VNCProxyUsername, VNCProxyPassword, " & _
                                               "VNCColors, VNCSmartSizeMode, VNCViewOnly, " & _
                                               "RDGatewayUsageMethod, RDGatewayHostname, RDGatewayUseConnectionCredentials, RDGatewayUsername, RDGatewayPassword, RDGatewayDomain, " & _
                                               "InheritCacheBitmaps, InheritColors, " & _
                                               "InheritDescription, InheritDisplayThemes, InheritDisplayWallpaper, InheritEnableFontSmoothing, InheritEnableDesktopComposition, InheritDomain, " & _
                                               "InheritIcon, InheritPanel, InheritPassword, InheritPort, " & _
                                               "InheritProtocol, InheritPuttySession, InheritRedirectDiskDrives, " & _
                                               "InheritRedirectKeys, InheritRedirectPorts, InheritRedirectPrinters, " & _
                                               "InheritRedirectSmartCards, InheritRedirectSound, InheritResolution, " & _
                                               "InheritUseConsoleSession, InheritRenderingEngine, InheritUsername, InheritICAEncryptionStrength, InheritRDPAuthenticationLevel, " & _
                                               "InheritPreExtApp, InheritPostExtApp, InheritMacAddress, InheritUserField, InheritExtApp, InheritVNCCompression, InheritVNCEncoding, " & _
                                               "InheritVNCAuthMode, InheritVNCProxyType, InheritVNCProxyIP, InheritVNCProxyPort, " & _
                                               "InheritVNCProxyUsername, InheritVNCProxyPassword, InheritVNCColors, " & _
                                               "InheritVNCSmartSizeMode, InheritVNCViewOnly, " & _
                                               "InheritRDGatewayUsageMethod, InheritRDGatewayHostname, InheritRDGatewayUseConnectionCredentials, InheritRDGatewayUsername, InheritRDGatewayPassword, InheritRDGatewayDomain, " & _
                                               "PositionID, ParentID, ConstantID, LastChange)" & _
                                               "VALUES (", SqlConnection)

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Or Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then
                        'xW.WriteStartElement("Node")
                        sqlQuery.CommandText &= "'" & PrepareValueForDB(node.Text) & "'," 'Name
                        sqlQuery.CommandText &= "'" & Tree.Node.GetNodeType(node).ToString & "'," 'Type
                    End If

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then 'container
                        sqlQuery.CommandText &= "'" & ContainerList(node.Tag).IsExpanded & "'," 'Expanded
                        curConI = ContainerList(node.Tag).ConnectionInfo
                        SaveConnectionFieldsSQL(curConI)

                        sqlQuery.CommandText = PrepareForDB(sqlQuery.CommandText)
                        sqlQuery.ExecuteNonQuery()
                        'parentID = gIndex
                        SaveNodesSQL(node.Nodes)
                        'xW.WriteEndElement()
                    End If

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Then
                        sqlQuery.CommandText &= "'" & False & "',"
                        curConI = ConnectionList(node.Tag)
                        SaveConnectionFieldsSQL(curConI)
                        'xW.WriteEndElement()
                        sqlQuery.CommandText = PrepareForDB(sqlQuery.CommandText)
                        sqlQuery.ExecuteNonQuery()
                    End If

                    'parentID = 0
                Next
            End Sub

            Private Sub SaveConnectionFieldsSQL(ByVal curConI As Connection.Info)
                With curConI
                    sqlQuery.CommandText &= "'" & PrepareValueForDB(.Description) & "',"
                    sqlQuery.CommandText &= "'" & PrepareValueForDB(.Icon) & "',"
                    sqlQuery.CommandText &= "'" & PrepareValueForDB(.Panel) & "',"

                    If SaveSecurity.Username = True Then
                        sqlQuery.CommandText &= "'" & PrepareValueForDB(.Username) & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    If SaveSecurity.Domain = True Then
                        sqlQuery.CommandText &= "'" & PrepareValueForDB(.Domain) & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    If SaveSecurity.Password = True Then
                        sqlQuery.CommandText &= "'" & PrepareValueForDB(Security.Crypt.Encrypt(.Password, pW)) & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    sqlQuery.CommandText &= "'" & PrepareValueForDB(.Hostname) & "',"
                    sqlQuery.CommandText &= "'" & .Protocol.ToString & "',"
                    sqlQuery.CommandText &= "'" & PrepareValueForDB(.PuttySession) & "',"
                    sqlQuery.CommandText &= "'" & .Port & "',"
                    sqlQuery.CommandText &= "'" & .UseConsoleSession & "',"
                    sqlQuery.CommandText &= "'" & .RenderingEngine.ToString & "',"
                    sqlQuery.CommandText &= "'" & .ICAEncryption.ToString & "',"
                    sqlQuery.CommandText &= "'" & .RDPAuthenticationLevel.ToString & "',"
                    sqlQuery.CommandText &= "'" & .Colors.ToString & "',"
                    sqlQuery.CommandText &= "'" & .Resolution.ToString & "',"
                    sqlQuery.CommandText &= "'" & .DisplayWallpaper & "',"
                    sqlQuery.CommandText &= "'" & .DisplayThemes & "',"
                    sqlQuery.CommandText &= "'" & .EnableFontSmoothing & "',"
                    sqlQuery.CommandText &= "'" & .EnableDesktopComposition & "',"
                    sqlQuery.CommandText &= "'" & .CacheBitmaps & "',"
                    sqlQuery.CommandText &= "'" & .RedirectDiskDrives & "',"
                    sqlQuery.CommandText &= "'" & .RedirectPorts & "',"
                    sqlQuery.CommandText &= "'" & .RedirectPrinters & "',"
                    sqlQuery.CommandText &= "'" & .RedirectSmartCards & "',"
                    sqlQuery.CommandText &= "'" & .RedirectSound.ToString & "',"
                    sqlQuery.CommandText &= "'" & .RedirectKeys & "',"

                    If curConI.OpenConnections.Count > 0 Then
                        sqlQuery.CommandText &= 1 & ","
                    Else
                        sqlQuery.CommandText &= 0 & ","
                    End If

                    sqlQuery.CommandText &= "'" & .PreExtApp & "',"
                    sqlQuery.CommandText &= "'" & .PostExtApp & "',"
                    sqlQuery.CommandText &= "'" & .MacAddress & "',"
                    sqlQuery.CommandText &= "'" & .UserField & "',"
                    sqlQuery.CommandText &= "'" & .ExtApp & "',"

                    sqlQuery.CommandText &= "'" & .VNCCompression.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCEncoding.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCAuthMode.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCProxyType.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCProxyIP & "',"
                    sqlQuery.CommandText &= "'" & .VNCProxyPort & "',"
                    sqlQuery.CommandText &= "'" & .VNCProxyUsername & "',"
                    sqlQuery.CommandText &= "'" & Security.Crypt.Encrypt(.VNCProxyPassword, pW) & "',"
                    sqlQuery.CommandText &= "'" & .VNCColors.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCSmartSizeMode.ToString & "',"
                    sqlQuery.CommandText &= "'" & .VNCViewOnly & "',"

                    sqlQuery.CommandText &= "'" & .RDGatewayUsageMethod.ToString & "',"
                    sqlQuery.CommandText &= "'" & .RDGatewayHostname & "',"
                    sqlQuery.CommandText &= "'" & .RDGatewayUseConnectionCredentials.ToString & "',"

                    If SaveSecurity.Username = True Then
                        sqlQuery.CommandText &= "'" & .RDGatewayUsername & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    If SaveSecurity.Password = True Then
                        sqlQuery.CommandText &= "'" & .RDGatewayPassword & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    If SaveSecurity.Domain = True Then
                        sqlQuery.CommandText &= "'" & .RDGatewayDomain & "',"
                    Else
                        sqlQuery.CommandText &= "'" & "" & "',"
                    End If

                    With .Inherit
                        If SaveSecurity.Inheritance = True Then
                            sqlQuery.CommandText &= "'" & .CacheBitmaps & "',"
                            sqlQuery.CommandText &= "'" & .Colors & "',"
                            sqlQuery.CommandText &= "'" & .Description & "',"
                            sqlQuery.CommandText &= "'" & .DisplayThemes & "',"
                            sqlQuery.CommandText &= "'" & .DisplayWallpaper & "',"
                            sqlQuery.CommandText &= "'" & .EnableFontSmoothing & "',"
                            sqlQuery.CommandText &= "'" & .EnableDesktopComposition & "',"
                            sqlQuery.CommandText &= "'" & .Domain & "',"
                            sqlQuery.CommandText &= "'" & .Icon & "',"
                            sqlQuery.CommandText &= "'" & .Panel & "',"
                            sqlQuery.CommandText &= "'" & .Password & "',"
                            sqlQuery.CommandText &= "'" & .Port & "',"
                            sqlQuery.CommandText &= "'" & .Protocol & "',"
                            sqlQuery.CommandText &= "'" & .PuttySession & "',"
                            sqlQuery.CommandText &= "'" & .RedirectDiskDrives & "',"
                            sqlQuery.CommandText &= "'" & .RedirectKeys & "',"
                            sqlQuery.CommandText &= "'" & .RedirectPorts & "',"
                            sqlQuery.CommandText &= "'" & .RedirectPrinters & "',"
                            sqlQuery.CommandText &= "'" & .RedirectSmartCards & "',"
                            sqlQuery.CommandText &= "'" & .RedirectSound & "',"
                            sqlQuery.CommandText &= "'" & .Resolution & "',"
                            sqlQuery.CommandText &= "'" & .UseConsoleSession & "',"
                            sqlQuery.CommandText &= "'" & .RenderingEngine & "',"
                            sqlQuery.CommandText &= "'" & .Username & "',"
                            sqlQuery.CommandText &= "'" & .ICAEncryption & "',"
                            sqlQuery.CommandText &= "'" & .RDPAuthenticationLevel & "',"
                            sqlQuery.CommandText &= "'" & .PreExtApp & "',"
                            sqlQuery.CommandText &= "'" & .PostExtApp & "',"
                            sqlQuery.CommandText &= "'" & .MacAddress & "',"
                            sqlQuery.CommandText &= "'" & .UserField & "',"
                            sqlQuery.CommandText &= "'" & .ExtApp & "',"

                            sqlQuery.CommandText &= "'" & .VNCCompression & "',"
                            sqlQuery.CommandText &= "'" & .VNCEncoding & "',"
                            sqlQuery.CommandText &= "'" & .VNCAuthMode & "',"
                            sqlQuery.CommandText &= "'" & .VNCProxyType & "',"
                            sqlQuery.CommandText &= "'" & .VNCProxyIP & "',"
                            sqlQuery.CommandText &= "'" & .VNCProxyPort & "',"
                            sqlQuery.CommandText &= "'" & .VNCProxyUsername & "',"
                            sqlQuery.CommandText &= "'" & .VNCProxyPassword & "',"
                            sqlQuery.CommandText &= "'" & .VNCColors & "',"
                            sqlQuery.CommandText &= "'" & .VNCSmartSizeMode & "',"
                            sqlQuery.CommandText &= "'" & .VNCViewOnly & "',"

                            sqlQuery.CommandText &= "'" & .RDGatewayUsageMethod & "',"
                            sqlQuery.CommandText &= "'" & .RDGatewayHostname & "',"
                            sqlQuery.CommandText &= "'" & .RDGatewayUseConnectionCredentials & "',"
                            sqlQuery.CommandText &= "'" & .RDGatewayUsername & "',"
                            sqlQuery.CommandText &= "'" & .RDGatewayPassword & "',"
                            sqlQuery.CommandText &= "'" & .RDGatewayDomain & "',"
                        Else
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"

                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"
                            sqlQuery.CommandText &= "'" & False & "',"

                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayUsageMethod
                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayHostname
                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayUseConnectionCredentials
                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayUsername
                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayPassword
                            sqlQuery.CommandText &= "'" & False & "'," ' .RDGatewayDomain
                        End If
                    End With

                    .PositionID = gIndex

                    If .IsContainer = False Then
                        If .Parent IsNot Nothing Then
                            parentID = TryCast(.Parent, Container.Info).ConnectionInfo.ConstantID
                        Else
                            parentID = 0
                        End If
                    Else
                        If TryCast(.Parent, Container.Info).Parent IsNot Nothing Then
                            parentID = TryCast(TryCast(.Parent, Container.Info).Parent, Container.Info).ConnectionInfo.ConstantID
                        Else
                            parentID = 0
                        End If
                    End If

                    sqlQuery.CommandText &= gIndex & "," & parentID & "," & .ConstantID & ",'" & Tools.Misc.DBDate(Now) & "')"
                End With
            End Sub
#End Region

#Region "Utility"
            Private Shared Function PrepareForDB(ByVal Text As String) As String
                Text = Replace(Text, "'True'", "1", , , CompareMethod.Text)
                Text = Replace(Text, "'False'", "0", , , CompareMethod.Text)

                Return Text
            End Function

            Private Shared Function PrepareValueForDB(ByVal Text As String) As String
                Text = Replace(Text, "'", "''", , , CompareMethod.Text)

                Return Text
            End Function
#End Region

#End Region
        End Class
    End Namespace
End Namespace
