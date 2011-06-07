Imports System.Globalization
Imports System.Data.SqlClient
Imports mRemoteNG.App.Runtime
Imports mRemoteNG.My.Resources

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

#Region "Public Methods"
            Public Sub TestConnection()
                Dim sqlDataReader As SqlDataReader = Nothing
                Try
                    OpenConnection()

                    Dim sqlQuery As New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                    sqlDataReader = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)
                    sqlDataReader.Read()

                    MessageBox.Show(frmMain, "The database connection was successful.", "SQL Server Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(frmMain, ex.Message, "SQL Server Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    If sqlDataReader IsNot Nothing Then
                        If Not sqlDataReader.IsClosed Then sqlDataReader.Close()
                    End If
                    CloseConnection()
                End Try
            End Sub

            Public Sub CreateTables()
                Try
                    OpenConnection()

                    Dim sqlQuery As New SqlCommand(My.Resources.CreateTables, SqlConnection)
                    sqlQuery.ExecuteNonQuery()

                    If RootTreeNode Is Nothing Then
                        Tree.Node.ResetTree()
                        RootTreeNode = Windows.treeForm.tvConnections.Nodes(0)
                    End If
                    Dim rootNode As TreeNode
                    rootNode = RootTreeNode.Clone

                    Dim strProtected As String
                    If rootNode.Tag IsNot Nothing Then
                        If TryCast(rootNode.Tag, mRemoteNG.Root.Info).Password = True Then
                            pW = TryCast(rootNode.Tag, mRemoteNG.Root.Info).PasswordString
                            strProtected = Security.Crypt.Encrypt("ThisIsProtected", pW)
                        Else
                            strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", pW)
                        End If
                    Else
                        strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", pW)
                    End If

                    Dim data As New ColumnValueCollection

                    data.Add("Name", rootNode.Text)
                    data.Add("Export", False, True)
                    data.Add("Protected", strProtected)
                    data.Add("ConfVersion", App.Info.Connections.ConnectionFileVersion.ToString(CultureInfo.InvariantCulture))

                    sqlQuery = New SqlCommand(String.Format(CultureInfo.InvariantCulture, "INSERT INTO tblRoot ({0}) VALUES({1})", data.ColumnsString, data.ValuesString), SqlConnection)
                    sqlQuery.ExecuteNonQuery()
                Catch ex As Exception
                    Throw
                Finally
                    CloseConnection()
                End Try
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
                    Dim sqlDataReader As SqlDataReader = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)
                    sqlDataReader.Read()

                    Dim version As System.Version = sqlDataReader.Item("confVersion")

                    sqlDataReader.Close()
                    CloseConnection()

                    Return version
                End Get
                Set(ByVal value As System.Version)

                End Set
            End Property
#End Region

#Region "Private Variables"
            Private confVersion As Double

            Private sqlQuery As SqlCommand
            'Private sqlDataReader As SqlDataReader

            Private gIndex As Integer = 0
            Private parentID As String = 0
#End Region

#Region "Private Methods"

#Region "SQL Server Connection"
            Private Function OpenConnection() As Boolean
                If SqlConnection Is Nothing Then Return False

                ' Already open?
                If Not (SqlConnection.State = ConnectionState.Closed Or SqlConnection.State = ConnectionState.Broken) Then Return True

                While True
                    Try
                        SqlConnection.Open()
                        Return True
                    Catch ex As System.Exception
                        If MessageBox.Show(String.Format(strErrorDatabaseOpenConnectionFailed, ex.Message, vbNewLine), My.Application.Info.ProductName, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) = DialogResult.Cancel Then
                            Return False
                        End If
                    End Try
                End While
            End Function

            Private Sub CloseConnection()
                If SqlConnection Is Nothing Then Return
                If SqlConnection.State = ConnectionState.Closed Then Return
                SqlConnection.Close()
            End Sub
#End Region

#Region "Load"
            Public Overrides Function Load() As Boolean
                Try
                    App.Runtime.IsConnectionsFileLoaded = False

                    If Not OpenConnection() Then Return False

                    sqlQuery = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                    Dim sqlDataReader As SqlDataReader = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                    sqlDataReader.Read()

                    If sqlDataReader.HasRows = False Then
                        App.Runtime.SaveConnections()

                        sqlQuery = New SqlCommand("SELECT * FROM tblRoot", SqlConnection)
                        sqlDataReader = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)
                        sqlDataReader.Read()
                    End If

                    Me.confVersion = Convert.ToDouble(sqlDataReader.Item("confVersion"), CultureInfo.InvariantCulture)

                    Dim rootNode As TreeNode
                    rootNode = New TreeNode(sqlDataReader.Item("Name"))

                    Dim rInfo As New Root.Info(Root.Info.RootType.Connection)
                    rInfo.Name = rootNode.Text
                    rInfo.TreeNode = rootNode

                    rootNode.Tag = rInfo
                    rootNode.ImageIndex = Images.Enums.TreeImage.Root
                    rootNode.SelectedImageIndex = Images.Enums.TreeImage.Root

                    If Security.Crypt.Decrypt(sqlDataReader.Item("Protected"), pW) <> "ThisIsNotProtected" Then
                        If Authenticate(sqlDataReader.Item("Protected"), False, rInfo) = False Then
                            My.Settings.LoadConsFromCustomLocation = False
                            My.Settings.CustomConsPath = ""
                            rootNode.Remove()
                            Return False
                        End If
                    End If

                    'Me._RootTreeNode.Text = rootNode.Text
                    'Me._RootTreeNode.Tag = rootNode.Tag
                    'Me._RootTreeNode.ImageIndex = Images.Enums.TreeImage.Root
                    'Me._RootTreeNode.SelectedImageIndex = Images.Enums.TreeImage.Root

                    sqlDataReader.Close()

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

                    CloseConnection()

                    SetMainFormText(My.Resources.strSQLServer)
                    App.Runtime.IsConnectionsFileLoaded = True
                    'App.Runtime.Windows.treeForm.InitialRefresh()

                    Return True
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strLoadFromSqlFailed & vbNewLine & ex.Message, True)
                    Return False
                End Try
            End Function

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
                    Dim sqlDataReader As SqlDataReader = sqlQuery.ExecuteReader(CommandBehavior.CloseConnection)

                    If sqlDataReader.HasRows = False Then
                        Exit Sub
                    End If

                    Dim tNode As TreeNode

                    While sqlDataReader.Read
                        tNode = New TreeNode(sqlDataReader.Item("Name"))
                        'baseNode.Nodes.Add(tNode)

                        If Tree.Node.GetNodeTypeFromString(sqlDataReader.Item("Type")) = Tree.Node.Type.Connection Then
                            Dim conI As Connection.Info = GetConnectionInfoFromSQL(sqlDataReader)
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
                        ElseIf Tree.Node.GetNodeTypeFromString(sqlDataReader.Item("Type")) = Tree.Node.Type.Container Then
                            Dim contI As New Container.Info
                            'If tNode.Parent IsNot Nothing Then
                            '    If Tree.Node.GetNodeType(tNode.Parent) = Tree.Node.Type.Container Then
                            '        contI.Parent = tNode.Parent.Tag
                            '    End If
                            'End If
                            'prevCont = contI 'NEW
                            contI.TreeNode = tNode

                            contI.Name = sqlDataReader.Item("Name")

                            Dim conI As Connection.Info

                            conI = GetConnectionInfoFromSQL(sqlDataReader)

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
                                If sqlDataReader.Item("Expanded") = True Then
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

                        If sqlDataReader.Item("ParentID") <> 0 Then
                            Dim pNode As TreeNode = Tree.Node.GetNodeFromConstantID(sqlDataReader.Item("ParentID"))

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
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strAddNodesFromSqlFailed & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Function GetConnectionInfoFromSQL(ByRef sqlDataReader As SqlDataReader) As Connection.Info
                Try
                    Dim conI As New Connection.Info

                    With SqlDataReader
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
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, My.Resources.strGetConnectionInfoFromSqlFailed & vbNewLine & ex.Message, True)
                End Try

                Return Nothing
            End Function
#End Region

#Region "Save"
            Public Overrides Function Save() As Boolean
                If Not OpenConnection() Then Return False

                If Not VerifyDatabaseVersion(SqlConnection) Then
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, strErrorConnectionListSaveFailed)
                    Return False
                End If

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

                sqlQuery = New SqlCommand("INSERT INTO tblRoot (Name, Export, Protected, ConfVersion) VALUES('" & EscapeSingleQuotes(tN.Text) & "', 0, '" & strProtected & "'," & App.Info.Connections.ConnectionFileVersion.ToString(CultureInfo.InvariantCulture) & ")", SqlConnection)
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

                CloseConnection()

                SetMainFormText(My.Resources.strSQLServer)

                Return True
            End Function

            Private Sub SaveNodesSQL(ByVal treeNodes As TreeNodeCollection)
                Dim data As New ColumnValueCollection

                For Each node As TreeNode In treeNodes
                    gIndex += 1

                    Dim currentConnectionInfo As Connection.Info

                    data.Clear()

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Or Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then
                        data.Add("Name", node.Text)
                        data.Add("Type", Tree.Node.GetNodeType(node).ToString)
                    End If

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then 'container
                        data.Add("Expanded", ContainerList(node.Tag).IsExpanded)

                        currentConnectionInfo = ContainerList(node.Tag).ConnectionInfo
                        SaveConnectionFieldsSQL(currentConnectionInfo, data)

                        sqlQuery = New SqlCommand(String.Format("INSERT INTO tblCons ({0}) VALUES ({1})", data.ColumnsString, data.ValuesString), SqlConnection)
                        sqlQuery.ExecuteNonQuery()

                        SaveNodesSQL(node.Nodes)
                    End If

                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Then
                        data.Add("Expanded", False)

                        currentConnectionInfo = ConnectionList(node.Tag)
                        SaveConnectionFieldsSQL(currentConnectionInfo, data)

                        sqlQuery = New SqlCommand(String.Format("INSERT INTO tblCons ({0}) VALUES ({1})", data.ColumnsString, data.ValuesString), SqlConnection)
                        sqlQuery.ExecuteNonQuery()
                    End If
                Next
            End Sub

            Private Sub SaveConnectionFieldsSQL(ByRef currentConnectionInfo As Connection.Info, ByRef data As ColumnValueCollection)
                With currentConnectionInfo
                    data.Add("Description", .Description)
                    data.Add("Icon", .Icon)
                    data.Add("Panel", .Panel)

                    If SaveSecurity.Username Then
                        data.Add("Username", .Username)
                    Else
                        data.Add("Username")
                    End If

                    If SaveSecurity.Domain Then
                        data.Add("DomainName", .Domain)
                    Else
                        data.Add("DomainName")
                    End If

                    If SaveSecurity.Password Then
                        data.Add("Password", Security.Crypt.Encrypt(.Password, pW))
                    Else
                        data.Add("Password")
                    End If

                    data.Add("Hostname", .Hostname)
                    data.Add("Protocol", .Protocol.ToString)
                    data.Add("PuttySession", .PuttySession)
                    data.Add("Port", .Port)
                    data.Add("ConnectToConsole", .UseConsoleSession, True)
                    data.Add("RenderingEngine", .RenderingEngine)
                    data.Add("ICAEncryptionStrength", .ICAEncryption)
                    data.Add("RDPAuthenticationLevel", .RDPAuthenticationLevel)
                    data.Add("Colors", .Colors)
                    data.Add("Resolution", .Resolution)
                    data.Add("DisplayWallpaper", .DisplayWallpaper, True)
                    data.Add("DisplayThemes", .DisplayThemes, True)
                    data.Add("EnableFontSmoothing", .EnableFontSmoothing, True)
                    data.Add("EnableDesktopComposition", .EnableDesktopComposition, True)
                    data.Add("CacheBitmaps", .CacheBitmaps, True)
                    data.Add("RedirectDiskDrives", .RedirectDiskDrives, True)
                    data.Add("RedirectPorts", .RedirectPorts, True)
                    data.Add("RedirectPrinters", .RedirectPrinters, True)
                    data.Add("RedirectSmartCards", .RedirectSmartCards, True)
                    data.Add("RedirectSound", .RedirectSound, True)
                    data.Add("RedirectKeys", .RedirectKeys, True)

                    If currentConnectionInfo.OpenConnections.Count > 0 Then
                        data.Add("Connected", True, True)
                    Else
                        data.Add("Connected", False, True)
                    End If

                    data.Add("PreExtApp", .PreExtApp)
                    data.Add("PostExtApp", .PostExtApp)
                    data.Add("MacAddress", .MacAddress)
                    data.Add("UserField", .UserField)
                    data.Add("ExtApp", .ExtApp)

                    data.Add("VNCCompression", .VNCCompression)
                    data.Add("VNCEncoding", .VNCEncoding)
                    data.Add("VNCAuthMode", .VNCAuthMode)
                    data.Add("VNCProxyType", .VNCProxyType)
                    data.Add("VNCProxyIP", .VNCProxyIP)
                    data.Add("VNCProxyPort", .VNCProxyPort)
                    data.Add("VNCProxyUsername", .VNCProxyUsername)
                    data.Add("VNCProxyPassword", Security.Crypt.Encrypt(.VNCProxyPassword, pW))
                    data.Add("VNCColors", .VNCColors)
                    data.Add("VNCSmartSizeMode", .VNCSmartSizeMode)
                    data.Add("VNCViewOnly", .VNCViewOnly, True)

                    data.Add("RDGatewayUsageMethod", .RDGatewayUsageMethod)
                    data.Add("RDGatewayHostname", .RDGatewayHostname)
                    data.Add("RDGatewayUseConnectionCredentials", .RDGatewayUseConnectionCredentials)

                    If SaveSecurity.Username Then
                        data.Add("RDGatewayUsername", .RDGatewayUsername)
                    Else
                        data.Add("RDGatewayUsername")
                    End If

                    If SaveSecurity.Password = True Then
                        data.Add("RDGatewayPassword", .RDGatewayPassword)
                    Else
                        data.Add("RDGatewayPassword")
                    End If

                    If SaveSecurity.Domain = True Then
                        data.Add("RDGatewayDomain", .RDGatewayDomain)
                    Else
                        data.Add("RDGatewayDomain")
                    End If

                    With .Inherit
                        If SaveSecurity.Inheritance = True Then
                            data.Add("InheritCacheBitmaps", .CacheBitmaps, True)
                            data.Add("InheritColors", .Colors, True)
                            data.Add("InheritDescription", .Description, True)
                            data.Add("InheritDisplayThemes", .DisplayThemes, True)
                            data.Add("InheritDisplayWallpaper", .DisplayWallpaper, True)
                            data.Add("InheritEnableFontSmoothing", .EnableFontSmoothing, True)
                            data.Add("InheritEnableDesktopComposition", .EnableDesktopComposition, True)
                            data.Add("InheritDomain", .Domain, True)
                            data.Add("InheritIcon", .Icon, True)
                            data.Add("InheritPanel", .Panel, True)
                            data.Add("InheritPassword", .Password, True)
                            data.Add("InheritPort", .Port, True)
                            data.Add("InheritProtocol", .Protocol, True)
                            data.Add("InheritPuttySession", .PuttySession, True)
                            data.Add("InheritRedirectDiskDrives", .RedirectDiskDrives, True)
                            data.Add("InheritRedirectKeys", .RedirectKeys, True)
                            data.Add("InheritRedirectPorts", .RedirectPorts, True)
                            data.Add("InheritRedirectPrinters", .RedirectPrinters, True)
                            data.Add("InheritRedirectSmartCards", .RedirectSmartCards, True)
                            data.Add("InheritRedirectSound", .RedirectSound, True)
                            data.Add("InheritResolution", .Resolution, True)
                            data.Add("InheritUseConsoleSession", .UseConsoleSession, True)
                            data.Add("InheritRenderingEngine", .RenderingEngine, True)
                            data.Add("InheritUsername", .Username, True)
                            data.Add("InheritICAEncryptionStrength", .ICAEncryption, True)
                            data.Add("InheritRDPAuthenticationLevel", .RDPAuthenticationLevel, True)
                            data.Add("InheritPreExtApp", .PreExtApp, True)
                            data.Add("InheritPostExtApp", .PostExtApp, True)
                            data.Add("InheritMacAddress", .MacAddress, True)
                            data.Add("InheritUserField", .UserField, True)
                            data.Add("InheritExtApp", .ExtApp, True)

                            data.Add("InheritVNCCompression", .VNCCompression, True)
                            data.Add("InheritVNCEncoding", .VNCEncoding, True)
                            data.Add("InheritVNCAuthMode", .VNCAuthMode, True)
                            data.Add("InheritVNCProxyType", .VNCProxyType, True)
                            data.Add("InheritVNCProxyIP", .VNCProxyIP, True)
                            data.Add("InheritVNCProxyPort", .VNCProxyPort, True)
                            data.Add("InheritVNCProxyUsername", .VNCProxyUsername, True)
                            data.Add("InheritVNCProxyPassword", .VNCProxyPassword, True)
                            data.Add("InheritVNCColors", .VNCColors, True)
                            data.Add("InheritVNCSmartSizeMode", .VNCSmartSizeMode, True)
                            data.Add("InheritVNCViewOnly", .VNCViewOnly, True)

                            data.Add("InheritRDGatewayUsageMethod", .RDGatewayUsageMethod, True)
                            data.Add("InheritRDGatewayHostname", .RDGatewayHostname, True)
                            data.Add("InheritRDGatewayUseConnectionCredentials", .RDGatewayUseConnectionCredentials, True)
                            data.Add("InheritRDGatewayUsername", .RDGatewayUsername, True)
                            data.Add("InheritRDGatewayPassword", .RDGatewayPassword, True)
                            data.Add("InheritRDGatewayDomain", .RDGatewayDomain, True)
                        Else
                            data.Add("InheritCacheBitmaps", False, True)
                            data.Add("InheritColors", False, True)
                            data.Add("InheritDescription", False, True)
                            data.Add("InheritDisplayThemes", False, True)
                            data.Add("InheritDisplayWallpaper", False, True)
                            data.Add("InheritEnableFontSmoothing", False, True)
                            data.Add("InheritEnableDesktopComposition", False, True)
                            data.Add("InheritDomain", False, True)
                            data.Add("InheritIcon", False, True)
                            data.Add("InheritPanel", False, True)
                            data.Add("InheritPassword", False, True)
                            data.Add("InheritPort", False, True)
                            data.Add("InheritProtocol", False, True)
                            data.Add("InheritPuttySession", False, True)
                            data.Add("InheritRedirectDiskDrives", False, True)
                            data.Add("InheritRedirectKeys", False, True)
                            data.Add("InheritRedirectPorts", False, True)
                            data.Add("InheritRedirectPrinters", False, True)
                            data.Add("InheritRedirectSmartCards", False, True)
                            data.Add("InheritRedirectSound", False, True)
                            data.Add("InheritResolution", False, True)
                            data.Add("InheritUseConsoleSession", False, True)
                            data.Add("InheritRenderingEngine", False, True)
                            data.Add("InheritUsername", False, True)
                            data.Add("InheritICAEncryption", False, True)
                            data.Add("InheritRDPAuthenticationLevel", False, True)
                            data.Add("InheritPreExtApp", False, True)
                            data.Add("InheritPostExtApp", False, True)
                            data.Add("InheritMacAddress", False, True)
                            data.Add("InheritUserField", False, True)
                            data.Add("InheritExtApp", False, True)

                            data.Add("InheritVNCCompression", False, True)
                            data.Add("InheritVNCEncoding", False, True)
                            data.Add("InheritVNCAuthMode", False, True)
                            data.Add("InheritVNCCompression", False, True)
                            data.Add("InheritVNCEncoding", False, True)
                            data.Add("InheritVNCAuthMode", False, True)
                            data.Add("InheritVNCProxyType", False, True)
                            data.Add("InheritVNCProxyIP", False, True)
                            data.Add("InheritVNCProxyPort", False, True)
                            data.Add("InheritVNCProxyUsername", False, True)
                            data.Add("InheritVNCProxyPassword", False, True)
                            data.Add("InheritVNCColors", False, True)
                            data.Add("InheritVNCSmartSizeMode", False, True)
                            data.Add("InheritVNCViewOnly", False, True)

                            data.Add("InheritRDGatewayUsageMethod", False, True)
                            data.Add("InheritRDGatewayHostname", False, True)
                            data.Add("InheritRDGatewayUseConnectionCredentials", False, True)
                            data.Add("InheritRDGatewayUsername", False, True)
                            data.Add("InheritRDGatewayPassword", False, True)
                            data.Add("InheritRDGatewayDomain", False, True)
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

                    data.Add("PositionID", gIndex)
                    data.Add("ParentID", parentID)
                    data.Add("ConstantID", .ConstantID)
                    data.Add("LastChange", Tools.Misc.DBDate(Now))
                End With
            End Sub
#End Region

#Region "Utility"
            Private Function VerifyDatabaseVersion(ByVal sqlConnection As SqlConnection) As Boolean
                Dim isVerified As Boolean = False
                Dim sqlDataReader As SqlDataReader = Nothing
                Dim databaseVersion As System.Version = Nothing
                Try
                    Dim sqlCommand As New SqlCommand("SELECT * FROM tblRoot", sqlConnection)
                    sqlDataReader = sqlCommand.ExecuteReader()
                    sqlDataReader.Read()

                    databaseVersion = New System.Version(Convert.ToDouble(sqlDataReader.Item("confVersion"), CultureInfo.InvariantCulture))

                    sqlDataReader.Close()

                    If databaseVersion.CompareTo(New System.Version(2, 2)) = 0 Then ' 2.2
                        MessageCollector.AddMessage(Messages.MessageClass.InformationMsg, String.Format("Upgrading database from version {0} to version {1}.", databaseVersion.ToString, "2.3"))
                        sqlCommand = New SqlCommand("ALTER TABLE tblCons ADD EnableFontSmoothing bit NOT NULL DEFAULT 0, EnableDesktopComposition bit NOT NULL DEFAULT 0, InheritEnableFontSmoothing bit NOT NULL DEFAULT 0, InheritEnableDesktopComposition bit NOT NULL DEFAULT 0;", sqlConnection)
                        sqlCommand.ExecuteNonQuery()
                        databaseVersion = New System.Version(2, 3)
                    End If

                    If databaseVersion.CompareTo(New System.Version(2, 3)) = 0 Then ' 2.3
                        isVerified = True
                    End If

                    If isVerified = False Then
                        MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, String.Format(strErrorBadDatabaseVersion, databaseVersion.ToString, My.Application.Info.ProductName))
                    End If
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, String.Format(strErrorVerifyDatabaseVersionFailed, ex.Message))
                Finally
                    If sqlDataReader IsNot Nothing Then
                        If Not sqlDataReader.IsClosed Then sqlDataReader.Close()
                    End If
                End Try
                Return isVerified
            End Function

            Private Shared Function EscapeSingleQuotes(ByVal text As String) As String
                Return Replace(text, "'", "''", , , CompareMethod.Text)
            End Function

            Private Shared Function ConvertBoolean(ByVal value As Boolean) As String
                If value Then Return "1" Else Return "0"
            End Function

            Private Class ColumnValueCollection
#Region "Public Methods"
                Public Sub Clear()
                    _columns.Clear()
                    _values.Clear()
                End Sub

                Public Sub Add(ByVal column As String, Optional ByVal value As String = Nothing, Optional ByVal isBoolean As Boolean = False)
                    _columns.Add(column)
                    If String.IsNullOrEmpty(value) Then
                        _values.Add("''")
                    Else
                        If isBoolean Then
                            _values.Add(EscapeSingleQuotes(ConvertBoolean(value)))
                        Else
                            _values.Add("'" & EscapeSingleQuotes(value) & "'")
                        End If
                    End If
                End Sub

                Public ReadOnly Property ColumnsString() As String
                    Get
                        Return String.Join(",", _columns.ToArray)
                    End Get
                End Property

                Public ReadOnly Property ValuesString() As String
                    Get
                        Return String.Join(",", _values.ToArray)
                    End Get
                End Property
#End Region

#Region "Private Variables"
                Private _columns As New List(Of String)
                Private _values As New List(Of String)
#End Region
            End Class
#End Region
#End Region
        End Class
    End Namespace
End Namespace
