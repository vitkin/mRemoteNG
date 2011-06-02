Namespace Config
    Namespace Connections
        Public MustInherit Class Base
#Region "Public Properties"
            Private _ConnectionFileName As String
            Public Property ConnectionFileName() As String
                Get
                    Return Me._ConnectionFileName
                End Get
                Set(ByVal value As String)
                    Me._ConnectionFileName = value
                End Set
            End Property

            Private _PreviousSelected As String
            Public Property PreviousSelected() As String
                Get
                    Return _PreviousSelected
                End Get
                Set(ByVal value As String)
                    _PreviousSelected = value
                End Set
            End Property

            Private _RootTreeNode As TreeNode
            Public Property RootTreeNode() As TreeNode
                Get
                    Return Me._RootTreeNode
                End Get
                Set(ByVal value As TreeNode)
                    Me._RootTreeNode = value
                End Set
            End Property

            Private _Import As Boolean
            Public Property Import() As Boolean
                Get
                    Return Me._Import
                End Get
                Set(ByVal value As Boolean)
                    Me._Import = value
                End Set
            End Property

            Private _ConnectionList As Connection.List
            Public Property ConnectionList() As Connection.List
                Get
                    Return Me._ConnectionList
                End Get
                Set(ByVal value As Connection.List)
                    Me._ConnectionList = value
                End Set
            End Property

            Private _ContainerList As Container.List
            Public Property ContainerList() As Container.List
                Get
                    Return Me._ContainerList
                End Get
                Set(ByVal value As Container.List)
                    Me._ContainerList = value
                End Set
            End Property

            Private _PreviousConnectionList As Connection.List
            Public Property PreviousConnectionList() As Connection.List
                Get
                    Return _PreviousConnectionList
                End Get
                Set(ByVal value As Connection.List)
                    _PreviousConnectionList = value
                End Set
            End Property

            Private _PreviousContainerList As Container.List
            Public Property PreviousContainerList() As Container.List
                Get
                    Return _PreviousContainerList
                End Get
                Set(ByVal value As Container.List)
                    _PreviousContainerList = value
                End Set
            End Property

            Private _SaveSecurity As Security.Save
            Public Property SaveSecurity() As Security.Save
                Get
                    Return Me._SaveSecurity
                End Get
                Set(ByVal value As Security.Save)
                    Me._SaveSecurity = value
                End Set
            End Property
#End Region

#Region "Public Methods"
            Public Overridable Function Load() As Boolean
                Throw New System.NotImplementedException
            End Function

            Public Overridable Function Save() As Boolean
                Throw New System.NotImplementedException
            End Function
#End Region

#Region "Protected Variables"
            Protected pW As String = App.Info.General.EncryptionKey
            Protected selNode As TreeNode
#End Region

#Region "Protected Functions"
            Protected Function Authenticate(ByVal Value As String, ByVal CompareToOriginalValue As Boolean, Optional ByVal RootInfo As mRemoteNG.Root.Info = Nothing) As Boolean
                If CompareToOriginalValue Then
                    Do Until Security.Crypt.Decrypt(Value, pW) <> Value
                        pW = Tools.Misc.PasswordDialog(False)

                        If pW = "" Then
                            Return False
                        End If
                    Loop
                Else
                    Do Until Security.Crypt.Decrypt(Value, pW) = "ThisIsProtected"
                        pW = Tools.Misc.PasswordDialog(False)

                        If pW = "" Then
                            Return False
                        End If
                    Loop

                    RootInfo.Password = True
                    RootInfo.PasswordString = pW
                End If

                Return True
            End Function
#End Region
        End Class
    End Namespace
End Namespace
