Imports System.IO

Namespace Config
    Namespace Connections
        Public Class VRDCSV
            Inherits Base
#Region "Public Methods"
            Public Overrides Function Load() As Boolean
                Throw New System.NotImplementedException
            End Function

            Public Overrides Function Save() As Boolean
                SaveTovRDCSV()
                Return True
            End Function
#End Region

#Region "Private Variables"
            Private csvWr As StreamWriter
#End Region

#Region "Private Methods"
            Private Sub SaveTovRDCSV()
                If App.Runtime.IsConnectionsFileLoaded = False Then
                    Exit Sub
                End If

                Dim tN As TreeNode
                tN = RootTreeNode.Clone

                Dim tNC As TreeNodeCollection
                tNC = tN.Nodes

                csvWr = New StreamWriter(ConnectionFileName)

                SaveNodevRDCSV(tNC)

                csvWr.Close()
            End Sub

            Private Sub SaveNodevRDCSV(ByVal tNC As TreeNodeCollection)
                For Each node As TreeNode In tNC
                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Then
                        Dim curConI As Connection.Info = node.Tag

                        If curConI.Protocol = Connection.Protocol.Protocols.RDP Then
                            WritevRDCSVLine(curConI)
                        End If
                    ElseIf Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then
                        SaveNodevRDCSV(node.Nodes)
                    End If
                Next
            End Sub

            Private Sub WritevRDCSVLine(ByVal con As Connection.Info)
                Dim nodePath As String = con.TreeNode.FullPath

                Dim firstSlash As Integer = nodePath.IndexOf("\")
                nodePath = nodePath.Remove(0, firstSlash + 1)
                Dim lastSlash As Integer = nodePath.LastIndexOf("\")

                If lastSlash > 0 Then
                    nodePath = nodePath.Remove(lastSlash)
                Else
                    nodePath = ""
                End If

                csvWr.WriteLine(con.Name & ";" & con.Hostname & ";" & con.MacAddress & ";;" & con.Port & ";" & con.UseConsoleSession & ";" & nodePath)
            End Sub
#End Region
        End Class
    End Namespace
End Namespace
