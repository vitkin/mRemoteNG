Imports System.IO

Namespace Config
    Namespace Connections
        Public Class CSV
            Inherits Base
#Region "Public Methods"
            Public Overrides Function Load() As Boolean
                Throw New System.NotImplementedException
            End Function

            Public Overrides Function Save() As Boolean
                SaveTomRCSV()
                Return True
            End Function
#End Region

#Region "Private Variables"
            Private csvWr As StreamWriter
#End Region

#Region "Private Methods"
            Private Sub SaveTomRCSV()
                If Not App.Runtime.IsConnectionsFileLoaded Then Return

                Dim tN As TreeNode
                tN = RootTreeNode.Clone

                Dim tNC As TreeNodeCollection
                tNC = tN.Nodes

                csvWr = New StreamWriter(ConnectionFileName)


                Dim csvLn As String = String.Empty

                csvLn += "Name;Folder;Description;Icon;Panel;"

                If SaveSecurity.Username Then
                    csvLn += "Username;"
                End If

                If SaveSecurity.Password Then
                    csvLn += "Password;"
                End If

                If SaveSecurity.Domain Then
                    csvLn += "Domain;"
                End If

                csvLn += "Hostname;Protocol;PuttySession;Port;ConnectToConsole;RenderingEngine;ICAEncryptionStrength;RDPAuthenticationLevel;Colors;Resolution;DisplayWallpaper;DisplayThemes;EnableFontSmoothing;EnableDesktopComposition;CacheBitmaps;RedirectDiskDrives;RedirectPorts;RedirectPrinters;RedirectSmartCards;RedirectSound;RedirectKeys;PreExtApp;PostExtApp;MacAddress;UserField;ExtApp;VNCCompression;VNCEncoding;VNCAuthMode;VNCProxyType;VNCProxyIP;VNCProxyPort;VNCProxyUsername;VNCProxyPassword;VNCColors;VNCSmartSizeMode;VNCViewOnly;RDGatewayUsageMethod;RDGatewayHostname;RDGatewayUseConnectionCredentials;RDGatewayUsername;RDGatewayPassword;RDGatewayDomain;"

                If SaveSecurity.Inheritance Then
                    csvLn += "InheritCacheBitmaps;InheritColors;InheritDescription;InheritDisplayThemes;InheritDisplayWallpaper;InheritEnableFontSmoothing;InheritEnableDesktopComposition;InheritDomain;InheritIcon;InheritPanel;InheritPassword;InheritPort;InheritProtocol;InheritPuttySession;InheritRedirectDiskDrives;InheritRedirectKeys;InheritRedirectPorts;InheritRedirectPrinters;InheritRedirectSmartCards;InheritRedirectSound;InheritResolution;InheritUseConsoleSession;InheritRenderingEngine;InheritUsername;InheritICAEncryptionStrength;InheritRDPAuthenticationLevel;InheritPreExtApp;InheritPostExtApp;InheritMacAddress;InheritUserField;InheritExtApp;InheritVNCCompression;InheritVNCEncoding;InheritVNCAuthMode;InheritVNCProxyType;InheritVNCProxyIP;InheritVNCProxyPort;InheritVNCProxyUsername;InheritVNCProxyPassword;InheritVNCColors;InheritVNCSmartSizeMode;InheritVNCViewOnly;InheritRDGatewayUsageMethod;InheritRDGatewayHostname;InheritRDGatewayUseConnectionCredentials;InheritRDGatewayUsername;InheritRDGatewayPassword;InheritRDGatewayDomain"
                End If

                csvWr.WriteLine(csvLn)

                SaveNodemRCSV(tNC)

                csvWr.Close()
            End Sub

            Private Sub SaveNodemRCSV(ByVal tNC As TreeNodeCollection)
                For Each node As TreeNode In tNC
                    If Tree.Node.GetNodeType(node) = Tree.Node.Type.Connection Then
                        Dim curConI As Connection.Info = node.Tag

                        WritemRCSVLine(curConI)
                    ElseIf Tree.Node.GetNodeType(node) = Tree.Node.Type.Container Then
                        SaveNodemRCSV(node.Nodes)
                    End If
                Next
            End Sub

            Private Sub WritemRCSVLine(ByVal con As Connection.Info)
                Dim nodePath As String = con.TreeNode.FullPath

                Dim firstSlash As Integer = nodePath.IndexOf("\")
                nodePath = nodePath.Remove(0, firstSlash + 1)
                Dim lastSlash As Integer = nodePath.LastIndexOf("\")

                If lastSlash > 0 Then
                    nodePath = nodePath.Remove(lastSlash)
                Else
                    nodePath = ""
                End If

                Dim csvLn As String = String.Empty

                csvLn += con.Name & ";" & nodePath & ";" & con.Description & ";" & con.Icon & ";" & con.Panel & ";"

                If SaveSecurity.Username Then
                    csvLn += con.Username & ";"
                End If

                If SaveSecurity.Password Then
                    csvLn += con.Password & ";"
                End If

                If SaveSecurity.Domain Then
                    csvLn += con.Domain & ";"
                End If

                csvLn += con.Hostname & ";" & con.Protocol.ToString & ";" & con.PuttySession & ";" & con.Port & ";" & con.UseConsoleSession & ";" & con.RenderingEngine.ToString & ";" & con.ICAEncryption.ToString & ";" & con.RDPAuthenticationLevel.ToString & ";" & con.Colors.ToString & ";" & con.Resolution.ToString & ";" & con.DisplayWallpaper & ";" & con.DisplayThemes & ";" & con.EnableFontSmoothing & ";" & con.EnableDesktopComposition & ";" & con.CacheBitmaps & ";" & con.RedirectDiskDrives & ";" & con.RedirectPorts & ";" & con.RedirectPrinters & ";" & con.RedirectSmartCards & ";" & con.RedirectSound.ToString & ";" & con.RedirectKeys & ";" & con.PreExtApp & ";" & con.PostExtApp & ";" & con.MacAddress & ";" & con.UserField & ";" & con.ExtApp & ";" & con.VNCCompression.ToString & ";" & con.VNCEncoding.ToString & ";" & con.VNCAuthMode.ToString & ";" & con.VNCProxyType.ToString & ";" & con.VNCProxyIP & ";" & con.VNCProxyPort & ";" & con.VNCProxyUsername & ";" & con.VNCProxyPassword & ";" & con.VNCColors.ToString & ";" & con.VNCSmartSizeMode.ToString & ";" & con.VNCViewOnly & ";"

                If SaveSecurity.Inheritance Then
                    csvLn += con.Inherit.CacheBitmaps & ";" & con.Inherit.Colors & ";" & con.Inherit.Description & ";" & con.Inherit.DisplayThemes & ";" & con.Inherit.DisplayWallpaper & ";" & con.Inherit.EnableFontSmoothing & ";" & con.Inherit.EnableDesktopComposition & ";" & con.Inherit.Domain & ";" & con.Inherit.Icon & ";" & con.Inherit.Panel & ";" & con.Inherit.Password & ";" & con.Inherit.Port & ";" & con.Inherit.Protocol & ";" & con.Inherit.PuttySession & ";" & con.Inherit.RedirectDiskDrives & ";" & con.Inherit.RedirectKeys & ";" & con.Inherit.RedirectPorts & ";" & con.Inherit.RedirectPrinters & ";" & con.Inherit.RedirectSmartCards & ";" & con.Inherit.RedirectSound & ";" & con.Inherit.Resolution & ";" & con.Inherit.UseConsoleSession & ";" & con.Inherit.RenderingEngine & ";" & con.Inherit.Username & ";" & con.Inherit.ICAEncryption & ";" & con.Inherit.RDPAuthenticationLevel & ";" & con.Inherit.PreExtApp & ";" & con.Inherit.PostExtApp & ";" & con.Inherit.MacAddress & ";" & con.Inherit.UserField & ";" & con.Inherit.ExtApp & ";" & con.Inherit.VNCCompression & ";" & con.Inherit.VNCEncoding & ";" & con.Inherit.VNCAuthMode & ";" & con.Inherit.VNCProxyType & ";" & con.Inherit.VNCProxyIP & ";" & con.Inherit.VNCProxyPort & ";" & con.Inherit.VNCProxyUsername & ";" & con.Inherit.VNCProxyPassword & ";" & con.Inherit.VNCColors & ";" & con.Inherit.VNCSmartSizeMode & ";" & con.Inherit.VNCViewOnly
                End If

                csvWr.WriteLine(csvLn)
            End Sub
#End Region
        End Class
    End Namespace
End Namespace
