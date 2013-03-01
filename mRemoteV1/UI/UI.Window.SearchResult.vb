Imports WeifenLuo.WinFormsUI.Docking
Imports mRemoteNG.App.Runtime
Imports System.IO

Namespace UI
    Namespace Window
        Public Class SearchResults
            Inherits UI.Window.Base

#Region "Form Init"
            Private components As System.ComponentModel.IContainer
            Friend WithEvents imgList As System.Windows.Forms.ImageList
            Friend WithEvents cMenScreenshotCopy As System.Windows.Forms.ToolStripMenuItem
            Friend WithEvents cMenScreenshotSave As System.Windows.Forms.ToolStripMenuItem
            Friend WithEvents dlgSaveSingleImage As System.Windows.Forms.SaveFileDialog
            Friend WithEvents dlgSaveAllImages As System.Windows.Forms.FolderBrowserDialog
            Friend WithEvents lvSearchResults As System.Windows.Forms.ListView
            Friend WithEvents clmMessage As System.Windows.Forms.ColumnHeader
            Public Results As List(Of TreeNode)

            Private Sub InitializeComponent()
                Me.components = New System.ComponentModel.Container()
                Me.lvSearchResults = New System.Windows.Forms.ListView()
                Me.clmMessage = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
                Me.imgList = New System.Windows.Forms.ImageList(Me.components)
                Me.SuspendLayout()
                '
                'lvSearchResults
                '
                Me.lvSearchResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.lvSearchResults.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.lvSearchResults.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.clmMessage})
                Me.lvSearchResults.FullRowSelect = True
                Me.lvSearchResults.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
                Me.lvSearchResults.HideSelection = False
                Me.lvSearchResults.Location = New System.Drawing.Point(-2, -3)
                Me.lvSearchResults.Name = "lvSearchResults"
                Me.lvSearchResults.ShowGroups = False
                Me.lvSearchResults.Size = New System.Drawing.Size(547, 230)
                Me.lvSearchResults.SmallImageList = Me.imgList
                Me.lvSearchResults.TabIndex = 11
                Me.lvSearchResults.UseCompatibleStateImageBehavior = False
                Me.lvSearchResults.View = System.Windows.Forms.View.Details
                '
                'clmMessage
                '
                Me.clmMessage.Text = Global.mRemoteNG.My.Language.strColumnMessage
                Me.clmMessage.Width = 500
                '
                'imgList
                '
                Me.imgList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
                Me.imgList.ImageSize = New System.Drawing.Size(16, 16)
                Me.imgList.TransparentColor = System.Drawing.Color.Transparent
                '
                'SearchResults
                '
                Me.ClientSize = New System.Drawing.Size(542, 223)
                Me.Controls.Add(Me.lvSearchResults)
                Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.HideOnClose = True
                Me.Icon = Global.mRemoteNG.My.Resources.Resources.Screenshot_Icon
                Me.Name = "SearchResults"
                Me.TabText = "Search Results"
                Me.Text = "Search results"
                Me.ResumeLayout(False)

            End Sub
#End Region

#Region "Form Stuff"
            Private Sub FillImageList()
                Try
                    Me.imgList.Images.Add(My.Resources.Root)
                    Me.imgList.Images.Add(My.Resources.Folder)
                    Me.imgList.Images.Add(My.Resources.Play)
                    Me.imgList.Images.Add(My.Resources.Pause)
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "FillImageList (UI.Window.SearchResult) failed" & vbNewLine & ex.Message, True)
                End Try
            End Sub

            Private Sub SearchResult_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
                FillImageList()
            End Sub

            Private Sub result_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvSearchResults.Click
                If lvSearchResults.SelectedIndices.Count = 0 Then Return
                Windows.treeForm.tvConnections.SelectedNode = Results(lvSearchResults.SelectedIndices(0))
            End Sub

            Private Sub result_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvSearchResults.DoubleClick
                If lvSearchResults.SelectedIndices.Count = 0 Then Return
                Windows.treeForm.tvConnections.SelectedNode = Results(lvSearchResults.SelectedIndices(0))
                App.Runtime.OpenConnection()
            End Sub

            Private Sub lvSearchResults_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvSearchResults.KeyDown
                If lvSearchResults.SelectedIndices.Count = 0 Then Return
                If e.KeyCode = Keys.Enter Then
                    For Each selectedIndex As Integer In lvSearchResults.SelectedIndices
                        Windows.treeForm.tvConnections.SelectedNode = Results(selectedIndex)
                        App.Runtime.OpenConnection()
                    Next
                End If
            End Sub
#End Region


#Region "Public Methods"
            Public Sub New(ByVal Panel As DockContent)
                Me.WindowType = Type.SearchResult
                Me.DockPnl = Panel
                Me.InitializeComponent()
            End Sub
#End Region

        End Class
    End Namespace
End Namespace