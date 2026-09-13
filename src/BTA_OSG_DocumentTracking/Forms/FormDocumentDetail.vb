Imports System.Drawing
Imports System.Windows.Forms


Namespace BTA_OSG
    Public Class FormDocumentDetail
        Inherits Form

        Private _document As Document
        
        Private tcMain As TabControl
        Private tpOverview As TabPage
        Private tpDirectives As TabPage
        Private tpRouting As TabPage
        Private tpMovements As TabPage
        
        Public Sub New(doc As Document)
            _document = doc
            InitializeComponent()
            ApplyTheme()
            LoadData()
        End Sub

        Private Sub InitializeComponent()
            Me.Size = New Size(800, 600)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Text = "Document Detail - " & _document.DocCode
            
            Dim pnlHeader = New Panel() With {.Dock = DockStyle.Top, .Height = 60}
            Dim lblTitle = New Label() With {.Text = _document.DocCode & " - " & _document.Title, .Font = New Font("Segoe UI", 14, FontStyle.Bold), .Dock = DockStyle.Fill}
            Dim btnPdf = New Button() With {.Text = "Launch PDF", .Dock = DockStyle.Right, .Width = 100}
            pnlHeader.Controls.Add(lblTitle)
            pnlHeader.Controls.Add(btnPdf)
            
            tcMain = New TabControl() With {.Dock = DockStyle.Fill}
            tpOverview = New TabPage("Overview")
            tpDirectives = New TabPage("Directives")
            tpRouting = New TabPage("Routing")
            tpMovements = New TabPage("Storage Movements")
            
            tcMain.TabPages.Add(tpOverview)
            tcMain.TabPages.Add(tpDirectives)
            tcMain.TabPages.Add(tpRouting)
            tcMain.TabPages.Add(tpMovements)
            
            Me.Controls.Add(tcMain)
            Me.Controls.Add(pnlHeader)
        End Sub

        Private Sub ApplyTheme()
            Me.BackColor = Color.FromArgb(15, 23, 42)
            tcMain.BackColor = Color.FromArgb(30, 41, 59)
            ' Style tabs and grids here
        End Sub

        Private Sub LoadData()
            ' Populate tabs using services based on _document.DocumentId
        End Sub
    End Class
End Namespace
