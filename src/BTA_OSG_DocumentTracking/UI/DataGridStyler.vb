Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class DataGridStyler
        Private Sub New()
        End Sub

        Public Shared Sub ApplyCivicStyle(dgv As DataGridView)
            If dgv Is Nothing Then Return

            dgv.AutoGenerateColumns = False
            dgv.EnableHeadersVisualStyles = False
            dgv.BackgroundColor = CivicCalmTheme.ColorSurface
            dgv.BorderStyle = BorderStyle.FixedSingle
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            dgv.GridColor = CivicCalmTheme.ColorBorder
            dgv.ColumnHeadersHeight = 32
            dgv.RowTemplate.Height = 28
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = False
            dgv.RowHeadersVisible = False

            ' Enable DoubleBuffered via reflection to prevent flicker during fast scrolling
            Dim prop = GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            If prop IsNot Nothing Then
                prop.SetValue(dgv, True, Nothing)
            End If

            ' Column Headers (Tier 2 Recessed Well)
            dgv.ColumnHeadersDefaultCellStyle.BackColor = CivicCalmTheme.ColorWell
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.ColumnHeadersDefaultCellStyle.Font = CivicCalmTheme.FontSectionHeader

            ' Default Row Styling (Pure White)
            dgv.DefaultCellStyle.BackColor = CivicCalmTheme.ColorSurface
            dgv.DefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
            dgv.DefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.DefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary

            ' Alternating Row Styling (Soft Off-White Tint #F9FAFB)
            dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F9FAFB")
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary
        End Sub

        Public Shared Sub SetEmptyState(dgv As DataGridView, watermarkLabel As Label, Optional customMessage As String = Nothing)
            If dgv IsNot Nothing Then
                dgv.Visible = True
            End If
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = If(customMessage, "No documents match the current filter criteria. Press Alt+C to clear filters.")
                watermarkLabel.ForeColor = CivicCalmTheme.ColorInkMuted
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetLoadingState(dgv As DataGridView, watermarkLabel As Label)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = "Loading document records..."
                watermarkLabel.ForeColor = CivicCalmTheme.ColorInkMuted
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetErrorState(dgv As DataGridView, watermarkLabel As Label, errorMessage As String)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Text = "Unable to retrieve records: " & errorMessage
                watermarkLabel.ForeColor = CivicCalmTheme.ColorDanger
                watermarkLabel.Visible = True
                watermarkLabel.BringToFront()
            End If
        End Sub

        Public Shared Sub SetPopulatedState(dgv As DataGridView, watermarkLabel As Label)
            If watermarkLabel IsNot Nothing Then
                watermarkLabel.Visible = False
            End If
            If dgv IsNot Nothing Then
                dgv.Visible = True
            End If
        End Sub
    End Class
End Namespace
