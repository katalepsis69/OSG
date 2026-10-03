Option Explicit On
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace BTA_OSG
    Public Class FormInputPrompt
        Inherits Form

        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property PromptValue As String = ""

        Private txtInput As TextBox
        Private btnSubmit As Button
        Private btnCancel As Button

        Public Sub New(titleText As String, promptText As String, Optional defaultValue As String = "", Optional isMultiline As Boolean = True)
            AppAssets.ApplyFormIcon(Me)
            Me.Text = titleText
            Me.Size = If(isMultiline, New Size(520, 320), New Size(520, 220))
            Me.MinimumSize = If(isMultiline, New Size(440, 280), New Size(440, 200))
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = CivicCalmTheme.ColorCanvas
            Me.Font = CivicCalmTheme.FontBody
            Me.KeyPreview = True

            Dim pnlCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = CivicCalmTheme.ColorSurface,
                .Padding = New Padding(18)
            }

            Dim lblHeader As New Label With {
                .Text = titleText,
                .Font = CivicCalmTheme.FontSectionHeader,
                .ForeColor = CivicCalmTheme.ColorInk,
                .Dock = DockStyle.Top,
                .Height = 26
            }

            Dim lblPrompt As New Label With {
                .Text = promptText,
                .Font = CivicCalmTheme.FontFieldLabel,
                .ForeColor = CivicCalmTheme.ColorInkMuted,
                .Dock = DockStyle.Top,
                .Height = 32,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim pnlInput As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 6, 0, 10)
            }

            txtInput = New TextBox With {
                .Dock = DockStyle.Fill,
                .Multiline = isMultiline,
                .AcceptsReturn = isMultiline,
                .ScrollBars = If(isMultiline, ScrollBars.Vertical, ScrollBars.None),
                .Font = CivicCalmTheme.FontBody,
                .BackColor = CivicCalmTheme.ColorSurface,
                .ForeColor = CivicCalmTheme.ColorInk,
                .BorderStyle = BorderStyle.FixedSingle,
                .Text = defaultValue
            }
            pnlInput.Controls.Add(txtInput)

            Dim flwButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .Height = 42,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .Padding = New Padding(0, 4, 0, 0)
            }

            btnCancel = New Button With {
                .Text = "&Cancel",
                .Size = New Size(90, 32),
                .BackColor = CivicCalmTheme.ColorWell,
                .ForeColor = CivicCalmTheme.ColorInk,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(8, 0, 0, 0)
            }
            btnCancel.FlatAppearance.BorderColor = CivicCalmTheme.ColorBorder
            AddHandler btnCancel.Click, Sub()
                                            Me.DialogResult = DialogResult.Cancel
                                            Me.Close()
                                        End Sub

            btnSubmit = New Button With {
                .Text = "&Submit",
                .Size = New Size(95, 32),
                .BackColor = CivicCalmTheme.ColorPrimary,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = CivicCalmTheme.FontFieldLabel,
                .Cursor = Cursors.Hand,
                .Margin = New Padding(0)
            }
            btnSubmit.FlatAppearance.BorderSize = 0
            AddHandler btnSubmit.Click, AddressOf OnSubmit

            flwButtons.Controls.AddRange(New Control() {btnCancel, btnSubmit})

            pnlCard.Controls.Add(pnlInput)
            pnlCard.Controls.Add(flwButtons)
            pnlCard.Controls.Add(lblPrompt)
            pnlCard.Controls.Add(lblHeader)

            Me.Controls.Add(pnlCard)
            Me.AcceptButton = btnSubmit
            Me.CancelButton = btnCancel
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            If txtInput IsNot Nothing Then
                txtInput.Focus()
                txtInput.SelectAll()
            End If
        End Sub

        Private Sub OnSubmit(sender As Object, e As EventArgs)
            Dim val = txtInput.Text.Trim()
            If String.IsNullOrWhiteSpace(val) Then
                MessageBox.Show("Please enter the requested information.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtInput.Focus()
                Return
            End If

            PromptValue = val
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub
    End Class
End Namespace
