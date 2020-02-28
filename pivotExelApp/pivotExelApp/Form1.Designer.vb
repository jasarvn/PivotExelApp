<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.txtRawData = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtReferenceData = New System.Windows.Forms.TextBox()
        Me.btnSelectRaw = New System.Windows.Forms.Button()
        Me.btnSelectReference = New System.Windows.Forms.Button()
        Me.btnCreatePivot = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtRawData
        '
        Me.txtRawData.Location = New System.Drawing.Point(28, 45)
        Me.txtRawData.Name = "txtRawData"
        Me.txtRawData.Size = New System.Drawing.Size(349, 20)
        Me.txtRawData.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DimGray
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pivot Table Creator"
        '
        'txtReferenceData
        '
        Me.txtReferenceData.Location = New System.Drawing.Point(28, 71)
        Me.txtReferenceData.Name = "txtReferenceData"
        Me.txtReferenceData.Size = New System.Drawing.Size(349, 20)
        Me.txtReferenceData.TabIndex = 2
        '
        'btnSelectRaw
        '
        Me.btnSelectRaw.Location = New System.Drawing.Point(383, 42)
        Me.btnSelectRaw.Name = "btnSelectRaw"
        Me.btnSelectRaw.Size = New System.Drawing.Size(143, 23)
        Me.btnSelectRaw.TabIndex = 3
        Me.btnSelectRaw.Text = "Select Raw Data"
        Me.btnSelectRaw.UseVisualStyleBackColor = True
        '
        'btnSelectReference
        '
        Me.btnSelectReference.Location = New System.Drawing.Point(383, 69)
        Me.btnSelectReference.Name = "btnSelectReference"
        Me.btnSelectReference.Size = New System.Drawing.Size(143, 23)
        Me.btnSelectReference.TabIndex = 4
        Me.btnSelectReference.Text = "Select Reference Table"
        Me.btnSelectReference.UseVisualStyleBackColor = True
        '
        'btnCreatePivot
        '
        Me.btnCreatePivot.Location = New System.Drawing.Point(152, 97)
        Me.btnCreatePivot.Name = "btnCreatePivot"
        Me.btnCreatePivot.Size = New System.Drawing.Size(143, 23)
        Me.btnCreatePivot.TabIndex = 5
        Me.btnCreatePivot.Text = "Create Pivot Table"
        Me.btnCreatePivot.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(7, 173)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(69, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(557, 198)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnCreatePivot)
        Me.Controls.Add(Me.btnSelectReference)
        Me.Controls.Add(Me.btnSelectRaw)
        Me.Controls.Add(Me.txtReferenceData)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtRawData)
        Me.Name = "Form1"
        Me.Text = "Pivot Table Creator"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtRawData As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtReferenceData As TextBox
    Friend WithEvents btnSelectRaw As Button
    Friend WithEvents btnSelectReference As Button
    Friend WithEvents btnCreatePivot As Button
    Friend WithEvents PictureBox1 As PictureBox
End Class
