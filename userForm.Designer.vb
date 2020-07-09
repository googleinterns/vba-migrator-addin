'Copyright 2020 Google LLC
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'   https://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class userForm
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
        Me.ClientIdInput = New System.Windows.Forms.TextBox()
        Me.ClientSecretIdInput = New System.Windows.Forms.TextBox()
        Me.ClientId = New System.Windows.Forms.Label()
        Me.ClientSecretId = New System.Windows.Forms.Label()
        Me.Submit = New System.Windows.Forms.Button()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pathInput = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ClientIdInput
        '
        Me.ClientIdInput.Location = New System.Drawing.Point(35, 55)
        Me.ClientIdInput.Name = "ClientIdInput"
        Me.ClientIdInput.Size = New System.Drawing.Size(422, 22)
        Me.ClientIdInput.TabIndex = 0
        '
        'ClientSecretIdInput
        '
        Me.ClientSecretIdInput.Location = New System.Drawing.Point(35, 131)
        Me.ClientSecretIdInput.Name = "ClientSecretIdInput"
        Me.ClientSecretIdInput.Size = New System.Drawing.Size(422, 22)
        Me.ClientSecretIdInput.TabIndex = 1
        '
        'ClientId
        '
        Me.ClientId.AutoSize = True
        Me.ClientId.Location = New System.Drawing.Point(32, 26)
        Me.ClientId.Name = "ClientId"
        Me.ClientId.Size = New System.Drawing.Size(54, 17)
        Me.ClientId.TabIndex = 2
        Me.ClientId.Text = "ClientId"
        '
        'ClientSecretId
        '
        Me.ClientSecretId.AutoSize = True
        Me.ClientSecretId.Location = New System.Drawing.Point(32, 102)
        Me.ClientSecretId.Name = "ClientSecretId"
        Me.ClientSecretId.Size = New System.Drawing.Size(103, 17)
        Me.ClientSecretId.TabIndex = 3
        Me.ClientSecretId.Text = "Client Secret Id"
        '
        'Submit
        '
        Me.Submit.Location = New System.Drawing.Point(266, 281)
        Me.Submit.Name = "Submit"
        Me.Submit.Size = New System.Drawing.Size(79, 34)
        Me.Submit.TabIndex = 4
        Me.Submit.Text = "Submit"
        Me.Submit.UseVisualStyleBackColor = True
        '
        'Cancel
        '
        Me.Cancel.Location = New System.Drawing.Point(376, 281)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(81, 34)
        Me.Cancel.TabIndex = 5
        Me.Cancel.Text = "Cancel"
        Me.Cancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 171)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Provide path to copy temp file"
        '
        'pathInput
        '
        Me.pathInput.Location = New System.Drawing.Point(35, 200)
        Me.pathInput.Name = "pathInput"
        Me.pathInput.Size = New System.Drawing.Size(422, 22)
        Me.pathInput.TabIndex = 7
        '
        'userForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(469, 327)
        Me.Controls.Add(Me.pathInput)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.Submit)
        Me.Controls.Add(Me.ClientSecretId)
        Me.Controls.Add(Me.ClientId)
        Me.Controls.Add(Me.ClientSecretIdInput)
        Me.Controls.Add(Me.ClientIdInput)
        Me.Name = "userForm"
        Me.Text = "User Credentials Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ClientIdInput As Windows.Forms.TextBox
    Friend WithEvents ClientSecretIdInput As Windows.Forms.TextBox
    Friend WithEvents ClientId As Windows.Forms.Label
    Friend WithEvents ClientSecretId As Windows.Forms.Label
    Friend WithEvents Submit As Windows.Forms.Button
    Friend WithEvents Cancel As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents pathInput As Windows.Forms.TextBox
End Class
