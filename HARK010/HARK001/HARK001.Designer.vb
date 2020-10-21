'/*-----------------------------------------------------------------------------
' * COPYRIGHT(C) ITI CORPORATION 2020
' * ITI CONFIDENTIAL AND PROPRIETARY
' *
' * All rights reserved by ITI Corporation.
' *-----------------------------------------------------------------------------/
Imports NAppUpdate.Framework

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class HARK001
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DateYearField1 As GrapeCity.Win.Editors.Fields.DateYearField = New GrapeCity.Win.Editors.Fields.DateYearField()
        Dim DateLiteralField1 As GrapeCity.Win.Editors.Fields.DateLiteralField = New GrapeCity.Win.Editors.Fields.DateLiteralField()
        Dim DateMonthField1 As GrapeCity.Win.Editors.Fields.DateMonthField = New GrapeCity.Win.Editors.Fields.DateMonthField()
        Dim DateLiteralField2 As GrapeCity.Win.Editors.Fields.DateLiteralField = New GrapeCity.Win.Editors.Fields.DateLiteralField()
        Dim DateDayField1 As GrapeCity.Win.Editors.Fields.DateDayField = New GrapeCity.Win.Editors.Fields.DateDayField()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HARK001))
        Me.SSBar = New System.Windows.Forms.StatusStrip()
        Me.TSSVersion = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TSSRowCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TLayoutContainer = New GrapeCity.Win.Containers.GcTableLayoutContainer()
        Me.TableColumn1 = New GrapeCity.Win.Containers.TableColumn()
        Me.TableColumn2 = New GrapeCity.Win.Containers.TableColumn()
        Me.TableRow1 = New GrapeCity.Win.Containers.TableRow()
        Me.TableRow2 = New GrapeCity.Win.Containers.TableRow()
        Me.Flc検索 = New GrapeCity.Win.Containers.GcFlowLayoutContainer()
        Me.Container検索 = New GrapeCity.Win.Containers.GcContainer()
        Me.Bt印刷 = New GrapeCity.Win.Buttons.GcButton()
        Me.txt需要先 = New GrapeCity.Win.Editors.GcTextBox(Me.components)
        Me.Lbl需要先コード指定 = New System.Windows.Forms.Label()
        Me.txt得意先 = New GrapeCity.Win.Editors.GcTextBox(Me.components)
        Me.Lbl得意先コード指定 = New System.Windows.Forms.Label()
        Me.txt商品コード = New GrapeCity.Win.Editors.GcTextBox(Me.components)
        Me.txt相手先品番 = New GrapeCity.Win.Editors.GcTextBox(Me.components)
        Me.Lbl相手先品番 = New System.Windows.Forms.Label()
        Me.Lbl商品 = New System.Windows.Forms.Label()
        Me.Lbl注意２ = New System.Windows.Forms.Label()
        Me.Lbl注意１ = New System.Windows.Forms.Label()
        Me.Cmb需要先 = New System.Windows.Forms.ComboBox()
        Me.Cmb得意先 = New System.Windows.Forms.ComboBox()
        Me.Lbl得意先 = New System.Windows.Forms.Label()
        Me.txtDate = New GrapeCity.Win.Editors.GcDate(Me.components)
        Me.DropDownButton6 = New GrapeCity.Win.Editors.DropDownButton()
        Me.Lbl対象日 = New System.Windows.Forms.Label()
        Me.Btクリア = New GrapeCity.Win.Buttons.GcButton()
        Me.Bt検索 = New GrapeCity.Win.Buttons.GcButton()
        Me.Lbl需要先 = New System.Windows.Forms.Label()
        Me.Flc汎用 = New GrapeCity.Win.Containers.GcFlowLayoutContainer()
        Me.Container汎用 = New GrapeCity.Win.Containers.GcContainer()
        Me.LblTitle = New System.Windows.Forms.Label()
        Me.Cmb事業所 = New System.Windows.Forms.ComboBox()
        Me.Cmb汎用 = New System.Windows.Forms.ComboBox()
        Me.Dgv = New System.Windows.Forms.DataGridView()
        Me.VBReport = New AdvanceSoftware.VBReport8.CellReport(Me.components)
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.SSBar.SuspendLayout()
        CType(Me.TLayoutContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TLayoutContainer.SuspendLayout()
        Me.Flc検索.SuspendLayout()
        Me.Container検索.SuspendLayout()
        CType(Me.txt需要先, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txt得意先, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txt商品コード, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txt相手先品番, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Flc汎用.SuspendLayout()
        Me.Container汎用.SuspendLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SSBar
        '
        Me.SSBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSSVersion, Me.TSSRowCount})
        Me.SSBar.Location = New System.Drawing.Point(0, 704)
        Me.SSBar.Name = "SSBar"
        Me.SSBar.Size = New System.Drawing.Size(1008, 25)
        Me.SSBar.TabIndex = 2
        Me.SSBar.Text = "SSBar"
        '
        'TSSVersion
        '
        Me.TSSVersion.AutoSize = False
        Me.TSSVersion.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TSSVersion.Name = "TSSVersion"
        Me.TSSVersion.Size = New System.Drawing.Size(100, 20)
        Me.TSSVersion.Text = "TSSVersion"
        '
        'TSSRowCount
        '
        Me.TSSRowCount.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TSSRowCount.Name = "TSSRowCount"
        Me.TSSRowCount.Size = New System.Drawing.Size(99, 20)
        Me.TSSRowCount.Text = "TSSRowCount"
        '
        'TLayoutContainer
        '
        Me.TLayoutContainer.Columns.AddRange(New GrapeCity.Win.Containers.TableColumn() {Me.TableColumn1, Me.TableColumn2})
        Me.TLayoutContainer.Rows.AddRange(New GrapeCity.Win.Containers.TableRow() {Me.TableRow1, Me.TableRow2})
        Me.TLayoutContainer.CellInfos.AddRange(New GrapeCity.Win.Containers.CellInfo() {New GrapeCity.Win.Containers.CellInfo(New GrapeCity.Win.Containers.CellPosition(0, 0), 2, 1)})
        Me.TLayoutContainer.Controls.Add(Me.Flc検索, 0, 1)
        Me.TLayoutContainer.Controls.Add(Me.Flc汎用, 0, 0)
        Me.TLayoutContainer.Controls.Add(Me.Dgv, 1, 1)
        Me.TLayoutContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TLayoutContainer.Location = New System.Drawing.Point(0, 0)
        Me.TLayoutContainer.Name = "TLayoutContainer"
        Me.TLayoutContainer.Size = New System.Drawing.Size(1008, 704)
        Me.TLayoutContainer.TabIndex = 3
        '
        'TableColumn1
        '
        Me.TableColumn1.SizeType = System.Windows.Forms.SizeType.Absolute
        Me.TableColumn1.Width = 248.0!
        '
        'TableColumn2
        '
        Me.TableColumn2.Width = 100.0!
        '
        'TableRow1
        '
        Me.TableRow1.Height = 68.0!
        Me.TableRow1.SizeType = System.Windows.Forms.SizeType.Absolute
        '
        'TableRow2
        '
        Me.TableRow2.Height = 100.0!
        '
        'Flc検索
        '
        Me.Flc検索.Controls.Add(Me.Container検索)
        Me.Flc検索.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Flc検索.Location = New System.Drawing.Point(5, 73)
        Me.Flc検索.Name = "Flc検索"
        Me.Flc検索.Size = New System.Drawing.Size(242, 626)
        Me.Flc検索.TabIndex = 2
        '
        'Container検索
        '
        Me.Container検索.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(217, Byte), Integer), CType(CType(217, Byte), Integer))
        Me.Container検索.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Container検索.Controls.Add(Me.ComboBox1)
        Me.Container検索.Controls.Add(Me.Bt印刷)
        Me.Container検索.Controls.Add(Me.txt需要先)
        Me.Container検索.Controls.Add(Me.Lbl需要先コード指定)
        Me.Container検索.Controls.Add(Me.txt得意先)
        Me.Container検索.Controls.Add(Me.Lbl得意先コード指定)
        Me.Container検索.Controls.Add(Me.txt商品コード)
        Me.Container検索.Controls.Add(Me.txt相手先品番)
        Me.Container検索.Controls.Add(Me.Lbl相手先品番)
        Me.Container検索.Controls.Add(Me.Lbl商品)
        Me.Container検索.Controls.Add(Me.Lbl注意２)
        Me.Container検索.Controls.Add(Me.Lbl注意１)
        Me.Container検索.Controls.Add(Me.Cmb需要先)
        Me.Container検索.Controls.Add(Me.Cmb得意先)
        Me.Container検索.Controls.Add(Me.Lbl得意先)
        Me.Container検索.Controls.Add(Me.txtDate)
        Me.Container検索.Controls.Add(Me.Lbl対象日)
        Me.Container検索.Controls.Add(Me.Btクリア)
        Me.Container検索.Controls.Add(Me.Bt検索)
        Me.Container検索.Controls.Add(Me.Lbl需要先)
        Me.Container検索.Location = New System.Drawing.Point(3, 3)
        Me.Container検索.Name = "Container検索"
        Me.Container検索.Size = New System.Drawing.Size(239, 602)
        Me.Container検索.TabIndex = 0
        '
        'Bt印刷
        '
        Me.Bt印刷.BackColor = System.Drawing.SystemColors.Control
        Me.Bt印刷.Location = New System.Drawing.Point(12, 568)
        Me.Bt印刷.Name = "Bt印刷"
        Me.Bt印刷.Size = New System.Drawing.Size(94, 25)
        Me.Bt印刷.TabIndex = 103
        Me.Bt印刷.Text = "印刷"
        Me.Bt印刷.UseVisualStyleBackColor = False
        '
        'txt需要先
        '
        Me.txt需要先.ContentAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.txt需要先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txt需要先.Format = "9"
        Me.txt需要先.Location = New System.Drawing.Point(91, 180)
        Me.txt需要先.MaxLength = 10
        Me.txt需要先.MaxLengthUnit = GrapeCity.Win.Editors.LengthUnit.[Byte]
        Me.txt需要先.Name = "txt需要先"
        Me.txt需要先.ShowOverflowTip = True
        Me.txt需要先.Size = New System.Drawing.Size(131, 21)
        Me.txt需要先.TabIndex = 9
        '
        'Lbl需要先コード指定
        '
        Me.Lbl需要先コード指定.AutoSize = True
        Me.Lbl需要先コード指定.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl需要先コード指定.Location = New System.Drawing.Point(3, 181)
        Me.Lbl需要先コード指定.Name = "Lbl需要先コード指定"
        Me.Lbl需要先コード指定.Size = New System.Drawing.Size(68, 18)
        Me.Lbl需要先コード指定.TabIndex = 31
        Me.Lbl需要先コード指定.Text = "【需要先】"
        '
        'txt得意先
        '
        Me.txt得意先.ContentAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.txt得意先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txt得意先.Format = "9"
        Me.txt得意先.Location = New System.Drawing.Point(91, 153)
        Me.txt得意先.MaxLength = 10
        Me.txt得意先.MaxLengthUnit = GrapeCity.Win.Editors.LengthUnit.[Byte]
        Me.txt得意先.Name = "txt得意先"
        Me.txt得意先.ShowOverflowTip = True
        Me.txt得意先.Size = New System.Drawing.Size(131, 21)
        Me.txt得意先.TabIndex = 8
        '
        'Lbl得意先コード指定
        '
        Me.Lbl得意先コード指定.AutoSize = True
        Me.Lbl得意先コード指定.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl得意先コード指定.Location = New System.Drawing.Point(2, 154)
        Me.Lbl得意先コード指定.Name = "Lbl得意先コード指定"
        Me.Lbl得意先コード指定.Size = New System.Drawing.Size(68, 18)
        Me.Lbl得意先コード指定.TabIndex = 29
        Me.Lbl得意先コード指定.Text = "【得意先】"
        '
        'txt商品コード
        '
        Me.txt商品コード.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.txt商品コード.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txt商品コード.Format = "9"
        Me.txt商品コード.Location = New System.Drawing.Point(91, 98)
        Me.txt商品コード.MaxLength = 60
        Me.txt商品コード.MaxLengthUnit = GrapeCity.Win.Editors.LengthUnit.[Byte]
        Me.txt商品コード.Name = "txt商品コード"
        Me.txt商品コード.ShowOverflowTip = True
        Me.txt商品コード.Size = New System.Drawing.Size(131, 21)
        Me.txt商品コード.TabIndex = 6
        '
        'txt相手先品番
        '
        Me.txt相手先品番.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.txt相手先品番.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txt相手先品番.Format = "H"
        Me.txt相手先品番.Location = New System.Drawing.Point(91, 126)
        Me.txt相手先品番.MaxLength = 60
        Me.txt相手先品番.MaxLengthUnit = GrapeCity.Win.Editors.LengthUnit.[Byte]
        Me.txt相手先品番.Name = "txt相手先品番"
        Me.txt相手先品番.ShowOverflowTip = True
        Me.txt相手先品番.Size = New System.Drawing.Size(131, 21)
        Me.txt相手先品番.TabIndex = 7
        '
        'Lbl相手先品番
        '
        Me.Lbl相手先品番.AutoSize = True
        Me.Lbl相手先品番.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl相手先品番.Location = New System.Drawing.Point(3, 127)
        Me.Lbl相手先品番.Name = "Lbl相手先品番"
        Me.Lbl相手先品番.Size = New System.Drawing.Size(92, 18)
        Me.Lbl相手先品番.TabIndex = 26
        Me.Lbl相手先品番.Text = "【相手先品番】"
        '
        'Lbl商品
        '
        Me.Lbl商品.AutoSize = True
        Me.Lbl商品.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl商品.Location = New System.Drawing.Point(3, 99)
        Me.Lbl商品.Name = "Lbl商品"
        Me.Lbl商品.Size = New System.Drawing.Size(56, 18)
        Me.Lbl商品.TabIndex = 25
        Me.Lbl商品.Text = "【商品】"
        '
        'Lbl注意２
        '
        Me.Lbl注意２.AutoSize = True
        Me.Lbl注意２.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl注意２.ForeColor = System.Drawing.Color.Blue
        Me.Lbl注意２.Location = New System.Drawing.Point(102, 512)
        Me.Lbl注意２.Name = "Lbl注意２"
        Me.Lbl注意２.Size = New System.Drawing.Size(74, 20)
        Me.Lbl注意２.TabIndex = 24
        Me.Lbl注意２.Text = "青字：任意"
        '
        'Lbl注意１
        '
        Me.Lbl注意１.AutoSize = True
        Me.Lbl注意１.Font = New System.Drawing.Font("メイリオ", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl注意１.ForeColor = System.Drawing.Color.Red
        Me.Lbl注意１.Location = New System.Drawing.Point(16, 512)
        Me.Lbl注意１.Name = "Lbl注意１"
        Me.Lbl注意１.Size = New System.Drawing.Size(74, 20)
        Me.Lbl注意１.TabIndex = 23
        Me.Lbl注意１.Text = "赤字：必須"
        '
        'Cmb需要先
        '
        Me.Cmb需要先.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cmb需要先.DropDownWidth = 250
        Me.Cmb需要先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Cmb需要先.FormattingEnabled = True
        Me.Cmb需要先.Location = New System.Drawing.Point(90, 39)
        Me.Cmb需要先.Name = "Cmb需要先"
        Me.Cmb需要先.Size = New System.Drawing.Size(132, 26)
        Me.Cmb需要先.TabIndex = 4
        Me.Cmb需要先.Tag = "ID2"
        '
        'Cmb得意先
        '
        Me.Cmb得意先.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cmb得意先.DropDownWidth = 250
        Me.Cmb得意先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Cmb得意先.FormattingEnabled = True
        Me.Cmb得意先.Location = New System.Drawing.Point(90, 11)
        Me.Cmb得意先.Name = "Cmb得意先"
        Me.Cmb得意先.Size = New System.Drawing.Size(132, 26)
        Me.Cmb得意先.TabIndex = 3
        Me.Cmb得意先.Tag = "ID1"
        '
        'Lbl得意先
        '
        Me.Lbl得意先.AutoSize = True
        Me.Lbl得意先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl得意先.Location = New System.Drawing.Point(3, 15)
        Me.Lbl得意先.Name = "Lbl得意先"
        Me.Lbl得意先.Size = New System.Drawing.Size(68, 18)
        Me.Lbl得意先.TabIndex = 22
        Me.Lbl得意先.Text = "【得意先】"
        '
        'txtDate
        '
        Me.txtDate.ContentAlignment = System.Drawing.ContentAlignment.MiddleCenter
        DateLiteralField1.Text = "/"
        DateLiteralField2.Text = "/"
        Me.txtDate.Fields.AddRange(New GrapeCity.Win.Editors.Fields.DateField() {DateYearField1, DateLiteralField1, DateMonthField1, DateLiteralField2, DateDayField1})
        Me.txtDate.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtDate.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.txtDate.Location = New System.Drawing.Point(91, 70)
        Me.txtDate.MaxDate = New Date(2999, 12, 31, 23, 59, 59, 0)
        Me.txtDate.MaxValue = New Date(2999, 12, 31, 23, 59, 59, 0)
        Me.txtDate.MinDate = New Date(2000, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MinValue = New Date(2000, 1, 1, 0, 0, 0, 0)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.SideButtons.AddRange(New GrapeCity.Win.Editors.SideButtonBase() {Me.DropDownButton6})
        Me.txtDate.Size = New System.Drawing.Size(131, 21)
        Me.txtDate.TabIndex = 5
        Me.txtDate.Tag = ""
        Me.txtDate.Value = Nothing
        '
        'DropDownButton6
        '
        Me.DropDownButton6.Name = "DropDownButton6"
        '
        'Lbl対象日
        '
        Me.Lbl対象日.AutoSize = True
        Me.Lbl対象日.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl対象日.Location = New System.Drawing.Point(3, 71)
        Me.Lbl対象日.Name = "Lbl対象日"
        Me.Lbl対象日.Size = New System.Drawing.Size(68, 18)
        Me.Lbl対象日.TabIndex = 5
        Me.Lbl対象日.Text = "【対象日】"
        '
        'Btクリア
        '
        Me.Btクリア.BackColor = System.Drawing.SystemColors.Control
        Me.Btクリア.Location = New System.Drawing.Point(128, 536)
        Me.Btクリア.Name = "Btクリア"
        Me.Btクリア.Size = New System.Drawing.Size(94, 25)
        Me.Btクリア.TabIndex = 102
        Me.Btクリア.Text = "クリア"
        Me.Btクリア.UseVisualStyleBackColor = False
        '
        'Bt検索
        '
        Me.Bt検索.BackColor = System.Drawing.SystemColors.Control
        Me.Bt検索.Location = New System.Drawing.Point(12, 536)
        Me.Bt検索.Name = "Bt検索"
        Me.Bt検索.Size = New System.Drawing.Size(94, 25)
        Me.Bt検索.TabIndex = 101
        Me.Bt検索.Text = "検索"
        Me.Bt検索.UseVisualStyleBackColor = False
        '
        'Lbl需要先
        '
        Me.Lbl需要先.AutoSize = True
        Me.Lbl需要先.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Lbl需要先.Location = New System.Drawing.Point(3, 43)
        Me.Lbl需要先.Name = "Lbl需要先"
        Me.Lbl需要先.Size = New System.Drawing.Size(68, 18)
        Me.Lbl需要先.TabIndex = 0
        Me.Lbl需要先.Text = "【需要先】"
        '
        'Flc汎用
        '
        Me.Flc汎用.Controls.Add(Me.Container汎用)
        Me.Flc汎用.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Flc汎用.Location = New System.Drawing.Point(5, 5)
        Me.Flc汎用.Name = "Flc汎用"
        Me.Flc汎用.Size = New System.Drawing.Size(998, 62)
        Me.Flc汎用.TabIndex = 1
        '
        'Container汎用
        '
        Me.Container汎用.Controls.Add(Me.LblTitle)
        Me.Container汎用.Controls.Add(Me.Cmb事業所)
        Me.Container汎用.Controls.Add(Me.Cmb汎用)
        Me.Container汎用.Location = New System.Drawing.Point(3, 3)
        Me.Container汎用.Name = "Container汎用"
        Me.Container汎用.Size = New System.Drawing.Size(794, 59)
        Me.Container汎用.TabIndex = 0
        '
        'LblTitle
        '
        Me.LblTitle.AutoSize = True
        Me.LblTitle.Font = New System.Drawing.Font("メイリオ", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblTitle.Location = New System.Drawing.Point(6, 12)
        Me.LblTitle.Name = "LblTitle"
        Me.LblTitle.Size = New System.Drawing.Size(175, 41)
        Me.LblTitle.TabIndex = 2
        Me.LblTitle.Text = "DAYTONA2"
        '
        'Cmb事業所
        '
        Me.Cmb事業所.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cmb事業所.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Cmb事業所.FormattingEnabled = True
        Me.Cmb事業所.Location = New System.Drawing.Point(187, 19)
        Me.Cmb事業所.Name = "Cmb事業所"
        Me.Cmb事業所.Size = New System.Drawing.Size(151, 26)
        Me.Cmb事業所.TabIndex = 1
        Me.Cmb事業所.Tag = "ID1"
        '
        'Cmb汎用
        '
        Me.Cmb汎用.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cmb汎用.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Cmb汎用.FormattingEnabled = True
        Me.Cmb汎用.Location = New System.Drawing.Point(345, 19)
        Me.Cmb汎用.Name = "Cmb汎用"
        Me.Cmb汎用.Size = New System.Drawing.Size(429, 26)
        Me.Cmb汎用.TabIndex = 2
        Me.Cmb汎用.Tag = "ID2"
        '
        'Dgv
        '
        Me.Dgv.AllowUserToAddRows = False
        Me.Dgv.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        Me.Dgv.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Dgv.Location = New System.Drawing.Point(253, 73)
        Me.Dgv.Name = "Dgv"
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.LightCyan
        Me.Dgv.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.Dgv.RowTemplate.Height = 21
        Me.Dgv.Size = New System.Drawing.Size(750, 626)
        Me.Dgv.TabIndex = 0
        '
        'VBReport
        '
        Me.VBReport.TemporaryPath = Nothing
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(95, 215)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(126, 20)
        Me.ComboBox1.TabIndex = 104
        '
        'HARK001
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.Controls.Add(Me.TLayoutContainer)
        Me.Controls.Add(Me.SSBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1024, 768)
        Me.Name = "HARK001"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Title"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SSBar.ResumeLayout(False)
        Me.SSBar.PerformLayout()
        CType(Me.TLayoutContainer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TLayoutContainer.ResumeLayout(False)
        Me.Flc検索.ResumeLayout(False)
        Me.Container検索.ResumeLayout(False)
        Me.Container検索.PerformLayout()
        CType(Me.txt需要先, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txt得意先, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txt商品コード, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txt相手先品番, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Flc汎用.ResumeLayout(False)
        Me.Container汎用.ResumeLayout(False)
        Me.Container汎用.PerformLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents SSBar As StatusStrip
    Private WithEvents TLayoutContainer As GrapeCity.Win.Containers.GcTableLayoutContainer
    Private WithEvents TableColumn1 As GrapeCity.Win.Containers.TableColumn
    Private WithEvents TableColumn2 As GrapeCity.Win.Containers.TableColumn
    Private WithEvents Dgv As DataGridView
    Private WithEvents TableRow1 As GrapeCity.Win.Containers.TableRow
    Private WithEvents TableRow2 As GrapeCity.Win.Containers.TableRow
    Private WithEvents Flc汎用 As GrapeCity.Win.Containers.GcFlowLayoutContainer
    Private WithEvents Container汎用 As GrapeCity.Win.Containers.GcContainer
    Private WithEvents Cmb汎用 As ComboBox
    Private WithEvents Flc検索 As GrapeCity.Win.Containers.GcFlowLayoutContainer
    Private WithEvents Container検索 As GrapeCity.Win.Containers.GcContainer
    Private WithEvents Lbl需要先 As Label
    Private WithEvents Btクリア As GrapeCity.Win.Buttons.GcButton
    Private WithEvents Bt検索 As GrapeCity.Win.Buttons.GcButton
    Private WithEvents Lbl対象日 As Label
    Private WithEvents txtDate As GrapeCity.Win.Editors.GcDate
    Private WithEvents DropDownButton6 As GrapeCity.Win.Editors.DropDownButton
    Private WithEvents TSSVersion As ToolStripStatusLabel
    Private WithEvents Cmb事業所 As ComboBox
    Private WithEvents TSSRowCount As ToolStripStatusLabel
    Private WithEvents LblTitle As Label
    Private WithEvents Lbl得意先 As Label
    Private WithEvents Cmb得意先 As ComboBox
    Private WithEvents Cmb需要先 As ComboBox
    Private WithEvents Lbl注意１ As Label
    Private WithEvents Lbl注意２ As Label
    Private WithEvents Lbl相手先品番 As Label
    Private WithEvents Lbl商品 As Label
    Private WithEvents txt商品コード As GrapeCity.Win.Editors.GcTextBox
    Private WithEvents txt相手先品番 As GrapeCity.Win.Editors.GcTextBox
    Private WithEvents txt需要先 As GrapeCity.Win.Editors.GcTextBox
    Private WithEvents Lbl需要先コード指定 As Label
    Private WithEvents txt得意先 As GrapeCity.Win.Editors.GcTextBox
    Private WithEvents Lbl得意先コード指定 As Label
    Private WithEvents Bt印刷 As GrapeCity.Win.Buttons.GcButton
    Private WithEvents VBReport As AdvanceSoftware.VBReport8.CellReport

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private xxxstrProgram_ID As String
    Private xxxint事業所コード As Integer
    Private xxxlng得意先コード As Long
    Private xxxlng需要先コード As Long
    Private Viewer As HARK990 = Nothing

    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        'SelectedValueChangedイベントハンドラの追加
        AddHandler Cmb事業所.SelectedValueChanged, AddressOf HeaderCmb_SelectedValueChanged
        AddHandler Cmb汎用.SelectedValueChanged, AddressOf HeaderCmb_SelectedValueChanged

        'SelectedValueChangedイベントハンドラの追加
        AddHandler Cmb得意先.SelectedValueChanged, AddressOf Cmb_SelectedValueChanged
        AddHandler Cmb需要先.SelectedValueChanged, AddressOf Cmb_SelectedValueChanged

        '
        AddHandler Cmb得意先.KeyDown, AddressOf Txt_KeyDown
        AddHandler Cmb事業所.KeyDown, AddressOf Txt_KeyDown
        AddHandler Cmb汎用.KeyDown, AddressOf Txt_KeyDown
        AddHandler Cmb需要先.KeyDown, AddressOf Txt_KeyDown
        AddHandler txtDate.KeyDown, AddressOf Txt_KeyDown
        AddHandler txt商品コード.KeyDown, AddressOf Txt_KeyDown
        AddHandler txt相手先品番.KeyDown, AddressOf Txt_KeyDown
        AddHandler txt得意先.KeyDown, AddressOf Txt_KeyDown
        AddHandler txt需要先.KeyDown, AddressOf Txt_KeyDown


        UpdateManager.Instance.UpdateSource = New Sources.SimpleWebSource(My.Settings.WebSource, My.Settings.WebProxy)
        UpdateManager.Instance.ReinstateIfRestarted()

    End Sub

    Friend WithEvents ComboBox1 As ComboBox


End Class
