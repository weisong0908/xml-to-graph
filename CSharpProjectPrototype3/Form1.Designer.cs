namespace CSharpProjectPrototype3
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonImportXML = new System.Windows.Forms.Button();
            this.buttonExportTrendReport = new System.Windows.Forms.Button();
            this.buttonGenerateGraph = new System.Windows.Forms.Button();
            this.xmlPreviewer = new System.Windows.Forms.WebBrowser();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.comboBoxXMLType = new System.Windows.Forms.ComboBox();
            this.groupBoxImportXMLFile = new System.Windows.Forms.GroupBox();
            this.textBoxXmlFileName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBoxExportToTrendReport = new System.Windows.Forms.GroupBox();
            this.groupBoxForMultipleDwgOnly = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxOrder10 = new System.Windows.Forms.TextBox();
            this.textBoxOrder9 = new System.Windows.Forms.TextBox();
            this.textBoxOrder8 = new System.Windows.Forms.TextBox();
            this.textBoxOrder7 = new System.Windows.Forms.TextBox();
            this.textBoxOrder6 = new System.Windows.Forms.TextBox();
            this.textBoxOrder5 = new System.Windows.Forms.TextBox();
            this.textBoxOrder4 = new System.Windows.Forms.TextBox();
            this.textBoxOrder3 = new System.Windows.Forms.TextBox();
            this.textBoxOrder2 = new System.Windows.Forms.TextBox();
            this.textBoxOrder1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxGenerateGraphs = new System.Windows.Forms.GroupBox();
            this.groupBoxSaveIn = new System.Windows.Forms.GroupBox();
            this.buttonBrowseFolder = new System.Windows.Forms.Button();
            this.textBoxSavePath = new System.Windows.Forms.TextBox();
            this.checkBoxSaveInOtherPlace = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxFileName = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBoxImportXMLFile.SuspendLayout();
            this.groupBoxExportToTrendReport.SuspendLayout();
            this.groupBoxForMultipleDwgOnly.SuspendLayout();
            this.groupBoxGenerateGraphs.SuspendLayout();
            this.groupBoxSaveIn.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonImportXML
            // 
            this.buttonImportXML.Location = new System.Drawing.Point(58, 19);
            this.buttonImportXML.Name = "buttonImportXML";
            this.buttonImportXML.Size = new System.Drawing.Size(200, 50);
            this.buttonImportXML.TabIndex = 0;
            this.buttonImportXML.Text = "Import XML file";
            this.buttonImportXML.UseVisualStyleBackColor = true;
            this.buttonImportXML.Click += new System.EventHandler(this.buttonImportXML_Click);
            // 
            // buttonExportTrendReport
            // 
            this.buttonExportTrendReport.Location = new System.Drawing.Point(54, 513);
            this.buttonExportTrendReport.Name = "buttonExportTrendReport";
            this.buttonExportTrendReport.Size = new System.Drawing.Size(200, 50);
            this.buttonExportTrendReport.TabIndex = 0;
            this.buttonExportTrendReport.Text = "Export to Trend Report";
            this.buttonExportTrendReport.UseVisualStyleBackColor = true;
            this.buttonExportTrendReport.Click += new System.EventHandler(this.buttonExportTrendReport_Click);
            // 
            // buttonGenerateGraph
            // 
            this.buttonGenerateGraph.Location = new System.Drawing.Point(66, 513);
            this.buttonGenerateGraph.Name = "buttonGenerateGraph";
            this.buttonGenerateGraph.Size = new System.Drawing.Size(200, 50);
            this.buttonGenerateGraph.TabIndex = 0;
            this.buttonGenerateGraph.Text = "Generate Graphs";
            this.buttonGenerateGraph.UseVisualStyleBackColor = true;
            this.buttonGenerateGraph.Click += new System.EventHandler(this.buttonGenerateGraph_Click);
            // 
            // xmlPreviewer
            // 
            this.xmlPreviewer.Location = new System.Drawing.Point(6, 122);
            this.xmlPreviewer.MinimumSize = new System.Drawing.Size(20, 20);
            this.xmlPreviewer.Name = "xmlPreviewer";
            this.xmlPreviewer.Size = new System.Drawing.Size(305, 441);
            this.xmlPreviewer.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "Open XML file";
            this.openFileDialog1.Filter = "XML files|*.xml";
            this.openFileDialog1.Title = "Open XML file";
            // 
            // comboBoxXMLType
            // 
            this.comboBoxXMLType.FormattingEnabled = true;
            this.comboBoxXMLType.Items.AddRange(new object[] {
            "Single and Multiple",
            "Windows XP 32 bit",
            "Windows 7 64 bit"});
            this.comboBoxXMLType.Location = new System.Drawing.Point(92, 32);
            this.comboBoxXMLType.Name = "comboBoxXMLType";
            this.comboBoxXMLType.Size = new System.Drawing.Size(121, 21);
            this.comboBoxXMLType.TabIndex = 6;
            this.comboBoxXMLType.SelectedIndexChanged += new System.EventHandler(this.comboBoxXMLType_SelectedIndexChanged);
            // 
            // groupBoxImportXMLFile
            // 
            this.groupBoxImportXMLFile.Controls.Add(this.textBoxXmlFileName);
            this.groupBoxImportXMLFile.Controls.Add(this.label3);
            this.groupBoxImportXMLFile.Controls.Add(this.xmlPreviewer);
            this.groupBoxImportXMLFile.Controls.Add(this.buttonImportXML);
            this.groupBoxImportXMLFile.Location = new System.Drawing.Point(12, 12);
            this.groupBoxImportXMLFile.Name = "groupBoxImportXMLFile";
            this.groupBoxImportXMLFile.Size = new System.Drawing.Size(317, 569);
            this.groupBoxImportXMLFile.TabIndex = 7;
            this.groupBoxImportXMLFile.TabStop = false;
            this.groupBoxImportXMLFile.Text = "1. Import XML file";
            // 
            // textBoxXmlFileName
            // 
            this.textBoxXmlFileName.Location = new System.Drawing.Point(89, 80);
            this.textBoxXmlFileName.Name = "textBoxXmlFileName";
            this.textBoxXmlFileName.ReadOnly = true;
            this.textBoxXmlFileName.Size = new System.Drawing.Size(221, 20);
            this.textBoxXmlFileName.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "XML file name:";
            // 
            // groupBoxExportToTrendReport
            // 
            this.groupBoxExportToTrendReport.Controls.Add(this.groupBoxForMultipleDwgOnly);
            this.groupBoxExportToTrendReport.Controls.Add(this.label1);
            this.groupBoxExportToTrendReport.Controls.Add(this.comboBoxXMLType);
            this.groupBoxExportToTrendReport.Controls.Add(this.buttonExportTrendReport);
            this.groupBoxExportToTrendReport.Enabled = false;
            this.groupBoxExportToTrendReport.Location = new System.Drawing.Point(335, 12);
            this.groupBoxExportToTrendReport.Name = "groupBoxExportToTrendReport";
            this.groupBoxExportToTrendReport.Size = new System.Drawing.Size(308, 569);
            this.groupBoxExportToTrendReport.TabIndex = 8;
            this.groupBoxExportToTrendReport.TabStop = false;
            this.groupBoxExportToTrendReport.Text = "2. Export to TrendReport";
            // 
            // groupBoxForMultipleDwgOnly
            // 
            this.groupBoxForMultipleDwgOnly.AutoSize = true;
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.label5);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.label4);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder10);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder9);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder8);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder7);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder6);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder5);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder4);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder3);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder2);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBoxOrder1);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.label2);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox10);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox9);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox8);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox7);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox6);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox5);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox4);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox3);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox2);
            this.groupBoxForMultipleDwgOnly.Controls.Add(this.textBox1);
            this.groupBoxForMultipleDwgOnly.Enabled = false;
            this.groupBoxForMultipleDwgOnly.Location = new System.Drawing.Point(17, 59);
            this.groupBoxForMultipleDwgOnly.Name = "groupBoxForMultipleDwgOnly";
            this.groupBoxForMultipleDwgOnly.Size = new System.Drawing.Size(285, 398);
            this.groupBoxForMultipleDwgOnly.TabIndex = 9;
            this.groupBoxForMultipleDwgOnly.TabStop = false;
            this.groupBoxForMultipleDwgOnly.Text = "For multiple drawings only";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(95, 100);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "Order in xml:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(176, 100);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Number of DWGs:";
            // 
            // textBoxOrder10
            // 
            this.textBoxOrder10.Location = new System.Drawing.Point(114, 359);
            this.textBoxOrder10.Name = "textBoxOrder10";
            this.textBoxOrder10.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder10.TabIndex = 19;
            // 
            // textBoxOrder9
            // 
            this.textBoxOrder9.Location = new System.Drawing.Point(114, 332);
            this.textBoxOrder9.Name = "textBoxOrder9";
            this.textBoxOrder9.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder9.TabIndex = 19;
            // 
            // textBoxOrder8
            // 
            this.textBoxOrder8.Location = new System.Drawing.Point(114, 305);
            this.textBoxOrder8.Name = "textBoxOrder8";
            this.textBoxOrder8.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder8.TabIndex = 19;
            // 
            // textBoxOrder7
            // 
            this.textBoxOrder7.Location = new System.Drawing.Point(114, 278);
            this.textBoxOrder7.Name = "textBoxOrder7";
            this.textBoxOrder7.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder7.TabIndex = 19;
            // 
            // textBoxOrder6
            // 
            this.textBoxOrder6.Location = new System.Drawing.Point(114, 252);
            this.textBoxOrder6.Name = "textBoxOrder6";
            this.textBoxOrder6.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder6.TabIndex = 19;
            // 
            // textBoxOrder5
            // 
            this.textBoxOrder5.Location = new System.Drawing.Point(114, 224);
            this.textBoxOrder5.Name = "textBoxOrder5";
            this.textBoxOrder5.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder5.TabIndex = 19;
            // 
            // textBoxOrder4
            // 
            this.textBoxOrder4.Location = new System.Drawing.Point(114, 197);
            this.textBoxOrder4.Name = "textBoxOrder4";
            this.textBoxOrder4.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder4.TabIndex = 19;
            // 
            // textBoxOrder3
            // 
            this.textBoxOrder3.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textBoxOrder3.Location = new System.Drawing.Point(114, 170);
            this.textBoxOrder3.Name = "textBoxOrder3";
            this.textBoxOrder3.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder3.TabIndex = 19;
            // 
            // textBoxOrder2
            // 
            this.textBoxOrder2.Location = new System.Drawing.Point(114, 142);
            this.textBoxOrder2.Name = "textBoxOrder2";
            this.textBoxOrder2.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder2.TabIndex = 19;
            // 
            // textBoxOrder1
            // 
            this.textBoxOrder1.Enabled = false;
            this.textBoxOrder1.Location = new System.Drawing.Point(114, 116);
            this.textBoxOrder1.Name = "textBoxOrder1";
            this.textBoxOrder1.Size = new System.Drawing.Size(22, 20);
            this.textBoxOrder1.TabIndex = 19;
            this.textBoxOrder1.Text = "1";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(21, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(246, 51);
            this.label2.TabIndex = 18;
            this.label2.Text = "Please type in the number of drawings in textboxes on your right and order of the" +
                " drawing in xml file in textboxes on yout left";
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(179, 359);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(100, 20);
            this.textBox10.TabIndex = 17;
            this.textBox10.Text = "No value";
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(179, 332);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(100, 20);
            this.textBox9.TabIndex = 16;
            this.textBox9.Text = "No value";
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(179, 305);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(100, 20);
            this.textBox8.TabIndex = 15;
            this.textBox8.Text = "No value";
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(179, 278);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(100, 20);
            this.textBox7.TabIndex = 14;
            this.textBox7.Text = "No value";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(179, 251);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(100, 20);
            this.textBox6.TabIndex = 13;
            this.textBox6.Text = "500";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(179, 224);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(100, 20);
            this.textBox5.TabIndex = 12;
            this.textBox5.Text = "5000";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(179, 197);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(100, 20);
            this.textBox4.TabIndex = 11;
            this.textBox4.Text = "10";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(179, 170);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(100, 20);
            this.textBox3.TabIndex = 10;
            this.textBox3.Text = "100";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(179, 142);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(100, 20);
            this.textBox2.TabIndex = 9;
            this.textBox2.Text = "1000";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(179, 116);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(100, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.Text = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Type of xml file:";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "Select Excel Workbook";
            this.openFileDialog2.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";
            this.openFileDialog2.Title = "Select Excel Workbook";
            // 
            // groupBoxGenerateGraphs
            // 
            this.groupBoxGenerateGraphs.Controls.Add(this.groupBoxSaveIn);
            this.groupBoxGenerateGraphs.Controls.Add(this.checkBoxSaveInOtherPlace);
            this.groupBoxGenerateGraphs.Controls.Add(this.label7);
            this.groupBoxGenerateGraphs.Controls.Add(this.label6);
            this.groupBoxGenerateGraphs.Controls.Add(this.textBoxFileName);
            this.groupBoxGenerateGraphs.Controls.Add(this.buttonGenerateGraph);
            this.groupBoxGenerateGraphs.Location = new System.Drawing.Point(649, 12);
            this.groupBoxGenerateGraphs.Name = "groupBoxGenerateGraphs";
            this.groupBoxGenerateGraphs.Size = new System.Drawing.Size(333, 569);
            this.groupBoxGenerateGraphs.TabIndex = 9;
            this.groupBoxGenerateGraphs.TabStop = false;
            this.groupBoxGenerateGraphs.Text = "3. Generate graphs";
            // 
            // groupBoxSaveIn
            // 
            this.groupBoxSaveIn.Controls.Add(this.buttonBrowseFolder);
            this.groupBoxSaveIn.Controls.Add(this.textBoxSavePath);
            this.groupBoxSaveIn.Enabled = false;
            this.groupBoxSaveIn.Location = new System.Drawing.Point(66, 168);
            this.groupBoxSaveIn.Name = "groupBoxSaveIn";
            this.groupBoxSaveIn.Size = new System.Drawing.Size(230, 100);
            this.groupBoxSaveIn.TabIndex = 7;
            this.groupBoxSaveIn.TabStop = false;
            this.groupBoxSaveIn.Text = "Save in ...";
            // 
            // buttonBrowseFolder
            // 
            this.buttonBrowseFolder.Location = new System.Drawing.Point(135, 45);
            this.buttonBrowseFolder.Name = "buttonBrowseFolder";
            this.buttonBrowseFolder.Size = new System.Drawing.Size(89, 23);
            this.buttonBrowseFolder.TabIndex = 7;
            this.buttonBrowseFolder.Text = "Browse Folder";
            this.buttonBrowseFolder.UseVisualStyleBackColor = true;
            this.buttonBrowseFolder.Click += new System.EventHandler(this.buttonBrowseFolder_Click);
            // 
            // textBoxSavePath
            // 
            this.textBoxSavePath.Location = new System.Drawing.Point(6, 19);
            this.textBoxSavePath.Name = "textBoxSavePath";
            this.textBoxSavePath.ReadOnly = true;
            this.textBoxSavePath.Size = new System.Drawing.Size(218, 20);
            this.textBoxSavePath.TabIndex = 6;
            // 
            // checkBoxSaveInOtherPlace
            // 
            this.checkBoxSaveInOtherPlace.AutoSize = true;
            this.checkBoxSaveInOtherPlace.Location = new System.Drawing.Point(66, 145);
            this.checkBoxSaveInOtherPlace.Name = "checkBoxSaveInOtherPlace";
            this.checkBoxSaveInOtherPlace.Size = new System.Drawing.Size(118, 17);
            this.checkBoxSaveInOtherPlace.TabIndex = 4;
            this.checkBoxSaveInOtherPlace.Text = "Save in other place";
            this.checkBoxSaveInOtherPlace.UseVisualStyleBackColor = true;
            this.checkBoxSaveInOtherPlace.CheckedChanged += new System.EventHandler(this.checkBoxSaveInOtherPlace_CheckedChanged);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(66, 98);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(230, 17);
            this.label7.TabIndex = 3;
            this.label7.Text = "By default this file is saved in \"My Documents\".";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(40, 55);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(193, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Enter a new file name for Trend Report:";
            // 
            // textBoxFileName
            // 
            this.textBoxFileName.Location = new System.Drawing.Point(66, 75);
            this.textBoxFileName.Name = "textBoxFileName";
            this.textBoxFileName.Size = new System.Drawing.Size(230, 20);
            this.textBoxFileName.TabIndex = 1;
            this.textBoxFileName.Text = "EvoPerfTestTrendReport_b";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(994, 593);
            this.Controls.Add(this.groupBoxGenerateGraphs);
            this.Controls.Add(this.groupBoxExportToTrendReport);
            this.Controls.Add(this.groupBoxImportXMLFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "XML2Graph (C# Project Prototype 3)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBoxImportXMLFile.ResumeLayout(false);
            this.groupBoxImportXMLFile.PerformLayout();
            this.groupBoxExportToTrendReport.ResumeLayout(false);
            this.groupBoxExportToTrendReport.PerformLayout();
            this.groupBoxForMultipleDwgOnly.ResumeLayout(false);
            this.groupBoxForMultipleDwgOnly.PerformLayout();
            this.groupBoxGenerateGraphs.ResumeLayout(false);
            this.groupBoxGenerateGraphs.PerformLayout();
            this.groupBoxSaveIn.ResumeLayout(false);
            this.groupBoxSaveIn.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonImportXML;
        private System.Windows.Forms.Button buttonExportTrendReport;
        private System.Windows.Forms.Button buttonGenerateGraph;
        private System.Windows.Forms.WebBrowser xmlPreviewer;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ComboBox comboBoxXMLType;
        private System.Windows.Forms.GroupBox groupBoxImportXMLFile;
        private System.Windows.Forms.GroupBox groupBoxExportToTrendReport;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBoxForMultipleDwgOnly;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.GroupBox groupBoxGenerateGraphs;
        private System.Windows.Forms.TextBox textBoxOrder10;
        private System.Windows.Forms.TextBox textBoxOrder9;
        private System.Windows.Forms.TextBox textBoxOrder8;
        private System.Windows.Forms.TextBox textBoxOrder7;
        private System.Windows.Forms.TextBox textBoxOrder6;
        private System.Windows.Forms.TextBox textBoxOrder5;
        private System.Windows.Forms.TextBox textBoxOrder4;
        private System.Windows.Forms.TextBox textBoxOrder3;
        private System.Windows.Forms.TextBox textBoxOrder2;
        private System.Windows.Forms.TextBox textBoxOrder1;
        private System.Windows.Forms.TextBox textBoxXmlFileName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxFileName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBoxSaveIn;
        private System.Windows.Forms.Button buttonBrowseFolder;
        private System.Windows.Forms.TextBox textBoxSavePath;
        private System.Windows.Forms.CheckBox checkBoxSaveInOtherPlace;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;

    }
}

