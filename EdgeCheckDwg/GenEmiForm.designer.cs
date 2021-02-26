namespace EdgeCheckDwg
{
    partial class GenEmiForm
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GenEmiForm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pdfCheckbox = new System.Windows.Forms.CheckBox();
            this.labelChooseFolder = new System.Windows.Forms.Label();
            this.chooseFolder = new System.Windows.Forms.Button();
            this.pianiRadio = new System.Windows.Forms.RadioButton();
            this.listaRadio = new System.Windows.Forms.RadioButton();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rO = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.arasExcelCheckbox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.pubblicaCheckbox = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pubblicaCheckbox);
            this.groupBox1.Controls.Add(this.arasExcelCheckbox);
            this.groupBox1.Controls.Add(this.pdfCheckbox);
            this.groupBox1.Controls.Add(this.labelChooseFolder);
            this.groupBox1.Controls.Add(this.chooseFolder);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(341, 115);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // pdfCheckbox
            // 
            this.pdfCheckbox.AutoSize = true;
            this.pdfCheckbox.Location = new System.Drawing.Point(6, 48);
            this.pdfCheckbox.Name = "pdfCheckbox";
            this.pdfCheckbox.Size = new System.Drawing.Size(79, 17);
            this.pdfCheckbox.TabIndex = 4;
            this.pdfCheckbox.Text = "Genera pdf";
            this.pdfCheckbox.UseVisualStyleBackColor = true;
            // 
            // labelChooseFolder
            // 
            this.labelChooseFolder.AutoSize = true;
            this.labelChooseFolder.Location = new System.Drawing.Point(87, 24);
            this.labelChooseFolder.Name = "labelChooseFolder";
            this.labelChooseFolder.Size = new System.Drawing.Size(142, 13);
            this.labelChooseFolder.TabIndex = 3;
            this.labelChooseFolder.Text = "Nessuna cartella selezionata";
            // 
            // chooseFolder
            // 
            this.chooseFolder.Location = new System.Drawing.Point(6, 19);
            this.chooseFolder.Name = "chooseFolder";
            this.chooseFolder.Size = new System.Drawing.Size(75, 23);
            this.chooseFolder.TabIndex = 2;
            this.chooseFolder.Text = "Choose";
            this.chooseFolder.UseVisualStyleBackColor = true;
            this.chooseFolder.Click += new System.EventHandler(this.chooseFolder_Click);
            // 
            // pianiRadio
            // 
            this.pianiRadio.AutoSize = true;
            this.pianiRadio.Location = new System.Drawing.Point(76, 19);
            this.pianiRadio.Name = "pianiRadio";
            this.pianiRadio.Size = new System.Drawing.Size(48, 17);
            this.pianiRadio.TabIndex = 1;
            this.pianiRadio.Text = "Piani";
            this.pianiRadio.UseVisualStyleBackColor = true;
            // 
            // listaRadio
            // 
            this.listaRadio.AutoSize = true;
            this.listaRadio.Checked = true;
            this.listaRadio.Location = new System.Drawing.Point(6, 19);
            this.listaRadio.Name = "listaRadio";
            this.listaRadio.Size = new System.Drawing.Size(47, 17);
            this.listaRadio.TabIndex = 0;
            this.listaRadio.TabStop = true;
            this.listaRadio.Text = "Liste";
            this.listaRadio.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 254);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(335, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Genera";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.start_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.rO);
            this.groupBox2.Controls.Add(this.radioButton2);
            this.groupBox2.Location = new System.Drawing.Point(12, 133);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(341, 50);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Destinazione Distribuzione";
            // 
            // rO
            // 
            this.rO.AutoSize = true;
            this.rO.Location = new System.Drawing.Point(76, 18);
            this.rO.Name = "rO";
            this.rO.Size = new System.Drawing.Size(61, 17);
            this.rO.TabIndex = 5;
            this.rO.Text = "Officina";
            this.rO.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Checked = true;
            this.radioButton2.Location = new System.Drawing.Point(6, 19);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(64, 17);
            this.radioButton2.TabIndex = 4;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Cantiere";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 283);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(365, 23);
            this.progressBar1.TabIndex = 7;
            // 
            // arasExcelCheckbox
            // 
            this.arasExcelCheckbox.AutoSize = true;
            this.arasExcelCheckbox.Location = new System.Drawing.Point(6, 71);
            this.arasExcelCheckbox.Name = "arasExcelCheckbox";
            this.arasExcelCheckbox.Size = new System.Drawing.Size(140, 17);
            this.arasExcelCheckbox.TabIndex = 5;
            this.arasExcelCheckbox.Text = "Genera Excel Emissione";
            this.arasExcelCheckbox.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(168, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Destinazione Distribuzione";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(168, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Tipo Distribuzione";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.listaRadio);
            this.groupBox3.Controls.Add(this.pianiRadio);
            this.groupBox3.Location = new System.Drawing.Point(12, 189);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(341, 45);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Divisione Distribuzione";
            // 
            // pubblicaCheckbox
            // 
            this.pubblicaCheckbox.AutoSize = true;
            this.pubblicaCheckbox.Location = new System.Drawing.Point(6, 93);
            this.pubblicaCheckbox.Name = "pubblicaCheckbox";
            this.pubblicaCheckbox.Size = new System.Drawing.Size(92, 17);
            this.pubblicaCheckbox.TabIndex = 6;
            this.pubblicaCheckbox.Text = "Pubblica Dwg";
            this.pubblicaCheckbox.UseVisualStyleBackColor = true;
            // 
            // GenEmiForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 306);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "GenEmiForm";
            this.Text = "Genera Emissione";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton pianiRadio;
        private System.Windows.Forms.RadioButton listaRadio;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rO;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label labelChooseFolder;
        private System.Windows.Forms.Button chooseFolder;
        public System.Windows.Forms.CheckBox pdfCheckbox;
        public System.Windows.Forms.CheckBox arasExcelCheckbox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        public System.Windows.Forms.CheckBox pubblicaCheckbox;
    }
}

