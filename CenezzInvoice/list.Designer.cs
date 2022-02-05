namespace CenezzInvoice
{
    partial class list
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(list));
            this.lister = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.finit = new System.Windows.Forms.DateTimePicker();
            this.init = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button9 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.nume = new System.Windows.Forms.TextBox();
            this.numc = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.selecto = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.idinvo = new System.Windows.Forms.TextBox();
            this.iplcont = new System.Windows.Forms.DataGridView();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lefte = new System.Windows.Forms.TextBox();
            this.supo = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label10 = new System.Windows.Forms.Label();
            this.ejer = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.numee = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.palletkgs = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.logpesos = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.label22 = new System.Windows.Forms.Label();
            this.deces = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.lister)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iplcont)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.deces)).BeginInit();
            this.SuspendLayout();
            // 
            // lister
            // 
            this.lister.AllowUserToAddRows = false;
            this.lister.AllowUserToDeleteRows = false;
            this.lister.AllowUserToResizeRows = false;
            this.lister.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.lister.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.lister.Location = new System.Drawing.Point(12, 112);
            this.lister.MultiSelect = false;
            this.lister.Name = "lister";
            this.lister.ReadOnly = true;
            this.lister.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.lister.RowHeadersVisible = false;
            this.lister.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.lister.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.lister.Size = new System.Drawing.Size(809, 256);
            this.lister.TabIndex = 0;
            this.lister.TabStop = false;
            this.lister.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.lister_CellClick);
            this.lister.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.lister_CellDoubleClick);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.Image = global::CenezzInvoice.Properties.Resources.mail_find;
            this.button1.Location = new System.Drawing.Point(384, 14);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(40, 36);
            this.button1.TabIndex = 1;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // finit
            // 
            this.finit.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.finit.Location = new System.Drawing.Point(297, 28);
            this.finit.Name = "finit";
            this.finit.Size = new System.Drawing.Size(80, 20);
            this.finit.TabIndex = 2;
            // 
            // init
            // 
            this.init.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.init.Location = new System.Drawing.Point(208, 28);
            this.init.Name = "init";
            this.init.Size = new System.Drawing.Size(81, 20);
            this.init.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(207, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Inicial:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(297, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Final:";
            // 
            // button9
            // 
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.Location = new System.Drawing.Point(457, 41);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(63, 67);
            this.button9.TabIndex = 254;
            this.button9.Text = "Excel";
            this.button9.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(17, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(334, 23);
            this.label3.TabIndex = 255;
            this.label3.Text = "Examinar Invoices y Packing Lists";
            // 
            // nume
            // 
            this.nume.Location = new System.Drawing.Point(148, 28);
            this.nume.Name = "nume";
            this.nume.Size = new System.Drawing.Size(54, 20);
            this.nume.TabIndex = 256;
            this.nume.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // numc
            // 
            this.numc.Location = new System.Drawing.Point(57, 28);
            this.numc.Name = "numc";
            this.numc.Size = new System.Drawing.Size(41, 20);
            this.numc.TabIndex = 257;
            this.numc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 258;
            this.label4.Text = "# cliente";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(99, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 13);
            this.label5.TabIndex = 259;
            this.label5.Text = "# emisor";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.numc);
            this.groupBox1.Controls.Add(this.nume);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.init);
            this.groupBox1.Controls.Add(this.finit);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(12, 47);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(439, 56);
            this.groupBox1.TabIndex = 260;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Herramientas de busqueda";
            // 
            // selecto
            // 
            this.selecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.selecto.Location = new System.Drawing.Point(809, 51);
            this.selecto.Name = "selecto";
            this.selecto.ReadOnly = true;
            this.selecto.Size = new System.Drawing.Size(141, 29);
            this.selecto.TabIndex = 260;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(730, 59);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 13);
            this.label6.TabIndex = 260;
            this.label6.Text = "Folio activo:";
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Location = new System.Drawing.Point(583, 41);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(63, 67);
            this.button2.TabIndex = 261;
            this.button2.Text = "Cancelar";
            this.button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // idinvo
            // 
            this.idinvo.Location = new System.Drawing.Point(1214, 88);
            this.idinvo.Name = "idinvo";
            this.idinvo.Size = new System.Drawing.Size(50, 20);
            this.idinvo.TabIndex = 262;
            this.idinvo.Visible = false;
            // 
            // iplcont
            // 
            this.iplcont.AllowUserToAddRows = false;
            this.iplcont.AllowUserToDeleteRows = false;
            this.iplcont.AllowUserToResizeRows = false;
            this.iplcont.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.iplcont.Location = new System.Drawing.Point(833, 112);
            this.iplcont.MultiSelect = false;
            this.iplcont.Name = "iplcont";
            this.iplcont.ReadOnly = true;
            this.iplcont.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.iplcont.RowHeadersVisible = false;
            this.iplcont.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.iplcont.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.iplcont.Size = new System.Drawing.Size(456, 256);
            this.iplcont.TabIndex = 0;
            this.iplcont.TabStop = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(965, 84);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(219, 23);
            this.label7.TabIndex = 264;
            this.label7.Text = "Contenido Invoice/PL";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 20);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 13);
            this.label8.TabIndex = 265;
            this.label8.Text = "Margen izq.:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(154, 20);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(69, 13);
            this.label9.TabIndex = 266;
            this.label9.Text = "Margen sup.:";
            // 
            // lefte
            // 
            this.lefte.Location = new System.Drawing.Point(72, 16);
            this.lefte.Name = "lefte";
            this.lefte.Size = new System.Drawing.Size(75, 20);
            this.lefte.TabIndex = 267;
            this.lefte.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress_1);
            // 
            // supo
            // 
            this.supo.Location = new System.Drawing.Point(227, 16);
            this.supo.Name = "supo";
            this.supo.Size = new System.Drawing.Size(75, 20);
            this.supo.TabIndex = 268;
            this.supo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress_1);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(310, 15);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(114, 23);
            this.button3.TabIndex = 269;
            this.button3.Text = "Guardar márgenes";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.supo);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.lefte);
            this.groupBox2.Location = new System.Drawing.Point(12, 374);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(439, 49);
            this.groupBox2.TabIndex = 270;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Configuración de margenes (en Milimetros)";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(728, 90);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(99, 13);
            this.label10.TabIndex = 271;
            this.label10.Text = "Ejercicio actual:";
            // 
            // ejer
            // 
            this.ejer.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ejer.Location = new System.Drawing.Point(833, 82);
            this.ejer.Name = "ejer";
            this.ejer.ReadOnly = true;
            this.ejer.Size = new System.Drawing.Size(117, 29);
            this.ejer.TabIndex = 272;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(752, 28);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(54, 13);
            this.label11.TabIndex = 273;
            this.label11.Text = "Número:";
            // 
            // numee
            // 
            this.numee.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numee.Location = new System.Drawing.Point(809, 20);
            this.numee.Name = "numee";
            this.numee.ReadOnly = true;
            this.numee.Size = new System.Drawing.Size(141, 29);
            this.numee.TabIndex = 274;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.palletkgs);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Location = new System.Drawing.Point(457, 374);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(205, 49);
            this.groupBox3.TabIndex = 271;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Configuración de pesos de Taras (kgs)";
            // 
            // palletkgs
            // 
            this.palletkgs.Location = new System.Drawing.Point(125, 16);
            this.palletkgs.Name = "palletkgs";
            this.palletkgs.ReadOnly = true;
            this.palletkgs.Size = new System.Drawing.Size(74, 20);
            this.palletkgs.TabIndex = 268;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 20);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(113, 13);
            this.label13.TabIndex = 266;
            this.label13.Text = "Tarima con embalajes:";
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Location = new System.Drawing.Point(520, 41);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(63, 67);
            this.button4.TabIndex = 275;
            this.button4.Text = "Factura";
            this.button4.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // logpesos
            // 
            this.logpesos.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logpesos.Location = new System.Drawing.Point(689, 383);
            this.logpesos.Multiline = true;
            this.logpesos.Name = "logpesos";
            this.logpesos.ReadOnly = true;
            this.logpesos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.logpesos.Size = new System.Drawing.Size(600, 53);
            this.logpesos.TabIndex = 276;
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Location = new System.Drawing.Point(1214, 15);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(63, 67);
            this.button5.TabIndex = 277;
            this.button5.Text = "PackingL";
            this.button5.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Visible = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Location = new System.Drawing.Point(646, 41);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(63, 67);
            this.button6.TabIndex = 278;
            this.button6.Text = "Resumen";
            this.button6.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.Location = new System.Drawing.Point(453, 9);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(205, 20);
            this.label22.TabIndex = 312;
            this.label22.Text = "Decimales para imprimir:";
            this.label22.Visible = false;
            // 
            // deces
            // 
            this.deces.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.deces.Location = new System.Drawing.Point(662, 3);
            this.deces.Maximum = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.deces.Minimum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.deces.Name = "deces";
            this.deces.Size = new System.Drawing.Size(44, 31);
            this.deces.TabIndex = 311;
            this.deces.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.deces.Visible = false;
            this.deces.ValueChanged += new System.EventHandler(this.deces_ValueChanged);
            // 
            // list
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1298, 446);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.deces);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.logpesos);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.numee);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.ejer);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.iplcont);
            this.Controls.Add(this.idinvo);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.selecto);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lister);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "list";
            this.Text = "Listado de Invoices";
            this.Deactivate += new System.EventHandler(this.list_Deactivate);
            this.Load += new System.EventHandler(this.list_Load);
            this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.list_KeyUp);
            ((System.ComponentModel.ISupportInitialize)(this.lister)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iplcont)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.deces)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView lister;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DateTimePicker finit;
        private System.Windows.Forms.DateTimePicker init;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox nume;
        private System.Windows.Forms.TextBox numc;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox selecto;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox idinvo;
        private System.Windows.Forms.DataGridView iplcont;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox lefte;
        private System.Windows.Forms.TextBox supo;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox ejer;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox numee;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox palletkgs;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox logpesos;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.NumericUpDown deces;
    }
}