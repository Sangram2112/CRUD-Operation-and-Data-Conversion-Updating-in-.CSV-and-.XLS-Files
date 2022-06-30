namespace Fifa
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtPlayerName = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Email = new System.Windows.Forms.Label();
            this.txtPlayerClub = new System.Windows.Forms.TextBox();
            this.txtCountry = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtJersyNo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btCsv = new System.Windows.Forms.Button();
            this.btExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Black;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(40, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "Player Name";
            // 
            // txtPlayerName
            // 
            this.txtPlayerName.Location = new System.Drawing.Point(172, 89);
            this.txtPlayerName.Name = "txtPlayerName";
            this.txtPlayerName.Size = new System.Drawing.Size(125, 20);
            this.txtPlayerName.TabIndex = 1;
            this.txtPlayerName.TextChanged += new System.EventHandler(this.txtUserName_TextChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaption;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.ActiveCaption;
            this.dataGridView1.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView1.Location = new System.Drawing.Point(148, 278);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(531, 150);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            // 
            // Email
            // 
            this.Email.AutoSize = true;
            this.Email.BackColor = System.Drawing.Color.Black;
            this.Email.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Email.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Email.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Email.Location = new System.Drawing.Point(67, 140);
            this.Email.Name = "Email";
            this.Email.Size = new System.Drawing.Size(40, 18);
            this.Email.TabIndex = 3;
            this.Email.Text = "Club";
            // 
            // txtPlayerClub
            // 
            this.txtPlayerClub.Location = new System.Drawing.Point(172, 138);
            this.txtPlayerClub.Name = "txtPlayerClub";
            this.txtPlayerClub.Size = new System.Drawing.Size(125, 20);
            this.txtPlayerClub.TabIndex = 4;
            // 
            // txtCountry
            // 
            this.txtCountry.Location = new System.Drawing.Point(172, 182);
            this.txtCountry.Name = "txtCountry";
            this.txtCountry.Size = new System.Drawing.Size(125, 20);
            this.txtCountry.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label2.Location = new System.Drawing.Point(56, 185);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 18);
            this.label2.TabIndex = 5;
            this.label2.Text = "Country";
            // 
            // txtJersyNo
            // 
            this.txtJersyNo.Location = new System.Drawing.Point(172, 229);
            this.txtJersyNo.Name = "txtJersyNo";
            this.txtJersyNo.Size = new System.Drawing.Size(125, 20);
            this.txtJersyNo.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Black;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label3.Location = new System.Drawing.Point(53, 232);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(69, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "Jersy No";
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnSave.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnSave.FlatAppearance.BorderSize = 2;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(351, 89);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(99, 33);
            this.btnSave.TabIndex = 9;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdate.Location = new System.Drawing.Point(351, 156);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(99, 33);
            this.btnUpdate.TabIndex = 10;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.Location = new System.Drawing.Point(350, 222);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(99, 32);
            this.btnDelete.TabIndex = 11;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Fifa.Properties.Resources.ec82a51e2d701406118f5d010b7fa0ab;
            this.pictureBox1.Location = new System.Drawing.Point(208, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(389, 63);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 15;
            this.pictureBox1.TabStop = false;
            // 
            // btCsv
            // 
            this.btCsv.BackgroundImage = global::Fifa.Properties.Resources.icons8_export_csv_96;
            this.btCsv.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btCsv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btCsv.Image = ((System.Drawing.Image)(resources.GetObject("btCsv.Image")));
            this.btCsv.Location = new System.Drawing.Point(517, 185);
            this.btCsv.Name = "btCsv";
            this.btCsv.Size = new System.Drawing.Size(115, 41);
            this.btCsv.TabIndex = 14;
            this.btCsv.Text = "Export To CSV";
            this.btCsv.UseVisualStyleBackColor = true;
            this.btCsv.Click += new System.EventHandler(this.btCsv_Click);
            // 
            // btExcel
            // 
            this.btExcel.BackgroundImage = global::Fifa.Properties.Resources.icons8_microsoft_excel_96;
            this.btExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btExcel.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.btExcel.FlatAppearance.BorderSize = 4;
            this.btExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btExcel.Image = global::Fifa.Properties.Resources.icons8_microsoft_excel_2019_96;
            this.btExcel.Location = new System.Drawing.Point(517, 114);
            this.btExcel.Name = "btExcel";
            this.btExcel.Size = new System.Drawing.Size(115, 44);
            this.btExcel.TabIndex = 12;
            this.btExcel.Text = "Export To Excel";
            this.btExcel.UseVisualStyleBackColor = true;
            this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gray;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btCsv);
            this.Controls.Add(this.btExcel);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtJersyNo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtCountry);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPlayerClub);
            this.Controls.Add(this.Email);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.txtPlayerName);
            this.Controls.Add(this.label1);
            this.DoubleBuffered = true;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = " ";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPlayerName;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label Email;
        private System.Windows.Forms.TextBox txtPlayerClub;
        private System.Windows.Forms.TextBox txtCountry;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtJersyNo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btExcel;
        private System.Windows.Forms.Button btCsv;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

