namespace Transformación_de_Datos_de_Fichada
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtFile = new System.Windows.Forms.TextBox();
            this.btnChooseFile = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.btnChooseDirOutput = new System.Windows.Forms.Button();
            this.btnProcessFile = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(38, 49);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(455, 20);
            this.txtFile.TabIndex = 0;
            // 
            // btnChooseFile
            // 
            this.btnChooseFile.Location = new System.Drawing.Point(499, 46);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Size = new System.Drawing.Size(180, 23);
            this.btnChooseFile.TabIndex = 1;
            this.btnChooseFile.Text = "Seleccionar Excel de sumario";
            this.btnChooseFile.UseVisualStyleBackColor = true;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(38, 105);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(455, 20);
            this.textBox2.TabIndex = 2;
            // 
            // btnChooseDirOutput
            // 
            this.btnChooseDirOutput.Location = new System.Drawing.Point(499, 102);
            this.btnChooseDirOutput.Name = "btnChooseDirOutput";
            this.btnChooseDirOutput.Size = new System.Drawing.Size(180, 23);
            this.btnChooseDirOutput.TabIndex = 3;
            this.btnChooseDirOutput.Text = "Elegir carpeta de salida";
            this.btnChooseDirOutput.UseVisualStyleBackColor = true;
            this.btnChooseDirOutput.Click += new System.EventHandler(this.btnChooseDirOutput_Click);
            // 
            // btnProcessFile
            // 
            this.btnProcessFile.Location = new System.Drawing.Point(173, 145);
            this.btnProcessFile.Name = "btnProcessFile";
            this.btnProcessFile.Size = new System.Drawing.Size(320, 44);
            this.btnProcessFile.TabIndex = 4;
            this.btnProcessFile.Text = "Procesar";
            this.btnProcessFile.UseVisualStyleBackColor = true;
            this.btnProcessFile.Click += new System.EventHandler(this.btnProcessFile_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(38, 195);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(641, 159);
            this.dataGridView1.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 395);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnProcessFile);
            this.Controls.Add(this.btnChooseDirOutput);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btnChooseFile);
            this.Controls.Add(this.txtFile);
            this.Name = "Form1";
            this.Text = "Procesar Fichadas";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btnChooseDirOutput;
        private System.Windows.Forms.Button btnProcessFile;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

