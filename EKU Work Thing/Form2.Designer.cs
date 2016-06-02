namespace EKU_Work_Thing
{
    partial class Form2
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
            this.invcolDGV = new System.Windows.Forms.DataGridView();
            this.invcolBuilding = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.invcolRoom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.invcolView = new System.Windows.Forms.DataGridViewButtonColumn();
            this.invClose = new System.Windows.Forms.Button();
            this.invDelete = new System.Windows.Forms.Button();
            this.invEdit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // invcolDGV
            // 
            this.invcolDGV.AllowUserToAddRows = false;
            this.invcolDGV.AllowUserToDeleteRows = false;
            this.invcolDGV.AllowUserToResizeColumns = false;
            this.invcolDGV.AllowUserToResizeRows = false;
            this.invcolDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.invcolDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.invcolBuilding,
            this.invcolRoom,
            this.invcolView});
            this.invcolDGV.Location = new System.Drawing.Point(12, 12);
            this.invcolDGV.Name = "invcolDGV";
            this.invcolDGV.Size = new System.Drawing.Size(420, 309);
            this.invcolDGV.TabIndex = 0;
            this.invcolDGV.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // invcolBuilding
            // 
            this.invcolBuilding.HeaderText = "Building";
            this.invcolBuilding.MinimumWidth = 175;
            this.invcolBuilding.Name = "invcolBuilding";
            this.invcolBuilding.ReadOnly = true;
            this.invcolBuilding.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.invcolBuilding.Width = 175;
            // 
            // invcolRoom
            // 
            this.invcolRoom.HeaderText = "Room";
            this.invcolRoom.Name = "invcolRoom";
            this.invcolRoom.ReadOnly = true;
            this.invcolRoom.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // invcolView
            // 
            this.invcolView.HeaderText = "View";
            this.invcolView.Name = "invcolView";
            this.invcolView.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.invcolView.Text = "View";
            // 
            // invClose
            // 
            this.invClose.Location = new System.Drawing.Point(357, 327);
            this.invClose.Name = "invClose";
            this.invClose.Size = new System.Drawing.Size(75, 23);
            this.invClose.TabIndex = 1;
            this.invClose.Text = "Close";
            this.invClose.UseVisualStyleBackColor = true;
            this.invClose.Click += new System.EventHandler(this.invClose_Click);
            // 
            // invDelete
            // 
            this.invDelete.Location = new System.Drawing.Point(276, 327);
            this.invDelete.Name = "invDelete";
            this.invDelete.Size = new System.Drawing.Size(75, 23);
            this.invDelete.TabIndex = 2;
            this.invDelete.Text = "Delete";
            this.invDelete.UseVisualStyleBackColor = true;
            // 
            // invEdit
            // 
            this.invEdit.Location = new System.Drawing.Point(195, 327);
            this.invEdit.Name = "invEdit";
            this.invEdit.Size = new System.Drawing.Size(75, 23);
            this.invEdit.TabIndex = 3;
            this.invEdit.Text = "Edit";
            this.invEdit.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AcceptButton = this.invClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 362);
            this.Controls.Add(this.invEdit);
            this.Controls.Add(this.invDelete);
            this.Controls.Add(this.invClose);
            this.Controls.Add(this.invcolDGV);
            this.Name = "Form2";
            this.Text = "Inventory Collected";
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView invcolDGV;
        private System.Windows.Forms.Button invClose;
        private System.Windows.Forms.Button invDelete;
        private System.Windows.Forms.Button invEdit;
        private System.Windows.Forms.DataGridViewTextBoxColumn invcolBuilding;
        private System.Windows.Forms.DataGridViewTextBoxColumn invcolRoom;
        private System.Windows.Forms.DataGridViewButtonColumn invcolView;
    }
}