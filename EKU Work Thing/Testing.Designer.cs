namespace EKU_Report_Reader
{
    partial class Testing
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
            this.invClear = new System.Windows.Forms.Button();
            this.invDelete = new System.Windows.Forms.Button();
            this.invClose = new System.Windows.Forms.Button();
            this.invcolDGV = new System.Windows.Forms.DataGridView();
            this.tBuilding = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tRoom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tView = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tEdit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tSelected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // invClear
            // 
            this.invClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.invClear.Location = new System.Drawing.Point(12, 327);
            this.invClear.Name = "invClear";
            this.invClear.Size = new System.Drawing.Size(55, 23);
            this.invClear.TabIndex = 8;
            this.invClear.Text = "Clear";
            this.invClear.UseVisualStyleBackColor = true;
            // 
            // invDelete
            // 
            this.invDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.invDelete.Location = new System.Drawing.Point(73, 327);
            this.invDelete.Name = "invDelete";
            this.invDelete.Size = new System.Drawing.Size(55, 23);
            this.invDelete.TabIndex = 7;
            this.invDelete.Text = "Delete";
            this.invDelete.UseVisualStyleBackColor = true;
            // 
            // invClose
            // 
            this.invClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.invClose.Location = new System.Drawing.Point(393, 327);
            this.invClose.Name = "invClose";
            this.invClose.Size = new System.Drawing.Size(55, 23);
            this.invClose.TabIndex = 6;
            this.invClose.Text = "Close";
            this.invClose.UseVisualStyleBackColor = true;
            // 
            // invcolDGV
            // 
            this.invcolDGV.AllowUserToAddRows = false;
            this.invcolDGV.AllowUserToDeleteRows = false;
            this.invcolDGV.AllowUserToResizeColumns = false;
            this.invcolDGV.AllowUserToResizeRows = false;
            this.invcolDGV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.invcolDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.invcolDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.tBuilding,
            this.tRoom,
            this.tView,
            this.tEdit,
            this.tSelected});
            this.invcolDGV.Location = new System.Drawing.Point(13, 12);
            this.invcolDGV.Name = "invcolDGV";
            this.invcolDGV.Size = new System.Drawing.Size(435, 309);
            this.invcolDGV.TabIndex = 5;
            // 
            // tBuilding
            // 
            this.tBuilding.HeaderText = "Building";
            this.tBuilding.MinimumWidth = 175;
            this.tBuilding.Name = "tBuilding";
            this.tBuilding.ReadOnly = true;
            this.tBuilding.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.tBuilding.Width = 175;
            // 
            // tRoom
            // 
            this.tRoom.HeaderText = "Room";
            this.tRoom.Name = "tRoom";
            this.tRoom.ReadOnly = true;
            this.tRoom.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // tView
            // 
            this.tView.HeaderText = "View";
            this.tView.MinimumWidth = 35;
            this.tView.Name = "tView";
            this.tView.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.tView.Text = "View";
            this.tView.Width = 35;
            // 
            // tEdit
            // 
            this.tEdit.HeaderText = "Edit";
            this.tEdit.Name = "tEdit";
            this.tEdit.Width = 40;
            // 
            // tSelected
            // 
            this.tSelected.HeaderText = "Select";
            this.tSelected.MinimumWidth = 40;
            this.tSelected.Name = "tSelected";
            this.tSelected.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.tSelected.Width = 40;
            // 
            // Testing
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 362);
            this.Controls.Add(this.invClear);
            this.Controls.Add(this.invDelete);
            this.Controls.Add(this.invClose);
            this.Controls.Add(this.invcolDGV);
            this.Name = "Testing";
            this.Text = "Testing Data Collected";
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button invClear;
        private System.Windows.Forms.Button invDelete;
        private System.Windows.Forms.Button invClose;
        private System.Windows.Forms.DataGridView invcolDGV;
        private System.Windows.Forms.DataGridViewTextBoxColumn tBuilding;
        private System.Windows.Forms.DataGridViewTextBoxColumn tRoom;
        private System.Windows.Forms.DataGridViewButtonColumn tView;
        private System.Windows.Forms.DataGridViewButtonColumn tEdit;
        private System.Windows.Forms.DataGridViewCheckBoxColumn tSelected;
    }
}