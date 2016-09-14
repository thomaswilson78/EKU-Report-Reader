namespace EKU_Work_Thing
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
            this.tClear = new System.Windows.Forms.Button();
            this.tDelete = new System.Windows.Forms.Button();
            this.tClose = new System.Windows.Forms.Button();
            this.tDeleteAll = new System.Windows.Forms.Button();
            this.tDGV = new System.Windows.Forms.DataGridView();
            this.tBuilding = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tRoom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tExport = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tEdit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tSelect = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.tDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // tClear
            // 
            this.tClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tClear.Location = new System.Drawing.Point(12, 327);
            this.tClear.Name = "tClear";
            this.tClear.Size = new System.Drawing.Size(55, 23);
            this.tClear.TabIndex = 8;
            this.tClear.Text = "Clear";
            this.tClear.UseVisualStyleBackColor = true;
            this.tClear.Click += new System.EventHandler(this.tClear_Click);
            // 
            // tDelete
            // 
            this.tDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tDelete.Location = new System.Drawing.Point(73, 327);
            this.tDelete.Name = "tDelete";
            this.tDelete.Size = new System.Drawing.Size(55, 23);
            this.tDelete.TabIndex = 7;
            this.tDelete.Text = "Delete";
            this.tDelete.UseVisualStyleBackColor = true;
            this.tDelete.Click += new System.EventHandler(this.tDelete_Click);
            // 
            // tClose
            // 
            this.tClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.tClose.Location = new System.Drawing.Point(393, 327);
            this.tClose.Name = "tClose";
            this.tClose.Size = new System.Drawing.Size(55, 23);
            this.tClose.TabIndex = 6;
            this.tClose.Text = "Close";
            this.tClose.UseVisualStyleBackColor = true;
            this.tClose.Click += new System.EventHandler(this.tClose_Click);
            // 
            // tDeleteAll
            // 
            this.tDeleteAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tDeleteAll.Location = new System.Drawing.Point(134, 327);
            this.tDeleteAll.Name = "tDeleteAll";
            this.tDeleteAll.Size = new System.Drawing.Size(66, 23);
            this.tDeleteAll.TabIndex = 9;
            this.tDeleteAll.Text = "Delete All";
            this.tDeleteAll.UseVisualStyleBackColor = true;
            this.tDeleteAll.Click += new System.EventHandler(this.tDeleteAll_Click);
            // 
            // tDGV
            // 
            this.tDGV.AllowUserToAddRows = false;
            this.tDGV.AllowUserToDeleteRows = false;
            this.tDGV.AllowUserToResizeColumns = false;
            this.tDGV.AllowUserToResizeRows = false;
            this.tDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.tBuilding,
            this.tRoom,
            this.tExport,
            this.tEdit,
            this.tSelect});
            this.tDGV.Location = new System.Drawing.Point(12, 12);
            this.tDGV.Name = "tDGV";
            this.tDGV.Size = new System.Drawing.Size(436, 309);
            this.tDGV.TabIndex = 10;
            this.tDGV.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tDGV_CellContentClick);
            // 
            // tBuilding
            // 
            this.tBuilding.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.tBuilding.HeaderText = "Building";
            this.tBuilding.MinimumWidth = 170;
            this.tBuilding.Name = "tBuilding";
            this.tBuilding.ReadOnly = true;
            this.tBuilding.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.tBuilding.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.tBuilding.Width = 170;
            // 
            // tRoom
            // 
            this.tRoom.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.tRoom.HeaderText = "Room";
            this.tRoom.MinimumWidth = 75;
            this.tRoom.Name = "tRoom";
            this.tRoom.ReadOnly = true;
            this.tRoom.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.tRoom.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.tRoom.Width = 75;
            // 
            // tExport
            // 
            this.tExport.HeaderText = "Export";
            this.tExport.MinimumWidth = 50;
            this.tExport.Name = "tExport";
            this.tExport.Width = 50;
            // 
            // tEdit
            // 
            this.tEdit.HeaderText = "Edit";
            this.tEdit.MinimumWidth = 50;
            this.tEdit.Name = "tEdit";
            this.tEdit.Width = 50;
            // 
            // tSelect
            // 
            this.tSelect.HeaderText = "Select";
            this.tSelect.MinimumWidth = 50;
            this.tSelect.Name = "tSelect";
            this.tSelect.Width = 50;
            // 
            // Testing
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 362);
            this.Controls.Add(this.tDGV);
            this.Controls.Add(this.tDeleteAll);
            this.Controls.Add(this.tClear);
            this.Controls.Add(this.tDelete);
            this.Controls.Add(this.tClose);
            this.Name = "Testing";
            this.Text = "Testing Data Collected";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Testing_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.tDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button tClear;
        private System.Windows.Forms.Button tDelete;
        private System.Windows.Forms.Button tClose;
        private System.Windows.Forms.Button tDeleteAll;
        private System.Windows.Forms.DataGridView tDGV;
        private System.Windows.Forms.DataGridViewTextBoxColumn tBuilding;
        private System.Windows.Forms.DataGridViewTextBoxColumn tRoom;
        private System.Windows.Forms.DataGridViewButtonColumn tExport;
        private System.Windows.Forms.DataGridViewButtonColumn tEdit;
        private System.Windows.Forms.DataGridViewCheckBoxColumn tSelect;
    }
}