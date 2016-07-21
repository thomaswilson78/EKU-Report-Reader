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
            this.invcolSelected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.invClose = new System.Windows.Forms.Button();
            this.invDelete = new System.Windows.Forms.Button();
            this.invClear = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).BeginInit();
            this.SuspendLayout();
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
            this.invcolBuilding,
            this.invcolRoom,
            this.invcolView,
            this.invcolSelected});
            this.invcolDGV.Location = new System.Drawing.Point(12, 12);
            this.invcolDGV.Name = "invcolDGV";
            this.invcolDGV.Size = new System.Drawing.Size(393, 309);
            this.invcolDGV.TabIndex = 0;
            this.invcolDGV.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.invcolDGV_CellContentClick);
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
            this.invcolView.MinimumWidth = 35;
            this.invcolView.Name = "invcolView";
            this.invcolView.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.invcolView.Text = "View";
            this.invcolView.Width = 35;
            // 
            // invcolSelected
            // 
            this.invcolSelected.HeaderText = "Select";
            this.invcolSelected.MinimumWidth = 40;
            this.invcolSelected.Name = "invcolSelected";
            this.invcolSelected.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.invcolSelected.Width = 40;
            // 
            // invClose
            // 
            this.invClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.invClose.Location = new System.Drawing.Point(350, 327);
            this.invClose.Name = "invClose";
            this.invClose.Size = new System.Drawing.Size(55, 23);
            this.invClose.TabIndex = 1;
            this.invClose.Text = "Close";
            this.invClose.UseVisualStyleBackColor = true;
            this.invClose.Click += new System.EventHandler(this.invClose_Click);
            // 
            // invDelete
            // 
            this.invDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.invDelete.Location = new System.Drawing.Point(73, 327);
            this.invDelete.Name = "invDelete";
            this.invDelete.Size = new System.Drawing.Size(55, 23);
            this.invDelete.TabIndex = 2;
            this.invDelete.Text = "Delete";
            this.invDelete.UseVisualStyleBackColor = true;
            this.invDelete.Click += new System.EventHandler(this.invDelete_Click);
            // 
            // invClear
            // 
            this.invClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.invClear.Location = new System.Drawing.Point(12, 327);
            this.invClear.Name = "invClear";
            this.invClear.Size = new System.Drawing.Size(55, 23);
            this.invClear.TabIndex = 4;
            this.invClear.Text = "Clear";
            this.invClear.UseVisualStyleBackColor = true;
            this.invClear.Click += new System.EventHandler(this.invClear_Click);
            // 
            // Form2
            // 
            this.AcceptButton = this.invClose;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(417, 362);
            this.Controls.Add(this.invClear);
            this.Controls.Add(this.invDelete);
            this.Controls.Add(this.invClose);
            this.Controls.Add(this.invcolDGV);
            this.MaximumSize = new System.Drawing.Size(433, 720);
            this.MinimumSize = new System.Drawing.Size(433, 400);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Inventory Collected";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.invcolDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView invcolDGV;
        private System.Windows.Forms.Button invClose;
        private System.Windows.Forms.Button invDelete;
        private System.Windows.Forms.Button invClear;
        private System.Windows.Forms.DataGridViewTextBoxColumn invcolBuilding;
        private System.Windows.Forms.DataGridViewTextBoxColumn invcolRoom;
        private System.Windows.Forms.DataGridViewButtonColumn invcolView;
        private System.Windows.Forms.DataGridViewCheckBoxColumn invcolSelected;
    }
}