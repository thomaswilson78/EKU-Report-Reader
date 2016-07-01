namespace EKU_Work_Thing
{
    partial class Form3
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
            this.BuildingCB = new System.Windows.Forms.ComboBox();
            this.RoomCB = new System.Windows.Forms.ComboBox();
            this.pullDataBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BuildingCB
            // 
            this.BuildingCB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.BuildingCB.FormattingEnabled = true;
            this.BuildingCB.Items.AddRange(new object[] {
            "Alumni Coliseum",
            "Ashland Building",
            "Begley Building",
            "Burrier Building",
            "Cammack Building",
            "Campbell Building",
            "Carter Building",
            "Case Annex",
            "Coates Administration Building",
            "Combs Classroom",
            "Crabbe Library",
            "Dizney Building",
            "Foster Music Building",
            "Gentry Building",
            "Keith Building",
            "McCreary Building",
            "Memorial Science",
            "Moberly Building",
            "Moore Building",
            "New Science Building",
            "Perkins Building",
            "Powell Building",
            "Roark Building",
            "Rowlett Building",
            "Stratton Building",
            "University Building",
            "Wallace Building",
            "Weaver Health",
            "Whalin Complex",
            "Whitlock Building"});
            this.BuildingCB.Location = new System.Drawing.Point(12, 12);
            this.BuildingCB.Name = "BuildingCB";
            this.BuildingCB.Size = new System.Drawing.Size(260, 21);
            this.BuildingCB.TabIndex = 0;
            this.BuildingCB.SelectedIndexChanged += new System.EventHandler(this.BuildingCB_SelectedIndexChanged);
            // 
            // RoomCB
            // 
            this.RoomCB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.RoomCB.FormattingEnabled = true;
            this.RoomCB.Location = new System.Drawing.Point(278, 12);
            this.RoomCB.Name = "RoomCB";
            this.RoomCB.Size = new System.Drawing.Size(121, 21);
            this.RoomCB.TabIndex = 1;
            // 
            // pullDataBtn
            // 
            this.pullDataBtn.Location = new System.Drawing.Point(324, 59);
            this.pullDataBtn.Name = "pullDataBtn";
            this.pullDataBtn.Size = new System.Drawing.Size(75, 23);
            this.pullDataBtn.TabIndex = 2;
            this.pullDataBtn.Text = "Pull";
            this.pullDataBtn.UseVisualStyleBackColor = true;
            this.pullDataBtn.Click += new System.EventHandler(this.pullDataBtn_Click);
            // 
            // Form3
            // 
            this.AcceptButton = this.pullDataBtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(415, 94);
            this.Controls.Add(this.pullDataBtn);
            this.Controls.Add(this.RoomCB);
            this.Controls.Add(this.BuildingCB);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Form3";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Pull From...";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox BuildingCB;
        private System.Windows.Forms.ComboBox RoomCB;
        private System.Windows.Forms.Button pullDataBtn;
    }
}