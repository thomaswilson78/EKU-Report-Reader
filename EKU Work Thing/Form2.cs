using System;
using System.Windows.Forms;
using System.Data.SQLite;

namespace EKU_Work_Thing
{
    public partial class Form2 : Form
    {
        Form1 f1 = new Form1();
        //initialization
        public Form2(Form1 parent)
        {
            f1 = parent;
            InitializeComponent();
            loadTable();
        }
        //close form window
        private void invClose_Click(object sender, EventArgs e)
        {
            Hide();
        }//done
        //load data into table
        public void loadTable()
        {
            invcolDGV.Rows.Clear();
            //load data into table
            SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;");
            try
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand("SELECT Building,Room FROM inventory_collected;", conn);
                SQLiteDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    invcolDGV.Rows.Add(reader["Building"].ToString(), reader["Room"].ToString(), "Edit");
                }
                cmd.Dispose();
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        //load information into inventory tab
        private void invcolDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex==2)
            {
                DataGridView getDGVInfo = (DataGridView)sender;
                int i = e.RowIndex;
                string b = getDGVInfo.Rows[i].Cells[0].Value.ToString();
                string r = getDGVInfo.Rows[i].Cells[1].Value.ToString();
                SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;");
                try
                {
                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM inventory_collected WHERE Building = @b AND Room = @r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    SQLiteDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        f1.addBuildingComBox.SelectedItem = b;
                        f1.addRoomTB.Text = r;
                        f1.addContComBox.SelectedItem = reader["Controller"].ToString();
                        f1.addAudioComBox.SelectedItem = reader["Audio"].ToString();
                        f1.addDockCB.Checked = reader["Dock"].ToString().Equals("1");
                        f1.addDCCB.Checked = reader["Doc_Cam"].ToString().Equals("1");
                        f1.addCamCB.Checked = reader["Camera"].ToString().Equals("1");
                        f1.addMicCB.Checked = reader["Mic"].ToString().Equals("1");
                        f1.addBRCB.Checked = reader["Bluray"].ToString().Equals("1");
                        f1.addDVDCB.Checked = reader["DVD"].ToString().Equals("1");
                        f1.addHDMICB.Checked = reader["HDMI_Pull"].ToString().Equals("1");
                        f1.addVGACB.Checked = reader["VGA_Pull"].ToString().Equals("1");
                        f1.addAVCB.Checked = reader["AV_Panel"].ToString().Equals("1");
                        f1.addSolCB.Checked = reader["Solstice"].ToString().Equals("1");
                        f1.addMMTB1.Text = reader["D1MakeModel"].ToString();
                        f1.addSerialTB1.Text = reader["D1Serial"].ToString();
                        f1.addScrTB1.Text = reader["D1Screen"].ToString();
                        f1.addIPTB1.Text = reader["D1IP"].ToString();
                        f1.addMACTB1.Text = reader["D1MAC"].ToString();
                        f1.addBulbTB1.Text = reader["D1Bulb"].ToString();
                        f1.addMMTB2.Text = reader["D2MakeModel"].ToString();
                        f1.addSerialTB2.Text = reader["D2Serial"].ToString();
                        f1.addScrTB2.Text = reader["D2Screen"].ToString();
                        f1.addIPTB2.Text = reader["D2IP"].ToString();
                        f1.addMACTB2.Text = reader["D2MAC"].ToString();
                        f1.addBulbTB2.Text = reader["D2Bulb"].ToString();
                        f1.addMMTB3.Text = reader["D3MakeModel"].ToString();
                        f1.addSerialTB3.Text = reader["D3Serial"].ToString();
                        f1.addScrTB3.Text = reader["D3Screen"].ToString();
                        f1.addIPTB3.Text = reader["D3IP"].ToString();
                        f1.addMACTB3.Text = reader["D3MAC"].ToString();
                        f1.addBulbTB3.Text = reader["D3Bulb"].ToString();
                        f1.addMMTB4.Text = reader["D4MakeModel"].ToString();
                        f1.addSerialTB4.Text = reader["D4Serial"].ToString();
                        f1.addScrTB4.Text = reader["D4Screen"].ToString();
                        f1.addIPTB4.Text = reader["D4IP"].ToString();
                        f1.addMACTB4.Text = reader["D4MAC"].ToString();
                        f1.addBulbTB4.Text = reader["D4Bulb"].ToString();
                        DateTime t;
                        DateTime.TryParse(reader["Filter"].ToString(), out t);
                        if (!t.ToString("MM/dd/yyyy").Equals("01/01/0001"))
                            f1.addFilter.Text = t.ToString("MM/dd/yyyy");
                        else
                            f1.addFilter.Text = "";
                        f1.addPCModTB.Text = reader["PCModel"].ToString();
                        f1.addPCSerialTB.Text = reader["PCSerial"].ToString();
                        f1.addNUCIPTB.Text = reader["NUCIP"].ToString();
                        f1.addNUCMACTB.Text = reader["NUCMAC"].ToString();
                        f1.addCatVidTB.Text = reader["Cat6Video"].ToString();
                        f1.addNetTB.Text = reader["NetworkPorts"].ToString();
                        DateTime.TryParse(reader["SolsticeDate"].ToString(), out t);
                        if (!t.ToString("MM/dd/yyyy").Equals("01/01/0001"))
                            f1.addSolActTB.Text = t.ToString("MM/dd/yyyy");
                        else
                            f1.addSolActTB.Text = "";
                        f1.addSolLicTB.Text = reader["SolsticeLicense"].ToString();
                        f1.addDscrptTB.Text = reader["Notes"].ToString();
                        
                        f1.Refresh();
                    }
                    reader.Close();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    conn.Close();
                }
            }
        }//unfinished
        //delete data from database and remove from table
        private void invDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to delete these items?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm.ToString().Equals("OK")) {
                bool success = false;
                bool error = false;
                SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;");
                try
                {
                    conn.Open();
                    for (int i = 0; i < invcolDGV.Rows.Count; i++)
                    {
                        if (invcolDGV.Rows[i].Cells[3].Value != null)
                            if (invcolDGV.Rows[i].Cells[3].Value.ToString().Equals("True"))
                            {
                                SQLiteCommand cmd = new SQLiteCommand("DELETE FROM inventory_collected WHERE Building=@B AND Room=@R", conn);
                                cmd.Parameters.AddWithValue("@B", invcolDGV.Rows[i].Cells[0].Value.ToString());
                                cmd.Parameters.AddWithValue("@R", invcolDGV.Rows[i].Cells[1].Value.ToString());
                                cmd.ExecuteNonQuery();
                                invcolDGV.Rows.RemoveAt(i);
                                success = true;
                                i--;
                            }
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    error = true;
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                if (success)
                    MessageBox.Show("Item(s) successfully removed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                else if (error)
                    MessageBox.Show("Error occured when removing items, not all items were deleted.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show("No items selected for deletion.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                invcolDGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
                conn.Close();
            }
        }//Done

        private void invClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < invcolDGV.Rows.Count; i++)
            {
                if (invcolDGV.Rows[i].Cells[3].Value != null)
                    if ((bool)invcolDGV.Rows[i].Cells[3].Value == true)
                        invcolDGV.Rows[i].Cells[3].Value = false;
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }
    }
}