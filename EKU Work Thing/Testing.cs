using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
//have to reference, now that I'm making the title more official (that or completely redo the project).
//PLANS:
//Datagrid view will be filled with information collected from "testingmain" (database table) : NOT DONE
//Have an edit button that will fill in all fields on the testing tab in the main form. : NOT DONE
//Have an export button that will export data into a testing sheet, needs the testing sheet in the Template folder : NOT DONE
//Delete button will delete any selected records : NOT DONE : NOTE BELOW IMPORTANT
/*
    Prefacing this right now: MAKE SURE TO INCLUDE "foreign keys=true" IN THE CONNECTION STRING!!!!
    SQLITE IS DUMB AND HAS THEM OFF BY DEFAULT, MAKING DELETION WAY MORE DIFFICULT THAN IT SHOULD BE!!!!
    This way by having foreign keys set, I only need to delete from "testingmain" and all other records will be deleted.
*/
//Clear will clear any selected items. : NOT DONE
//Possibly add a "Delete All" button that will delete all records in order to start from scratch : NOT DONE: ALSO SHOULD DO IN THE INVENTORY FORM AS WELL
namespace EKU_Work_Thing
{
    public partial class Testing : Form
    {
        Form1 f1 = new Form1();
        public Testing(Form1 parent)
        {
            f1 = parent;
            InitializeComponent();
            loadTable();
            //will initialize table here
        }

        public void loadTable()
        {
            tDGV.Rows.Clear();
            //load data into table
            SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;");
            try
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand("SELECT Building,Room FROM testingmain;", conn);
                SQLiteDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                    tDGV.Rows.Add(reader["Building"].ToString(), reader["Room"].ToString(), "Export", "Edit", false);
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

        private void Testing_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void tDeleteAll_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to delete all items?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm.ToString().Equals("OK"))
            {
                confirm = MessageBox.Show("Are you certain? This action cannot be undone.", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (confirm.ToString().Equals("OK"))
                {
                    bool success = false;
                    SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;foreign keys=true");
                    try
                    {
                        tDGV.Rows.Clear();
                        conn.Open();
                        SQLiteCommand cmd = new SQLiteCommand("DELETE FROM testingmain", conn);
                        cmd.ExecuteNonQuery();
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    if (success)
                        MessageBox.Show("All items successfully removed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else
                        MessageBox.Show("Error occured when removing items.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tDGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    conn.Dispose();
                }

            }
        }

        private void tDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to delete these items?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm.ToString().Equals("OK"))
            {
                bool success = false;
                bool error = false;
                SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;foreign keys=true");
                try
                {
                    conn.Open();
                    for (int i = 0; i < tDGV.Rows.Count; i++)
                    {
                        if (tDGV.Rows[i].Cells[4].Value.Equals(true))
                        {
                            SQLiteCommand cmd = new SQLiteCommand("DELETE FROM testingmain WHERE Building=@B AND Room=@R", conn);
                            cmd.Parameters.AddWithValue("@B", tDGV.Rows[i].Cells[0].Value.ToString());
                            cmd.Parameters.AddWithValue("@R", tDGV.Rows[i].Cells[1].Value.ToString());
                            cmd.ExecuteNonQuery();
                            tDGV.Rows.RemoveAt(i);
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
                tDGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
                conn.Dispose();
            }
        }

        private void tClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tDGV.Rows.Count; i++)
                tDGV.Rows[i].Cells[3].Value = false;
        }

        private void tDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //for exporting
            if(e.ColumnIndex == 2)
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                    MessageBox.Show("Microsoft Office and Excel needs to be installed to export reports.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\", "testing sheet.xlsx");
                    Workbook wb = xlApp.Workbooks.Add(path);
                    Worksheet ws = (Worksheet)wb.Worksheets[1];
                    SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3");
                    try
                    {
                        conn.Open();
                        string b = tDGV.Rows[e.RowIndex].Cells[0].Value.ToString();
                        string r = tDGV.Rows[e.RowIndex].Cells[1].Value.ToString();

                        //main
                        SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM testingmain WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        SQLiteDataReader reader = cmd.ExecuteReader();
                        reader.Read();
                        ws.Cells[36, 1] = "Building: " + b;
                        ws.Cells[36, 6] = r;
                        ws.Cells[38, 1] = "Agent Name: " + reader["Agent"].ToString();
                        ws.Cells[38, 6] = reader["DateCol"].ToString();
                        ws.Cells[67, 1] = reader["Notes"].ToString();
                        
                        //general
                        cmd = new SQLiteCommand("SELECT * FROM testinggeneral WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        string[] data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        int pos = 2;
                        for (int y=5; y<21; y++)
                        {
                            ws.Cells[y, 1 + int.Parse(reader[data[pos]].ToString())] = "X";
                            pos++;
                        }
                        for(int y=5;y<21;y++)
                        {
                            ws.Cells[y, 5] = reader[data[pos]].ToString();
                            pos++;
                        }
                        ws.Range[ws.Cells[5, 2], ws.Cells[20, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[5 ,5], ws.Cells[20, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        //vidaud
                        cmd = new SQLiteCommand("SELECT * FROM testingvideoaudio WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        pos = 2;
                        for (int y = 24; y < 30; y++)
                        {
                            for (int x = 2; x < 8; x++)
                            {
                                ws.Cells[y,x] = reader[data[pos]].ToString();
                                pos++;
                            }
                        }
                        ws.Range[ws.Cells[24, 2], ws.Cells[29, 6]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[24, 7], ws.Cells[29, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        //mic
                        cmd = new SQLiteCommand("SELECT * FROM testingmic WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        pos = 2;
                        for (int y = 33; y < 35; y++)
                        {
                            ws.Cells[y, 1 + int.Parse(reader[data[pos]].ToString())] = "X";
                            pos++;
                        }
                        for (int y = 33; y < 35; y++)
                        {
                            ws.Cells[y, 5] = reader[data[pos]].ToString();
                            pos++;
                        }
                        ws.Range[ws.Cells[33, 2], ws.Cells[34, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[33, 5], ws.Cells[34, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        //doc cam
                        cmd = new SQLiteCommand("SELECT * FROM testingdoccam WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        pos = 2;
                        for (int y = 45; y < 49; y++)
                        {
                            ws.Cells[y, 1 + int.Parse(reader[data[pos]].ToString())] = "X";
                            pos++;
                        }
                        for (int y = 45; y < 49; y++)
                        {
                            ws.Cells[y, 5] = reader[data[pos]].ToString();
                            pos++;
                        }
                        ws.Range[ws.Cells[45, 2], ws.Cells[48, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[45, 5], ws.Cells[48, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        //dvd/blu
                        cmd = new SQLiteCommand("SELECT * FROM testingdvdblu WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        pos = 2;
                        for (int y = 52; y < 56; y++)
                        {
                            ws.Cells[y, 1 + int.Parse(reader[data[pos]].ToString())] = "X";
                            pos++;
                        }
                        for (int y = 52; y < 56; y++)
                        {
                            ws.Cells[y, 5] = reader[data[pos]].ToString();
                            pos++;
                        }
                        ws.Range[ws.Cells[52, 2], ws.Cells[55, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[52, 5], ws.Cells[55, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        //iptv
                        cmd = new SQLiteCommand("SELECT * FROM testingiptv WHERE Building=@b AND Room=@r", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        reader = cmd.ExecuteReader();
                        reader.Read();

                        data = new string[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            data[i] = reader.GetName(i);

                        pos = 2;
                        for (int y = 59; y < 64; y++)
                        {
                            ws.Cells[y, 1 + int.Parse(reader[data[pos]].ToString())] = "X";
                            pos++;
                        }
                        for (int y = 59; y < 64; y++)
                        {
                            ws.Cells[y, 5] = reader[data[pos]].ToString();
                            pos++;
                        }
                        ws.Range[ws.Cells[59, 2], ws.Cells[63, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ws.Range[ws.Cells[59, 5], ws.Cells[63, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    catch (Exception ex)
                    {
                        wb.Close(0);
                        xlApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                        xlApp = null;
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    conn.Close();
                    if(xlApp!=null)
                    {
                        xlApp.WindowState = XlWindowState.xlMaximized;
                        xlApp.Visible = true;
                    }
                }
            }
            //for editing
            if (e.ColumnIndex == 3)
            {
                f1.resetTestingTables();
                DataGridView getDGVInfo = (DataGridView)sender;
                int i = e.RowIndex;
                string b = getDGVInfo.Rows[i].Cells[0].Value.ToString();
                string r = getDGVInfo.Rows[i].Cells[1].Value.ToString();
                SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;");
                try
                {
                    //main info
                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM testingmain WHERE Building = @b AND Room = @r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    SQLiteDataReader reader = cmd.ExecuteReader();
                    reader.Read();
                    f1.testBuilding.Text = reader["Building"].ToString();
                    f1.testRoom.Text = reader["Room"].ToString();
                    f1.testName.Text = reader["Agent"].ToString();
                    f1.testBuilding.Text = reader["Building"].ToString();
                    f1.testNotesTB.Text = reader["Notes"].ToString();
                    reader.Close();

                    //general info
                    //only commenting the first one since every one (except vid/aud) works the same.
                    cmd = new SQLiteCommand("SELECT * FROM testinggeneral WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    string[] data = new string[reader.FieldCount];//column names extracted to string array (way easier than writing them out manually)
                    for (int t = 0; t < reader.FieldCount; t++)//fills in string with column names
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();//reads the data (important, otherwise no data will be collected. also no need for a while or if since there's only one instance of data).
                    for(int x = 0;x<f1.testGeneralDGV.Rows.Count;x++)
                        f1.testGeneralDGV.Rows[x].Cells[int.Parse(reader[data[x+2].ToString()].ToString())].Value = true;//sets value based on number recorded
                    for (int x = 0; x < f1.testGeneralDGV.Rows.Count; x++)//gets notes
                        f1.testGeneralDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                    reader.Close();//closed so reader can be used again

                    //vid/aud info
                    cmd = new SQLiteCommand("SELECT * FROM testingvideoaudio WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    data = new string[reader.FieldCount];
                    for (int t = 0; t < reader.FieldCount; t++)
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();
                    int k = 2;
                    for (int y = 0; y < f1.testVidAudDGV.Rows.Count; y++)
                        for (int x = 1; x < f1.testVidAudDGV.Columns.Count; x++)
                        {
                            f1.testVidAudDGV.Rows[y].Cells[x].Value = reader[k].ToString();
                            k++;
                        }
                    reader.Close();

                    //mic info
                    cmd = new SQLiteCommand("SELECT * FROM testingmic WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    data = new string[reader.FieldCount];
                    for (int t = 0; t < reader.FieldCount; t++)
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();
                    for (int x = 0; x < f1.testMicDGV.Rows.Count; x++)
                        f1.testMicDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                    for (int x = 0; x < f1.testMicDGV.Rows.Count; x++)
                        f1.testMicDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                    reader.Close();

                    //doccam info
                     cmd = new SQLiteCommand("SELECT * FROM testingdoccam WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    data = new string[reader.FieldCount];
                    for (int t = 0; t < reader.FieldCount; t++)
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();
                    for (int x = 0; x < f1.testDocDGV.Rows.Count; x++)
                        f1.testDocDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                    for (int x = 0; x < f1.testDocDGV.Rows.Count; x++)
                        f1.testDocDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                    reader.Close();

                    //bluray/dvd info
                    cmd = new SQLiteCommand("SELECT * FROM testingdvdblu WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    data = new string[reader.FieldCount];
                    for (int t = 0; t < reader.FieldCount; t++)
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();
                    for (int x = 0; x < f1.testDVDDGV.Rows.Count; x++)
                        f1.testDVDDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                    for (int x = 0; x < f1.testDVDDGV.Rows.Count; x++)
                        f1.testDVDDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                    reader.Close();

                    //iptv info
                    cmd = new SQLiteCommand("SELECT * FROM testingiptv WHERE Building=@b AND Room=@r;", conn);
                    cmd.Parameters.AddWithValue("@b", b);
                    cmd.Parameters.AddWithValue("@r", r);
                    reader = cmd.ExecuteReader();
                    data = new string[reader.FieldCount];
                    for (int t = 0; t < reader.FieldCount; t++)
                        data[t] = reader.GetName(t).ToString();
                    reader.Read();
                    for (int x = 0; x < f1.testIPTVDGV.Rows.Count; x++)
                        f1.testIPTVDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                    for (int x = 0; x < f1.testIPTVDGV.Rows.Count; x++)
                        f1.testIPTVDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                    reader.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    conn.Dispose();
                    f1.testFocus();
                }
            }
        }
    }
}
