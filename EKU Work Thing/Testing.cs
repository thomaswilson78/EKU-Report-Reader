using System;
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

        //load data into table
        public void loadTable()
        {
            tDGV.Rows.Clear();
            //pulls data from the database to load into the table
            using (SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;"))
            {
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
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
        }
        //hides form to be pulled up later
        private void Testing_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }
        //delete all items from both the datagridview (program table) and from the database
        private void tDeleteAll_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to delete all items?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm.ToString().Equals("OK"))
            {
                confirm = MessageBox.Show("Are you certain you want to delete all the data? This action cannot be undone.", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation); //prompt twice to prevent accidents
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
        //deletes selected items from the datagridview (program's table) and database
        private void tDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to delete these items?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm.ToString().Equals("OK"))
            {
                bool success = false;
                bool error = false;
                using (SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;foreign keys=true"))
                {
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
                    catch (Exception)
                    {
                        success = false;
                        error = true;
                    }
                    if (success)
                        MessageBox.Show("Item(s) successfully removed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else if (error)
                        MessageBox.Show("Error occured when removing items, not all items were deleted.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                        MessageBox.Show("No items selected for deletion.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    tDGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void tClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tDGV.Rows.Count; i++)
                tDGV.Rows[i].Cells[3].Value = false;
        }
        //handles all instances where the cell is clicked. Mainly for handling when the export and import buttons are clicked
        private void tDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //for exporting
            if(e.ColumnIndex == 2)
            {
                Type officeType = Type.GetTypeFromProgID("Excel.Application");
                if (officeType == null)
                    MessageBox.Show("Microsoft Office and Excel needs to be installed to export testing reports.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Templates\", "testing sheet.xlsx");
                    Workbooks wbs = xlApp.Workbooks; //this is BEYOND the stupidest fucking thing ever. You have to make a Workbooks object, otherwise when you close everything else out, this will still remain and excel will still be open.
                    Workbook wb = wbs.Add(path);
                    Worksheet ws = (Worksheet)wb.Worksheets[1];
                    using (SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3"))
                    {
                        try
                        {
                            conn.Open();
                            //gets the building and room number for later
                            string b = tDGV.Rows[e.RowIndex].Cells[0].Value.ToString();
                            string r = tDGV.Rows[e.RowIndex].Cells[1].Value.ToString();
                            ws.Name = b + " " + r;

                            //main, fills in basic information (building, room, name, date)
                            SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM testingmain WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                ws.Cells[36, 1] = "Building: " + b;
                                ws.Cells[36, 6] = r;
                                ws.Cells[38, 1] = "Agent Name: " + reader["Agent"].ToString();
                                string dt = reader["DateCol"].ToString();
                                dt = dt.Substring(0, dt.Length - 5);//removes time from date
                                ws.Cells[38, 6] = dt;
                                ws.Cells[67, 1] = reader["Notes"].ToString();
                            }
                            string[] data; //collects the column names in the database to make looping easier when needing to fetch a column's data from an SQL data reader
                            string[,] mark; //indicates where to mark in the excel spreadsheet
                            string[,] notes; //simply collects notes 
                            int pos; //tracks position for reading data (usually info is stored at the beginning and notes are stored at the end in the database, excepting being video/audio).
                            //(cont.) Since two loops need to be used for two arrays, felt it was easier to track using a variable than figuring where the data needed to start in 2nd loop manually

                            //general
                            cmd = new SQLiteCommand("SELECT * FROM testinggeneral WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);

                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);
                                //much faster for interop to apply a data set from a range using an array than enter data in one cell at at time
                                mark = new string[16, 3];
                                for (int i = 0; i < 16; i++)
                                    for (int k = 0; k < 3; k++)
                                        mark[i, k] = "";
                                notes = new string[16, 1];
                                for (int i = 0; i < 16; i++)
                                    notes[i, 0] = "";
                                pos = 2;
                                for (int y = 0; y < 16; y++)
                                {
                                    mark[y, int.Parse(reader[data[pos]].ToString()) - 1] = "X";
                                    pos++;
                                }
                                for (int y = 0; y < 16; y++)
                                {
                                    notes[y, 0] = reader[data[pos]].ToString();
                                    pos++;
                                }
                                ws.Range[ws.Cells[5, 2], ws.Cells[20, 4]].Value = mark;
                                ws.Range[ws.Cells[5, 5], ws.Cells[20, 5]].Value = notes;
                                ws.Range[ws.Cells[5, 2], ws.Cells[20, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[5, 5], ws.Cells[20, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            }//everything follows the same general principal (except being vid/aud)
                            //if ever want to streamline, turn this into a function.

                            //vid/aud
                            cmd = new SQLiteCommand("SELECT * FROM testingvideoaudio WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();

                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);

                                mark = new string[6, 6];
                                for (int i = 0; i < 6; i++)
                                    for (int k = 0; k < 6; k++)
                                        mark[i, k] = "";
                                //no need for notes, everything is linear
                                pos = 2;
                                for (int y = 0; y < 6; y++)
                                    for (int x = 0; x < 6; x++)
                                    {
                                        mark[y, x] = reader[data[pos]].ToString();
                                        pos++;
                                    }

                                ws.Range[ws.Cells[24, 2], ws.Cells[29, 7]].Value = mark;
                                ws.Range[ws.Cells[24, 2], ws.Cells[29, 6]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[24, 7], ws.Cells[29, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            }

                            //mic
                            cmd = new SQLiteCommand("SELECT * FROM testingmic WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);
                                mark = new string[2, 3];
                                for (int i = 0; i < 2; i++)
                                    for (int k = 0; k < 3; k++)
                                        mark[i, k] = "";
                                notes = new string[2, 1];
                                for (int i = 0; i < 2; i++)
                                    notes[i, 0] = "";

                                pos = 2;
                                for (int y = 0; y < 2; y++)
                                {
                                    mark[y, int.Parse(reader[data[pos]].ToString()) - 1] = "X";
                                    pos++;
                                }
                                for (int y = 0; y < 2; y++)
                                {
                                    notes[y, 0] = reader[data[pos]].ToString();
                                    pos++;
                                }
                                ws.Range[ws.Cells[33, 2], ws.Cells[34, 4]].Value = mark;
                                ws.Range[ws.Cells[33, 5], ws.Cells[34, 5]].Value = notes;
                                ws.Range[ws.Cells[33, 2], ws.Cells[34, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[33, 5], ws.Cells[34, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            }

                            //doc cam
                            cmd = new SQLiteCommand("SELECT * FROM testingdoccam WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);
                                mark = new string[4, 3];
                                for (int i = 0; i < 4; i++)
                                    for (int k = 0; k < 3; k++)
                                        mark[i, k] = "";
                                notes = new string[4, 1];
                                for (int i = 0; i < 4; i++)
                                    notes[i, 0] = "";

                                pos = 2;
                                for (int y = 0; y < 4; y++)
                                {
                                    mark[y, int.Parse(reader[data[pos]].ToString()) - 1] = "X";
                                    pos++;
                                }
                                for (int y = 0; y < 4; y++)
                                {
                                    notes[y, 0] = reader[data[pos]].ToString();
                                    pos++;
                                }
                                ws.Range[ws.Cells[45, 2], ws.Cells[48, 4]].Value = mark;
                                ws.Range[ws.Cells[45, 5], ws.Cells[48, 5]].Value = notes;
                                ws.Range[ws.Cells[45, 2], ws.Cells[48, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[45, 5], ws.Cells[48, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            }

                            //dvd/blu
                            cmd = new SQLiteCommand("SELECT * FROM testingdvdblu WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);
                                mark = new string[4, 3];
                                for (int i = 0; i < 4; i++)
                                    for (int k = 0; k < 3; k++)
                                        mark[i, k] = "";
                                notes = new string[4, 1];
                                for (int i = 0; i < 4; i++)
                                    notes[i, 0] = "";

                                pos = 2;
                                for (int y = 0; y < 4; y++)
                                {
                                    mark[y, int.Parse(reader[data[pos]].ToString()) - 1] = "X";
                                    pos++;
                                }
                                for (int y = 0; y < 4; y++)
                                {
                                    notes[y, 0] = reader[data[pos]].ToString();
                                    pos++;
                                }
                                ws.Range[ws.Cells[52, 2], ws.Cells[55, 4]].Value = mark;
                                ws.Range[ws.Cells[52, 5], ws.Cells[55, 5]].Value = notes;
                                ws.Range[ws.Cells[52, 2], ws.Cells[55, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[52, 5], ws.Cells[55, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            }

                            //iptv
                            cmd = new SQLiteCommand("SELECT * FROM testingiptv WHERE Building=@b AND Room=@r", conn);
                            cmd.Parameters.AddWithValue("@b", b);
                            cmd.Parameters.AddWithValue("@r", r);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                reader.Read();
                                data = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                    data[i] = reader.GetName(i);
                                mark = new string[5, 3];
                                for (int i = 0; i < 5; i++)
                                    for (int k = 0; k < 3; k++)
                                        mark[i, k] = "";
                                notes = new string[5, 1];
                                for (int i = 0; i < 5; i++)
                                    notes[i, 0] = "";

                                pos = 2;
                                for (int y = 0; y < 5; y++)
                                {
                                    mark[y, int.Parse(reader[data[pos]].ToString()) - 1] = "X";
                                    pos++;
                                }
                                for (int y = 0; y < 5; y++)
                                {
                                    notes[y, 0] = reader[data[pos]].ToString();
                                    pos++;
                                }
                                ws.Range[ws.Cells[59, 2], ws.Cells[63, 4]].Value = mark;
                                ws.Range[ws.Cells[59, 5], ws.Cells[63, 5]].Value = notes;
                                ws.Range[ws.Cells[59, 2], ws.Cells[63, 4]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ws.Range[ws.Cells[59, 5], ws.Cells[63, 10]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                reader.Close();
                            }

                            //main part done, attempts to save completed file
                            xlApp.DisplayAlerts = false; //used because excel is stupid an will prompt again if you want to replace the file (even though s.f.d will already ask you that).

                            SaveFileDialog sfd = new SaveFileDialog();
                            sfd.FileName = tDGV.Rows[e.RowIndex].Cells[0].Value.ToString() + " " + tDGV.Rows[e.RowIndex].Cells[1].Value.ToString(); //set default filename which will consist of building and room #
                            sfd.Filter = "Excel Spreadsheet (*.xlsx)|*.xlsx"; //so it saves as an excel file only
                            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //defaults the directory to Documents
                            if (sfd.ShowDialog() == DialogResult.OK) //even occurs if the ok button is pressed
                            {
                                try
                                {
                                    wb.Close(SaveChanges: true, Filename: sfd.FileName.ToString()); //Filename will included specified path
                                }
                                catch (Exception)
                                {
                                    wb.Close(0);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        //release all excel related objects to close excel, otherwise process will remain in memory unless closed via task manager.
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(ws) > 0) ;
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(wb) > 0) ;
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs) > 0) ;
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) > 0) ;
                        ws = null;
                        wb = null;
                        wbs = null;
                        xlApp = null;
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                }
            }
            //for editing
            if (e.ColumnIndex == 3)
            {
                f1.resetTestingTables();//completely resets the data in the testing tables before adding data in
                DataGridView getDGVInfo = (DataGridView)sender;//converts the sender object into data datagridview variable
                int i = e.RowIndex;//get's row index for next two variables
                string b = getDGVInfo.Rows[i].Cells[0].Value.ToString();//get's building from row clicked
                string r = getDGVInfo.Rows[i].Cells[1].Value.ToString();//get's room # from row clicked
                getDGVInfo = null;//object not use again, set to null for garbage collection later
                using (SQLiteConnection conn = new SQLiteConnection("Data Source=ReportDB.sqlite;Version=3;"))
                {
                    try
                    {
                        //main info
                        conn.Open();
                        SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM testingmain WHERE Building = @b AND Room = @r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);

                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            reader.Read();
                            f1.testBuilding.Text = reader["Building"].ToString();
                            f1.testRoom.Text = reader["Room"].ToString();
                            f1.testName.Text = reader["Agent"].ToString();
                            f1.testDate.Text = reader["DateCol"].ToString();
                            f1.testNotesTB.Text = reader["Notes"].ToString();
                        }


                        //general info
                        //only commenting the first one since every one (except vid/aud) works the same.
                        
                        cmd = new SQLiteCommand("SELECT * FROM testinggeneral WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);

                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];//column names extracted to string array (way easier than writing them out manually)
                            for (int t = 0; t < reader.FieldCount; t++)//fills in string with column names
                                data[t] = reader.GetName(t).ToString();
                            reader.Read();//reads the data (important, otherwise no data will be collected. also no need for a while or if since there's only one instance of data).
                            for (int x = 0; x < f1.testGeneralDGV.Rows.Count; x++)
                                f1.testGeneralDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;//sets value based on number recorded
                            for (int x = 0; x < f1.testGeneralDGV.Rows.Count; x++)//gets notes
                                f1.testGeneralDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                        }

                        //vid/aud info
                        cmd = new SQLiteCommand("SELECT * FROM testingvideoaudio WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];
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
                        }

                        //mic info
                        cmd = new SQLiteCommand("SELECT * FROM testingmic WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];
                            for (int t = 0; t < reader.FieldCount; t++)
                                data[t] = reader.GetName(t).ToString();
                            reader.Read();
                            for (int x = 0; x < f1.testMicDGV.Rows.Count; x++)
                                f1.testMicDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                            for (int x = 0; x < f1.testMicDGV.Rows.Count; x++)
                                f1.testMicDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                        }

                        //doccam info
                        cmd = new SQLiteCommand("SELECT * FROM testingdoccam WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];
                            for (int t = 0; t < reader.FieldCount; t++)
                                data[t] = reader.GetName(t).ToString();
                            reader.Read();
                            for (int x = 0; x < f1.testDocDGV.Rows.Count; x++)
                                f1.testDocDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                            for (int x = 0; x < f1.testDocDGV.Rows.Count; x++)
                                f1.testDocDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                        }

                        //bluray/dvd info
                        cmd = new SQLiteCommand("SELECT * FROM testingdvdblu WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];
                            for (int t = 0; t < reader.FieldCount; t++)
                                data[t] = reader.GetName(t).ToString();
                            reader.Read();
                            for (int x = 0; x < f1.testDVDDGV.Rows.Count; x++)
                                f1.testDVDDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                            for (int x = 0; x < f1.testDVDDGV.Rows.Count; x++)
                                f1.testDVDDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                        }

                        //iptv info
                        cmd = new SQLiteCommand("SELECT * FROM testingiptv WHERE Building=@b AND Room=@r;", conn);
                        cmd.Parameters.AddWithValue("@b", b);
                        cmd.Parameters.AddWithValue("@r", r);
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            string[] data = new string[reader.FieldCount];
                            for (int t = 0; t < reader.FieldCount; t++)
                                data[t] = reader.GetName(t).ToString();
                            reader.Read();
                            for (int x = 0; x < f1.testIPTVDGV.Rows.Count; x++)
                                f1.testIPTVDGV.Rows[x].Cells[int.Parse(reader[data[x + 2].ToString()].ToString())].Value = true;
                            for (int x = 0; x < f1.testIPTVDGV.Rows.Count; x++)
                                f1.testIPTVDGV.Rows[x].Cells[4].Value = reader["Notes" + (x + 1)].ToString();
                        }
                        f1.testFocus();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                } 
            }
        }

        private void tClose_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
