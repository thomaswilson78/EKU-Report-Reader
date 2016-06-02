using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Office.Interop.Excel;

/*To Do:
    -Need to adjust footprints data. Some rooms have projector but do not have screen data, so they'll be missed in the report. Also some
        rooms still do not have screens.
    -Need to improve way of getting district data, especially for excel data usage. Likely will store into a 2D array.
    -Need to export data into an excel document that can be used for reporting data. Data will be for districts.
*/
namespace EKU_Work_Thing
{
    public partial class Form1 : Form
    {
        List<roomInfo> campusData = new List<roomInfo>();
        //used to populate the location listbox. If possible, need to find a better way without searching every object in campusData
        
        string[] districtArray = new string[] {
            "Crabbe Library","University Building","Combs Classroom","Keith Building","McCreary Building","Weaver Health",//library: 6
            "Cammack Building","Moore Building","Memorial Science","Roark Building",//oldsci: 4
            "New Science Building","Dizney Building","Rowlett Building",//newsci: 3
            "Wallace Building","Case Annex","Powell Building",//central: 3
            "Stratton Building","Ashland Building","Perkins Building","Carter Building",//justice: 4
            "Whitlock Building",//services: 1
            "Coates Administration Building","Jones Building",//admin: 2
            "Burrier Building","Campbell Building","Whalin Complex","Foster Music Building",//arts: 4
            "Alumni Coliseum","Begley Building","Moberly Building","Gentry Building"//fitness: 4
            };
        public Form1()
        {
            InitializeComponent();
            tabControl1.TabPages.Remove(Display2);
            tabControl1.TabPages.Remove(Display3);
            tabControl1.TabPages.Remove(Display4);
            tabControl1.TabPages.Remove(OtherDevices);
            tabControl1.TabPages.Remove(Description);
            buildLB.SetSelected(0, true);
            districtLB.SetSelected(0, true);

            for (int i = 0; i < 15; i++)
                testGeneralDGV.Rows.Add();
            testGeneralDGV.Rows[0].Cells[0].Value = "Display(s) power on";
            testGeneralDGV.Rows[1].Cells[0].Value = "(CP) Warm up screen alots enough time for projector warm up light to stop.";
            testGeneralDGV.Rows[2].Cells[0].Value = "No Filter or Bulb Life error message.";
            testGeneralDGV.Rows[3].Cells[0].Value = "Image is properly centered, zoomed, \"squared\" (i.e. not jutting corners), and focused.";
            testGeneralDGV.Rows[4].Cells[0].Value = "Equipment, cart, and cables are neat and organized.";
            testGeneralDGV.Rows[5].Cells[0].Value = "(CP) Quick guide is displayed after powering on the system.";
            testGeneralDGV.Rows[6].Cells[0].Value = "(CP) Time/Date are properly displayed and are correct.";
            testGeneralDGV.Rows[7].Cells[0].Value = "(CP) \"Help\" button displays IT contact information when pressed.";
            testGeneralDGV.Rows[8].Cells[0].Value = "(CP) Video mute button highlights when selected and mutes the display.";
            testGeneralDGV.Rows[9].Cells[0].Value = "(CP) \"Auto Image\" button auto images and stays highlighted until complete.";
            testGeneralDGV.Rows[10].Cells[0].Value = "Screen controls ▲, ▼, ■  work properly and screen is not marked/damaged.";
            testGeneralDGV.Rows[11].Cells[0].Value = "Lighting system properly dims and brightens light.";
            testGeneralDGV.Rows[12].Cells[0].Value = "(CP) Cool down screen alots enough time for projector cool down light to stop.";
            testGeneralDGV.Rows[13].Cells[0].Value = "(CP) Powering down shuts down all other equipment (Doc Cam, etc).";
            testGeneralDGV.Rows[14].Cells[0].Value = "Check filter for damage (yes if damaged, no if good).";

            for (int i = 0; i < 6; i++)
                testVidAudDGV.Rows.Add();
            testVidAudDGV.Rows[0].Cells[0].Value = "(CP) Button stays highlighted with properly label when selected and held.";
            testVidAudDGV.Rows[1].Cells[0].Value = "Output displays properly without issues (e.g. no flickering, vibrating, etc.)";
            testVidAudDGV.Rows[2].Cells[0].Value = "Audio works properly without issues (e.g. no static, poor quality, etc.)";
            testVidAudDGV.Rows[3].Cells[0].Value = "Volume adjusts properly on speakers.";
            testVidAudDGV.Rows[4].Cells[0].Value = "(CP) Mute button mutes audio and highlights when selected.";
            testVidAudDGV.Rows[5].Cells[0].Value = "(CP) \"All\", \"Left\", \"Right\", and \"Center\" display options show proper source.";

            for (int i = 0; i < 2; i++)
                testMicDGV.Rows.Add();
            testMicDGV.Rows[0].Cells[0].Value = "Microphone output sounds clear (e.g. no popping, static, feedback, etc).";
            testMicDGV.Rows[1].Cells[0].Value = "Microphone volume level is approriately loud enough.";

            for (int i = 0; i < 4; i++)
                testDocDGV.Rows.Add();
            testDocDGV.Rows[0].Cells[0].Value = "Zoom + and - fuctions change image accordingly";
            testDocDGV.Rows[1].Cells[0].Value = "Focus ▲, ▼ and \"Auto\" functions change image accordingly.";
            testDocDGV.Rows[2].Cells[0].Value = "Iris ▲, ▼ and \"Normal\" functions change image accordingly.";
            testDocDGV.Rows[3].Cells[0].Value = "(CP) All soft keys are properly labeled & highlight properly when being held down.";

            for (int i = 0; i < 4; i++)
                testDVDDGV.Rows.Add();
            testDVDDGV.Rows[0].Cells[0].Value = "(CP) \"Menu\" and \"Title\" buttons bring up their respective screens.";
            testDVDDGV.Rows[1].Cells[0].Value = "(CP) Arrow keys navigate properly through the menus.";
            testDVDDGV.Rows[2].Cells[0].Value = "(CP) All soft key controls for play, stop, fast-forward, etc. work accordingly.";
            testDVDDGV.Rows[3].Cells[0].Value = "(CP) All soft keys are properly labeled & highlight properly when being held down.";

            for (int i = 0; i < 5; i++)
                testIPTVDGV.Rows.Add();
            testIPTVDGV.Rows[0].Cells[0].Value = "(CP) Arrow keys navigate properly through the menus.";
            testIPTVDGV.Rows[1].Cells[0].Value = "(CP) Channel up and down buttons work correctly.";
            testIPTVDGV.Rows[2].Cells[0].Value = "(CP) Channel number can be directly input to navigate to channel.";
            testIPTVDGV.Rows[3].Cells[0].Value = "(CP) \"Last\" button goes to the IPTV main menu.";
            testIPTVDGV.Rows[4].Cells[0].Value = "(CP) All soft keys are properly labeled & highlight properly when being held down.";
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }//Done

        //loads a footprints .csv file into the program
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                //using is a way of assigning functionality without needing to use multiple statements such as "ofd.Filter="CSV|*.csv"" 
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV|*.csv", ValidateNames = true, Multiselect = false })
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        //removes old data to prevent data overlapping
                        if(campusData.Count>0)
                        {
                            campusData.RemoveRange(0,campusData.Count-1);
                        }
                        //TextFieldParser works like a StreamReaser, but parses .csv data properly.
                        //Requires a reference to the Microsoft.VisualBasic .dll file
                        //and then needs the "using Microsoft.VisualBasic.FileIO" line to utilize it.
                        TextFieldParser parser = new TextFieldParser(ofd.FileName);
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        while (!parser.EndOfData)
                        {
                            roomInfo newRoom = new roomInfo();//Object collects data about room
                            //seperate the .csv data by ','
                            string[] lines = parser.ReadFields();

                            if (!lines[0].Equals("Building Equipment Resides In")) //store all relevant data as attributes in an object
                            {
                                newRoom.Building = lines[0];
                                newRoom.Room = lines[1];
                                newRoom.display1 = lines[2];
                                newRoom.display2 = lines[3];
                                newRoom.display3 = lines[4];
                                newRoom.display4 = lines[5];
                                newRoom.serial1 = lines[6];
                                newRoom.serial2 = lines[7];
                                newRoom.serial3 = lines[8];
                                newRoom.serial4 = lines[9];
                                newRoom.screen1 = lines[10];
                                newRoom.screen2 = lines[11];
                                newRoom.screen3 = lines[12];
                                newRoom.screen4 = lines[13];
                                newRoom.ip1 = lines[14];
                                newRoom.ip2 = lines[15];
                                newRoom.ip3 = lines[16];
                                newRoom.ip4 = lines[17];
                                newRoom.mac1 = lines[18];
                                newRoom.mac2 = lines[19];
                                newRoom.mac3 = lines[20];
                                newRoom.mac4 = lines[21];
                                newRoom.bulb1 = lines[22];
                                newRoom.bulb2 = lines[23];
                                newRoom.bulb3 = lines[24];
                                newRoom.bulb4 = lines[25];
                                newRoom.dock = lines[26].Equals("On");
                                newRoom.docCam = lines[27].Equals("Yes");
                                newRoom.DVD = lines[28].Equals("On");
                                newRoom.Bluray = lines[29].Equals("On");
                                newRoom.camera = lines[30].Equals("On");
                                newRoom.mic = lines[31].Equals("On");
                                newRoom.vga = lines[32].Equals("On");
                                newRoom.hdmi = lines[33].Equals("On");
                                newRoom.audio = lines[34];
                                newRoom.control = lines[35];
                                lines[36] = lines[36].Replace("\n", "; ");
                                newRoom.other = lines[36];
                                newRoom.description = lines[37];
                                DateTime.TryParse(lines[38], out newRoom.filter);
                                DateTime.TryParse(lines[39], out newRoom.alarm);
                                newRoom.av = lines[40].Equals("On");
                                campusData.Add(newRoom);//Add object into collection of objects (rooms)
                            }
                        }
                        parser.Close();
                        ofd.Dispose();

                        int dispCount = 0;
                        int filCount = 0;
                        foreach (var room in campusData)
                        {
                            //Set district for each building
                            if (room.Building.Equals("Crabbe Library") || room.Building.Equals("University Building") || room.Building.Equals("Combs Classroom")
                                || room.Building.Equals("Keith Building") || room.Building.Equals("McCreary Building"))
                                room.District = "Library District";
                            else if (room.Building.Equals("Cammack Building") || room.Building.Equals("Moore Building") || room.Building.Equals("Memorial Science")
                                || room.Building.Equals("Roark Building"))
                                room.District = "Old Science District";
                            else if (room.Building.Equals("New Science Building") || room.Building.Equals("Dizney Building") || room.Building.Equals("Rowlett Building"))
                                room.District = "New Science District";
                            else if (room.Building.Equals("Wallace Building") || room.Building.Equals("Case Annex") || room.Building.Equals("Powell Building"))
                                room.District = "Central Campus Area";
                            else if (room.Building.Equals("Stratton Building") || room.Building.Equals("Ashland Building") || room.Building.Equals("Perkins Building")
                                || room.Building.Equals("Carter Building"))
                                room.District = "Justice District";
                            else if (room.Building.Equals("Whitlock Building"))
                                room.District = "Services District";
                            else if (room.Building.Equals("Coates Administration Building") || room.Building.Equals("Jones Building"))
                                room.District = "Administrative District";
                            else if (room.Building.Equals("Campbell Building") || room.Building.Equals("Foster Music Building") || room.Building.Equals("Burrier Building")
                                || room.Building.Equals("Whalin Complex"))
                                room.District = "Arts District";
                            else if (room.Building.Equals("Alumni Coliseum") || room.Building.Equals("Begley Building") || room.Building.Equals("Moberly Building"))
                                room.District = "Fitness District";
                            
                            //counts number of rooms and projectors
                            if (!room.display1.Equals(""))
                                dispCount++;
                            if (!room.display2.Equals(""))
                                dispCount++;
                            if (!room.display3.Equals(""))
                                dispCount++;
                            if (!room.display4.Equals(""))
                                dispCount++;
                            int temp = (DateTime.Now - room.filter).Days;

                            //check if filter is older than 3 months and report it if last filter clean date is older than that
                            if (temp >= 90)
                            {
                                filCount++;
                                int test = dataGridView1.Rows.Count;
                                DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                                row.Cells[0].Value = room.Building;
                                row.Cells[1].Value = room.Room;
                                if (room.filter.ToShortDateString().Equals("1/1/0001"))
                                    row.Cells[2].Value = "N/A";
                                else
                                    row.Cells[2].Value = room.filter.ToShortDateString();
                                if (room.alarm.ToShortDateString().Equals("1/1/0001"))
                                    row.Cells[3].Value = "N/A";
                                else
                                    row.Cells[3].Value = room.alarm.ToShortDateString();
                                dataGridView1.Rows.Add(row);
                            }
                        }
                        //sort filter data in table and prints out display/filter data
                        dataGridView1.Sort(roomDG, ListSortDirection.Ascending);
                        dataGridView1.Sort(buildingDG, ListSortDirection.Ascending);
                        //add rooms to the list based on the currently selected building
                        roomsLB.Items.Clear();
                        if (buildLB.SelectedIndex >= 0)
                        {
                            foreach (var rooms in campusData)
                            {
                                if (buildLB.SelectedItem.ToString().Equals(rooms.Building))
                                {
                                    roomsLB.Items.Add(rooms.Room);
                                }
                            }
                        }
                        int disRooms = 0;
                        int invCollect = 0;
                        //checks that inventory information has been entered since the new hires started to closely
                        //approximate the number of inventory information that has been collected.
                        foreach (var dist in campusData)
                        {
                            if (districtLB.SelectedItem.ToString().Equals(dist.District))
                            {
                                disRooms++;
                                if (dist.filter > DateTime.Parse("03/13/2016"))
                                {
                                    invCollect++;
                                }
                            }
                        }
                        //prints all data last in case of errors
                        totalDisplaysTB.Text = dispCount.ToString();
                        mainNeededTB.Text = filCount.ToString();
                        campusData = campusData.OrderBy(o => o.Room).ToList();//Order by room first
                        campusData = campusData.OrderBy(o => o.Building).ToList();//Then order by building
                        totalRoomsTB.Text = campusData.Count.ToString();//Prints total number of rooms
                        disTotalTB.Text = disRooms.ToString();
                        disInvTB.Text = invCollect.ToString();
                        exportToolStripMenuItem.Enabled = true;
                    }
            }
            catch (Exception)
            {
                disInvTB.Text = "0";
                disTotalTB.Text = "0";
                totalRoomsTB.Clear();
                makeModelTB1.Text = "";
                makeModelTB2.Text = "";
                makeModelTB3.Text = "";
                makeModelTB4.Text = "";
                serialTB1.Text = "";
                serialTB2.Text = "";
                serialTB3.Text = "";
                serialTB4.Text = "";
                screenTB1.Text = "";
                screenTB2.Text = "";
                screenTB3.Text = "";
                screenTB4.Text = "";
                ipTB1.Text = "";
                ipTB2.Text = "";
                ipTB3.Text = "";
                ipTB4.Text = "";
                macTB1.Text = "";
                macTB2.Text = "";
                macTB3.Text = "";
                macTB4.Text = "";
                bulbTB1.Text = "";
                bulbTB2.Text = "";
                bulbTB3.Text = "";
                bulbTB4.Text = "";
                otherTB.Text = "";
                descriptionTB.Text = "";
                filterTB.Text = "";
                alarmTB.Text = "";
                controlTB.Text = "";
                audioTB.Text = "";
                dsCB.Checked = false;
                lcCB.Checked = false;
                avcpCB.Checked = false;
                dvdCB.Checked = false;
                brCB.Checked = false;
                dcCB.Checked = false;
                micCB.Checked = false;
                vgaCB.Checked = false;
                hdmiCB.Checked = false;
                avcpCB.Checked = false;
                tabControl1.TabPages.Remove(Display2);
                tabControl1.TabPages.Remove(Display3);
                tabControl1.TabPages.Remove(Display4);
                tabControl1.TabPages.Remove(OtherDevices);
                tabControl1.TabPages.Remove(Description);
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                totalDisplaysTB.Clear();
                mainNeededTB.Clear();
                roomsLB.Items.Clear();
                if (campusData.Count>0)
                    campusData.RemoveRange(0, campusData.Count - 1);
                MessageBox.Show("File could not be loaded. Make sure that the proper Footprints report is being pulled (\"#EKU REPORTING SOFTWARE REPORT\").", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                exportToolStripMenuItem.Enabled = false;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }//Done

        //updates the rooms listbox when a building is selected.
        private void buildLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            roomsLB.Items.Clear();
            foreach (var rooms in campusData)
                if (buildLB.SelectedItem.ToString().Equals(rooms.Building))
                    roomsLB.Items.Add(rooms.Room);
        }//Done

        //grabs inventory information for the building/room
        private void roomsLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (var rooms in campusData)
            {
                if (buildLB.SelectedItem != null && roomsLB.SelectedItem != null)//ensure the program doesn't crash by select a null value
                {
                    if (buildLB.SelectedItem.ToString().Equals(rooms.Building) && roomsLB.SelectedItem.ToString().Equals(rooms.Room))
                    {
                        makeModelTB1.Text = rooms.display1;
                        makeModelTB2.Text = rooms.display2;
                        makeModelTB3.Text = rooms.display3;
                        makeModelTB4.Text = rooms.display4;
                        serialTB1.Text = rooms.serial1;
                        serialTB2.Text = rooms.serial2;
                        serialTB3.Text = rooms.serial3;
                        serialTB4.Text = rooms.serial4;
                        screenTB1.Text = rooms.screen1;
                        screenTB2.Text = rooms.screen2;
                        screenTB3.Text = rooms.screen3;
                        screenTB4.Text = rooms.screen4;
                        ipTB1.Text = rooms.ip1;
                        ipTB2.Text = rooms.ip2;
                        ipTB3.Text = rooms.ip3;
                        ipTB4.Text = rooms.ip4;
                        macTB1.Text = rooms.mac1;
                        macTB2.Text = rooms.mac2;
                        macTB3.Text = rooms.mac3;
                        macTB4.Text = rooms.mac4;
                        bulbTB1.Text = rooms.bulb1;
                        bulbTB2.Text = rooms.bulb2;
                        bulbTB3.Text = rooms.bulb3;
                        bulbTB4.Text = rooms.bulb4;
                        otherTB.Text = rooms.other;
                        descriptionTB.Text = rooms.description;
                        filterTB.Text = rooms.filter.ToString("MM/dd/yyyy");
                        alarmTB.Text = rooms.alarm.ToString("MM/dd/yyyy");
                        controlTB.Text = rooms.control;
                        audioTB.Text = rooms.audio;
                        dsCB.Checked = rooms.dock;
                        lcCB.Checked = rooms.camera;
                        avcpCB.Checked = rooms.av;//not included yet
                        dvdCB.Checked = rooms.DVD;
                        brCB.Checked = rooms.Bluray;
                        dcCB.Checked = rooms.docCam;
                        micCB.Checked = rooms.mic;
                        vgaCB.Checked = rooms.vga;
                        hdmiCB.Checked = rooms.hdmi;
                        avcpCB.Checked = rooms.av;
                        break;
                    }
                }
            }
            //add/removes tabs based on if there's information regarding the room
            tabControl1.TabPages.Remove(Display2);
            tabControl1.TabPages.Remove(Display3);
            tabControl1.TabPages.Remove(Display4);
            tabControl1.TabPages.Remove(OtherDevices);
            tabControl1.TabPages.Remove(Description);

            if (!makeModelTB2.Text.Equals(""))
                if (tabControl1.TabPages.IndexOf(Display2) < 0)
                    tabControl1.TabPages.Insert(tabControl1.TabCount, Display2);

            if (!makeModelTB3.Text.Equals(""))
                if (tabControl1.TabPages.IndexOf(Display3) < 0)
                    tabControl1.TabPages.Insert(tabControl1.TabCount, Display3);

            if (!makeModelTB4.Text.Equals(""))
                if (tabControl1.TabPages.IndexOf(Display4) < 0)
                    tabControl1.TabPages.Insert(tabControl1.TabCount, Display4);

            if (!otherTB.Text.Equals(""))
                if (tabControl1.TabPages.IndexOf(OtherDevices) < 0)
                    tabControl1.TabPages.Insert(tabControl1.TabCount, OtherDevices);
            
            tabControl1.TabPages.Insert(tabControl1.TabCount, Description);
        }//Done

        //adds buildings to the location listbox based on the district selected
        private void districtLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(districtLB.SelectedIndex)
            {
                case 0:
                    locationsLB.Items.Clear();
                    for(int i=0;i<6;i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 1:
                    locationsLB.Items.Clear();
                    for (int i = 6; i < 10; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 2:
                    locationsLB.Items.Clear();
                    for (int i = 10; i < 13; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 3:
                    locationsLB.Items.Clear();
                    for (int i = 13; i < 16; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 4:
                    locationsLB.Items.Clear();
                    for (int i = 16; i < 20; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 5:
                    locationsLB.Items.Clear();
                    locationsLB.Items.Add(districtArray[20]);
                    break;
                case 6:
                    locationsLB.Items.Clear();
                    for (int i = 21; i < 23; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 7:
                    locationsLB.Items.Clear();
                    for (int i = 23; i < 27; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
                case 8:
                    locationsLB.Items.Clear();
                    for (int i = 27; i < 31; i++)
                        locationsLB.Items.Add(districtArray[i]);
                    break;
            }
            //displays information for the rooms for each district and inventory that has been collected since the new hires.
            int disRooms = 0;
            int invCollect = 0;

            foreach (var dist in campusData)
            {
                if (districtLB.SelectedItem.ToString().Equals(dist.District))
                {
                    disRooms++;
                    if (dist.filter>DateTime.Parse("03/13/2016"))
                    {
                        invCollect++;
                    }
                }
            }
            disTotalTB.Text = disRooms.ToString();
            disInvTB.Text = invCollect.ToString();
        }//Done


        //Export the data into an excel spreadsheet that charts the data.
        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //create an excel application object to open excel
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if(xlApp==null)
            {
                MessageBox.Show("Exporting requires Microsoft Office and Excel to be installed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (campusData.Count == 0)
                {
                    MessageBox.Show("No .csv file loaded.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try {
                        string test = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Template\", "template.xlsx");
                        Workbook wb = xlApp.Workbooks.Add(test);
                        Worksheet ws = (Worksheet)wb.Worksheets[1];
                        //Extracts data from objects and loads it into the excel spreadsheet
                        for (int i = 0; i < districtLB.Items.Count; i++)
                        {
                            string t = districtLB.Items[i].ToString();
                            ws.Cells[i + 2, 1] = t;

                            int disRooms = 0;
                            int invCollect = 0;
                            foreach (var dist in campusData)
                            {
                                if (t.Equals(dist.District))
                                {
                                    disRooms++;
                                    if (dist.filter > DateTime.Parse("03/13/2016"))
                                    {
                                        invCollect++;
                                    }
                                }
                            }
                            ws.Cells[i + 2, 2] = disRooms;
                            ws.Cells[i + 2, 3] = invCollect;
                        }
                        xlApp.Visible = true;

                        releaseObject(wb);
                        releaseObject(ws);
                        releaseObject(xlApp);
                    }
                    catch(Exception)
                    {
                        MessageBox.Show("Template file not found.", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }   
                }
                
            }
        }//Prints data to excel, need to create chart

        private void showForm2()
        {
            Form2 f2 = new Form2();
            f2.Show();
        }//Done

        private void viewInvCollectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showForm2();
        }//Done

        private void addEditBtn_Click(object sender, EventArgs e)
        {
            showForm2();
        }//Done

        private void addAddUpdateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (!addBuildingComBox.Text.Equals(""))
                {
                    if (!addRoomTB.Text.Equals(""))
                    {
                        if (!addMMTB1.Text.Equals(""))
                        {
                            //Do the thing
                        }
                        else
                            MessageBox.Show("Please enter information for Display 1.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                        MessageBox.Show("Please enter the room number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                    MessageBox.Show("Please Select a building.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }//In progress, need SQL


        //custom functions
        //prevents user from clicking on multiple checkboxes in the same row. works for all datagridviews using this event
        private void testPreventMultiCB(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView d = (DataGridView)sender;
            d.CommitEdit(DataGridViewDataErrorContexts.Commit);
            if (e.ColumnIndex > 0 && e.ColumnIndex < 4)
            {
                int y = e.RowIndex;
                if (e.ColumnIndex == 1)
                {
                    d.Rows[y].Cells[2].Value = null;
                    d.Rows[y].Cells[3].Value = null;
                }
                else if (e.ColumnIndex == 2)
                {
                    d.Rows[y].Cells[1].Value = null;
                    d.Rows[y].Cells[3].Value = null;
                }
                else if (e.ColumnIndex == 3)
                {
                    d.Rows[y].Cells[2].Value = null;
                    d.Rows[y].Cells[1].Value = null;
                }
            }
        }//done
        //garbage collection
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }//done

        private void followUpFocus(object sender, MouseEventArgs e)
        {
            testNotesTB.Focus();
        }
    }

    //save information about a room into an object
    public class roomInfo
    {
        public string Building { get; set; }
        public string Room { get; set; }
        public string District { get; set; }
        public string display1 { get; set; }
        public string display2 { get; set; }
        public string display3 { get; set; }
        public string display4 { get; set; }
        public string serial1 { get; set; }
        public string serial2 { get; set; }
        public string serial3 { get; set; }
        public string serial4 { get; set; }
        public string screen1 { get; set; }
        public string screen2 { get; set; }
        public string screen3 { get; set; }
        public string screen4 { get; set; }
        public string ip1 { get; set; }
        public string ip2 { get; set; }
        public string ip3 { get; set; }
        public string ip4 { get; set; }
        public string mac1 { get; set; }
        public string mac2 { get; set; }
        public string mac3 { get; set; }
        public string mac4 { get; set; }
        public string bulb1 { get; set; }
        public string bulb2 { get; set; }
        public string bulb3 { get; set; }
        public string bulb4 { get; set; }
        public string audio { get; set; }
        public string control { get; set; }
        public string other { get; set; }
        public string description { get; set; }
        public DateTime filter;
        public DateTime alarm;
        public bool dock = false;
        public bool docCam = false;
        public bool DVD = false;
        public bool Bluray = false;
        public bool camera = false;
        public bool mic = false;
        public bool vga = false;
        public bool hdmi = false;
        public bool av = false;
    }//Done
}

/*
                    //!!!Works, but does not make a good chart, nice for reference
                    //creates a workbook and worksheet file.
                    Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    Worksheet ws = (Worksheet)wb.Worksheets[1];
                    //Headers
                    ws.Cells[1, 1] = "District";
                    ws.Cells[1, 2] = "Total Rooms";
                    ws.Cells[1, 3] = "Completed";
                    ws.Cells[1, 4] = "Percent Finished";
                    //Extracts data from objects and loads it into the excel spreadsheet
                    for (int i = 0; i < districtLB.Items.Count; i++)
                    {
                        string t = districtLB.Items[i].ToString();
                        ws.Cells[i + 2, 1] = t;

                        int disRooms = 0;
                        int invCollect = 0;
                        foreach (var dist in campusData)
                        {
                            if (t.Equals(dist.District))
                            {
                                disRooms++;
                                if (dist.filter > DateTime.Parse("03/13/2016"))
                                {
                                    invCollect++;
                                }
                            }
                        }
                        ws.Cells[i + 2, 2] = disRooms;
                        ws.Cells[i + 2, 3] = invCollect;
                        ws.Cells[i + 2, 4] = ((double)invCollect / (double)disRooms);
                        ws.Cells[i + 2, 4].NumberFormat = "0.00%";
                    }
                    //Creates a chart in the worksheet
                    ChartObjects cObjs = (ChartObjects)ws.ChartObjects();
                    ChartObject cObj = cObjs.Add(5, 200, 600, 300);
                    Chart c = cObj.Chart;
                    c.HasTitle = true;
                    c.ChartTitle.Text = "Maintenance Completed";
                    //Extracts information from cells to add into the chart
                    SeriesCollection seriesCollection = c.SeriesCollection();

                    Series series1 = seriesCollection.NewSeries();
                    Range xValues = ws.Range["A2", "A10"];
                    Range values = ws.Range["B2", "B10"];
                    series1.XValues = xValues;
                    series1.Values = values;

                    Series series2 = seriesCollection.NewSeries();
                    values = ws.Range["C2", "C10"];
                    series2.Values = values;

                    series1.Name = "Total Rooms";
                    series2.Name = "Completed";


                    series1.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowLabel,false,true,false,false,false,true,false,true,true);

                    xlApp.Visible = true;

                    releaseObject(wb);
                    releaseObject(ws);
                    releaseObject(xlApp);
                    */
