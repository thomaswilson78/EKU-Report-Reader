using System;
using System.Windows.Forms;

namespace EKU_Work_Thing
{
    public partial class Form3 : Form
    {
        //inherit data (campusData and textbox information) from form 1
        Form1 f1 = new Form1();
        public Form3(Form1 parent)
        {
            f1 = parent;
            
            InitializeComponent();
            BuildingCB.SelectedIndex = 0;//defaults to first item in list
            addRooms();
            RoomCB.SelectedIndex = 0;
        }
        //add rooms to the listboxs
        private void addRooms()
        {
            RoomCB.Items.Clear();
            if (f1.campusData.Count > 0)
            {
                foreach (var room in f1.campusData)
                    if (room.Building.Equals(BuildingCB.Text))
                    RoomCB.Items.Add(room.Room);
            }
            if(RoomCB.Items.Count>0)
                RoomCB.SelectedIndex = 0;
        }
        //takes building and room information from selected values and fills the values from the .csv report into the inventory tab of the main form
        private void pullDataBtn_Click(object sender, EventArgs e)
        {
            f1.addBuildingComBox.SelectedItem = BuildingCB.Text;
            f1.addRoomTB.Text = RoomCB.Text;
            roomInfo exactRoom = new roomInfo();
            foreach (var room in f1.campusData)
            {
                if (room.Building.Equals(BuildingCB.Text) && room.Room.Equals(RoomCB.Text))
                {
                    exactRoom = room;
                    break;
                }
            }
            f1.addContComBox.SelectedItem = exactRoom.control;
            f1.addAudioComBox.SelectedItem = exactRoom.audio;
            f1.addDockCB.Checked = exactRoom.dock;
            f1.addDCCB.Checked = exactRoom.docCam;
            f1.addCamCB.Checked = exactRoom.camera;
            f1.addMicCB.Checked = exactRoom.mic;
            f1.addBRCB.Checked = exactRoom.Bluray;
            f1.addDVDCB.Checked = exactRoom.DVD;
            f1.addHDMICB.Checked = exactRoom.hdmi;
            f1.addVGACB.Checked = exactRoom.vga;
            f1.addAVCB.Checked = exactRoom.av;
            f1.addSolCB.Checked = exactRoom.sol;
            f1.addMMTB1.Text = exactRoom.display1;
            f1.addSerialTB1.Text = exactRoom.serial1;
            f1.addScrTB1.Text = exactRoom.screen1;
            f1.addIPTB1.Text = exactRoom.ip1;
            f1.addMACTB1.Text = exactRoom.mac1;
            f1.addBulbTB1.Text = exactRoom.bulb1;
            f1.addMMTB2.Text = exactRoom.display2;
            f1.addSerialTB2.Text = exactRoom.serial2;
            f1.addScrTB2.Text = exactRoom.screen2;
            f1.addIPTB2.Text = exactRoom.ip2;
            f1.addMACTB2.Text = exactRoom.mac2;
            f1.addBulbTB2.Text = exactRoom.bulb3;
            f1.addMMTB3.Text = exactRoom.display3;
            f1.addSerialTB3.Text = exactRoom.serial3;
            f1.addScrTB3.Text = exactRoom.screen3;
            f1.addIPTB3.Text = exactRoom.ip3;
            f1.addMACTB3.Text = exactRoom.mac3;
            f1.addBulbTB3.Text = exactRoom.bulb3;
            f1.addMMTB4.Text = exactRoom.display4;
            f1.addSerialTB4.Text = exactRoom.serial4;
            f1.addScrTB4.Text = exactRoom.screen4;
            f1.addIPTB4.Text = exactRoom.ip4;
            f1.addMACTB4.Text = exactRoom.mac4;
            f1.addBulbTB4.Text = exactRoom.bulb4;
            if (exactRoom.filter.ToString("MM/dd/yyyy") != "01/01/0001")
                f1.addFilter.Text = exactRoom.filter.ToString("MM/dd/yyyy");
            else
                f1.addFilter.Text = "";
            f1.addPCModTB.Text = exactRoom.PCModel;
            f1.addPCSerialTB.Text = exactRoom.PCSerial;
            f1.addNUCIPTB.Text = exactRoom.nucip;
            f1.addNUCMACTB.Text = exactRoom.nucmac;
            f1.addCatVidTB.Text = exactRoom.Cat6.ToString();
            f1.addNetTB.Text = exactRoom.NetPorts.ToString();
            if (exactRoom.solDate.ToString("MM/dd/yyyy") != "01/01/0001")
                f1.addSolActTB.Text = exactRoom.solDate.ToString("MM/dd/yyyy");
            else
                f1.addSolActTB.Text = "";
            f1.addOtherTB.Text = exactRoom.other;
            f1.Refresh();
            this.Close();
        }
        //changes rooms based on building selected
        private void BuildingCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            addRooms();
        }
        //hide form instead of closing, help perserve last entered data
        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }
    }
}
