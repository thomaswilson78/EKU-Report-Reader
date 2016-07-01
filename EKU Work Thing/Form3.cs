using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EKU_Work_Thing
{
    public partial class Form3 : Form
    {
        BuildingData bd = new BuildingData();
        Form1 f1 = new Form1();
        public Form3(Form1 f)
        {
            f1 = f;
            InitializeComponent();
            BuildingCB.SelectedIndex = 0;
            addRooms();
            RoomCB.SelectedIndex = 0;

        }
        private void addRooms()
        {
            RoomCB.Items.Clear();
            if (f1.campusData.Count > 0)
            {
                foreach (var room in f1.campusData)
                    if (room.Building.Equals(BuildingCB.Text))
                    RoomCB.Items.Add(room.Room);
            }
            RoomCB.SelectedIndex = 0;
        }
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
            //f1.addSolCB.Checked = exactRoom.sol; need to add solstice to roomInfo, footprints report, and to load
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
            f1.addBulbTB3.Text = exactRoom.bulb3;
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
            f1.addFilter.Text = exactRoom.filter.ToString("MM/dd/yyyy");
            //f1.addPCModTB.Text = exactRoom; need to add PC to roomInfo, footprints report, and to load
            //f1.addPCSerialTB.Text = exactRoom; need to add PC serial to roomInfo, footprints report, and to load
            //f1.addNUCIPTB.Text = reader["NUCIP"].ToString();
            //f1.addCatVidTB.Text = reader["Cat6Video"].ToString();
            //f1.addNetTB.Text = reader["NetworkPorts"].ToString();
            f1.addSolActTB.Text = exactRoom.alarm.ToString("MM/dd/yyyy");
            //f1.addSolLicTB.Text = reader["SolsticeLicense"].ToString();
            //f1.addDscrptTB.Text = reader["Notes"].ToString();
            f1.Refresh();
            this.Close();
        }

        private void BuildingCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            addRooms();
        }
    }
}
