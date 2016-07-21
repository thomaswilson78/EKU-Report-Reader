using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EKU_Work_Thing;//have to reference, now that I'm making the title more official (that or completely redo the project).
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
namespace EKU_Report_Reader
{
    public partial class Testing : Form
    {
        Form1 f1 = new Form1();
        public Testing(Form1 parent)
        {
            f1 = parent;
            InitializeComponent();
            //will initialize table here
        }
    }
}
