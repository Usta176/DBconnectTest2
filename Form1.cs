using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CxDbLib;

namespace DBconnectTest
{
    public partial class Form1 : Form
    {
        // The class for the database communication
        private static DbRoute DbInterface = new DbRoute();

        public Form1()
        {
            Dictionary<string, string> ConfigValues = new Dictionary<string, string>();

            InitializeComponent();

            ConfigValues["database"] = "1";
            ConfigValues["sqlserver"] = "C:\\ProgramData\\Camaix\\Tools4Tools\\Tool.db";
            ConfigValues["sqluser"] = "sa";
            ConfigValues["sqlpassword"] = "cdiT21E#";
            ConfigValues["trustedconnection"] = "yes";

            ConfigValues["sqlitedatei"] = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\Camaix\\Tools4Tools\\Tool.db";

            DbInterface.SetConfigData(ConfigValues);

            DbInterface.InitDbSystem(1); // 1 means database system Sqlite

            Dictionary<string, string> DinKeys = DbInterface.GetDin4000Keys(1);// 1 means database system Sqlite

            SortedDictionary<string, string> DinNames = DbInterface.GetDin4000Names(1);// 1 means database system Sqlite

            

            for (int i = 0; i < DinKeys.Count; i++)
            {
                dataGridView1.Rows.Add(DinKeys.ElementAt(i).Key, DinKeys.ElementAt(i).Value);
            }

            for (int i = 0; i < DinNames.Count; i++)
            {
                dataGridView2.Rows.Add(DinNames.ElementAt(i).Key, DinNames.ElementAt(i).Value);
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
