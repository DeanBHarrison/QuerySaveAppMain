using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuerySaveApp
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();

            //set Datasource and Datamember, both of these are required to return data.
            dataGridView1.DataSource = XMLData.ReturnXMLDataset(2);
            dataGridView1.DataMember = "Authors";

            ////add a button column and then add buttons to it
            DataGridClass dataGridClass = new DataGridClass();
            //dataGridClass.addButtonColumn(dataGridView1, "Browse save location", "Run", "Savebutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Browse load location", "Run", "Loadbutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Run Report", "Run", "ReportButton", "Run");
            // dataGridClass.SetColumnsOrder(dataGridView1, "Stored_Procedure", "Save_location", "Browse save location", "Load_location", "Browse load location");
            dataGridClass.DisableTableSorting(dataGridView1);
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //SAVES THE DATASET.

            //what is this (dataset)
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath(2));
        }
    }
}
