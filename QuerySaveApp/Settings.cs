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
            //add a button column and then add buttons to it
            DataGridClass.DisableTableSorting(dataGridView1);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            richTextBox1.AppendText("bullshit");
            DataGridClass.LoadCombocells(dataGridView1);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //what is this (dataset)
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath(2));
            DataGridClass.LoadCombocells(dataGridView1);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
