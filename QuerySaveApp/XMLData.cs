using Excel = Microsoft.Office.Interop.Excel;
// excel object library is required to use the above, not the fucking shitty interlop piece of shit.
using System.Reflection;//fuck knows what this one is being used for.
// gotta use system not microsoft or it errors on using a certificate for some reason.
using System.Data.SqlClient;
using System;
using System.Data;

namespace QuerySaveApp
{
    public static class XMLData
    {
        
        public static string getFilePath()
        {

            string filePath = Environment.CurrentDirectory + @"\xmldata.xml";
            return filePath;
        }

         public static DataSet ReturnXMLDataset()
        {
           
            DataSet XMLDATASET = new DataSet();

            XMLDATASET.ReadXml(getFilePath());

            return XMLDATASET;
        }

        public static void SaveData(DataSet ds)
        {
            //DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath());
        }
    }
}
