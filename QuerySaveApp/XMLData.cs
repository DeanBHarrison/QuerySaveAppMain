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
        public static string getFilePath(int filenumber)
        {          

            //1 = FRONT PAGE
            //2 = SETTINGS

            if (filenumber == 1)
            {
                string filepath = Environment.CurrentDirectory + @"\xmldata.xml";
                return filepath;
            }
            else if (filenumber == 2)
            {
                string filepath = Environment.CurrentDirectory + @"\xmldata2.xml";
                return filepath;
            }
            else
            {
                return "Incorrect Filepath";
            }
        }

         public static DataSet ReturnXMLDataset(int filenumber)
        {          
            DataSet XMLDATASET = new DataSet();
            XMLDATASET.ReadXml(getFilePath(filenumber));
            return XMLDATASET;
        }

        public static void SaveData(DataSet ds, int XMLint)
        {
            //DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath(XMLint));
        }


        public static void test()
        {
            // Here your xml file
            string xmlFile = "Data.xml";

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlFile, XmlReadMode.InferSchema);

            // Then display informations to test
            foreach (DataTable table in dataSet.Tables)
            {
                Console.WriteLine(table);
                for (int i = 0; i < table.Columns.Count; ++i)
                    Console.Write("\t" + table.Columns[i].ColumnName.Substring(0, Math.Min(6, table.Columns[i].ColumnName.Length)));
                Console.WriteLine();
                foreach (var row in table.AsEnumerable())
                {
                    for (int i = 0; i < table.Columns.Count; ++i)
                    {
                        Console.Write("\t" + row[i]);
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}
