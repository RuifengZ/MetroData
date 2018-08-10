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
using System.Data.OleDb;
using System.Xml.Linq;

namespace MetroData
{
    public partial class Form1 : Form
    {
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Documents\RuifengSummerJob\MetroData\MetroDatabase.accdb;Persist Security Info=False;";
        private OleDbConnection connection;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            connectToolStripMenuItem.PerformClick();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void connectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                connection = new OleDbConnection(connectionString);
                connection.Open();
                disconnectToolStripMenuItem.Enabled = true;
                connectToolStripMenuItem.Enabled = false;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void disconnectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                connection = new OleDbConnection(connectionString);
                connection.Close();
                disconnectToolStripMenuItem.Enabled = false;
                connectToolStripMenuItem.Enabled = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = connection;
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT ID, SERIAL_NBR, LATITUDE, LONGITUDE, RouteID, DateStamp, SequenceID, Time, URL FROM [2018-3-24]", connection);
                DataTable data = new DataTable();
                dataAdapter.Fill(data);
                // For each row, check SERIAL_NBR
                object sNNR = 0;
                int count = 0;
                DataRow prevRow = data.Rows[0];
                DataRow firstRow = data.Rows[0];
                string xmlll = @"http://api.metrocloudalliance.com/route/?from_place=34.07275,-118.045138888889&to_place=34.1405555555556,-118.141388888889&date=3/10/2017&time=06:54:02&mode=TRANSIT,WALK&max_itineraries=4&output_format=xml&api_key=Lvj4Y3icznHjPSWT8HjU";
                XDocument.Load(xmlll);
                foreach (DataRow row in data.Rows)
                {
                    row[2] = dmsToDegLat(row[2].ToString());
                    row[3] = dmsToDegLon(row[3].ToString());
                    if (row[1].Equals(sNNR))
                    {
                        count++;
                        row[8] = parseURL(prevRow[2], prevRow[3], row[2], row[3], prevRow[5], prevRow[7]);
                        //cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].URL = " + "\"" + row[8] + "\"" + ", " + "[2018-3-24].XML = " + "\"" + @parseXml(XDocument.Load(row[8].ToString())) + "\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                    }
                    else
                    {
                        count = 1;
                        sNNR = row[1];
                        row[8] = parseURL(prevRow[2], prevRow[3], firstRow[2], firstRow[3], prevRow[5], prevRow[7]);
                        //cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].URL = " + "\"" + row[8] + "\"" + ", " + "[2018-3-24].XML = " + "\"" + @parseXml(XDocument.Load(row[8].ToString())) + "\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                        firstRow = row;
                    }
                    Console.WriteLine(row[8].ToString());
                    Console.WriteLine(row[4].ToString());
                    findRoute(XDocument.Load(row[8].ToString()), row[4].ToString());
                    cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].URL = " + "\"" + row[8] + "\"" + ", " + "[2018-3-24].XML = " + "\"HEY\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                    //cmd.ExecuteNonQuery();
                    cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].SequenceID = " + count + ", " + "[2018-3-24].LATITUDE_C = " + row[2] + ", " + "[2018-3-24].LONGITUDE_C = " + row[3] + " WHERE [2018-3-24].[ID] = " + row[0] + ";";
                    //cmd.ExecuteNonQuery();
                    prevRow = row;
                }
                prevRow = data.Rows[data.Rows.Count - 1];
                prevRow[8] = parseURL(prevRow[2], prevRow[3], firstRow[2], firstRow[3], prevRow[5], prevRow[7]);
                //cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].URL = " + "\"" + prevRow[8] + "\"" + ", " + "[2018-3-24].XML = " + "\"" + @parseXml(XDocument.Load(prevRow[8].ToString())) + "\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                //cmd.ExecuteNonQuery();
                //dataAdapter.Update(data);
                tableDisplay.DataSource = data;
                tableDisplay.AutoResizeColumns();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private double dmsToDegLat(string a)
        {
            string[] dms = a.Split(' ');
            if (dms.Length == 1)
                return Convert.ToDouble(a);
            return Math.Round(Convert.ToDouble(dms[0]) + Convert.ToDouble(dms[1]) / 60.0 + Convert.ToDouble(dms[2]) / 3600.0, 5);
        }

        private double dmsToDegLon(string a)
        {
            string[] dms = a.Split(' ');
            if (dms.Length == 1)
                return Convert.ToDouble(a);
            return Math.Round(Convert.ToDouble(dms[0]) - Convert.ToDouble(dms[1]) / 60.0 - Convert.ToDouble(dms[2]) / 3600.0, 5);
        }


        private string findRoute(XDocument xml, string route)
        {
            if (xml.Descendants("itineraries").Count() < 1)
                return "FAIL";
            var itineraries = xml.Descendants("itineraries").First();
            //Console.WriteLine(itineraries);
            var routenum = route.Substring(6).PadLeft(3, '0');

            Console.WriteLine("---------------------------------------------------------------------");

            foreach (var item in itineraries.Elements())
            {
                Console.WriteLine(item.Descendants("route_short_name").Count());
                //Console.WriteLine(item);
                //Console.WriteLine(items.Descendants("route"));
                foreach (var routes in item.Descendants("route_short_name"))
                {
                    Console.WriteLine("R1: " + routes.Value);
                    Console.WriteLine("R2: " + routenum);
                    if (routes.Value.Equals(routenum))
                    {
                        if(item.Descendants("route_short_name").Count() == 1)
                        {
                            return "route";
                        }
                        else
                        {
                            Console.WriteLine(item.Descendants("route_short_name"));
                        }

                    }
                    //Console.WriteLine(routes.Value);
                }
                //Console.WriteLine("NEXT");
            }
            //itineraries.Element("item0");
            return "";
        }

        private string parseXml(XDocument xml)
        {
            //StringReader reader = new StringReader(xml);
            //string xmlLine, xmlFull = "";
            //while (true)
            //{
            //    xmlLine = reader.ReadLine();
            //    if (xmlLine != null)
            //    {
            //        xmlFull = xmlFull + xmlLine + "\r\n";
            //    }
            //    else
            //    {
            //        break;
            //    }
            //}
            //return xmlFull;
            return "";
        }

        private string parseURL(object lat1, object lon1, object lat2, object lon2, object date, object time)
        {
            return "http://api.metrocloudalliance.com/route/?from_place=" + lat1 + "," + lon1 + "&to_place=" + lat2 + "," + lon2 + "&date=" + date.ToString().Split(' ')[0] + "&time=" + time + "&mode=TRANSIT,WALK&max_itineraries=4&output_format=xml&api_key=Lvj4Y3icznHjPSWT8HjU";
        }

        private void tableDisplay_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
