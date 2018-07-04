using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

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
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT ID, SERIAL_NBR, [Transaction DTM], LATITUDE, LONGITUDE, RouteID, DateStamp, SequenceID, Time, URL FROM [2018-3-24]", connection);
                DataTable data = new DataTable();
                dataAdapter.Fill(data);
                // For each row, check SERIAL_NBR
                object sNNR = 0;
                int count = 0;
                DataRow prevRow = data.Rows[0];
                DataRow firstRow = data.Rows[0];
                //Console.WriteLine(prevNBR);
                foreach (DataRow row in data.Rows)
                {
                    if (row[1].Equals(sNNR))
                    {
                        count++;
                        cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].SequenceID = " + count + " WHERE [2018-3-24].[ID] = " + row[0] + ";";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "UPDATE[2018-3-24]" + " SET[2018-3-24].URL = " + "\"http://api.metrocloudalliance.com/route/?from_place=" + prevRow[3] + "," + prevRow[4] + "&to_place=" + row[3] + "," + row[4] + "&date=" + prevRow[6] + "&time=" + prevRow[8] + "&mode=TRANSIT,WALK&max_itineraries=4&output_format=xml&api_key=Lvj4Y3icznHjPSWT8HjU\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                        //Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        count = 1;
                        sNNR = row[1];
                        cmd.CommandText = "UPDATE[2018-3-24] SET[2018-3-24].SequenceID = " + count + " WHERE [2018-3-24].[ID] = " + row[0] + ";";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "UPDATE[2018-3-24]" + " SET[2018-3-24].URL = " + "\"http://api.metrocloudalliance.com/route/?from_place=" + prevRow[3] + "," + prevRow[4] + "&to_place=" + firstRow[3] + "," + firstRow[4] + "&date=" + prevRow[6] + "&time=" + prevRow[8] + "&mode=TRANSIT,WALK&max_itineraries=4&output_format=xml&api_key=Lvj4Y3icznHjPSWT8HjU\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                        //Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                        firstRow = row;
                    }
                    prevRow = row;
                }
                prevRow = data.Rows[data.Rows.Count - 1];
                cmd.CommandText = "UPDATE[2018-3-24]" + " SET[2018-3-24].URL = " + "\"http://api.metrocloudalliance.com/route/?from_place=" + prevRow[3] + "," + prevRow[4] + "&to_place=" + firstRow[3] + "," + firstRow[4] + "&date=" + prevRow[6] + "&time=" + prevRow[8] + "&mode=TRANSIT,WALK&max_itineraries=4&output_format=xml&api_key=Lvj4Y3icznHjPSWT8HjU\"" + " WHERE [2018-3-24].[ID] = " + prevRow[0] + ";";
                cmd.ExecuteNonQuery();
                //dataAdapter.Update(data);
                tableDisplay.DataSource = data;
                tableDisplay.AutoResizeColumns();

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            
        }

        private void tableDisplay_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
