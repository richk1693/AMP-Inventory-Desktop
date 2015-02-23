using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;
using System.Configuration;
using System.IO;

namespace InventoryDesktop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private StringReader myReader;
        SqlCeConnection conn = new SqlCeConnection();
        //SqlCeConnection conn = new SqlCeConnection(@"Data Source = C:\Users\Dawg\Documents\Visual Studio 2008\Projects\InventoryDesktop\InventoryDesktop\AppDatabase2.sdf");
        //SqlCeConnection conn = new SqlCeConnection(@"Data Source = E:\Documents and Settings\Larry\My Documents\WINDOWSMOBILE15 My Documents\AppDatabase2.sdf");
        //SqlCeConnection conn = new SqlCeConnection(@"Data Source = C:\Users\Larry\Documents\Documents on SS 1\AppDatabase2.sdf");


        //All of the on load events weren't made by me so i'm scared to mess with them.
        private void Form1_Load(object sender, EventArgs e)
        {
            /*// TODO: This line of code loads data into the 'appDatabase2DataSet.Table2' table. You can move, or remove it, as needed.
            this.table2TableAdapter.Fill(this.appDatabase2DataSet.Table2);
            // TODO: This line of code loads data into the 'appDatabase2DataSet.Table2' table. You can move, or remove it, as needed.
            this.table2TableAdapter.Fill(this.appDatabase2DataSet.Table2);
            // TODO: This line of code loads data into the 'appDatabase2DataSet.Table2' table. You can move, or remove it, as needed.
            this.table2TableAdapter.Fill(this.appDatabase2DataSet.Table2);
            // TODO: This line of code loads data into the 'appDatabase2DataSet.Table2' table. You can move, or remove it, as needed.
            this.table2TableAdapter.Fill(this.appDatabase2DataSet.Table2);
             */
            if(File.Exists(@"C:\Users\Dawg\Documents\Visual Studio 2008\Projects\InventoryDesktop\InventoryDesktop\AppDatabase2.sdf"))
                conn.ConnectionString = @"Data Source = C:\Users\Dawg\Documents\Visual Studio 2008\Projects\InventoryDesktop\InventoryDesktop\AppDatabase2.sdf";
            if (File.Exists(@"C:\Users\clem\My Documents\Documents on SS 1\AppDatabase2.sdf"))
                conn.ConnectionString = @"Data Source = C:\Users\clem\My Documents\Documents on SS 1\AppDatabase2.sdf";
            if (File.Exists(@"C:\Users\Larry\Documents\Documents on SS 1\AppDatabase2.sdf"))
                conn.ConnectionString = @"Data Source = C:\Users\Larry\Documents\Documents on SS 1\AppDatabase2.sdf";
            reloadGrid();
        }


        //Attempt to reload the grid.
        private void reloadGrid() {
            SqlCeDataAdapter dataAdapter = new SqlCeDataAdapter("SELECT * FROM Table2", conn);
            DataTable table = new DataTable();
            dataAdapter.Fill(table);
            table.DefaultView.Sort="Type,barcode";
            dataGridView1.DataSource = table;
            
            //view.Sort= "Barcode, min";
            //dataGridView1.DataSource = view;

        }
        //Auto-generated method. I'm not touchin it.
        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }

       
        //Export the database as an Excel file on button click.
        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //ToCsV(dataGridView1, @"c:\export.xls");
                ToCsV(dataGridView1, sfd.FileName); // Here dataGridview1 is your grid view name 
            }  
        }


        //Save changes in the current row to the database
        private void button3_Click_1(object sender, EventArgs e)
        {
            
            string test = dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
            //Giant query is just an update for each database value
            string query = @"UPDATE Table2 SET onHand = @1, madeYTD = @2, soldYTD = @3, Min = @4,
                            EOQ = @5, ytdMoAvg = @6, Cost = @7, Value = @8, type = @9 WHERE [Barcode] = (@0);";
            conn.Open();
            SqlCeCommand cmd = new SqlCeCommand(query, conn);
            //Each line is a column from the dataGridView, establishing the relationship between them and the db.
            cmd.Parameters.AddWithValue("@0", dataGridView1[0,dataGridView1.CurrentCell.RowIndex].Value.ToString());
            cmd.Parameters.AddWithValue("@1", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@2", dataGridView1[2, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@3", dataGridView1[3, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@4", dataGridView1[4, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@5", dataGridView1[5, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@6", dataGridView1[6, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@7", dataGridView1[7, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@8", dataGridView1[8, dataGridView1.CurrentCell.RowIndex].Value);
            cmd.Parameters.AddWithValue("@9", dataGridView1[10, dataGridView1.CurrentCell.RowIndex].Value.ToString());
            cmd.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("Save Successful");
        }


        //Fill listBox1 with reorder information when "Reorder" button is clicked.
        private void button4_Click(object sender, EventArgs e)
        {
            SqlCeCommand cmd = new SqlCeCommand("SELECT * FROM Table2", conn);
            conn.Open();
            SqlCeDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                int a = rdr.GetInt32(1);
                int b = rdr.GetInt32(4);

                //If amount on hand is <= reorder then add it to the reorder list.
                if (rdr.GetInt32(1) <= rdr.GetInt32(4))
                {
                    listBox1.Items.Add(rdr.GetString(0) +"           On Hand:"+ rdr.GetInt32(1).ToString() + "           Reorder: " + rdr.GetInt32(4).ToString() + "           EOQ: " + rdr.GetInt32(5));
                }

            }
            conn.Close();
        }



        //Function to print listBox from online tutorial. I'm assuming it's magic cuz i dont understand it.
        protected void printDocument1_PrintPage(object sender,System.Drawing.Printing.PrintPageEventArgs ev)
  {
      float linesPerPage = 0;
      float yPosition = 0;
      int count = 0;
      float leftMargin = ev.MarginBounds.Left;
      float topMargin = ev.MarginBounds.Top;
      string line = null;
     //Font printFont = this.listBox1.Font;
       Font printFont = new Font("times New Roman", 12);
     SolidBrush myBrush = new SolidBrush(Color.Black);
 
     // Work out the number of lines per page, using the MarginBounds.
      linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);
 
     // Iterate over the string using the StringReader, printing each line.
      while (count < linesPerPage && ((line = myReader.ReadLine()) != null))
     {
         // calculate the next line position based on
         // the height of the font according to the printing device
          yPosition = topMargin + (count * printFont.GetHeight(ev.Graphics));
 
         // draw the next line in the rich edit control
  
         ev.Graphics.DrawString(line, printFont,
                                myBrush, leftMargin,
                                yPosition, new StringFormat());
         count++;
     }
      
 
     // If there are more lines, print another page.
     if (line != null)
         ev.HasMorePages = true;
     else
         ev.HasMorePages = false;
 
     myBrush.Dispose();
 }
        
        //When the print button is clicked, print the document! (Online tutorial)
        private void button5_Click(object sender, EventArgs e)
        {
           printDialog1.Document = printDocument1;
            string strText = "";
            foreach (object x in listBox1.Items)
            {
                 strText = strText + x.ToString() + "\n";
            }
 
            myReader = new StringReader(strText);
            if (printDialog1.ShowDialog() == DialogResult.OK)
             {
                this.printDocument1.Print();
             }

        }

        //Add a record to the database
        private void addButton_Click(object sender, EventArgs e)
        {
            bool flag = true;
            for (int i = 0; i < 11; i++)
            {
                if (dataGridView1[i, dataGridView1.CurrentCell.RowIndex].Value.ToString() == "")
                {
                    flag = false;
                    MessageBox.Show("Please insert a value for every field before adding");
                    break;
                }
            }
            if (flag)
            {
                conn.Open();
                SqlCeCommand cmd = new SqlCeCommand(@"INSERT INTO Table2 (barcode, onHand, madeYTD, soldYTD,
                                                min, EOQ, ytdMoAvg, cost, value, onHand1231, type) VALUES (@0,@1,
                                                @2,@3,@4,@5,@6,@7,@8,@9,@10);", conn);
                cmd.Parameters.AddWithValue("@0", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                cmd.Parameters.AddWithValue("@1", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@2", dataGridView1[2, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@3", dataGridView1[3, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@4", dataGridView1[4, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@5", dataGridView1[5, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@6", dataGridView1[6, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@7", dataGridView1[7, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@8", dataGridView1[8, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@9", dataGridView1[9, dataGridView1.CurrentCell.RowIndex].Value);
                cmd.Parameters.AddWithValue("@10", dataGridView1[10, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        //Remove a record from the database
        private void deleteButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you wish to delete row " + dataGridView1[0,dataGridView1.CurrentCell.RowIndex].Value.ToString(), "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                conn.Open();
                SqlCeCommand cmd = new SqlCeCommand("DELETE FROM Table2 WHERE [barcode] = @0", conn);
                cmd.Parameters.AddWithValue("@0", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                cmd.ExecuteNonQuery();
                reloadGrid();
                conn.Close();
                MessageBox.Show("Record Deleted");
            }
            
        }

        //End of year changes. made and sold YTD = 0. 1231 = current onhand.
        private void button2_Click(object sender, EventArgs e)
        {

            var result = MessageBox.Show("Are you sure you wish to run end of year?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                conn.Open();
                SqlCeCommand cmd = new SqlCeCommand("SELECT * FROM Table2", conn);
                SqlCeCommand cmd2 = new SqlCeCommand("UPDATE Table2 SET madeYTD = 0, soldYTD = 0, onHand1231 = @1 WHERE barcode = @2", conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    cmd2.Parameters.AddWithValue("@1", rdr.GetInt32(1).ToString());
                    cmd2.Parameters.AddWithValue("@2", rdr.GetString(0));
                    cmd2.ExecuteNonQuery();
                    cmd2.Parameters.Clear();
                }
                reloadGrid();
                conn.Close();
                MessageBox.Show("End of year executed successfully.");
            }
        }



        
    }
}
