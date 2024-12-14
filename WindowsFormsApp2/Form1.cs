using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using System.IO;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        private List<string[]> rows = null;
        private List<string[]> Filteredrows = null;
        public SqlConnection SqlConnection = null;


        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Stroka_podklych_2"].ConnectionString);

            SqlConnection.Open();


            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("Select * from Product",SqlConnection);
  
            DataSet ds = new DataSet();

            sqlDataAdapter.Fill(ds);

            dataGridView3.DataSource = ds.Tables[0];

            //if (SqlConnection.State == ConnectionState.Open)
            //{ MessageBox.Show("Подключение Устоновлено"); }

            //-----------------------------------------------------------------


            SqlDataReader Datareader_1 = null;

            rows = new List<string[]>();

            string[] row = null;

            try
            {
                SqlCommand sqlCommand = new SqlCommand("Select * from data_Buyer order by Id_Buyer", SqlConnection);

                Datareader_1 = sqlCommand.ExecuteReader();

                while (Datareader_1.Read())
                {
                    row = new string[]
                          {Convert.ToString(Datareader_1["Id_Buyer"])
                          ,Convert.ToString(Datareader_1["Login_buyer"])
                          ,Convert.ToString(Datareader_1["Password_buyer"])
                          ,Convert.ToString(Datareader_1["Discount_buyer"])
                          ,Convert.ToString(Datareader_1["Country"])
                          ,Convert.ToString(Datareader_1["City"])
                          ,Convert.ToString(Datareader_1["Address_buyer"])
                          ,Convert.ToString(Datareader_1["Name_buyer_status"])
                          ,Convert.ToString(Datareader_1["Buyer_category"])
                          ,Convert.ToString(Datareader_1["Email"])
                          ,Convert.ToString(Datareader_1["Mobil_phone"])

                    };
                    rows.Add(row);
                }
            }
            catch (Exception ty)
            {
                MessageBox.Show(ty.Message);
            }
            finally
            {
                if (Datareader_1 != null && !Datareader_1.IsClosed)
                {
                    Datareader_1.Close();
                }
            }

            RefreshList(rows);
            //------------------------------------------------------------------------------------------------


        }


        private void RefreshList(List<string[]> list)
        {
            listView2.Items.Clear();
            foreach (string[] s in list)
            {
                listView2.Items.Add(new ListViewItem(s));           
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            SqlDataAdapter dataAdapter = new SqlDataAdapter(
           "Select * from DATA_Purchase", SqlConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            dataGridView1.DataSource = dataSet.Tables[0];


        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                textBox1.Text, SqlConnection);

            DataSet dataSet2 = new DataSet();

            dataAdapter.Fill(dataSet2);

            dataGridView2.DataSource = dataSet2.Tables[0];

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand
            ($"begin tran  exec ManufacturerAdd N'{textBox2.Text}',N'{textBox3.Text}',N'{textBox4.Text}',{textBox5.Text}," +
            $"N'{textBox6.Text}',N'{textBox7.Text}',N'{textBox8.Text}' commit", SqlConnection);
            MessageBox.Show(command.ExecuteNonQuery().ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            SqlCommand command2 = new SqlCommand
            ($"begin tran exec AttributesAdd N'{textBox9.Text}',N'{textBox10.Text}',N'{textBox11.Text}',N'{textBox12.Text}'," +
            $"N'{textBox13.Text}',N'{textBox14.Text}' commit"
            , SqlConnection);

            DateTime date = DateTime.Parse(textBox14.Text);


            command2.Parameters.AddWithValue("Color", textBox9.Text);
            command2.Parameters.AddWithValue("Volume", textBox10.Text);
            command2.Parameters.AddWithValue("Memory", textBox11.Text);
            command2.Parameters.AddWithValue("Clock_speed", textBox12.Text);
            command2.Parameters.AddWithValue("Weight_Product", textBox13.Text);
            command2.Parameters.AddWithValue("Production_date", $"{date.Year}/{date.Month}/{date.Day}");

            MessageBox.Show(command2.ExecuteNonQuery().ToString());


        }

        private void button5_Click(object sender, EventArgs e)
        {

            listView1.Items.Clear(); //Отчиска listView1

            SqlDataReader Datareader = null;

            try
               {
                SqlCommand sqlCommand = new SqlCommand("Select * from Attributes order by ID_Attributes", SqlConnection);

                Datareader = sqlCommand.ExecuteReader();

                ListViewItem item = null;


                while (Datareader.Read())
                {
                    item = new ListViewItem(new string[] {Convert.ToString(Datareader["Id_Attributes"]) 
                          ,Convert.ToString(Datareader["Color"])
                          ,Convert.ToString(Datareader["Volume"])
                          ,Convert.ToString(Datareader["Memory"])
                          ,Convert.ToString(Datareader["Clock_speed"])
                          ,Convert.ToString(Datareader["Weight_Product"])
                          ,Convert.ToString(Datareader["Production_date"]) 
                    });

                    listView1.Items.Add(item); //Запись в listView1
                }
            }
            catch (Exception ex) 
               { 
                
                MessageBox.Show(ex.Message);
            
               }
            finally 
               {
                if (Datareader != null && !Datareader.IsClosed)
                
                { 
                    
                    Datareader.Close(); 
                
                } 
                        
            }         

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Name_product like '%{textBox15.Text}%'";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case  0:
                     (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Quantity <= 100"; 
                    break;
                case  1:
                    (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Quantity >= 100 and Quantity <= 500";
                    break;
                case 2:
                    (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Quantity >= 500 and Quantity <= 100000";
                    break;
                case  3:
                    (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Quantity >= 100000 and Quantity <= 1000000";
                    break;
                case  4:
                    (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"Quantity >= 1000000 and Quantity <= 1000000000";
                    break;
                case  5:
                    (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = $"";
                    break;
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            Filteredrows = rows.Where((x) => x[7].ToLower().Contains(textBox16.Text.ToLower())).ToList();
            RefreshList(Filteredrows);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case 0:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) <= 5).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 1:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) >= 5 && double.Parse(x[3]) <= 8).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 2:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) >= 8 && double.Parse(x[3]) <= 12).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 3:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) >= 12 && double.Parse(x[3]) <= 20).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 4:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) >= 20 && double.Parse(x[3]) <= 40).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 5:
                    Filteredrows = rows.Where((x) => double.Parse(x[3]) >= 40 && double.Parse(x[3]) <= 65).ToList();
                    RefreshList(Filteredrows);
                    break;
                case 6:
                    RefreshList(Filteredrows);
                    break;
                
            }
        }
        //  выгрузка файла в Эксель. Снужными показателями
        private void button6_Click(object sender, EventArgs e)
        {
            var ExelApp = new Excel.Application();

            ExelApp.Visible = true;
            var wb = ExelApp.Workbooks.Add(1);
            var ws = wb.Worksheets[1];
            int columns = 1;
            int Rows = 1;
            foreach (ListViewItem lvi in listView2.Items)
            {
                columns = 1;
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    ws.Cells[Rows, columns] = lvs.Text;
                    columns++;
                }
                Rows++;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            using (Magazin_3Entities db = new Magazin_3Entities())
            {
                List<Buyer> List_2 = db.Buyers.ToList();
                foreach (Buyer b in List_2)
                {
                    ListViewItem item_1 = new ListViewItem(b.Id_buyer.ToString());
                    item_1.SubItems.Add(b.Id_Buyer_category.ToString());
                    item_1.SubItems.Add(b.Id_Contact_details.ToString());
                    item_1.SubItems.Add(b.Id_Buyer_status.ToString());
                    item_1.SubItems.Add(b.Login_buyer);
                    item_1.SubItems.Add(b.Password_buyer.ToString());
                    item_1.SubItems.Add(b.Discount_buyer.ToString());
                    item_1.SubItems.Add(b.Country);
                    item_1.SubItems.Add(b.City);
                    item_1.SubItems.Add(b.Address_buyer.ToString());
                    listView3.Items.Add(item_1);
                }
                Cursor.Current = Cursors.Default;
            }

        }

        private async void button8_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "CSV|*.csv", ValidateNames = true })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (StreamWriter sw = new StreamWriter(new FileStream(saveFileDialog.FileName, FileMode.Create), Encoding.UTF8))
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("Id_buyer,Id_Buyer_category,Id_Contact_details,Id_Buyer_status," +
                                      "Login_buyer,Password_buyer,Discount_buyer,Country,City,Address_buyer");
                        foreach (ListViewItem item in listView3.Items)
                        {
                            sb.AppendLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}",
                                item.SubItems[0].Text, item.SubItems[1].Text,
                                item.SubItems[2].Text, item.SubItems[3].Text,
                                item.SubItems[4].Text, item.SubItems[5].Text,
                                item.SubItems[6].Text, item.SubItems[7].Text,
                                item.SubItems[8].Text, item.SubItems[9].Text
                                ));

                        }
                        await sw.WriteLineAsync(sb.ToString());
                        MessageBox.Show("Ваши данные были успешно обработаны.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }


                
    }
}


