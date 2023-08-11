using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.IO;
using Microsoft.Data.SqlClient;

namespace Task_2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private async void открытьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                var res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileNames;
                    if(fileName.Length > 1)
                    {
                        foreach(var file in fileName)
                        {                   
                            toolStripComboBox1.Items.Add(file);
                            await OpenAndImportExcelFile(file);
                        }
                    }
                    else
                    {
                        Text = fileName[0];

                        await OpenAndImportExcelFile(Text);
                    }

                }
                else
                {
                    throw new Exception("Неверный файл!");
                }
                toolStripComboBox1.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task OpenAndImportExcelFile(string path)
        {
            using(FileStream fs = File.OpenRead(path))
            {
                
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs);

                DataSet ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });
                
                tableCollection = ds.Tables;

                foreach (DataTable table in tableCollection)
                {

                    var arr = table.Select();
                    using (SqlConnection connection = new SqlConnection(connectionStringToDB))
                    {
                        await connection.OpenAsync();
                        try
                        {
                            SqlCommand commandDefault = new SqlCommand();
                            commandDefault.Connection = connection;
                            commandDefault.CommandText = @"INSERT INTO Class_2 VALUES (@CHECKS, @OPENING_BALANCE_ACTIVE, @OPENING_BALANCE_PASSIVE, @TURNS_DEBIT, @TURNS_CREDIT, @OUTGOING_BALANCE_ACTIVE, @OUTGOING_BALANCE_PASSIVE)";
                            commandDefault.Parameters.Add("@CHECKS", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@OPENING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@OPENING_BALANCE_PASSIVE", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@TURNS_DEBIT", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@TURNS_CREDIT", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@OUTGOING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                            commandDefault.Parameters.Add("@OUTGOING_BALANCE_PASSIVE", SqlDbType.NVarChar);

                            commandDefault.Parameters["@CHECKS"].Value = arr[8][0].ToString();
                            commandDefault.Parameters["@OPENING_BALANCE_ACTIVE"].Value = DBNull.Value;
                            commandDefault.Parameters["@OPENING_BALANCE_PASSIVE"].Value = DBNull.Value;
                            commandDefault.Parameters["@TURNS_DEBIT"].Value = DBNull.Value;
                            commandDefault.Parameters["@TURNS_CREDIT"].Value = DBNull.Value;
                            commandDefault.Parameters["@OUTGOING_BALANCE_ACTIVE"].Value = DBNull.Value;
                            commandDefault.Parameters["@OUTGOING_BALANCE_PASSIVE"].Value = DBNull.Value;

                            await commandDefault.ExecuteNonQueryAsync();
                        }
                        catch (SqlException ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                        for (int i = 9; i < arr.Length - 1; i++)
                        {
                            if (arr[i][0].ToString() != "ПО КЛАССУ")
                            {
                                try
                                {
                                    SqlCommand command = new SqlCommand();
                                    command.Connection = connection;
                                    command.CommandText = @"INSERT INTO Class_2 VALUES (@CHECKS, @OPENING_BALANCE_ACTIVE, @OPENING_BALANCE_PASSIVE, @TURNS_DEBIT, @TURNS_CREDIT, @OUTGOING_BALANCE_ACTIVE, @OUTGOING_BALANCE_PASSIVE)";
                                    command.Parameters.Add("@CHECKS", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OPENING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OPENING_BALANCE_PASSIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@TURNS_DEBIT", SqlDbType.NVarChar);
                                    command.Parameters.Add("@TURNS_CREDIT", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OUTGOING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OUTGOING_BALANCE_PASSIVE", SqlDbType.NVarChar);

                                    command.Parameters["@CHECKS"].Value = arr[i][0].ToString();
                                    command.Parameters["@OPENING_BALANCE_ACTIVE"].Value = arr[i][1];
                                    command.Parameters["@OPENING_BALANCE_PASSIVE"].Value = arr[i][2];
                                    command.Parameters["@TURNS_DEBIT"].Value = arr[i][3];
                                    command.Parameters["@TURNS_CREDIT"].Value = arr[i][4];
                                    command.Parameters["@OUTGOING_BALANCE_ACTIVE"].Value = arr[i][5];
                                    command.Parameters["@OUTGOING_BALANCE_PASSIVE"].Value = arr[i][6];

                                    await command.ExecuteNonQueryAsync();
                                }
                                catch (SqlException ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }

                            }
                            else
                            {
                                i++;
                                if (i != arr.Length - 1)
                                {
                                    SqlCommand command = new SqlCommand();
                                    command.Connection = connection;
                                    command.CommandText = @"INSERT INTO Class_2 VALUES (@CHECKS, @OPENING_BALANCE_ACTIVE, @OPENING_BALANCE_PASSIVE, @TURNS_DEBIT, @TURNS_CREDIT, @OUTGOING_BALANCE_ACTIVE, @OUTGOING_BALANCE_PASSIVE)";
                                    command.Parameters.Add("@CHECKS", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OPENING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OPENING_BALANCE_PASSIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@TURNS_DEBIT", SqlDbType.NVarChar);
                                    command.Parameters.Add("@TURNS_CREDIT", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OUTGOING_BALANCE_ACTIVE", SqlDbType.NVarChar);
                                    command.Parameters.Add("@OUTGOING_BALANCE_PASSIVE", SqlDbType.NVarChar);

                                    command.Parameters["@CHECKS"].Value = arr[i][0].ToString();
                                    command.Parameters["@OPENING_BALANCE_ACTIVE"].Value = DBNull.Value;
                                    command.Parameters["@OPENING_BALANCE_PASSIVE"].Value = DBNull.Value;
                                    command.Parameters["@TURNS_DEBIT"].Value = DBNull.Value;
                                    command.Parameters["@TURNS_CREDIT"].Value = DBNull.Value;
                                    command.Parameters["@OUTGOING_BALANCE_ACTIVE"].Value = DBNull.Value;
                                    command.Parameters["@OUTGOING_BALANCE_PASSIVE"].Value = DBNull.Value;

                                    await command.ExecuteNonQueryAsync();
                                }
                            }

                        }
                    }
                }
            }
        }

        private async void toolStripButton1_Click(object sender, EventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionStringToDB))
            {
                await connection.OpenAsync();
                string command = @"SELECT * FROM Class_2";
                SqlDataAdapter adapter = new SqlDataAdapter(command, connection);

                DataSet ds = new DataSet();
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];

            }
        }

        private string[] fileName = null;
        private DataTableCollection tableCollection = null;
        private string connectionStringToDB = "Server=DESKTOP-4HNKCF1;Database=Training OCB 1;Trusted_Connection=True;TrustServerCertificate=True;";

    }
}
