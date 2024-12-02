using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace DownloadExcelOledb
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private DataTable GetDataFromSQL(string connectionString, string TableName)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                DateTime StartDate = FromDTpicker.Value;
                DateTime EndDate = ToDTpicker.Value;
                EndDate = EndDate.AddDays(1);



                string query = "";
                string ResultQuerry = "";
                string Shift = comboBox1.SelectedItem.ToString();
                if (Shift == "BOTH")
                {
                    ResultQuerry = CaseForQuerry(query,TableName,StartDate,EndDate);
                    //query = "Select * from " + TableName + " Where MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
                   
                   
                }
                else
                {

                    //query = "Select * from " + TableName + " where Shift = '" + Shift + "' And MfgDate is not Null MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
                   ResultQuerry = CaseForQuerryWithShift(query, TableName, StartDate, EndDate,Shift);
                }
                using (SqlCommand command = new SqlCommand(ResultQuerry, connection))
               //using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataTable.Load(reader);
                    }
                }
            }

            return dataTable;
        }

        private  void WriteDataToExcel(DataTable dataTable, string filePath, string sheetName)
        {
            sheetName = GetValidSheetName(sheetName);

            string connString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();

                OleDbCommand command = new OleDbCommand("CREATE TABLE [" + sheetName + "] (", conn);

                foreach (DataColumn column in dataTable.Columns)
                {
                    command.CommandText += "[" + column.ColumnName + "] TEXT,";
                }

                command.CommandText = command.CommandText.TrimEnd(',') + ")";
                command.ExecuteNonQuery();

                // Insert data into the worksheet
                foreach (DataRow row in dataTable.Rows)
                {
                    command = new OleDbCommand("INSERT INTO [" + sheetName + "$] (", conn);
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        command.CommandText += "[" + column.ColumnName + "],";
                    }

                    command.CommandText = command.CommandText.TrimEnd(',') + ") VALUES (";

                    foreach (var item in row.ItemArray)
                    {
                        command.CommandText += "'" + item.ToString().Replace("'", "''") + "',"; // Handle single quotes in data
                    }

                    command.CommandText = command.CommandText.TrimEnd(',') + ")";
                    command.ExecuteNonQuery();
                }
            }

            
        }

        private static string GetValidSheetName(string name)
        {
            string validName = name.Substring(8) ;
           
            return validName;
        }

        private void button1_Click(object sender, EventArgs e)
        {



            lblMess.Text = "";
            lblMess.Visible = true;
            lblMess.Text="Please Wait...";

            string ConnString = "data source=PMA-L-3;initial catalog=YY81;persist security info=True;user id=sa;password=pmal@123;";
            List<string> tableNames;

            if (checkBox3.Checked)
            {
                tableNames = new List<string>()
                {
                    "SequenceFB_SingleRH_Mig2",
                    "SequenceFB_SingleLH_Mig2",
                    "SequenceFB_DoubleRH_Mig1",
                    "SequenceFB_DoubleLH_Mig1",
                    "SequenceFC_NonLifterDoubleLH_Mig3",
                    "SequenceFC_LifterSingleRH_Mig5",
                    "SequenceFC_LifterDoubleRH_Mig5",
                    "SequenceFC_LifterDoubleLH_Mig5",
                    "SequenceFC_NonLiftersSingleLH_Mig3",
                    "SequenceFC_NonLifterDoubleRH_Mig4",
                };
            }
            else if(checkBox1.Checked)
                {
                tableNames = new List<string>()
                {
                    "SequenceFC_NonLifterDoubleLH_Mig3",
                    "SequenceFC_LifterSingleRH_Mig5",
                    "SequenceFC_LifterDoubleRH_Mig5",
                    "SequenceFC_LifterDoubleLH_Mig5",
                    "SequenceFC_NonLiftersSingleLH_Mig3",
                    "SequenceFC_NonLifterDoubleRH_Mig4",
                };
                }
            else if (checkBox2.Checked)
            {
                tableNames = new List<string>()
                {
                    "SequenceFB_SingleRH_Mig2",
                    "SequenceFB_SingleLH_Mig2",
                    "SequenceFB_DoubleRH_Mig1",
                    "SequenceFB_DoubleLH_Mig1",
                };
            }
            else
            {
                tableNames = new List<string>()
                {
                    "SequenceFB_SingleRH_Mig2",
                    "SequenceFB_SingleLH_Mig2",
                    "SequenceFB_DoubleRH_Mig1",
                    "SequenceFB_DoubleLH_Mig1",
                    "SequenceFC_NonLifterDoubleLH_Mig3",
                    "SequenceFC_LifterSingleRH_Mig5",
                    "SequenceFC_LifterDoubleRH_Mig5",
                    "SequenceFC_LifterDoubleLH_Mig5",
                    "SequenceFC_NonLiftersSingleLH_Mig3",
                    "SequenceFC_NonLifterDoubleRH_Mig4",
                };
            }
            try
            {
                if (comboBox2.SelectedItem.ToString() != "All")
                {
                    string KeepStr = comboBox2.SelectedItem.ToString().Replace("CELL", "");
                    for (int i = 0; i <= tableNames.Count; i++)
                    {
                        if (!tableNames[i].Contains(KeepStr))
                        { tableNames.RemoveAt(i); i--; }
                    }
                }
            }
            catch { }

           


            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        foreach (string tableName in tableNames)
                        {
                            DataTable dataTable = GetDataFromSQL(ConnString, tableName);
                            WriteDataToExcel(dataTable, sfd.FileName, tableName);
                        }
                        lblMess.Text = "Report Generated successfully";
                        MessageBox.Show("Data has been successfully exported to Excel.", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
            comboBox2.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox3.Checked = false;
            checkBox2.Checked = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox3.Checked = false;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
        }

        private string CaseForQuerry(string query, string TableName, DateTime StartDate, DateTime EndDate)
        {
            if (TableName == "SequenceFB_DoubleLH_Mig1" || TableName == "SequenceFB_DoubleRH_Mig1" || TableName == "SequenceFB_SingleLH_Mig2" || TableName == "SequenceFB_SingleRH_Mig2")
            {
                                                                                                                                                                                                                              //query = "Select * from " + TableName + " Where MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
                query = "Select Id,Date, Shift,SequenceNo,InnerRecliner,OuterRecliner,MfgDate,BuiltTktDate,FinalBarCode,EngravingData,InspectionStation as 'Mig Visual InspectionStation' ,CrimpingStatus,SpiralStatus,PrformanceTest, Gaugingstation, Rework,ReworkDone,Reject,RejectAt,Station from " + TableName + " Where MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null and MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }


            else if (TableName == "SequenceFC_NonLifterDoubleLH_Mig3" || TableName == "SequenceFC_NonLifterDoubleRH_Mig4" || TableName == "SequenceFC_NonLiftersSingleLH_Mig3")
            {
                query = "Select Id,Date,Shift,SequenceNo,MfgDate,BuiltTktDate,FinalBarCode,EngravingData,InspectionStation as 'Mig Visual InspectionStation',GaugingStation,Rework,ReworkDone,Reject,RejectAt,Station from " + TableName + " Where MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null and MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }

            else if (TableName == "SequenceFC_LifterDoubleLH_Mig5" || TableName == "SequenceFC_LifterDoubleRH_Mig5" || TableName == "SequenceFC_LifterSingleRH_Mig5")
            {
                query = "Select Id,Date,Shift,SequenceNo,MfgDate,BuiltTktDate,FinalBarCode, EngravingData, InspectionStation as 'Mig Visual InspectionStation' ,FlaringStatus,GaugingStation,Rework,ReworkDone,Reject,RejectAt,Station,MigGauging,Torque1,Torque2,Torque3,Torque4,Torque5,BreakAssyStatus from " + TableName + " Where MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null and MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }
            return query;
        }


        private string CaseForQuerryWithShift(string query, string TableName, DateTime StartDate, DateTime EndDate, string Shift)
        {
            if (TableName == "SequenceFB_DoubleLH_Mig1" || TableName == "SequenceFB_DoubleRH_Mig1" || TableName == "SequenceFB_SingleLH_Mig2" || TableName == "SequenceFB_SingleRH_Mig2")
            {
                query = "Select Id,Date, Shift,SequenceNo,InnerRecliner,OuterRecliner,MfgDate,BuiltTktDate,FinalBarCode,EngravingData,InspectionStation as 'Mig Visual InspectionStation' ,CrimpingStatus,SpiralStatus,PrformanceTest, Gaugingstation, Rework,ReworkDone,Reject,RejectAt,Station from " + TableName + " Where Shift = '" + Shift + "' and MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null and MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }


            else if (TableName == "SequenceFC_NonLifterDoubleLH_Mig3" || TableName == "SequenceFC_NonLifterDoubleRH_Mig4" || TableName == "SequenceFC_NonLiftersSingleLH_Mig3")
            {
                query = "Select Id,Date,Shift,SequenceNo,MfgDate,BuiltTktDate,FinalBarCode,EngravingData,InspectionStation as 'Mig Visual InspectionStation',GaugingStation,Rework,ReworkDone,Reject,RejectAt,Station from " + TableName + " Where Shift = '" + Shift + "' and MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate is not Null and MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }

            else if (TableName == "SequenceFC_LifterDoubleLH_Mig5" || TableName == "SequenceFC_LifterDoubleRH_Mig5" || TableName == "SequenceFC_LifterSingleRH_Mig5")
            {
                query = "Select Id,Date,Shift,SequenceNo,MfgDate,BuiltTktDate,FinalBarCode, EngravingData, InspectionStation as 'Mig Visual InspectionStation' ,FlaringStatus,GaugingStation,Rework,ReworkDone,Reject,RejectAt,Station,MigGauging,Torque1,Torque2,Torque3,Torque4,Torque5,BreakAssyStatus from " + TableName + " Where Shift = '" + Shift + "' and MfgDate is not Null and MfgDate >= '" + StartDate.ToString("MM/dd/yyyy HH:mm:ss") + "' And MfgDate <= '" + EndDate.ToString("MM/dd/yyyy HH:mm:ss") + "'";
            }
            return query;

        }




        }

    }

