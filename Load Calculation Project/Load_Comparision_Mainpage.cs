﻿using ClosedXML.Excel;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Load_Calculation_Project
{
    public partial class load_comparision_mainpage : Form
    {
      
        public load_comparision_mainpage()
        {
            InitializeComponent();
        }
        DataTableCollection tableCollection;
        private void btnbrowse_Click(object sender, EventArgs e) //Get Browser
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls"})
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    qvexcel.Text = openFileDialog.FileName;
                    using(var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable =(_)=>new ExcelDataTableConfiguration() { UseHeaderRow=true}
                            });
                            tableCollection = result.Tables;
                            qv_s1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                qv_s1.Items.Add(table.TableName);
                                qv_s2.Items.Add(table.TableName);
                                qv_s3.Items.Add(table.TableName);
                                qv_s4.Items.Add(table.TableName);
                                qv_s5.Items.Add(table.TableName);
                                qv_s6.Items.Add(table.TableName);
                                qv_s7.Items.Add(table.TableName);
                                qv_s8.Items.Add(table.TableName);
                            }
                               
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }


        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", Multiselect = false })

            {

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    DataTable dt_FCT = new DataTable();
                    fatigue_ct.Text = ofd.FileName;
                    using (XLWorkbook workbook = new XLWorkbook(ofd.FileName))
                    {
                        bool isFirstRow = true;
                        var rows = workbook.Worksheet(1).RowsUsed();
                        foreach (var row in rows)
                        {
                            if (isFirstRow)
                            {
                                foreach (IXLCell cell in row.Cells())
                                    dt_FCT.Columns.Add(cell.Value.ToString());
                                isFirstRow = false;
                            }
                            else
                            {
                                dt_FCT.Rows.Add();
                                int i = 0;
                                foreach (IXLCell cell in row.Cells())
                                    dt_FCT.Rows[dt_FCT.Rows.Count - 1][i++] = cell.Value.ToString();
                            }
                        }
                        dataGridView1.DataSource = dt_FCT.DefaultView;
                    }
                }
            }
            }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    additioanl_gen_frame.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            additioanl_gen_frame_sheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                additioanl_gen_frame_sheet.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DataTable dt_qv_s1 = tableCollection[qv_s1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s1;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    flop_ref_load.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            flop_ref_load_sheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                flop_ref_load_sheet.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DIVGL1.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            DIVGL1_sheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                DIVGL1_sheet.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DIVGL2.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            qv_s1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                qv_s1.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DIVGL3.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            qv_s1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                qv_s1.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DIVGL4.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            qv_s1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                qv_s1.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    DIVGL5.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            qv_s1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                qv_s1.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }
        }

        private void qv_s2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s2 = tableCollection[qv_s2.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s2;
        }

        private void qv_s3_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s3 = tableCollection[qv_s3.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s3;
        }

        private void qv_s4_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s4 = tableCollection[qv_s4.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s4;
        }

        private void qv_s5_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s5 = tableCollection[qv_s5.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s5;
        }

        private void qv_s7_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s7 = tableCollection[qv_s7.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s7;
        }

        private void qv_s6_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_qv_s7 = tableCollection[qv_s6.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_qv_s7;
        }

        private void additioanl_gen_frame_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            DataTable dt_additioanl_gen_frame_sheet = tableCollection[additioanl_gen_frame_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_additioanl_gen_frame_sheet;
        }

        private void flop_ref_load_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            DataTable dt_flop_ref_load_sheet = tableCollection[flop_ref_load_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_flop_ref_load_sheet;
        }

        private void DIVGL1_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            DataTable dt_DIVGL1_sheet = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_DIVGL1_sheet;
        }

        private void DIVGL2_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_DIVGL2_sheet = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_DIVGL2_sheet;
        }

        private void DIVGL3_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_DIVGL3_sheet = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_DIVGL3_sheet;
        }

        private void DIVGL4_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_DIVGL4_sheet = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_DIVGL4_sheet;
        }

        private void DIVGL5_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt_DIVGL5_sheet = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt_DIVGL5_sheet;
        }

        //public bool flag = false;

        public void form_validation()
        {
            if (String.IsNullOrEmpty(bartcode.Text))
            {
                MessageBox.Show("Enter BART Name");
      
            }
            else if (String.IsNullOrEmpty(pos1.Text) && String.IsNullOrEmpty(pos2.Text) && String.IsNullOrEmpty(pos3.Text) && String.IsNullOrEmpty(pos4.Text) && String.IsNullOrEmpty(pos5.Text))
            {
                MessageBox.Show("Enter Atleast 1 Position");

            }
            else if (String.IsNullOrEmpty(noofyears.Text))
            {
                MessageBox.Show("Enter No of Years");

            }
            else if (String.IsNullOrEmpty(qvexcel.Text))
            {
                MessageBox.Show("Import QuickerView Excel");
                
            }
            else if (String.IsNullOrEmpty(additioanl_gen_frame.Text))
            {
                MessageBox.Show("Import Additional Gentrator Frame Excel");
            }
            else if (String.IsNullOrEmpty(flop_ref_load.Text))
            {
                MessageBox.Show("Import Flop Reference Load Excel");
             
            }
            else if (String.IsNullOrEmpty(fatigue_ct.Text))
            {
                MessageBox.Show("Import Fatique Excel");
          
            }
            else if (String.IsNullOrEmpty(DIVGL1.Text) && String.IsNullOrEmpty(DIVGL2.Text) && String.IsNullOrEmpty(DIVGL3.Text) && String.IsNullOrEmpty(DIVGL4.Text) && String.IsNullOrEmpty(DIVGL5.Text))
            {
                MessageBox.Show("Import Atleast 1 DIVGL Excel");
            }
            else
            {
                MessageBox.Show("Something Went Wrong!");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //LOGIC IMPLEMENTATION
            if (String.IsNullOrEmpty(bartcode.Text) || String.IsNullOrEmpty(pos1.Text) || String.IsNullOrEmpty(qvexcel.Text)
                || String.IsNullOrEmpty(qvexcel.Text) || String.IsNullOrEmpty(additioanl_gen_frame.Text) 
                || String.IsNullOrEmpty(flop_ref_load.Text) || String.IsNullOrEmpty(fatigue_ct.Text) || String.IsNullOrEmpty(DIVGL1.Text))
            {
                form_validation();
            }
            else
            {
                MessageBox.Show("Welcome");
            }
       
            
            
            
            

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
    }
}