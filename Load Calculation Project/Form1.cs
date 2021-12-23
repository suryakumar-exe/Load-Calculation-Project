using ClosedXML.Excel;
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
    public partial class Form1 : Form
    {
        String file_name = "";
        public Form1()
        {
            InitializeComponent();
        }
        DataTableCollection tableCollection;
        private void btnbrowse_Click(object sender, EventArgs e)
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
                    DataTable dt = new DataTable();
                    file_name = ofd.FileName;
                    using (XLWorkbook workbook = new XLWorkbook(ofd.FileName))
                    {
                        bool isFirstRow = true;
                        var rows = workbook.Worksheet(1).RowsUsed();
                        foreach (var row in rows)
                        {
                            if (isFirstRow)
                            {
                                foreach (IXLCell cell in row.Cells())
                                    dt.Columns.Add(cell.Value.ToString());
                                isFirstRow = false;
                            }
                            else
                            {
                                dt.Rows.Add();
                                int i = 0;
                                foreach (IXLCell cell in row.Cells())
                                    dt.Rows[dt.Rows.Count - 1][i++] = cell.Value.ToString();
                            }
                        }
                        dataGridView1.DataSource = dt.DefaultView;
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
            DataTable dt = tableCollection[qv_s1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
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
            DataTable dt = tableCollection[qv_s2.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void qv_s3_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s3.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void qv_s4_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s4.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void qv_s5_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s5.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void qv_s7_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s7.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void qv_s6_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[qv_s6.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void additioanl_gen_frame_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            DataTable dt = tableCollection[additioanl_gen_frame_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void flop_ref_load_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                DataTable dt = tableCollection[flop_ref_load_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void DIVGL1_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                 DataTable dt = tableCollection[DIVGL1_sheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }
    }
}
