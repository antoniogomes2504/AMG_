using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadExcelFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = result.Tables[cboSheet.SelectedIndex];
        }

        DataSet result;
        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook |*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fs);
                    using (var rdr = ExcelReaderFactory.CreateOpenXmlReader(fs))
                    {
                        var conf = new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true //THIS IS WHAT YOU ARE AFTER
                            }
                        };

                        var ds = rdr.AsDataSet(conf); //THIS IS WHERE IT IS USED
                    }
                    result = reader.AsDataSet();
                    cboSheet.Items.Clear();
                    foreach (DataTable dt in result.Tables)
                    cboSheet.Items.Add(dt.TableName);
                    reader.Close();
                }
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
