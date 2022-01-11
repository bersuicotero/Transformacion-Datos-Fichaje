using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace Transformación_de_Datos_de_Fichada
{
    public partial class Form1 : Form
    {
        public class userData
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public List<String> WorkingHours { get; set; }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file 
            if (file.ShowDialog() == DialogResult.OK)
            {
                filePath = file.FileName;
                txtFile.Text = filePath;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0)
                {
                    DataSet dataExcel = GenerarDataSet(filePath);
                    DataTable dt = dataExcel.Tables[0];
                    List<String> Days = new List<String>();
                    List<String> numberOfDays = new List<String>();
                    List<userData> users = new List<userData>();
                    List<String> horas = new List<String>();


                    for (int i = 4; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i][0].ToString() != "")
                        {

                            for (int j = 3; j < dt.Columns.Count; j++)
                            {
                                if (j < dt.Columns.Count)
                                {
                                    horas.AddRange(new [] { dt.Rows[i][j].ToString(), i.ToString()});
                                }
                                else
                                {
                                    j = 0;
                                }
                            }
                            users.Add(new userData
                            {
                                Id = dt.Rows[i][0].ToString(),
                                Name = dt.Rows[i][1].ToString(),
                                WorkingHours = horas
                            });
                            
                        }
                    }

                    for (int i = 3; i < dt.Columns.Count; i++)
                    {
                        if (dt.Rows[1][i].ToString() != "")
                        {
                            numberOfDays.Add(dt.Rows[1][i].ToString());
                        }
                        if (dt.Rows[2][i].ToString() != "")
                        {
                            Days.Add(dt.Rows[2][i].ToString());
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Solo puede ingresar archivos excel con formato .xls");
                }
            }
            else
            {
                MessageBox.Show("Por favor seleccione un archivo con formato .xls");
            }

        }

        public DataSet GenerarDataSet(string filePath)
        {
            OleDbConnection oledbConn = new OleDbConnection();
            try
            {
                if (Path.GetExtension(filePath) == ".xls" || Path.GetExtension(filePath) == ".XLS")
                {
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"");
                }
                else if (Path.GetExtension(filePath) == ".xlsx")
                {
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");
                }

                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand(); ;
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();

                DataTable dbSchema = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dbSchema == null || dbSchema.Rows.Count < 1)
                {
                    throw new Exception("Error: No se puede determinar la primera hoja del excel.");
                }

                string firstSheetName = "Registros$";

                cmd.Connection = oledbConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + firstSheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds);

                return ds;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }
            finally
            {
                oledbConn.Close();
            }
        }
    }
}
