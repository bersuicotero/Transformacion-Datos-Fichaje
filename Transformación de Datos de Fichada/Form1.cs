using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Ookii.Dialogs.WinForms;
using SpreadsheetLight;

namespace Transformación_de_Datos_de_Fichada
{
    public partial class Form1 : Form
    {

        List<String> Days = new List<String>();
        List<String> numberOfDays = new List<String>();
        List<userData> users = new List<userData>();
        List<userPresenteeism> presenteeisms = new List<userPresenteeism>();


        public Form1()
        {
            InitializeComponent();
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                filePath = file.FileName;
                txtFile.Text = filePath;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0)
                {
                    DataSet dataExcel = GenerarDataSet(filePath);
                    DataTable dt = dataExcel.Tables[0];


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

                    createUserData(dt);
                    createDataToDraw(dt);
                    DataSet dataSet = GetDataTable(presenteeisms);
                    DataTable dtToDraw = dataSet.Tables[0];
                    dataGridView1.DataSource = dtToDraw;
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

        public DataSet GetDataTable(List<userPresenteeism> users)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");

            for (int i = 2; i < 34; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            foreach (var user in users)
            {
                DataRow dr = dt.NewRow();
                dr["Id"] = user.Id;
                dr["Name"] = user.Name;
                int index = 2;
                foreach (string valuetodraw in user.presenteeism)
                {

                    dr[index] = valuetodraw;
                    index++;
                }
                dt.Rows.Add(dr);
            }
            ds.Tables.Add(dt);
            return ds;
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

        public void createUserData(DataTable dt)
        {
            for (int i = 3; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString() != "")
                {

                    userData userdatanew = new userData();

                    userdatanew.Id = dt.Rows[i][0].ToString();
                    userdatanew.Name = dt.Rows[i][1].ToString();


                    for (int j = 3; j < 34; j++)
                    {
                        userdatanew.WorkingHours.Add(dt.Columns.Count > j ? dt.Rows[i][j].ToString().Replace(":", "") : "");
                    }

                    users.Add(userdatanew);

                }
            }
        }

        #region createDataToDraw
        public void createDataToDraw(DataTable dt)
        {
            //Acá es donde arma toda la logica de que valores debe dibujar en el excel de salida
            //Dependiendo de el horario de la fichada y si existe o no alguna fichada
            foreach (var user in users)
            {
                //Declaracion de variables
                #region
                userPresenteeism userpresnew = new userPresenteeism();
                userpresnew.Id = user.Id;
                userpresnew.Name = user.Name;

                foreach (string whour in user.WorkingHours)
                {

                    string[] splitedValue = whour.Split('\n');
                    string valuetodraw = string.Empty;

                    if (splitedValue.Length > 1)
                    {
                        if (splitedValue[0] != "" && splitedValue[1] != "")
                        {
                            int valueToEval1 = Convert.ToInt32(splitedValue[0]);
                            int valueToEval2 = Convert.ToInt32(splitedValue[1]);
                            if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                            {
                                valuetodraw = "P";
                            }
                            else
                            {
                                valuetodraw = "I";
                            }
                        }
                        else if (splitedValue[0] == "" || splitedValue[1] == "")
                        {
                            valuetodraw = "E";
                        }
                    }
                    else
                    {
                        if (numberOfDays.Count >= dt.Columns.Count)
                        {

                            valuetodraw = "I";
                        }
                        else
                        {

                            valuetodraw = "";
                        }
                    }
                    userpresnew.presenteeism.Add(valuetodraw);
                }
                presenteeisms.Add(userpresnew);

                #endregion
            }
        }


        #endregion

        private void btnProcessFile_Click(object sender, EventArgs e)
        {

            string path = textBox2.Text + "\\" +"tabla.xlsx";

            SLDocument doc = new SLDocument();
            int iNDays = 3;
            foreach (string nd in numberOfDays)
            {
                doc.SetCellValue(1, iNDays, nd);
                iNDays++;
            }
            int iDays = 3;
            foreach (string d in Days)
            {
                doc.SetCellValue(2, iDays, d);
                iDays++;
            }

            int iRow = 3;
            foreach (userPresenteeism s in presenteeisms)
            {
                doc.SetCellValue(iRow,1,s.Id);
                doc.SetCellValue(iRow,2,s.Name);

                int iCol = 3;
                foreach(string p in s.presenteeism)
                {
                    doc.SetCellValue(iRow, iCol, p);

                    iCol++;
                }
                iRow++;
            }

            doc.SaveAs(path);
        }

        private void btnChooseDirOutput_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.SelectedPath;
            }

        }
    }
}
