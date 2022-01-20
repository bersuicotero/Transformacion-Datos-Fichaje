using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Ookii.Dialogs.WinForms;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Transformación_de_Datos_de_Fichada
{
    public partial class Form1 : Form
    {

        List<String> Days = new List<String>();
        List<String> numberOfDays = new List<String>();
        List<userData> users = new List<userData>();
        List<userPresenteeism> presenteeisms = new List<userPresenteeism>();
        string perExcel = string.Empty;


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

                txtFile.ReadOnly = true;
                btnChooseFile.Enabled = false;

                if (fileExt.CompareTo(".xls") == 0)
                {
                    //GetDataWithSpreadSheet(filePath);
                    DataSet dataExcel = GenerarDataSet(filePath);
                    DataTable dt = dataExcel.Tables[0];

                    string periodo = dt.Rows[0][2].ToString();
                    string[] periodoParsed = periodo.Split('/');
                    perExcel = periodoParsed[1];


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
                    btnChooseFile.Enabled = true;
                    txtFile.ReadOnly = false;
                    txtFile.Text = null;

                }
            }
            else
            {
                MessageBox.Show("Por favor seleccione un archivo con formato .xls");
                btnChooseFile.Enabled = true;
                txtFile.ReadOnly = false;
                txtFile.Text = null;
            }

        }

        public void GetDataWithSpreadSheet(string path)
        {
            SLDocument doc = new SLDocument(path);




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
                oleda.SelectCommand = cmd;
                oleda.Fill(ds);

                oledbConn.Close();

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

                int i = 1;
                int daysCol = 0;
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
                        if (numberOfDays.Count >= i && splitedValue[0] == "" && Days[daysCol].ToString() != "Do")
                        {

                            valuetodraw = "I";
                        }
                        else
                        {

                            valuetodraw = "";
                        }

                    }
                    i++;
                    daysCol++;
                    userpresnew.presenteeism.Add(valuetodraw);
                }

                presenteeisms.Add(userpresnew);

                #endregion
            }
        }

        private SLThemeSettings BuildTheme()
        {
            SLThemeSettings theme = new SLThemeSettings();
            theme.ThemeName = "RDSColourTheme";
            theme.Light1Color = System.Drawing.Color.White;
            theme.Dark1Color = System.Drawing.Color.Black;
            theme.Light2Color = System.Drawing.Color.Gray;
            theme.Dark2Color = System.Drawing.Color.IndianRed;
            theme.Accent1Color = System.Drawing.Color.Red;
            theme.Accent2Color = System.Drawing.Color.Tomato;
            theme.Accent3Color = System.Drawing.Color.Yellow;
            theme.Accent4Color = System.Drawing.Color.LawnGreen;
            theme.Accent5Color = System.Drawing.Color.DeepSkyBlue;
            theme.Accent6Color = System.Drawing.Color.DarkViolet;
            theme.Hyperlink = System.Drawing.Color.Blue;
            theme.FollowedHyperlinkColor = System.Drawing.Color.Purple;
            return theme;
        }

        public bool exportToExcel(string path)
        {
            try
            {
                SLThemeSettings settings = BuildTheme();
                SLDocument doc = new SLDocument(settings);

                //Se definen los estilos para los headers
                SLStyle headerStyle = doc.CreateStyle();
                headerStyle.Font.FontColor = System.Drawing.Color.White;
                headerStyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

                //Se definen los estilos para los capos de presentismo
                SLStyle errorStyle = doc.CreateStyle();
                SLStyle absenceStyle = doc.CreateStyle();
                SLStyle SunDayStyle = doc.CreateStyle();

                absenceStyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent2Color);
                errorStyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent1Color);
                SunDayStyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Accent5Color);

                //Se setean los headers de las 2 primeras columnas
                doc.SetCellValue(2, 1, "ID");
                doc.SetCellStyle(2, 1, headerStyle);
                doc.SetCellValue(2, 2, "Empleado");
                doc.SetColumnWidth(2, 2, 20);
                doc.SetCellStyle(2, 2, headerStyle);

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
                    doc.SetCellStyle(2, iDays, headerStyle);
                    iDays++;
                }

                int iRow = 3;
                foreach (userPresenteeism s in presenteeisms)
                {
                    doc.SetCellValue(iRow, 1, s.Id);
                    doc.SetCellValue(iRow, 2, s.Name);

                    int iCol = 3;
                    int daysCol = 0;
                    foreach (string p in s.presenteeism)
                    {
                        doc.SetCellValue(iRow, iCol, p);
                        if (p == "E")
                        {
                            doc.SetCellStyle(iRow, iCol, errorStyle);
                        }
                        if (p == "I")
                        {
                            doc.SetCellStyle(iRow, iCol, absenceStyle);
                        }
                        if (p == "" && daysCol < Days.Count )
                        {
                            if (Days[daysCol].ToString() == "Do")
                            {
                                doc.SetCellStyle(iRow, iCol, SunDayStyle);

                            }
                        }
                        daysCol++;
                        iCol++;
                    }
                    iRow++;
                }

                doc.SaveAs(path);
                MessageBox.Show("Se exportó correctamente el archivo" + " " + perExcel + "Resultado.xlsx  en la ruta: " + textBox2.Text);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
        public void unlockForm()
        {
            dataGridView1.DataSource = null;
            numberOfDays.Clear();
            Days.Clear();
            users.Clear();
            presenteeisms.Clear();
            btnChooseFile.Enabled = true;
            txtFile.ReadOnly = false;
            txtFile.Text = null;
            btnChooseDirOutput.Enabled = true;
            textBox2.ReadOnly = false;
            textBox2.Text = null;

        }
        #endregion

        private void btnProcessFile_Click(object sender, EventArgs e)
        {
            try
            {
                string path = textBox2.Text + "\\" + perExcel + "Resultado.xlsx";
                var result = exportToExcel(path);

                if (result == true)
                {
                    unlockForm();
                }
                if (result == false)
                {
                    MessageBox.Show("Algo salió mal, intentelo de nuevo. Si anteriormente exportó el mismo archivo, asegurese de no tenerlo abierto");
                    unlockForm();
                }

            }
            catch (Exception)
            {

                throw;
            }

        }

        private void btnChooseDirOutput_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.SelectedPath;
                textBox2.ReadOnly = true;
                btnChooseDirOutput.Enabled = false;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnProcessFile.Enabled = false;
        }


        private void ReadOnlyChanged(object sender, EventArgs e)
        {
            if (txtFile.ReadOnly == true && textBox2.ReadOnly == true)
            {
                btnProcessFile.Enabled = true;
            }
            else
            {
                btnProcessFile.Enabled = false;
            }
        }
    }
}
