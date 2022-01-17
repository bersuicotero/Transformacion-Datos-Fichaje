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
                    createDataToDraw();
                    dataGridView1.DataSource = presenteeisms;
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

        public void createUserData(DataTable dt)
        {
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                if (i == 1)
                {
                    users.Add(new userData
                    {
                        Id = "",
                        Name = "",
                        WorkingHours1 = dt.Columns.Count > 3 ? dt.Rows[1][3].ToString() : "",
                        WorkingHours2 = dt.Columns.Count > 4 ? dt.Rows[1][4].ToString() : "",
                        WorkingHours3 = dt.Columns.Count > 5 ? dt.Rows[1][5].ToString() : "",
                        WorkingHours4 = dt.Columns.Count > 6 ? dt.Rows[1][6].ToString() : "",
                        WorkingHours5 = dt.Columns.Count > 7 ? dt.Rows[1][7].ToString() : "",
                        WorkingHours6 = dt.Columns.Count > 8 ? dt.Rows[1][8].ToString() : "",
                        WorkingHours7 = dt.Columns.Count > 9 ? dt.Rows[1][9].ToString() : "",
                        WorkingHours8 = dt.Columns.Count > 10 ? dt.Rows[1][10].ToString() : "",
                        WorkingHours9 = dt.Columns.Count > 11 ? dt.Rows[1][11].ToString() : "",
                        WorkingHours10 = dt.Columns.Count > 12 ? dt.Rows[1][12].ToString() : "",
                        WorkingHours11 = dt.Columns.Count > 13 ? dt.Rows[1][13].ToString() : "",
                        WorkingHours12 = dt.Columns.Count > 14 ? dt.Rows[1][14].ToString() : "",
                        WorkingHours13 = dt.Columns.Count > 15 ? dt.Rows[1][15].ToString() : "",
                        WorkingHours14 = dt.Columns.Count > 16 ? dt.Rows[1][16].ToString() : "",
                        WorkingHours15 = dt.Columns.Count > 17 ? dt.Rows[1][17].ToString() : "",
                        WorkingHours16 = dt.Columns.Count > 18 ? dt.Rows[1][18].ToString() : "",
                        WorkingHours17 = dt.Columns.Count > 19 ? dt.Rows[1][19].ToString() : "",
                        WorkingHours18 = dt.Columns.Count > 20 ? dt.Rows[1][20].ToString() : "",
                        WorkingHours19 = dt.Columns.Count > 21 ? dt.Rows[1][21].ToString() : "",
                        WorkingHours20 = dt.Columns.Count > 22 ? dt.Rows[1][22]?.ToString() : "",
                        WorkingHours21 = dt.Columns.Count > 23 ? dt.Rows[1][23]?.ToString() : "",
                        WorkingHours22 = dt.Columns.Count > 24 ? dt.Rows[1][24]?.ToString() : "",
                        WorkingHours23 = dt.Columns.Count > 25 ? dt.Rows[1][25]?.ToString() : "",
                        WorkingHours24 = dt.Columns.Count > 26 ? dt.Rows[1][26]?.ToString() : "",
                        WorkingHours25 = dt.Columns.Count > 27 ? dt.Rows[1][27]?.ToString() : "",
                        WorkingHours26 = dt.Columns.Count > 28 ? dt.Rows[1][28]?.ToString() : "",
                        WorkingHours27 = dt.Columns.Count > 29 ? dt.Rows[1][29]?.ToString() : "",
                        WorkingHours28 = dt.Columns.Count > 30 ? dt.Rows[1][30]?.ToString() : "",
                        WorkingHours29 = dt.Columns.Count > 31 ? dt.Rows[1][31]?.ToString() : "",
                        WorkingHours30 = dt.Columns.Count > 32 ? dt.Rows[1][32]?.ToString() : "",
                        WorkingHours31 = dt.Columns.Count > 33 ? dt.Rows[1][33]?.ToString() : ""
                    });
                }
                if (i == 2)
                {
                    users.Add(new userData
                    {
                        Id = "",
                        Name = "",
                        WorkingHours1 = dt.Columns.Count > 3 ? dt.Rows[2][3].ToString() : "",
                        WorkingHours2 = dt.Columns.Count > 4 ? dt.Rows[2][4].ToString() : "",
                        WorkingHours3 = dt.Columns.Count > 5 ? dt.Rows[2][5].ToString() : "",
                        WorkingHours4 = dt.Columns.Count > 6 ? dt.Rows[2][6].ToString() : "",
                        WorkingHours5 = dt.Columns.Count > 7 ? dt.Rows[2][7].ToString() : "",
                        WorkingHours6 = dt.Columns.Count > 8 ? dt.Rows[2][8].ToString() : "",
                        WorkingHours7 = dt.Columns.Count > 9 ? dt.Rows[2][9].ToString() : "",
                        WorkingHours8 = dt.Columns.Count > 10 ? dt.Rows[2][10].ToString() : "",
                        WorkingHours9 = dt.Columns.Count > 11 ? dt.Rows[2][11].ToString() : "",
                        WorkingHours10 = dt.Columns.Count > 12 ? dt.Rows[2][12].ToString() : "",
                        WorkingHours11 = dt.Columns.Count > 13 ? dt.Rows[2][13].ToString() : "",
                        WorkingHours12 = dt.Columns.Count > 14 ? dt.Rows[2][14].ToString() : "",
                        WorkingHours13 = dt.Columns.Count > 15 ? dt.Rows[2][15].ToString() : "",
                        WorkingHours14 = dt.Columns.Count > 16 ? dt.Rows[2][16].ToString() : "",
                        WorkingHours15 = dt.Columns.Count > 17 ? dt.Rows[2][17].ToString() : "",
                        WorkingHours16 = dt.Columns.Count > 18 ? dt.Rows[2][18].ToString() : "",
                        WorkingHours17 = dt.Columns.Count > 19 ? dt.Rows[2][19].ToString() : "",
                        WorkingHours18 = dt.Columns.Count > 20 ? dt.Rows[2][20].ToString() : "",
                        WorkingHours19 = dt.Columns.Count > 21 ? dt.Rows[2][21].ToString() : "",
                        WorkingHours20 = dt.Columns.Count > 22 ? dt.Rows[2][22]?.ToString() : "",
                        WorkingHours21 = dt.Columns.Count > 23 ? dt.Rows[2][23]?.ToString() : "",
                        WorkingHours22 = dt.Columns.Count > 24 ? dt.Rows[2][24]?.ToString() : "",
                        WorkingHours23 = dt.Columns.Count > 25 ? dt.Rows[2][25]?.ToString() : "",
                        WorkingHours24 = dt.Columns.Count > 26 ? dt.Rows[2][26]?.ToString() : "",
                        WorkingHours25 = dt.Columns.Count > 27 ? dt.Rows[2][27]?.ToString() : "",
                        WorkingHours26 = dt.Columns.Count > 28 ? dt.Rows[2][28]?.ToString() : "",
                        WorkingHours27 = dt.Columns.Count > 29 ? dt.Rows[2][29]?.ToString() : "",
                        WorkingHours28 = dt.Columns.Count > 30 ? dt.Rows[2][30]?.ToString() : "",
                        WorkingHours29 = dt.Columns.Count > 31 ? dt.Rows[2][31]?.ToString() : "",
                        WorkingHours30 = dt.Columns.Count > 32 ? dt.Rows[2][32]?.ToString() : "",
                        WorkingHours31 = dt.Columns.Count > 33 ? dt.Rows[2][33]?.ToString() : ""
                    });
                }
                else if (dt.Rows[i][0].ToString() != "" && i >= 3)
                {
                    users.Add(new userData
                    {
                        Id = dt.Rows[i][0].ToString(),
                        Name = dt.Rows[i][1].ToString(),
                        WorkingHours1 = dt.Columns.Count > 3 ? dt.Rows[i][3].ToString().Replace(":", "") : "",
                        WorkingHours2 = dt.Columns.Count > 4 ? dt.Rows[i][4].ToString().Replace(":", "") : "",
                        WorkingHours3 = dt.Columns.Count > 5 ? dt.Rows[i][5].ToString().Replace(":", "") : "",
                        WorkingHours4 = dt.Columns.Count > 6 ? dt.Rows[i][6].ToString().Replace(":", "") : "",
                        WorkingHours5 = dt.Columns.Count > 7 ? dt.Rows[i][7].ToString().Replace(":", "") : "",
                        WorkingHours6 = dt.Columns.Count > 8 ? dt.Rows[i][8].ToString().Replace(":", "") : "",
                        WorkingHours7 = dt.Columns.Count > 9 ? dt.Rows[i][9].ToString().Replace(":", "") : "",
                        WorkingHours8 = dt.Columns.Count > 10 ? dt.Rows[i][10].ToString().Replace(":", "") : "",
                        WorkingHours9 = dt.Columns.Count > 11 ? dt.Rows[i][11].ToString().Replace(":", "") : "",
                        WorkingHours10 = dt.Columns.Count > 12 ? dt.Rows[i][12].ToString().Replace(":", "") : "",
                        WorkingHours11 = dt.Columns.Count > 13 ? dt.Rows[i][13].ToString().Replace(":", "") : "",
                        WorkingHours12 = dt.Columns.Count > 14 ? dt.Rows[i][14].ToString().Replace(":", "") : "",
                        WorkingHours13 = dt.Columns.Count > 15 ? dt.Rows[i][15].ToString().Replace(":", "") : "",
                        WorkingHours14 = dt.Columns.Count > 16 ? dt.Rows[i][16].ToString().Replace(":", "") : "",
                        WorkingHours15 = dt.Columns.Count > 17 ? dt.Rows[i][17].ToString().Replace(":", "") : "",
                        WorkingHours16 = dt.Columns.Count > 18 ? dt.Rows[i][18].ToString().Replace(":", "") : "",
                        WorkingHours17 = dt.Columns.Count > 19 ? dt.Rows[i][19].ToString().Replace(":", "") : "",
                        WorkingHours18 = dt.Columns.Count > 20 ? dt.Rows[i][20].ToString().Replace(":", "") : "",
                        WorkingHours19 = dt.Columns.Count > 21 ? dt.Rows[i][21].ToString().Replace(":", "") : "",
                        WorkingHours20 = dt.Columns.Count > 22 ? dt.Rows[i][22]?.ToString().Replace(":", "") : "",
                        WorkingHours21 = dt.Columns.Count > 23 ? dt.Rows[i][23]?.ToString().Replace(":", "") : "",
                        WorkingHours22 = dt.Columns.Count > 24 ? dt.Rows[i][24].ToString().Replace(":", "") : "",
                        WorkingHours23 = dt.Columns.Count > 25 ? dt.Rows[i][25].ToString().Replace(":", "") : "",
                        WorkingHours24 = dt.Columns.Count > 26 ? dt.Rows[i][26].ToString().Replace(":", "") : "",
                        WorkingHours25 = dt.Columns.Count > 27 ? dt.Rows[i][27].ToString().Replace(":", "") : "",
                        WorkingHours26 = dt.Columns.Count > 28 ? dt.Rows[i][28].ToString().Replace(":", "") : "",
                        WorkingHours27 = dt.Columns.Count > 29 ? dt.Rows[i][29].ToString().Replace(":", "") : "",
                        WorkingHours28 = dt.Columns.Count > 30 ? dt.Rows[i][30].ToString().Replace(":", "") : "",
                        WorkingHours29 = dt.Columns.Count > 31 ? dt.Rows[i][31].ToString().Replace(":", "") : "",
                        WorkingHours30 = dt.Columns.Count > 32 ? dt.Rows[i][32].ToString().Replace(":", "") : "",
                        WorkingHours31 = dt.Columns.Count > 33 ? dt.Rows[i][33].ToString().Replace(":", "") : ""
                    });
                }
            }
        }

        #region createDataToDraw
        public void createDataToDraw()
        {
            //Acá es donde arma toda la logica de que valores debe dibujar en el excel de salida
            //Dependiendo de el horario de la fichada y si existe o no alguna fichada
            foreach (var user in users)
            {
                //Declaracion de variables
                #region
                string[] splitedValue1 = user.WorkingHours1.Split('\n');
                string[] splitedValue2 = user.WorkingHours2.Split('\n');
                string[] splitedValue3 = user.WorkingHours3.Split('\n');
                string[] splitedValue4 = user.WorkingHours4.Split('\n');
                string[] splitedValue5 = user.WorkingHours5.Split('\n');
                string[] splitedValue6 = user.WorkingHours6.Split('\n');
                string[] splitedValue7 = user.WorkingHours7.Split('\n');
                string[] splitedValue8 = user.WorkingHours8.Split('\n');
                string[] splitedValue9 = user.WorkingHours9.Split('\n');
                string[] splitedValue10 = user.WorkingHours10.Split('\n');
                string[] splitedValue11 = user.WorkingHours11.Split('\n');
                string[] splitedValue12 = user.WorkingHours12.Split('\n');
                string[] splitedValue13 = user.WorkingHours13.Split('\n');
                string[] splitedValue14 = user.WorkingHours14.Split('\n');
                string[] splitedValue15 = user.WorkingHours15.Split('\n');
                string[] splitedValue16 = user.WorkingHours16.Split('\n');
                string[] splitedValue17 = user.WorkingHours17.Split('\n');
                string[] splitedValue18 = user.WorkingHours18.Split('\n');
                string[] splitedValue19 = user.WorkingHours19.Split('\n');
                string[] splitedValue20 = user.WorkingHours20.Split('\n');
                string[] splitedValue21 = user.WorkingHours21.Split('\n');
                string[] splitedValue22 = user.WorkingHours22.Split('\n');
                string[] splitedValue23 = user.WorkingHours23.Split('\n');
                string[] splitedValue24 = user.WorkingHours24.Split('\n');
                string[] splitedValue25 = user.WorkingHours25.Split('\n');
                string[] splitedValue26 = user.WorkingHours26.Split('\n');
                string[] splitedValue27 = user.WorkingHours27.Split('\n');
                string[] splitedValue28 = user.WorkingHours28.Split('\n');
                string[] splitedValue29 = user.WorkingHours29.Split('\n');
                string[] splitedValue30 = user.WorkingHours30.Split('\n');
                string[] splitedValue31 = user.WorkingHours31.Split('\n');

                string valueToDraw1 = "";
                string valueToDraw2 = "";
                string valueToDraw3 = "";
                string valueToDraw4 = "";
                string valueToDraw5 = "";
                string valueToDraw6 = "";
                string valueToDraw7 = "";
                string valueToDraw8 = "";
                string valueToDraw9 = "";
                string valueToDraw10 = "";
                string valueToDraw11 = "";
                string valueToDraw12 = "";
                string valueToDraw13 = "";
                string valueToDraw14 = "";
                string valueToDraw15 = "";
                string valueToDraw16 = "";
                string valueToDraw17 = "";
                string valueToDraw18 = "";
                string valueToDraw19 = "";
                string valueToDraw20 = "";
                string valueToDraw21 = "";
                string valueToDraw22 = "";
                string valueToDraw23 = "";
                string valueToDraw24 = "";
                string valueToDraw25 = "";
                string valueToDraw26 = "";
                string valueToDraw27 = "";
                string valueToDraw28 = "";
                string valueToDraw29 = "";
                string valueToDraw30 = "";
                string valueToDraw31 = "";
                #endregion

                //Logica presentismo
                #region Logica IF
                if (splitedValue1.Length > 1)
                {
                    if (splitedValue1[0] != "" && splitedValue1[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue1[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue1[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw1 = "P";
                        }
                        else
                        {
                            valueToDraw1 = "I";
                        }
                    }
                    else if (splitedValue1[0] == "" || splitedValue1[1] == "")
                    {
                        valueToDraw1 = "E";
                    }

                }
                else
                {
                    if (numberOfDays.Count >= 1)
                    {

                        valueToDraw1 = "I";
                    }
                    else
                    {

                        valueToDraw1 = "";
                    }
                }

                if (splitedValue2.Length > 1)
                {
                    if (splitedValue2[0] != "" && splitedValue2[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue2[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue2[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw2 = "P";
                        }
                        else
                        {
                            valueToDraw2 = "I";
                        }
                    }
                    else if (splitedValue2[0] == "" || splitedValue2[1] == "")
                    {
                        valueToDraw2 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 2)
                    {
                        valueToDraw2 = "I";
                    }
                    else
                    {
                        valueToDraw2 = "";
                    }
                }

                if (splitedValue3.Length > 1)
                {
                    if (splitedValue3[0] != "" && splitedValue3[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue3[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue3[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw3 = "P";
                        }
                        else
                        {
                            valueToDraw3 = "I";
                        }
                    }
                    else if (splitedValue3[0] == "" || splitedValue3[1] == "")
                    {
                        valueToDraw3 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 3)
                    {
                        valueToDraw3 = "I";
                    }
                    else
                    {
                        valueToDraw3 = "";
                    }
                }

                if (splitedValue4.Length > 1)
                {
                    if (splitedValue4[0] != "" && splitedValue4[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue4[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue4[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw4 = "P";
                        }
                        else
                        {
                            valueToDraw4 = "I";
                        }
                    }
                    else if (splitedValue4[0] == "" || splitedValue4[1] == "")
                    {
                        valueToDraw4 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 4)
                    {
                        valueToDraw4 = "I";
                    }
                    else
                    {
                        valueToDraw4 = "";
                    }
                }

                if (splitedValue5.Length > 1)
                {
                    if (splitedValue5[0] != "" && splitedValue5[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue5[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue5[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw5 = "P";
                        }
                        else
                        {
                            valueToDraw5 = "I";
                        }
                    }
                    else if (splitedValue5[0] == "" || splitedValue5[1] == "")
                    {
                        valueToDraw5 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 5)
                    {
                        valueToDraw5 = "I";
                    }
                    else
                    {
                        valueToDraw5 = "";
                    }
                }

                if (splitedValue6.Length > 1)
                {
                    if (splitedValue6[0] != "" && splitedValue6[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue6[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue6[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw6 = "P";
                        }
                        else
                        {
                            valueToDraw6 = "I";
                        }
                    }
                    else if (splitedValue6[0] == "" || splitedValue6[1] == "")
                    {
                        valueToDraw6 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 6)
                    {
                        valueToDraw6 = "I";
                    }
                    else
                    {
                        valueToDraw6 = "";
                    }
                }

                if (splitedValue7.Length > 1)
                {
                    if (splitedValue7[0] != "" && splitedValue7[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue7[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue7[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw7 = "P";
                        }
                        else
                        {
                            valueToDraw7 = "I";
                        }
                    }
                    else if (splitedValue7[0] == "" || splitedValue7[1] == "")
                    {
                        valueToDraw7 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 7)
                    {
                        valueToDraw7 = "I";
                    }
                    else
                    {
                        valueToDraw7 = "";
                    }
                }

                if (splitedValue8.Length > 1)
                {
                    if (splitedValue8[0] != "" && splitedValue8[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue8[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue8[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw8 = "P";
                        }
                        else
                        {
                            valueToDraw8 = "I";
                        }
                    }
                    else if (splitedValue8[0] == "" || splitedValue8[1] == "")
                    {
                        valueToDraw8 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 8)
                    {
                        valueToDraw8 = "I";
                    }
                    else
                    {
                        valueToDraw8 = "";
                    }
                }

                if (splitedValue9.Length > 1)
                {
                    if (splitedValue9[0] != "" && splitedValue9[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue9[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue9[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw9 = "P";
                        }
                        else
                        {
                            valueToDraw9 = "I";
                        }
                    }
                    else if (splitedValue9[0] == "" || splitedValue9[1] == "")
                    {
                        valueToDraw9 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 9)
                    {
                        valueToDraw9 = "I";
                    }
                    else
                    {
                        valueToDraw9 = "";
                    }
                }

                if (splitedValue10.Length > 1)
                {
                    if (splitedValue10[0] != "" && splitedValue10[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue10[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue10[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw10 = "P";
                        }
                        else
                        {
                            valueToDraw10 = "I";
                        }
                    }
                    else if (splitedValue10[0] == "" || splitedValue10[1] == "")
                    {
                        valueToDraw10 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 10)
                    {
                        valueToDraw10 = "I";
                    }
                    else
                    {
                        valueToDraw10 = "";
                    }
                }

                if (splitedValue11.Length > 1)
                {
                    if (splitedValue11[0] != "" && splitedValue11[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue11[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue11[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw11 = "P";
                        }
                        else
                        {
                            valueToDraw11 = "I";
                        }
                    }
                    else if (splitedValue11[0] == "" || splitedValue11[1] == "")
                    {
                        valueToDraw11 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 11)
                    {
                        valueToDraw11 = "I";
                    }
                    else
                    {
                        valueToDraw11 = "";
                    }
                }

                if (splitedValue12.Length > 1)
                {
                    if (splitedValue12[0] != "" && splitedValue12[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue12[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue12[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw12 = "P";
                        }
                        else
                        {
                            valueToDraw12 = "I";
                        }
                    }
                    else if (splitedValue12[0] == "" || splitedValue12[1] == "")
                    {
                        valueToDraw12 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 12)
                    {
                        valueToDraw12 = "I";
                    }
                    else
                    {
                        valueToDraw12 = "";
                    }
                }

                if (splitedValue13.Length > 1)
                {
                    if (splitedValue13[0] != "" && splitedValue13[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue13[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue13[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw13 = "P";
                        }
                        else
                        {
                            valueToDraw13 = "I";
                        }
                    }
                    else if (splitedValue13[0] == "" || splitedValue13[1] == "")
                    {
                        valueToDraw13 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 13)
                    {
                        valueToDraw13 = "I";
                    }
                    else
                    {
                        valueToDraw13 = "";
                    }
                }

                if (splitedValue14.Length > 1)
                {
                    if (splitedValue14[0] != "" && splitedValue14[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue14[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue14[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw14 = "P";
                        }
                        else
                        {
                            valueToDraw14 = "I";
                        }
                    }
                    else if (splitedValue14[0] == "" || splitedValue14[1] == "")
                    {
                        valueToDraw14 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 14)
                    {
                        valueToDraw14 = "I";
                    }
                    else
                    {
                        valueToDraw14 = "";
                    }
                }

                if (splitedValue15.Length > 1)
                {
                    if (splitedValue15[0] != "" && splitedValue15[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue15[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue15[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw15 = "P";
                        }
                        else
                        {
                            valueToDraw15 = "I";
                        }
                    }
                    else if (splitedValue15[0] == "" || splitedValue15[1] == "")
                    {
                        valueToDraw15 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 15)
                    {
                        valueToDraw15 = "I";
                    }
                    else
                    {
                        valueToDraw15 = "";
                    }
                }

                if (splitedValue16.Length > 1)
                {
                    if (splitedValue16[0] != "" && splitedValue16[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue16[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue16[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw16 = "P";
                        }
                        else
                        {
                            valueToDraw16 = "I";
                        }
                    }
                    else if (splitedValue16[0] == "" || splitedValue16[1] == "")
                    {
                        valueToDraw16 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 16)
                    {
                        valueToDraw16 = "I";
                    }
                    else
                    {
                        valueToDraw16 = "";
                    }
                }

                if (splitedValue17.Length > 1)
                {
                    if (splitedValue17[0] != "" && splitedValue17[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue17[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue17[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw17 = "P";
                        }
                        else
                        {
                            valueToDraw17 = "I";
                        }
                    }
                    else if (splitedValue17[0] == "" || splitedValue17[1] == "")
                    {
                        valueToDraw17 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 17)
                    {
                        valueToDraw17 = "I";
                    }
                    else
                    {
                        valueToDraw17 = "";
                    }
                }

                if (splitedValue18.Length > 1)
                {
                    if (splitedValue18[0] != "" && splitedValue18[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue18[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue18[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw18 = "P";
                        }
                        else
                        {
                            valueToDraw18 = "I";
                        }
                    }
                    else if (splitedValue18[0] == "" || splitedValue18[1] == "")
                    {
                        valueToDraw18 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 18)
                    {
                        valueToDraw18 = "I";
                    }
                    else
                    {
                        valueToDraw18 = "";
                    }
                }

                if (splitedValue19.Length > 1)
                {
                    if (splitedValue19[0] != "" && splitedValue19[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue19[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue19[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw19 = "P";
                        }
                        else
                        {
                            valueToDraw19 = "I";
                        }
                    }
                    else if (splitedValue19[0] == "" || splitedValue19[1] == "")
                    {
                        valueToDraw19 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 19)
                    {
                        valueToDraw19 = "I";
                    }
                    else
                    {
                        valueToDraw19 = "";
                    }
                }

                if (splitedValue20.Length > 1)
                {
                    if (splitedValue20[0] != "" && splitedValue20[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue20[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue20[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw20 = "P";
                        }
                        else
                        {
                            valueToDraw20 = "I";
                        }
                    }
                    else if (splitedValue20[0] == "" || splitedValue20[1] == "")
                    {
                        valueToDraw20 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 20)
                    {
                        valueToDraw20 = "I";

                    }
                    else
                    {
                        valueToDraw20 = "";
                    }
                }

                if (splitedValue21.Length > 1)
                {
                    if (splitedValue21[0] != "" && splitedValue21[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue21[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue21[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw21 = "P";
                        }
                        else
                        {
                            valueToDraw21 = "I";
                        }
                    }
                    else if (splitedValue21[0] == "" || splitedValue21[1] == "")
                    {
                        valueToDraw21 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 21)
                    {
                        valueToDraw21 = "I";
                    }
                    else
                    {
                        valueToDraw21 = "";
                    }
                }

                if (splitedValue22.Length > 1)
                {
                    if (splitedValue22[0] != "" && splitedValue22[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue22[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue22[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw22 = "P";
                        }
                        else
                        {
                            valueToDraw22 = "I";
                        }
                    }
                    else if (splitedValue22[0] == "" || splitedValue22[1] == "")
                    {
                        valueToDraw22 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 22)
                    {
                        valueToDraw22 = "I";
                    }
                    else
                    {
                        valueToDraw22 = "";
                    }
                }

                if (splitedValue23.Length > 1)
                {
                    if (splitedValue23[0] != "" && splitedValue23[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue23[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue23[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw23 = "P";
                        }
                        else
                        {
                            valueToDraw23 = "I";
                        }
                    }
                    else if (splitedValue23[0] == "" || splitedValue23[1] == "")
                    {
                        valueToDraw23 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 23)
                    {
                        valueToDraw23 = "I";
                    }
                    else
                    {
                        valueToDraw23 = "";
                    }
                }

                if (splitedValue24.Length > 1)
                {
                    if (splitedValue24[0] != "" && splitedValue24[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue24[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue24[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw24 = "P";
                        }
                        else
                        {
                            valueToDraw24 = "I";
                        }
                    }
                    else if (splitedValue24[0] == "" || splitedValue24[1] == "")
                    {
                        valueToDraw24 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 24)
                    {
                        valueToDraw24 = "I";
                    }
                    else
                    {
                        valueToDraw24 = "";
                    }
                }

                if (splitedValue25.Length > 1)
                {
                    if (splitedValue25[0] != "" && splitedValue25[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue25[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue25[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw25 = "P";
                        }
                        else
                        {
                            valueToDraw25 = "I";
                        }
                    }
                    else if (splitedValue25[0] == "" || splitedValue25[1] == "")
                    {
                        valueToDraw25 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 25)
                    {
                        valueToDraw25 = "I";
                    }
                    else
                    {
                        valueToDraw25 = "";
                    }
                }

                if (splitedValue26.Length > 1)
                {
                    if (splitedValue26[0] != "" && splitedValue26[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue26[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue26[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw26 = "P";
                        }
                        else
                        {
                            valueToDraw26 = "I";
                        }
                    }
                    else if (splitedValue26[0] == "" || splitedValue26[1] == "")
                    {
                        valueToDraw26 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 26)
                    {
                        valueToDraw26 = "I";
                    }
                    else
                    {
                        valueToDraw26 = "";
                    }
                }

                if (splitedValue27.Length > 1)
                {
                    if (splitedValue27[0] != "" && splitedValue27[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue27[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue27[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw27 = "P";
                        }
                        else
                        {
                            valueToDraw27 = "I";
                        }
                    }
                    else if (splitedValue27[0] == "" || splitedValue27[1] == "")
                    {
                        valueToDraw27 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 27)
                    {
                        valueToDraw27 = "I";
                    }
                    else
                    {
                        valueToDraw27 = "";
                    }
                }

                if (splitedValue28.Length > 1)
                {
                    if (splitedValue28[0] != "" && splitedValue28[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue28[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue28[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw28 = "P";
                        }
                        else
                        {
                            valueToDraw28 = "I";
                        }
                    }
                    else if (splitedValue28[0] == "" || splitedValue28[1] == "")
                    {
                        valueToDraw28 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 28)
                    {
                        valueToDraw28 = "I";
                    }
                    else
                    {
                        valueToDraw28 = "";
                    }
                }

                if (splitedValue29.Length > 1)
                {
                    if (splitedValue29[0] != "" && splitedValue29[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue29[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue29[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw29 = "P";
                        }
                        else
                        {
                            valueToDraw29 = "I";
                        }
                    }
                    else if (splitedValue29[0] == "" || splitedValue29[1] == "")
                    {
                        valueToDraw29 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 29)
                    {
                        valueToDraw29 = "I";
                    }
                    else
                    {
                        valueToDraw29 = "";
                    }
                }

                if (splitedValue30.Length > 1)
                {
                    if (splitedValue30[0] != "" && splitedValue30[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue30[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue30[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw30 = "P";
                        }
                        else
                        {
                            valueToDraw30 = "I";
                        }
                    }
                    else if (splitedValue30[0] == "" || splitedValue30[1] == "")
                    {
                        valueToDraw30 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 30)
                    {
                        valueToDraw30 = "I";
                    }
                    else
                    {
                        valueToDraw30 = "";
                    }
                }

                if (splitedValue31.Length > 1)
                {
                    if (splitedValue31[0] != "" && splitedValue31[1] != "")
                    {
                        int valueToEval1 = Convert.ToInt32(splitedValue31[0]);
                        int valueToEval2 = Convert.ToInt32(splitedValue31[1]);
                        if (valueToEval1 <= 0530 && valueToEval2 >= 1130)
                        {
                            valueToDraw31 = "P";
                        }
                        else
                        {
                            valueToDraw31 = "I";
                        }
                    }
                    else if (splitedValue31[0] == "" || splitedValue31[1] == "")
                    {
                        valueToDraw31 = "E";
                    }
                }
                else
                {
                    if (numberOfDays.Count >= 31)
                    {
                        valueToDraw31 = "I";
                    }
                    else
                    {
                        valueToDraw31 = "";
                    }
                }
                #endregion





                //creación del objeto a representare en el excel de salida
                presenteeisms.Add(new userPresenteeism
                {
                    Id = user.Id,
                    Name = user.Name,
                    presenteeism1 = valueToDraw1,
                    presenteeism2 = valueToDraw2,
                    presenteeism3 = valueToDraw3,
                    presenteeism4 = valueToDraw4,
                    presenteeism5 = valueToDraw5,
                    presenteeism6 = valueToDraw6,
                    presenteeism7 = valueToDraw7,
                    presenteeism8 = valueToDraw8,
                    presenteeism9 = valueToDraw9,
                    presenteeism10 = valueToDraw10,
                    presenteeism11 = valueToDraw11,
                    presenteeism12 = valueToDraw12,
                    presenteeism13 = valueToDraw13,
                    presenteeism14 = valueToDraw14,
                    presenteeism15 = valueToDraw15,
                    presenteeism16 = valueToDraw16,
                    presenteeism17 = valueToDraw17,
                    presenteeism18 = valueToDraw18,
                    presenteeism19 = valueToDraw19,
                    presenteeism20 = valueToDraw20,
                    presenteeism21 = valueToDraw21,
                    presenteeism22 = valueToDraw22,
                    presenteeism23 = valueToDraw23,
                    presenteeism24 = valueToDraw24,
                    presenteeism25 = valueToDraw25,
                    presenteeism26 = valueToDraw26,
                    presenteeism27 = valueToDraw27,
                    presenteeism28 = valueToDraw28,
                    presenteeism29 = valueToDraw29,
                    presenteeism30 = valueToDraw30,
                    presenteeism31 = valueToDraw31
                });
            }
        }

        #endregion
    }
}
