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

                    for (int i = 3; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i][0].ToString() != "")
                        {
                            users.Add(new userData
                            {
                                Id = dt.Rows[i][0].ToString(),
                                Name = dt.Rows[i][1].ToString(),
                                WorkingHours1 = dt.Rows[i][3].ToString().Replace(":", ""),
                                WorkingHours2 = dt.Rows[i][4].ToString().Replace(":", ""),
                                WorkingHours3 = dt.Rows[i][5].ToString().Replace(":", ""),
                                WorkingHours4 = dt.Rows[i][6].ToString().Replace(":", ""),
                                WorkingHours5 = dt.Rows[i][7].ToString().Replace(":", ""),
                                WorkingHours6 = dt.Rows[i][8].ToString().Replace(":", ""),
                                WorkingHours7 = dt.Rows[i][9].ToString().Replace(":", ""),
                                WorkingHours8 = dt.Rows[i][10].ToString().Replace(":", ""),
                                WorkingHours9 = dt.Rows[i][11].ToString().Replace(":", ""),
                                WorkingHours10 = dt.Rows[i][12].ToString().Replace(":", ""),
                                WorkingHours11 = dt.Rows[i][13].ToString().Replace(":", ""),
                                WorkingHours12 = dt.Rows[i][14].ToString().Replace(":", ""),
                                WorkingHours13 = dt.Rows[i][15].ToString().Replace(":", ""),
                                WorkingHours14 = dt.Rows[i][16].ToString().Replace(":", ""),
                                WorkingHours15 = dt.Rows[i][17].ToString().Replace(":", ""),
                                WorkingHours16 = dt.Rows[i][18].ToString().Replace(":", ""),
                                WorkingHours17 = dt.Rows[i][19].ToString().Replace(":", ""),
                                WorkingHours18 = dt.Rows[i][20].ToString().Replace(":", ""),
                                WorkingHours19 = dt.Rows[i][21].ToString().Replace(":", ""),
                                WorkingHours20 = dt.Rows[i][22].ToString().Replace(":", ""),
                                WorkingHours21 = dt.Rows[i][23].ToString().Replace(":", ""),
                                WorkingHours22 = dt.Rows[i][24].ToString().Replace(":", ""),
                                WorkingHours23 = dt.Rows[i][25].ToString().Replace(":", ""),
                                WorkingHours24 = dt.Rows[i][26].ToString().Replace(":", ""),
                                WorkingHours25 = dt.Rows[i][27].ToString().Replace(":", ""),
                                WorkingHours26 = dt.Rows[i][28].ToString().Replace(":", ""),
                                WorkingHours27 = dt.Rows[i][29].ToString().Replace(":", ""),
                                WorkingHours28 = dt.Rows[i][30].ToString().Replace(":", ""),
                                WorkingHours29 = dt.Rows[i][31].ToString().Replace(":", ""),
                                WorkingHours30 = dt.Rows[i][32].ToString().Replace(":", ""),
                                WorkingHours31 = dt.Rows[i][33].ToString().Replace(":", "")
                            });

                        }
                    }
                    foreach (var user in users)
                    {
                        //Declaracion de variables
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

                        //Logica

                        if (splitedValue1.Length > 1)
                        {
                            if (splitedValue1[0] != "" && splitedValue1[1] != "")
                            {
                                int valueToEval1 = Convert.ToInt32(splitedValue1[0]);
                                int valueToEval2 = Convert.ToInt32(splitedValue1[1]);
                                if (valueToEval1 <= 0530 || valueToEval2 >= 1130)
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
                            valueToDraw1 = "I";
                        }

                        if (splitedValue2.Length > 1)
                        {
                            if (splitedValue2[0] != "" && splitedValue2[1] != "")
                            {
                                int valueToEval1 = Convert.ToInt32(splitedValue2[0]);
                                int valueToEval2 = Convert.ToInt32(splitedValue2[1]);
                                if (valueToEval1 <= 0530 || valueToEval2 >= 1130)
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
                            valueToDraw2 = "I";
                        }

                        if (splitedValue3.Length > 1)
                        {
                            if (splitedValue3[0] != "" && splitedValue3[1] != "")
                            {
                                int valueToEval1 = Convert.ToInt32(splitedValue3[0]);
                                int valueToEval2 = Convert.ToInt32(splitedValue3[1]);
                                if (valueToEval1 <= 0530 || valueToEval2 >= 1130)
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
                            valueToDraw3 = "I";
                        }

                        if (splitedValue4.Length > 1)
                        {
                            if (splitedValue4[0] != "" && splitedValue4[1] != "")
                            {
                                int valueToEval1 = Convert.ToInt32(splitedValue4[0]);
                                int valueToEval2 = Convert.ToInt32(splitedValue4[1]);
                                if (valueToEval1 <= 0530 || valueToEval2 >= 1130)
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
                            valueToDraw4 = "I";
                        }

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
    }
}
