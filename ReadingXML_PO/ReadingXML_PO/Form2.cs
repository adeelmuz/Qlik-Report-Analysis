using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadingXML_PO
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            StreamReader objSR = new StreamReader(@"C:\Users\adeel.muzaffar.SYSTEMS\Desktop\Report Analysis\Extraction_MSP_Direct.txt");
            string strODBC = "";
            string strODBCString = "";
            string strOLEBD = "";
            string strConString = "";
            string strAccessString = "";
            List<string> listTable = new List<string>();
            List<string> listSharePoint = new List<string>();
            List<string> listLet = new List<string>();
            List<string> listXls = new List<string>();
            List<string> listQvw = new List<string>();
            List<string> listCsv = new List<string>();
            List<string> listWrite = new List<string>();
            List<string> listRead = new List<string>();
            

            while (!objSR.EndOfStream)
            {
                string strLine = objSR.ReadLine();
                strLine = strLine.Trim('.').Trim();

                if (strLine.Length>=4)
                if (strLine.Substring(0,4).ToUpper().Equals("ODBC"))
                {
                    strODBC = "ODBC";
                    string[] strFields=strLine.Split(' ');
                    strODBCString = strFields[strFields.Length - 1];

                }

                if (strLine.Length >= 5)
                if (strLine.Substring(0, 5).ToUpper().Equals("OLEDB"))
                {
                    strOLEBD = "OLEDB";
                    string[] strFields=strLine.Split(' ');
                    strConString = strFields[strFields.Length - 1];
                }

                if (strLine.IndexOf(".accdb")>0)
                {
                    string[] strFields = strLine.Split(' ');
                    strAccessString = strFields[strFields.Length - 1];
                }

                //if (strLine.Length >= 4)
                //if (strLine.Substring(0,4).ToUpper().Equals("FROM"))
                //{
                //    if (strLine.IndexOf("https://sharepoint2013") == -1)
                //    {
                //        //string[] strFields = strLine.Split(' ');
                //        //listTable.Add(strFields[strFields.Length - 1]);
                //        listTable.Add(strLine);
                //    }
                //}

                //if (strLine.Length >= 4)
                //if (strLine.Substring(0, 4).ToUpper().Equals("FROM"))
                //{
                //    if (strLine.IndexOf("https://sharepoint2013") > 0)
                //    {
                //        //string[] strFields = strLine.Split(' ');
                //        //listSharePoint.Add(strFields[strFields.Length - 1]);
                //        listSharePoint.Add(strLine);
                //    }
                //}


                if (strLine.Length >= 3)
                    if (strLine.Substring(0, 3).ToUpper().Equals("SQL"))
                    {
                        while(!objSR.EndOfStream)
                        {
                            strLine = objSR.ReadLine();
                            strLine = strLine.Trim('.').Trim();
                            if (strLine.Length >= 4)
                                if (strLine.Substring(0, 4).ToUpper().Equals("FROM"))
                                {
                                    listTable.Add(strLine);
                                    break;
                                }
                        }
                    }

                if (strLine.IndexOf("https://sharepoint2013") > 0)
                {
                    //string[] strFields = strLine.Split(' ');
                    //listSharePoint.Add(strFields[strFields.Length - 1]);
                    listSharePoint.Add(strLine);
                }


                if (strLine.Length >= 3)
                if (strLine.Substring(0, 3).ToUpper().Equals("LET"))
                {
                    listLet.Add(strLine);
                }


                 if (strLine.ToUpper().IndexOf("XLS") > 0)
                 {
                     string[] strFields = strLine.Split(' ');
                     listXls.Add(strFields[strFields.Length - 1]);
                 }

                 if (strLine.ToUpper().IndexOf("QVW") > 0)
                 {
                     string[] strFields = strLine.Split(' ');
                     listQvw.Add(strFields[strFields.Length - 1]);
                 }


                 if (strLine.ToUpper().IndexOf("CSV") > 0)
                 {
                     //string[] strFields = strLine.Split(' ');
                     listCsv.Add(strLine);
                 }
                 
                if (strLine.Length >= 5)
                    if (strLine.Substring(0, 5).ToUpper().Equals("STORE"))
                    {
                        if (strLine.ToUpper().IndexOf("INTO")==-1)
                        {
                            while (!objSR.EndOfStream)
                            {
                                strLine = objSR.ReadLine();
                                strLine = strLine.Trim('.').Trim();
                                if (strLine.ToUpper().IndexOf("INTO") > 0)
                                    break;
                            }
                            
                        }
                        if (strLine.ToUpper().IndexOf("INTO") > 0)
                        {
                            string[] strFields = strLine.Split(' ');
                            if (!strFields[strFields.Length - 2].ToUpper().Equals("INTO"))
                            {
                                listWrite.Add(strFields[strFields.Length - 2]+" "+strFields[strFields.Length - 1]);
                            }
                            else
                            {
                                listWrite.Add(strFields[strFields.Length - 1]);
                            }

                            
                        }
                    }

            }


            StreamWriter objSW =new StreamWriter(@"C:\Users\adeel.muzaffar.SYSTEMS\Desktop\Report Analysis\test.txt");
            objSW.AutoFlush = true;


            if (strODBC != "")
            {
                objSW.WriteLine(strODBC);
                objSW.WriteLine(strODBCString);
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (strOLEBD != "")
            {
                objSW.WriteLine(strOLEBD);
                objSW.WriteLine(strConString);
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (strAccessString != "")
            {
                objSW.WriteLine(strAccessString);
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (listTable.Count>0)
            {
                objSW.WriteLine("Table:");
                foreach (string strLine in listTable)
                {
                    objSW.WriteLine(strLine);

                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }


            if (listSharePoint.Count > 0)
            {
                objSW.WriteLine("SharePoint:");
                foreach (string strLine in listSharePoint)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (listLet.Count > 0)
            {
                objSW.WriteLine("Folder Path:");
                foreach (string strLine in listLet)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (listXls.Count > 0)
            {
                objSW.WriteLine("Xls File:");
                foreach (string strLine in listXls)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }


            if (listQvw.Count > 0)
            {
                objSW.WriteLine("QVW File:");
                foreach (string strLine in listQvw)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (listCsv.Count > 0)
            {
                objSW.WriteLine("CSV File:");
                foreach (string strLine in listCsv)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }


            if (listWrite.Count > 0)
            {
                objSW.WriteLine("Write List:");
                foreach (string strLine in listWrite)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            if (listRead.Count > 0)
            {
                objSW.WriteLine("Read List:");
                foreach (string strLine in listRead)
                {
                    objSW.WriteLine(strLine);
                }
                objSW.WriteLine("");
                objSW.WriteLine("");
            }

            objSW.Close();
            

            
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string[] strFiles = Directory.GetFiles(textBox1.Text);
            DataTable objTable = new DataTable("Sheet");
            objTable.Columns.Add("FileName");
            objTable.Columns.Add("strODBC");
            objTable.Columns.Add("strODBCString");
            objTable.Columns.Add("strOLEBD");
            objTable.Columns.Add("strConString");
            objTable.Columns.Add("strAccessString");
            objTable.Columns.Add("listTable");
            objTable.Columns.Add("listSharePoint");
            objTable.Columns.Add("listLet");
            //objTable.Columns.Add("listXls");
            //objTable.Columns.Add("listCsv");
            objTable.Columns.Add("listWriteQvd");
            objTable.Columns.Add("listWriteCsv");
            objTable.Columns.Add("listWriteXls");
            objTable.Columns.Add("listReadQvd");
            objTable.Columns.Add("listReadCsv");
            objTable.Columns.Add("listReadXls");
            objTable.Columns.Add("listQvw");
            



            foreach (string strFilePath in strFiles)
            {

                StreamReader objSR = new StreamReader(strFilePath);
                string strFileName = Path.GetFileNameWithoutExtension(strFilePath);
                string strODBC = "";
                string strODBCString = "";
                string strOLEBD = "";
                string strConString = "";
                string strAccessString = "";


                List<string> listODBC = new List<string>();
                List<string> listODBCString = new List<string>();
                List<string> listOLEBD = new List<string>();
                List<string> listOLEBDString = new List<string>();
                List<string> listAccessString = new List<string>();

                List<string> listTable = new List<string>();
                List<string> listSharePoint = new List<string>();
                List<string> listLet = new List<string>();
                List<string> listXls = new List<string>();
                List<string> listQvw = new List<string>();
                List<string> listCsv = new List<string>();
                List<string> listWrite = new List<string>();
                List<string> listRead = new List<string>();


                while (!objSR.EndOfStream)
                {
                    string strLine = objSR.ReadLine();
                    strLine = strLine.Trim('.').Trim();
                    if (strLine.Trim().Equals("") || strLine.Trim().Equals(";") || strLine.Length==1)
                    {
                        continue;
                    }

                    if (strLine.Substring(0,2).Equals("//"))
                    {
                        continue;
                    }

                    if (strLine.Length >= 4)
                        if (strLine.Substring(0, 4).ToUpper().Equals("ODBC"))
                        {
                            listODBC.Add("ODBC");
                            listODBCString.Add(strLine.Replace("ODBC CONNECT TO", "").Trim());
                            //strODBC = "ODBC";
                            //strODBCString= strLine.Replace("ODBC CONNECT TO","").Trim();
                            
                            //string[] strFields = strLine.Split(' ');
                            //strODBCString = strFields[strFields.Length - 1];

                        }

                    if (strLine.Length >= 5)
                        if (strLine.Substring(0, 5).ToUpper().Equals("OLEDB"))
                        {
                            listOLEBD.Add("OLEDB");
                            listOLEBDString.Add(strLine.Replace("OLEDB CONNECT TO", "").Trim());
                            
                            //strOLEBD = "OLEDB";
                            //strConString = strLine.Replace("OLEDB CONNECT TO", "").Trim(); 
                            //string[] strFields = strLine.Split(' ');
                            //strConString = strFields[strFields.Length - 1];
                        }

                    if (strLine.IndexOf(".accdb") != -1)
                    {
                        //string[] strFields = strLine.Split(' ');
                        //strAccessString = strFields[strFields.Length - 1];
                        listAccessString.Add(strLine.Replace("OLEDB CONNECT TO", "").Replace("ODBC CONNECT TO","").Trim());
                    }

                    //if (strLine.Length >= 4)
                    //if (strLine.Substring(0,4).ToUpper().Equals("FROM"))
                    //{
                    //    if (strLine.IndexOf("https://sharepoint2013") == -1)
                    //    {
                    //        //string[] strFields = strLine.Split(' ');
                    //        //listTable.Add(strFields[strFields.Length - 1]);
                    //        listTable.Add(strLine);
                    //    }
                    //}

                    //if (strLine.Length >= 4)
                    //if (strLine.Substring(0, 4).ToUpper().Equals("FROM"))
                    //{
                    //    if (strLine.IndexOf("https://sharepoint2013") > 0)
                    //    {
                    //        //string[] strFields = strLine.Split(' ');
                    //        //listSharePoint.Add(strFields[strFields.Length - 1]);
                    //        listSharePoint.Add(strLine);
                    //    }
                    //}


                    if (strLine.Length >= 3)
                        if (strLine.Substring(0, 3).ToUpper().Equals("SQL"))
                        {
                            while (!objSR.EndOfStream)
                            {
                                strLine = objSR.ReadLine();
                                strLine = strLine.Trim('.').Trim();
                                if (strLine.Length >= 4)
                                    if (strLine.Substring(0, 4).ToUpper().Equals("FROM"))
                                    {
                                        listTable.Add(strLine);
                                        break;
                                    }
                            }
                        }

                    if (strLine.Length >= 6)
                        if (strLine.Substring(0, 6).ToUpper().Equals("SELECT"))
                        {
                            while (!objSR.EndOfStream)
                            {
                                if (strLine.ToUpper().Contains(" FROM"))
                                {
                                    listTable.Add(strLine);
                                    break;
                                }
                                strLine = objSR.ReadLine();
                                strLine = strLine.Trim('.').Trim();
                                if (strLine.Length >= 4)
                                    if (strLine.Substring(0, 4).ToUpper().Equals("FROM"))
                                    {
                                        listTable.Add(strLine);
                                        break;
                                    }
                            }
                        }




                    if (strLine.IndexOf("https://sharepoint2013") !=-1)
                    {
                        //string[] strFields = strLine.Split(' ');
                        //listSharePoint.Add(strFields[strFields.Length - 1]);
                        listSharePoint.Add(strLine);
                    }


                    if (strLine.Length >= 3)
                        if (strLine.Substring(0, 3).ToUpper().Equals("LET"))
                        {
                            listLet.Add(strLine);
                        }


                    //if (strLine.ToUpper().IndexOf("XLS") != -1)
                    //{
                    //    string[] strFields = strLine.Split(' ');
                    //    listXls.Add(strFields[strFields.Length - 1]);
                    //}

                    if (strLine.ToUpper().IndexOf(".QVW") != -1)
                    {
                        //string[] strFields = strLine.Split(' ');
                        listQvw.Add(strLine);
                    }


                    //if (strLine.ToUpper().IndexOf("CSV") != -1)
                    //{
                    //    //string[] strFields = strLine.Split(' ');
                    //    listCsv.Add(strLine);
                    //}

                    if (strLine.Length >= 5)
                        if (strLine.Substring(0, 5).ToUpper().Equals("STORE"))
                        {
                            if (strLine.ToUpper().IndexOf("INTO") == -1)
                            {
                                while (!objSR.EndOfStream)
                                {
                                    strLine = objSR.ReadLine();
                                    strLine = strLine.Trim('.').Trim();

                                    if (strLine.Trim().Contains("//"))
                                    {
                                        if (strLine.ToUpper().Substring(0, 2).Equals("//"))
                                        {
                                            continue;
                                        }
                                    }

                                    if (strLine.ToUpper().IndexOf("INTO") != -1)
                                        break;
                                }

                            }
                            if (strLine.ToUpper().IndexOf("INTO") != -1)
                            {
                                string[] strFields = strLine.Split(' ');
                                if (!strFields[strFields.Length - 2].ToUpper().Equals("INTO"))
                                {
                                    listWrite.Add(strFields[strFields.Length - 2] + " " + strFields[strFields.Length - 1]);
                                }
                                else
                                {
                                    listWrite.Add(strFields[strFields.Length - 1]);
                                }


                            }
                        }


                    string[] strLoadFields = strLine.Trim().ToUpper().Split(' ');
                    if (strLine.Trim().ToUpper().Equals("LOAD") || strLine.Trim().ToUpper().Equals("MAPPING LOAD") || strLine.Trim().ToUpper().Equals("LOAD DISTINCT") || strLoadFields[0].Equals("LOAD"))
                    {
                        //string all = strLine;
                        while (!objSR.EndOfStream)
                        {
                           // all += objSR.ReadLine();


                            strLine = objSR.ReadLine();
                            strLine = strLine.Trim('.').Trim();

                            if (strLine.Trim().Contains("//"))
                            {
                                if (strLine.ToUpper().Substring(0,2).Equals("//"))
                                {
                                    continue;
                                }
                            }

                            if (strLine.ToUpper().IndexOf(".QVD") != -1 || strLine.ToUpper().IndexOf(".CSV") != -1 || strLine.ToUpper().IndexOf(".XLS") != -1 || strLine.Contains("https://sharepoint201"))
                               break;
                            else if (strLine.ToUpper().IndexOf("Resident") != -1)
                            {
                                break;
                            }
                            else if (strLine.Trim().Equals(""))
                            {
                                continue;
                            }
                            
                            else if (strLine.Trim()[strLine.Trim().Length-1]==';')
                            {
                                break;
                            }
                        }
                        if(strLine.Contains("https://sharepoint201"))
                            listSharePoint.Add(strLine.Trim());
                        else
                            listRead.Add(strLine.Trim());
                    }


                }

                if (!Directory.Exists(Path.GetDirectoryName(strFilePath) + "\\Extract"))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(strFilePath) + "\\Extract");
                }
                StreamWriter objSW = new StreamWriter(Path.GetDirectoryName(strFilePath) + "\\Extract\\" + strFileName+"_Extract.txt");
                objSW.AutoFlush = true;

                DataRow objDR = objTable.NewRow();


                objSW.WriteLine(strFileName);
                objDR[0] = strFileName;
                objSW.WriteLine("");
                objSW.WriteLine("");

                

                if (listODBC.Count > 0)
                {
                    objSW.WriteLine("ODBC:");
                    string strValue = "";
                    foreach (string strLine in listODBC)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[1] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }


                if (listODBCString.Count > 0)
                {
                    objSW.WriteLine("ODBCString:");
                    string strValue = "";
                    foreach (string strLine in listODBCString)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[2] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                if (listOLEBD.Count > 0)
                {
                    objSW.WriteLine("OLEBD:");
                    string strValue = "";
                    foreach (string strLine in listOLEBD)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[3] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                if (listOLEBDString.Count > 0)
                {
                    objSW.WriteLine("OLEBDString:");
                    string strValue = "";
                    foreach (string strLine in listOLEBDString)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[4] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                if (listAccessString.Count > 0)
                {
                    objSW.WriteLine("AccessString:");
                    string strValue = "";
                    foreach (string strLine in listAccessString)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[5] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                /////////////////////


                if (listTable.Count > 0)
                {
                    objSW.WriteLine("Table:");
                    string strValue = "";
                    foreach (string strLine in listTable)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;

                    }
                    objDR[6] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }


                if (listSharePoint.Count > 0)
                {
                    objSW.WriteLine("SharePoint:");
                    string strValue = "";
                    foreach (string strLine in listSharePoint)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;
                    }
                    objDR[7] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                if (listLet.Count > 0)
                {
                    objSW.WriteLine("Folder Path:");
                    string strValue = "";
                    foreach (string strLine in listLet)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;
                    }
                    objDR[8] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }

                //if (listXls.Count > 0)
                //{
                //    objSW.WriteLine("Xls File:");
                //    string strValue = "";
                //    foreach (string strLine in listXls)
                //    {
                //        objSW.WriteLine(strLine);
                //        strValue = strValue + strLine + Environment.NewLine;
                //    }
                //    objDR[9] = strValue;
                //    objSW.WriteLine("");
                //    objSW.WriteLine("");
                //}


                

                //if (listCsv.Count > 0)
                //{
                //    objSW.WriteLine("CSV File:");
                //    string strValue = "";
                //    foreach (string strLine in listCsv)
                //    {
                //        objSW.WriteLine(strLine);
                //        strValue = strValue + strLine + Environment.NewLine;
                //    }
                //    objDR[11] = strValue;
                //    objSW.WriteLine("");
                //    objSW.WriteLine("");
                //}


                if (listWrite.Count > 0)
                {
                    objSW.WriteLine("Write List:");
                    string strValue = "";
                    foreach (string strLine in listWrite)
                    {
                        if (strLine.ToUpper().IndexOf(".QVD") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[9] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                    
                    strValue = "";
                    foreach (string strLine in listWrite)
                    {
                        if (strLine.ToUpper().IndexOf(".CSV") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[10] = strValue;

                    strValue = "";
                    foreach (string strLine in listWrite)
                    {
                        if (strLine.ToUpper().IndexOf(".XLS") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[11] = strValue;

                }

                if (listRead.Count > 0)
                {
                    objSW.WriteLine("Read List:");
                    string strValue = "";
                    foreach (string strLine in listRead)
                    {
                        if (strLine.ToUpper().IndexOf(".QVD") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[12] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");

                    strValue = "";
                    foreach (string strLine in listRead)
                    {
                        if (strLine.ToUpper().IndexOf(".CSV") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[13] = strValue;


                    strValue = "";
                    foreach (string strLine in listRead)
                    {
                        if (strLine.ToUpper().IndexOf(".XLS") != -1)
                        {
                            objSW.WriteLine(strLine);
                            strValue = strValue + strLine + Environment.NewLine;
                        }
                    }
                    objDR[14] = strValue;

                }

                if (listQvw.Count > 0)
                {
                    objSW.WriteLine("QVW File:");
                    string strValue = "";
                    foreach (string strLine in listQvw)
                    {
                        objSW.WriteLine(strLine);
                        strValue = strValue + strLine + Environment.NewLine;
                    }
                    objDR[15] = strValue;
                    objSW.WriteLine("");
                    objSW.WriteLine("");
                }


                objSW.Close();
                objTable.Rows.Add(objDR);

            }
            ExportDataSetToExcel(objTable);

            MessageBox.Show("Done....");
        }

        private void ExportDataSetToExcel(DataTable table)
        {

            if (File.Exists(textBox2.Text))
            {
                File.Delete(textBox2.Text);
                File.Copy(textBox3.Text, textBox2.Text);
            }
            else
            {
                File.Copy(textBox3.Text, textBox2.Text);
            }
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();
            

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(textBox2.Text);


            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = table.TableName;

            for (int i = 1; i < table.Columns.Count + 1; i++)
            {
                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
            }

            for (int j = 0; j < table.Rows.Count; j++)
            {
                for (int k = 0; k < table.Columns.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                }
            }


            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }
    }
}
