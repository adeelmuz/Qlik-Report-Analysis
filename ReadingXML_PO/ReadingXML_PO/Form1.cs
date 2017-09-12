using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace ReadingXML_PO
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            string strlastElement = null;
            XmlTextReader reader = new XmlTextReader(textBox1.Text);
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // The node is an element.
                        if (reader.Name.Equals("Address", StringComparison.OrdinalIgnoreCase))
                        {
                            
                        }
                        if (!objHT.ContainsKey(reader.Name))
                        {
                            objHT.Add(reader.Name, new List<string>());
                            objHT_ElementCounter.Add(reader.Name, 1);
                        }
                        else
                        {
                            int intCounter =(int) objHT_ElementCounter[reader.Name];
                            intCounter++;
                            objHT_ElementCounter[reader.Name] = intCounter;
                        }

                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        strValue = reader.Value.Replace("\r\n",",");
                        break;
                    case XmlNodeType.EndElement: //Display the end of the element.
                        List<string> strList=(List<string>) objHT[reader.Name];
                        strList.Add(strValue);

                        break;
                }
            }
            reader.Close();
            StreamWriter objStreamWriter = new StreamWriter("PO_XML_to_TEXT.txt");
            objStreamWriter.AutoFlush = true;
            string strLine = null;
            foreach (string strKey in objHT_ElementCounter.Keys)
            {
                strLine = null;
                strLine = strLine + strKey+ '\t';
                strLine = strLine + (string)objHT_ElementCounter[strKey].ToString() + '\t';
                List<string> strList = (List<string>)objHT[strKey];
                foreach (string strListValue in strList)
                {
                    strLine = strLine + strListValue + '\t';
                }
                objStreamWriter.WriteLine(strLine);
            }
            objStreamWriter.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            XmlTextReader reader = new XmlTextReader(textBox1.Text);
            int intFileCounter = 0;
            while (reader.Read())
            {

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("PurchLineAllVersions", StringComparison.OrdinalIgnoreCase))
                    {
                        intFileCounter++;
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (reader.Name.Equals("DefaultDimension", StringComparison.OrdinalIgnoreCase))
                                {
                                    //SkippThisElement(reader, "DefaultDimension");
                                    GetDefaultDimensionElementInFile(reader, "DefaultDimension", intFileCounter);
                                    continue;
                                }
                                if (reader.Name.Equals("PurchLineHistoryAddress", StringComparison.OrdinalIgnoreCase))
                                {
                                    //SkippThisElement(reader, "PurchLineHistoryAddress");
                                    GetThisElementInFile(reader, "PurchLineHistoryAddress", intFileCounter);
                                    continue;
                                }
                                if (reader.Name.Equals("InventReportDimHistory", StringComparison.OrdinalIgnoreCase))
                                {
                                    //SkippThisElement(reader, "InventReportDimHistory");
                                    GetThisElementInFile(reader, "InventReportDimHistory", intFileCounter);
                                    continue;
                                }
                                if (!objHT.ContainsKey(reader.Name))
                                {
                                    objHT.Add(reader.Name, new List<string>());
                                    objHT_ElementCounter.Add(reader.Name, 1);
                                }
                                else
                                {
                                    int intCounter = (int)objHT_ElementCounter[reader.Name];
                                    intCounter++;
                                    objHT_ElementCounter[reader.Name] = intCounter;
                                }
                            }

                            if (reader.NodeType == XmlNodeType.Text)
                            {
                                strValue = reader.Value.Replace("\r\n", ",");
                            }

                            if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                if (reader.Name.Equals("PurchLineAllVersions", StringComparison.OrdinalIgnoreCase))
                                {
                                    WriteDataToFile(intFileCounter, objHT, objHT_ElementCounter, "PurchLineAllVersions");
                                    objHT.Clear();
                                    objHT_ElementCounter.Clear();
                                    break;
                                }
                                List<string> strList = (List<string>)objHT[reader.Name];
                                strList.Add(strValue);
                            }


                        }
                    }
                }
            }
            reader.Close();
        }

        private void GetDefaultDimensionElementInFile(XmlTextReader reader, string strFileNamePrefix, int intFileCounter)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            while (reader.Read())
            {
                if (reader.Name.Equals(""))
                {
                    continue;
                }
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("Values", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name.Equals(strFileNamePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        WriteDataToFile(intFileCounter, objHT, objHT_ElementCounter, strFileNamePrefix);
                        objHT.Clear();
                        objHT_ElementCounter.Clear();
                        break;
                    }
                    if (reader.Name.Equals("Values", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }
                


                if (reader.Name.Equals("Value", StringComparison.OrdinalIgnoreCase))
                {
                    while (reader.Read())
                    {
                        
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            if (!objHT.ContainsKey(reader.Name))
                            {
                                objHT.Add(reader.Name, new List<string>());
                                objHT_ElementCounter.Add(reader.Name, 1);
                            }
                            else
                            {
                                int intCounter = (int)objHT_ElementCounter[reader.Name];
                                intCounter++;
                                objHT_ElementCounter[reader.Name] = intCounter;
                            }
                        }
                        if (reader.NodeType == XmlNodeType.Text)
                        {
                            strValue = reader.Value.Replace("\r\n", ",");
                        }

                        if (reader.NodeType == XmlNodeType.EndElement)
                        {
                            if (reader.Name.Equals("Values", StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }
                            List<string> strList = (List<string>)objHT[reader.Name];
                            strList.Add(strValue);
                            if (reader.Name.Equals("Value", StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }
                        }
                    }
                }
                

            }
           
        }

        private void GetThisElementInFile(XmlTextReader reader, string strFileNamePrefix, int intFileCounter)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.HasAttributes)
                    {
                        if (reader.GetAttribute(0).Equals("entity",StringComparison.OrdinalIgnoreCase))
                        {
                            GetThisElementInFile(reader, reader.Name, intFileCounter);
                            continue;
                        }
                       
                    }
                    if (!objHT.ContainsKey(reader.Name))
                    {
                        objHT.Add(reader.Name, new List<string>());
                        objHT_ElementCounter.Add(reader.Name, 1);
                    }
                    else
                    {
                        int intCounter = (int)objHT_ElementCounter[reader.Name];
                        intCounter++;
                        objHT_ElementCounter[reader.Name] = intCounter;
                    }
                }
                if (reader.NodeType == XmlNodeType.Text)
                {
                    strValue = reader.Value.Replace("\r\n", ",");
                }

                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name.Equals(strFileNamePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        WriteDataToFile(intFileCounter, objHT, objHT_ElementCounter, strFileNamePrefix);
                        objHT.Clear();
                        objHT_ElementCounter.Clear();
                        break;
                    }
                    List<string> strList = (List<string>)objHT[reader.Name];
                    strList.Add(strValue);
                    
                }
            }

        }

        private void WriteDataToFile(int intFileCounter, Hashtable objHT, Hashtable objHT_ElementCounter,string strFileNamePrefix)
        {
            StreamWriter objStreamWriter = new StreamWriter(strFileNamePrefix +"_"+ intFileCounter.ToString() + ".txt");
            objStreamWriter.AutoFlush = true;
            string strLine = null;
            foreach (string strKey in objHT_ElementCounter.Keys)
            {
                strLine = null;
                strLine = strLine + strKey + '\t';
                strLine = strLine + (string)objHT_ElementCounter[strKey].ToString() + '\t';
                List<string> strList = (List<string>)objHT[strKey];
                foreach (string strListValue in strList)
                {
                    strLine = strLine + strListValue + '\t';
                }
                objStreamWriter.WriteLine(strLine);
            }
            objStreamWriter.Close();
        }

        private void SkippThisElement(XmlTextReader reader, string pStrEndElement)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name.Equals(pStrEndElement, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            XmlTextReader reader = new XmlTextReader(textBox1.Text);
            int intFileCounter = 0;
            while (reader.Read())
            {

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("PurchTableAllVersions", StringComparison.OrdinalIgnoreCase))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (reader.Name.Equals("PurchTableHistoryAddress", StringComparison.OrdinalIgnoreCase))
                                {
                                    //SkippThisElement(reader, "PurchTableHistoryAddress");
                                    if (reader.HasAttributes)
                                    {
                                        if (reader.GetAttribute(0).Equals("entity", StringComparison.OrdinalIgnoreCase))
                                        {
                                            GetThisElementInFile(reader, reader.Name, intFileCounter);
                                            continue;
                                        }

                                    }
                                    continue;
                                }
                                if (reader.Name.Equals("PurchLineAllVersions", StringComparison.OrdinalIgnoreCase))
                                {
                                    SkippThisElement(reader, "PurchLineAllVersions");
                                    continue;
                                }
                                if (!objHT.ContainsKey(reader.Name))
                                {
                                    objHT.Add(reader.Name, new List<string>());
                                    objHT_ElementCounter.Add(reader.Name, 1);
                                }
                                else
                                {
                                    int intCounter = (int)objHT_ElementCounter[reader.Name];
                                    intCounter++;
                                    objHT_ElementCounter[reader.Name] = intCounter;
                                }
                            }

                            if (reader.NodeType == XmlNodeType.Text)
                            {
                                strValue = reader.Value.Replace("\r\n", ",");
                            }

                            if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                if (reader.Name.Equals("PurchTableAllVersions", StringComparison.OrdinalIgnoreCase))
                                {
                                    WriteDataToFile(0, objHT, objHT_ElementCounter, "PurchTableAllVersions");
                                    objHT.Clear();
                                    objHT_ElementCounter.Clear();
                                    break;
                                }
                                List<string> strList = (List<string>)objHT[reader.Name];
                                strList.Add(strValue);
                            }


                        }
                    }
                }
            }
            reader.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Hashtable objHT = new Hashtable();
            Hashtable objHT_ElementCounter = new Hashtable();
            string strValue = null;
            XmlTextReader reader = new XmlTextReader(textBox1.Text);
            int intFileCounter = 0;
            while (reader.Read())
            {

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("VendPurchOrderJour", StringComparison.OrdinalIgnoreCase))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (reader.Name.Equals("PurchTableAllVersions", StringComparison.OrdinalIgnoreCase))
                                {
                                    SkippThisElement(reader, "PurchTableAllVersions");
                                    continue;
                                }
                                if (!objHT.ContainsKey(reader.Name))
                                {
                                    objHT.Add(reader.Name, new List<string>());
                                    objHT_ElementCounter.Add(reader.Name, 1);
                                }
                                else
                                {
                                    int intCounter = (int)objHT_ElementCounter[reader.Name];
                                    intCounter++;
                                    objHT_ElementCounter[reader.Name] = intCounter;
                                }
                            }

                            if (reader.NodeType == XmlNodeType.Text)
                            {
                                strValue = reader.Value.Replace("\r\n", ",");
                            }

                            if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                if (reader.Name.Equals("VendPurchOrderJour", StringComparison.OrdinalIgnoreCase))
                                {
                                    WriteDataToFile(0, objHT, objHT_ElementCounter, "VendPurchOrderJour");
                                    objHT.Clear();
                                    objHT_ElementCounter.Clear();
                                    break;
                                }
                                List<string> strList = (List<string>)objHT[reader.Name];
                                strList.Add(strValue);
                            }


                        }
                    }
                }
            }
            reader.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            XmlDocument xd = new XmlDocument();
            xd.Load(@"C:\Users\adeel.muzaffar.SYSTEMS\Desktop\Regeneron\Tasks\Tasks\Task_0c6f860e-57db-4ad4-8fee-31b4100d2cd3.xml");

            XmlNodeList nodelist = xd.SelectNodes("/DistributeTask");
            string strName = "";
            string strEnabled = "";

            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    strName = node.Attributes.GetNamedItem("Name").Value;
                    strEnabled = node.Attributes.GetNamedItem("Enabled").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }


            string strID = "";
            nodelist = xd.SelectNodes("DistributeTask/StartTriggers/ScheduleTrigger");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    strID = node.Attributes.GetNamedItem("ID").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }


            string strFileName = "";
            nodelist = xd.SelectNodes("DistributeTask/SourceDocument");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    strFileName = node.Attributes.GetNamedItem("FileName").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }

            string strType = "";
            nodelist = xd.SelectNodes("DistributeTask/Resources/Resource");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    strType = node.Attributes.GetNamedItem("Type").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }

            string strEnableDateTime = "";
            string strExpireDateTime = "";
            string strRecurrenceType = "";
            string strRepeatEvery = "";
            string strHourStart = "";
            string strDayStart = "";
            nodelist = xd.SelectNodes("DistributeTask/StartTriggers/ScheduleTrigger");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    //strType = node.Attributes.GetNamedItem("Type").Value;
                    strEnableDateTime = node.Attributes.GetNamedItem("EnableDateTime").Value;
                    strExpireDateTime = node.Attributes.GetNamedItem("ExpireDateTime").Value;
                    strRecurrenceType = node.Attributes.GetNamedItem("RecurrenceType").Value;
                    strRepeatEvery = node.Attributes.GetNamedItem("RepeatEvery").Value;
                    strHourStart = node.Attributes.GetNamedItem("HourStart").Value;
                    strDayStart = node.Attributes.GetNamedItem("DayStart").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }
            if (nodelist.Count == 0)
            {
                strEnableDateTime = "N/A";
                strExpireDateTime = "N/A";
                strRecurrenceType = "N/A";
                strRepeatEvery = "N/A";
                strHourStart = "N/A";
                strDayStart = "N/A";
            }



            string strMonth = "";
            string strDayOfWeek = "";
            string strDay = "";
            string strHour = "";
            string strRecurrenceHit = "";
            nodelist = xd.SelectNodes("DistributeTask/StartTriggers/ScheduleTrigger/RecurrenceFilter");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    //strType = node.Attributes.GetNamedItem("Type").Value;
                    strMonth = node.Attributes.GetNamedItem("Month").Value;
                    strDayOfWeek = node.Attributes.GetNamedItem("DayOfWeek").Value;
                    strDay = node.Attributes.GetNamedItem("Day").Value;
                    strHour = node.Attributes.GetNamedItem("Hour").Value;
                    strRecurrenceHit = node.Attributes.GetNamedItem("RecurrenceHit").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }
            if (nodelist.Count == 0)
            {
                strMonth = "N/A";
                strDayOfWeek = "N/A";
                strDay = "N/A";
                strHour = "N/A";
                strRecurrenceHit = "N/A";
            }



        }

        private void button6_Click(object sender, EventArgs e)
        {
            //string strA = "a12121-21221-212121.qwv with hello";
            //strA.IndexOf(

            Hashtable objHT = new Hashtable();
            StreamWriter objSW = new StreamWriter(@"C:\Users\adeel.muzaffar.SYSTEMS\Desktop\Regeneron\Tasks\final.txt");
            objSW.AutoFlush = true;
            
            string[] strFiles = Directory.GetFiles(@"C:\Users\adeel.muzaffar.SYSTEMS\Desktop\Regeneron\Tasks\Tasks");
            foreach (string strFilePath in strFiles)
            {
                //string strFileName = Path.GetFileNameWithoutExtension(strFilePath);
                //strFileName = strFileName.Replace("Task_", "");

                XmlDocument xd = new XmlDocument();
                xd.Load(strFilePath);

                XmlNodeList nodelist = xd.SelectNodes("/DistributeTask");
                string strName = "";
                string strID = "";
                foreach (XmlNode node in nodelist) // for each <testcase> node
                {
                    strID = node.Attributes.GetNamedItem("ID").Value;
                    strName = node.Attributes.GetNamedItem("Name").Value;
                }

                objHT.Add(strID, strName);
                
            }

            int intCounter = 1;
            foreach (string strFilePath in strFiles)
            {
                XmlDocument xd = new XmlDocument();
                xd.Load(strFilePath);

                XmlNodeList nodelist = xd.SelectNodes("/DistributeTask");
                string strName = "";
                string strEnabled = "";

                foreach (XmlNode node in nodelist) // for each <testcase> node
                {
                    try
                    {
                        strName = node.Attributes.GetNamedItem("Name").Value;
                        strEnabled = node.Attributes.GetNamedItem("Enabled").Value;
                    }
                    catch (Exception objE)
                    {
                        MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                    }
                }


                string strID = "";
                string strPName = "";
                nodelist = xd.SelectNodes("DistributeTask/StartTriggers/TaskCompletedTrigger");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    strID = node.Attributes.GetNamedItem("TargetID").Value;
                    if (objHT.ContainsKey(strID))
                    {
                        strPName =(string) objHT[strID];
                    }

                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }

            string strEnableDateTime = "";
            string strExpireDateTime = "";
            string strRecurrenceType = "";
            string strRepeatEvery = "";
            string strHourStart = "";
            string strDayStart = "";
            nodelist = xd.SelectNodes("DistributeTask/StartTriggers/ScheduleTrigger");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    //strType = node.Attributes.GetNamedItem("Type").Value;
                    strEnableDateTime = node.Attributes.GetNamedItem("EnableDateTime").Value;
                    strExpireDateTime = node.Attributes.GetNamedItem("ExpireDateTime").Value;
                    strRecurrenceType = node.Attributes.GetNamedItem("RecurrenceType").Value;
                    strRepeatEvery = node.Attributes.GetNamedItem("RepeatEvery").Value;
                    strHourStart = node.Attributes.GetNamedItem("HourStart").Value;
                    strDayStart = node.Attributes.GetNamedItem("DayStart").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }
            if (nodelist.Count == 0)
            {
                strEnableDateTime = "N/A";
                strExpireDateTime = "N/A";
                strRecurrenceType = "N/A";
                strRepeatEvery = "N/A";
                strHourStart = "N/A";
                strDayStart = "N/A";
            }
            


            string strMonth = "";
            string strDayOfWeek = "";
            string strDay = "";
            string strHour = "";
            string strRecurrenceHit = "";
            nodelist = xd.SelectNodes("DistributeTask/StartTriggers/ScheduleTrigger/RecurrenceFilter");
            foreach (XmlNode node in nodelist) // for each <testcase> node
            {
                try
                {
                    //strType = node.Attributes.GetNamedItem("Type").Value;
                    strMonth = node.Attributes.GetNamedItem("Month").Value;
                    strDayOfWeek = node.Attributes.GetNamedItem("DayOfWeek").Value;
                    strDay = node.Attributes.GetNamedItem("Day").Value;
                    strHour = node.Attributes.GetNamedItem("Hour").Value;
                    strRecurrenceHit = node.Attributes.GetNamedItem("RecurrenceHit").Value;
                }
                catch (Exception objE)
                {
                    MessageBox.Show("Error in reading XML", "xmlError", MessageBoxButtons.OK);
                }
            }
            if (nodelist.Count == 0)
            {
                strMonth = "N/A";
                strDayOfWeek = "N/A";
                strDay = "N/A";
                strHour = "N/A";
                strRecurrenceHit = "N/A";
            }


            string[] strCName = strName.Split('\\');
            string[] strParName = strPName.Split('\\');

                //objSW.WriteLine(strName + ";" + strEnabled + ";" + strPName);

            string strOutPut = "";
            strOutPut = Path.GetFileNameWithoutExtension(strFilePath) + "\t" + intCounter.ToString() + "\t" + strName + "\t" + strCName[0] + "\t" + strCName[strCName.Length - 1] + "\t" + strEnabled + "\t" + strPName + "\t" + strParName[0] + "\t" + strParName[strParName.Length - 1];


            strOutPut = strOutPut + "\t" + strEnableDateTime;
            strOutPut = strOutPut + "\t" + strExpireDateTime;
            strOutPut = strOutPut + "\t" + strRecurrenceType;
            strOutPut = strOutPut + "\t" + strRepeatEvery;
            strOutPut = strOutPut + "\t" + strHourStart;
            strOutPut = strOutPut + "\t" + strDayStart;
            strOutPut = strOutPut + "\t" + strMonth;
            strOutPut = strOutPut + "\t" + strDayOfWeek;
            strOutPut = strOutPut + "\t" + strDay;
            strOutPut = strOutPut + "\t" + strHour;
            strOutPut = strOutPut + "\t" + strRecurrenceHit;


            objSW.WriteLine(strOutPut);
            intCounter++;

            }
            objSW.Close();
            MessageBox.Show("Done...");
        }
    }
}
