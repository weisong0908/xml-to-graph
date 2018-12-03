/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////                                                                 //////////////////////////////////
//////////////////////////////////    created by: Teng Wei Song (Intern) on 26th Spetember 2011    //////////////////////////////////
//////////////////////////////////                                                                 //////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpProjectPrototype3
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AboutBox1 abtBox = new AboutBox1();
            this.Text = string.Format("{0} version {1}", abtBox.AssemblyTitle,abtBox.AssemblyVersion);
        }

        //PREPARATIONS BEFORE 1st STEP/////////////////////////////////////////////////////////////////

        //CREATE XML DOCUMENT
        XmlDocument xmldoc = new XmlDocument();

        //PREPARATION DONE/////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////////////////////////
        //1st STEP: OPEN THE XML FILE//////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        private void buttonImportXML_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //xmlPreviewer clears content if any
                xmlPreviewer.GoHome();

                //shows xml file
                xmlPreviewer.Navigate(openFileDialog1.FileName);

                //read xml file in xmldoc
                xmldoc.Load(openFileDialog1.FileName);

                //display filename in the text box
                int xmlFileNameIndex = openFileDialog1.FileName.LastIndexOf(@"\");
                textBoxXmlFileName.Text = openFileDialog1.FileName.Substring(xmlFileNameIndex + 1);

                //enable groupbox "Export to Trend Report"
                groupBoxExportToTrendReport.Enabled = true;
            }
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////
        //1st STEP DONE////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        //PREPARATIONS BEFORE 2nd STEP/////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        //MISVALUE
        object misValue = System.Reflection.Missing.Value;

        //COMBINE TEST AND LABEL TO WORKFLOW IN XML
        private string combineTestAndLabelToMeasuredWorkflow(string test, string label)
        {
            string measuredWorkflow = test + ": " + label;
            return measuredWorkflow;
        }

        //DRAWING COUNTER//////////////////////////////////////////////////////////////////////////////
        static int xmlMultipleDwgOrder1_label_performance = 0;
        static int xmlMultipleDwgOrder2_test_performance = 0;
        static int xmlMultipleDwgOrder1_label_capacity = 0;
        static int xmlMultipleDwgOrder2_test_capacity = 0;

        //FIND MAXIMUM ROW/////////////////////////////////////////////////////////////////////////////
        private int findMaxRow(Excel.Worksheet currentWorksheet)
        {
            int i;
            Excel.Range range = currentWorksheet.UsedRange;

            for (i = 1; i <= 200; i++)
            {
                if ((string)(range.Cells[i, 1] as Excel.Range).Value2 == null)
                {
                    return i - 1;
                }
            }
            return 200;
        }

        //FIND MAXIMUM COLUMN//////////////////////////////////////////////////////////////////////////
        private int findMaxColumn(Excel.Worksheet currentWorksheet)
        {
            int i;
            Excel.Range range = currentWorksheet.UsedRange;

            for (i = 1; i <= 30; i++)
            {
                if ((string)(range.Cells[1, i] as Excel.Range).Value2 == null)
                {
                    return i - 1;
                }
            }
            return 0;
        }

        //FIND COLUMN BY BUILD VERSION/////////////////////////////////////////////////////////////////
        private int findColumn(Excel.Worksheet currentWorksheet, string currentBuildVersion)
        {
            int i;
            Excel.Range range = currentWorksheet.UsedRange;

            for (i = 1; i <= 30; i++)
            {
                if ((string)(range.Cells[1, i] as Excel.Range).Value2 == currentBuildVersion)
                {
                    return i;
                }
            }
            return 0;
        }

        //RELEASE OBJECT////////////////////////////////////////////////////////////////////////////////
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception occured during releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //PREPARATION DONE/////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////////////////////////
        //2nd STEP: READ AND WRITE VALUE TO EXCEL WORKBOOK "Trend Report.xls"//////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////
        
        private void comboBoxXMLType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //enable group box "for multiple only if user choose single and multiple"
            if (comboBoxXMLType.Text == "Single and Multiple")
            {
                groupBoxForMultipleDwgOnly.Enabled = true;
            }
            else
            {
                groupBoxForMultipleDwgOnly.Enabled = false;
            }
        }

        private void buttonExportTrendReport_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook;

            //try
            //{
                if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                }

                excelWorkbook = excelApp.Workbooks.Open(openFileDialog2.FileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                
                string typeOfXML = comboBoxXMLType.Text;

                //for SINGLE AND MULTIPLE//////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////////////////////////////////////

                if (typeOfXML == "Single and Multiple")
                {
                    //worksheets///////////////////////////////////////////////////////////////////////////
                    Excel.Worksheet ws_performanceSingle = excelWorkbook.Worksheets.get_Item(1);
                    Excel.Worksheet ws_capacitySingle = excelWorkbook.Worksheets.get_Item(2);
                    Excel.Worksheet ws_performance1000dwg = excelWorkbook.Worksheets.get_Item(3);
                    Excel.Worksheet ws_capacity1000dwg = excelWorkbook.Worksheets.get_Item(4);
                    Excel.Worksheet ws_performance100dwg = excelWorkbook.Worksheets.get_Item(5);
                    Excel.Worksheet ws_capacity100dwg = excelWorkbook.Worksheets.get_Item(6);
                    Excel.Worksheet ws_performance10dwg = excelWorkbook.Worksheets.get_Item(7);
                    Excel.Worksheet ws_capacity10dwg = excelWorkbook.Worksheets.get_Item(8);
                    Excel.Worksheet ws_performance5000dwg = excelWorkbook.Worksheets.get_Item(9);
                    Excel.Worksheet ws_capacity5000dwg = excelWorkbook.Worksheets.get_Item(10);
                    Excel.Worksheet ws_performance500dwg = excelWorkbook.Worksheets.get_Item(11);
                    Excel.Worksheet ws_capacity500dwg = excelWorkbook.Worksheets.get_Item(12);
                    Excel.Worksheet ws_performance7 = excelWorkbook.Worksheets.get_Item(13);
                    Excel.Worksheet ws_capacity7 = excelWorkbook.Worksheets.get_Item(14);
                    Excel.Worksheet ws_performance8 = excelWorkbook.Worksheets.get_Item(15);
                    Excel.Worksheet ws_capacity8 = excelWorkbook.Worksheets.get_Item(16);
                    Excel.Worksheet ws_performance9 = excelWorkbook.Worksheets.get_Item(17);
                    Excel.Worksheet ws_capacity9 = excelWorkbook.Worksheets.get_Item(18);
                    Excel.Worksheet ws_performance10 = excelWorkbook.Worksheets.get_Item(19);
                    Excel.Worksheet ws_capacity10 = excelWorkbook.Worksheets.get_Item(20);

                    Excel.Worksheet currentWorksheet;

                    //new multiple drawing
                    if (textBox7.Text != "No value")
                    {
                        ws_performance7.Name = "Performance " + textBox7.Text + " DWGs";
                        ws_capacity7.Name = "Capacity " + textBox7.Text + " DWGs";
                    }
                    if (textBox8.Text != "No value")
                    {
                        ws_performance8.Name = "Performance " + textBox7.Text + " DWGs";
                        ws_capacity8.Name = "Capacity " + textBox8.Text + " DWGs";
                    }
                    if (textBox9.Text != "No value")
                    {
                        ws_performance9.Name = "Performance " + textBox9.Text + " DWGs";
                        ws_capacity9.Name = "Capacity " + textBox9.Text + " DWGs";
                    }
                    if (textBox10.Text != "No value")
                    {
                        ws_performance10.Name = "Performance " + textBox10.Text + " DWGs";
                        ws_capacity10.Name = "Capacity " + textBox10.Text + " DWGs";
                    }

                    //build version//////////////////////////////////////////////////////////////////////////////
                    XmlNodeList xmlDocRow = xmldoc.SelectNodes("Document/Row");
                    int oneMark;
                    oneMark = xmlDocRow.Item(0).SelectSingleNode("Version").InnerText.IndexOf("1");
                    string buildVersion = xmlDocRow.Item(0).SelectSingleNode("Version").InnerText.Substring(oneMark, 9);

                    //order of drawings in xml file//////////////////////////////////////////////////////////////
                    TextBox[] textBox = { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7, textBox8, textBox9, textBox10 };
                    TextBox[] textBoxOrder = { textBoxOrder1, textBoxOrder2, textBoxOrder3, textBoxOrder4, textBoxOrder5, textBoxOrder6, textBoxOrder7, textBoxOrder8, textBoxOrder9, textBoxOrder10 };
                    Excel.Worksheet[] storedWorksheetPerformance = { ws_performanceSingle, ws_performance1000dwg, ws_performance100dwg, ws_performance10dwg, ws_performance5000dwg, ws_performance500dwg, ws_performance7, ws_performance8, ws_performance9, ws_performance10 };
                    Excel.Worksheet[] storedWorksheetCapacity = { ws_capacitySingle, ws_capacity1000dwg, ws_capacity100dwg, ws_capacity10dwg, ws_capacity5000dwg, ws_capacity500dwg, ws_capacity7, ws_capacity8, ws_capacity9, ws_capacity10 };

                    //sort performance worksheet
                    Excel.Worksheet[] sortedWorksheetPerformance = new Excel.Worksheet[10];
                    int orderForSortedWorksheetPerformance;
                    for (int i = 0; i < 10; i++)
                    {
                        int j;
                        for (j = 0; j < 10; j++)
                        {
                            int ii = i + 1;
                            if (textBoxOrder[j].Text == ii.ToString())
                            {
                                orderForSortedWorksheetPerformance = j;
                                sortedWorksheetPerformance[i] = storedWorksheetPerformance[j];
                            }
                        }
                    }

                    //sort capacity worksheet
                    Excel.Worksheet[] sortedWorksheetCapacity = new Excel.Worksheet[10];
                    int orderForSortedWorksheetCapacity;
                    for (int i = 0; i < 10; i++)
                    {
                        int j;
                        for (j = 0; j < 10; j++)
                        {
                            int ii = i + 1;
                            if (textBoxOrder[j].Text == ii.ToString())
                            {
                                orderForSortedWorksheetCapacity = j;
                                sortedWorksheetCapacity[i] = storedWorksheetCapacity[j];
                            }
                        }
                    }

                    //ACTUAL PROCESS/////////////////////////////////////////////////////////////////////////////
                    /////////////////////////////////////////////////////////////////////////////////////////////

                    int xmlItemCounter;
                    for (xmlItemCounter = 0; xmlDocRow.Item(xmlItemCounter) != null; xmlItemCounter++)
                    {
                        //performance////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Performance")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label);

                            //get the set of data from 1st part of xml file
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText == "RemovePdocTime")
                            {
                                xmlMultipleDwgOrder1_label_performance = xmlMultipleDwgOrder1_label_performance + 1;
                            }

                            if (xmlMultipleDwgOrder1_label_performance >= 1)
                            {
                                currentWorksheet = sortedWorksheetPerformance[xmlMultipleDwgOrder1_label_performance - 1];

                                int maxRow = findMaxRow(currentWorksheet);
                                int maxColumn = findMaxColumn(currentWorksheet);

                                //create a new row with buildnumber if did not exist
                                if (findColumn(currentWorksheet, buildVersion) == 0)
                                {
                                    currentWorksheet.Cells[1, maxColumn + 1] = buildVersion;
                                    maxColumn = maxColumn + 1;
                                }
                                else
                                {
                                    maxColumn = findColumn(currentWorksheet, buildVersion);
                                }

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInSeconds = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000;

                                        currentWorksheet.Cells[rowNumber, maxColumn] = Math.Round(valueInSeconds, 3, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }                                
                            
                            //get the set of data from 2nd part of xml file
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText == "Insert_first_line")
                            {
                                xmlMultipleDwgOrder2_test_performance = xmlMultipleDwgOrder2_test_performance + 1;
                            }

                            if (xmlMultipleDwgOrder2_test_performance >= 1)
                            {
                                currentWorksheet = sortedWorksheetPerformance[xmlMultipleDwgOrder2_test_performance - 1];

                                int maxRow = findMaxRow(currentWorksheet);
                                int thisColumn = findColumn(currentWorksheet, buildVersion);

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInSeconds = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000;

                                        currentWorksheet.Cells[rowNumber, thisColumn] = Math.Round(valueInSeconds, 3, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }
                            
                            //deal with startup
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Workflow").InnerText == "Startup")
                            {
                                currentWorksheet = excelWorkbook.Worksheets.get_Item(1);
                                int maxRow = findMaxRow(currentWorksheet);
                                int thisColumn = findColumn(currentWorksheet, buildVersion);

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInSeconds = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000;

                                        currentWorksheet.Cells[rowNumber, thisColumn] = Math.Round(valueInSeconds, 3, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }
                        }
                    }

                    for (xmlItemCounter = 0; xmlDocRow.Item(xmlItemCounter) != null; xmlItemCounter++)
                    {
                        //capacity////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Capacity")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label);

                            //get the set of data from 1st part of xml file
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText == "RemovePdocMemory")
                            {
                                xmlMultipleDwgOrder1_label_capacity = xmlMultipleDwgOrder1_label_capacity + 1;
                            }

                            if (xmlMultipleDwgOrder1_label_capacity >= 1)
                            {
                                currentWorksheet = sortedWorksheetCapacity[xmlMultipleDwgOrder1_label_capacity - 1];

                                int maxRow = findMaxRow(currentWorksheet);
                                int maxColumn = findMaxColumn(currentWorksheet);

                                //create a new row with buildnumber if did not exist
                                if (findColumn(currentWorksheet, buildVersion) == 0)
                                {
                                    currentWorksheet.Cells[1, maxColumn + 1] = buildVersion;
                                    maxColumn = maxColumn + 1;
                                }
                                else
                                {
                                    maxColumn = findColumn(currentWorksheet, buildVersion);
                                }

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInMegaBytes = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000000;

                                        currentWorksheet.Cells[rowNumber, maxColumn] = Math.Round(valueInMegaBytes, 2, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }
                            
                            //get the set of data from 2nd part of xml file
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText == "Insert_first_line")
                            {
                                xmlMultipleDwgOrder2_test_capacity = xmlMultipleDwgOrder2_test_capacity + 1;
                            }

                            if (xmlMultipleDwgOrder2_test_capacity >= 1)
                            {
                                currentWorksheet = sortedWorksheetCapacity[xmlMultipleDwgOrder2_test_capacity - 1];

                                int maxRow = findMaxRow(currentWorksheet);
                                int thisColumn = findColumn(currentWorksheet, buildVersion);

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInMegaBytes = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000000;

                                        currentWorksheet.Cells[rowNumber, thisColumn] = Math.Round(valueInMegaBytes, 2, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }
                            
                            //deal with startup
                            if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Workflow").InnerText == "Startup")
                            {
                                currentWorksheet = excelWorkbook.Worksheets.get_Item(2);
                                int maxRow = findMaxRow(currentWorksheet);
                                int thisColumn = findColumn(currentWorksheet, buildVersion);

                                int rowNumber;
                                for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                                {
                                    Excel.Range range = currentWorksheet.UsedRange;
                                    string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                    if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                    {
                                        double valueInMegaBytes = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000000;

                                        currentWorksheet.Cells[rowNumber, thisColumn] = Math.Round(valueInMegaBytes, 2, MidpointRounding.AwayFromZero);
                                    }
                                }
                            }
                        }
                    }

                    excelWorkbook.Save();
                    excelWorkbook.Close();
                    excelApp.Quit();

                    releaseObject(ws_performanceSingle);
                    releaseObject(ws_capacitySingle);
                    releaseObject(ws_performance10dwg);
                    releaseObject(ws_capacity10dwg);
                    releaseObject(ws_performance100dwg);
                    releaseObject(ws_capacity100dwg);
                    releaseObject(ws_performance500dwg);
                    releaseObject(ws_capacity500dwg);
                    releaseObject(ws_performance1000dwg);
                    releaseObject(ws_capacity1000dwg);
                    releaseObject(ws_performance5000dwg);
                    releaseObject(ws_capacity5000dwg);
                    releaseObject(ws_performance7);
                    releaseObject(ws_capacity7);
                    releaseObject(ws_performance8);
                    releaseObject(ws_capacity8);
                    releaseObject(ws_performance9);
                    releaseObject(ws_capacity9);
                    releaseObject(ws_performance10);
                    releaseObject(ws_capacity10);
                    releaseObject(excelWorkbook);
                    releaseObject(excelApp);

                    MessageBox.Show("Exported to: " + openFileDialog2.FileName, "Exported");
                }

                //for 32 BIT////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////////////////////////////////////

                if (typeOfXML == "Windows XP 32 bit")
                {
                    //worksheets/////////////////////////////////////////////////////////////////////////////////
                    Excel.Worksheet ws_performanceSingle = excelWorkbook.Worksheets.get_Item(1);
                    Excel.Worksheet ws_capacitySingle = excelWorkbook.Worksheets.get_Item(2);

                    //build version//////////////////////////////////////////////////////////////////////////////
                    XmlNodeList xmlDocRow = xmldoc.SelectNodes("Document/Row");
                    string buildVersion = xmlDocRow.Item(0).SelectSingleNode("Version").InnerText.Substring(4);

                    int xmlItemCounter;
                    for (xmlItemCounter = 0; xmlDocRow.Item(xmlItemCounter) != null; xmlItemCounter++)
                    {
                        //performance////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Performance")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label) + " (x32)";
                            
                            int maxRow = findMaxRow(ws_performanceSingle);
                            int maxColumn = findMaxColumn(ws_performanceSingle);

                            //create a new row with buildnumber if did not exist
                            if (findColumn(ws_performanceSingle, buildVersion) == 0)
                            {
                                ws_performanceSingle.Cells[1, maxColumn + 1] = buildVersion;
                                maxColumn = maxColumn + 1;
                            }
                            else
                            {
                                maxColumn = findColumn(ws_performanceSingle, buildVersion);
                            }

                            int rowNumber;
                            for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                            {
                                Excel.Range range = ws_performanceSingle.UsedRange;
                                string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                {
                                    double valueInSeconds = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000;

                                    ws_performanceSingle.Cells[rowNumber, maxColumn] = Math.Round(valueInSeconds, 3, MidpointRounding.AwayFromZero);
                                }
                            }
                        }

                        //capacity///////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Capacity")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label) + " (x32)";

                            int maxRow = findMaxRow(ws_capacitySingle);
                            int maxColumn = findMaxColumn(ws_capacitySingle);

                            //create a new row with buildnumber if did not exist
                            if (findColumn(ws_capacitySingle, buildVersion) == 0)
                            {
                                ws_capacitySingle.Cells[1, maxColumn + 1] = buildVersion;
                                maxColumn = maxColumn + 1;
                            }
                            else
                            {
                                maxColumn = findColumn(ws_capacitySingle, buildVersion);
                            }

                            int rowNumber;
                            for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                            {
                                Excel.Range range = ws_capacitySingle.UsedRange;
                                string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                {
                                    double valueInMegaBytes = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000000;

                                    ws_capacitySingle.Cells[rowNumber, maxColumn] = Math.Round(valueInMegaBytes, 2, MidpointRounding.AwayFromZero);
                                }
                            }
                        }
                    }

                    excelWorkbook.Save();
                    excelWorkbook.Close();
                    excelApp.Quit();

                    releaseObject(ws_performanceSingle);
                    releaseObject(ws_capacitySingle);
                    releaseObject(excelWorkbook);
                    releaseObject(excelApp);

                    MessageBox.Show("Exported to: " + openFileDialog2.FileName, "Exported");
                }

                //for 64 BIT////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////////////////////////////////////

                if (typeOfXML == "Windows 7 64 bit")
                {
                    //worksheets/////////////////////////////////////////////////////////////////////////////////
                    Excel.Worksheet ws_performanceSingle = excelWorkbook.Worksheets.get_Item(1);
                    Excel.Worksheet ws_capacitySingle = excelWorkbook.Worksheets.get_Item(2);

                    //build version//////////////////////////////////////////////////////////////////////////////
                    XmlNodeList xmlDocRow = xmldoc.SelectNodes("Document/Row");
                    string buildVersion = xmlDocRow.Item(0).SelectSingleNode("Version").InnerText.Substring(4);

                    int xmlItemCounter;
                    for (xmlItemCounter = 0; xmlDocRow.Item(xmlItemCounter) != null; xmlItemCounter++)
                    {
                        //performance////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Performance")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label) + " (x64)";

                            int maxRow = findMaxRow(ws_performanceSingle);
                            int maxColumn = findMaxColumn(ws_performanceSingle);

                            //create a new row with buildnumber if did not exist
                            if (findColumn(ws_performanceSingle, buildVersion) == 0)
                            {
                                ws_performanceSingle.Cells[1, maxColumn + 1] = buildVersion;
                                maxColumn = maxColumn + 1;
                            }
                            else
                            {
                                maxColumn = findColumn(ws_performanceSingle, buildVersion);
                            }

                            int rowNumber;
                            for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                            {
                                Excel.Range range = ws_performanceSingle.UsedRange;
                                string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                {
                                    double valueInSeconds = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000;

                                    ws_performanceSingle.Cells[rowNumber, maxColumn] = Math.Round(valueInSeconds, 3, MidpointRounding.AwayFromZero);
                                }
                            }
                        }

                        //capacity///////////////////////////////////////////////////////////////////////////////
                        if (xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Type").InnerText == "Capacity")
                        {
                            //combine test and label to measured workflow                    
                            string test = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Test").InnerText;
                            string label = xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Label").InnerText;
                            string xmlMeasuredWorkflow = combineTestAndLabelToMeasuredWorkflow(test, label) + " (x64)";

                            int maxRow = findMaxRow(ws_capacitySingle);
                            int maxColumn = findMaxColumn(ws_capacitySingle);

                            //create a new row with buildnumber if did not exist
                            if (findColumn(ws_capacitySingle, buildVersion) == 0)
                            {
                                ws_capacitySingle.Cells[1, maxColumn + 1] = buildVersion;
                                maxColumn = maxColumn + 1;
                            }
                            else
                            {
                                maxColumn = findColumn(ws_capacitySingle, buildVersion);
                            }

                            int rowNumber;
                            for (rowNumber = 2; rowNumber <= maxRow; rowNumber++)
                            {
                                Excel.Range range = ws_capacitySingle.UsedRange;
                                string excelWMeasureWorkflow = (string)(range.Cells[rowNumber, 1] as Excel.Range).Value2;

                                if (excelWMeasureWorkflow == xmlMeasuredWorkflow)
                                {
                                    double valueInMegaBytes = double.Parse(xmlDocRow.Item(xmlItemCounter).SelectSingleNode("Value").InnerText) / 1000000;

                                    ws_capacitySingle.Cells[rowNumber, maxColumn] = Math.Round(valueInMegaBytes, 2, MidpointRounding.AwayFromZero);
                                }
                            }
                        }
                    }

                    excelWorkbook.Save();
                    excelWorkbook.Close();
                    excelApp.Quit();

                    releaseObject(ws_performanceSingle);
                    releaseObject(ws_capacitySingle);
                    releaseObject(excelWorkbook);
                    releaseObject(excelApp);

                    MessageBox.Show("Exported to: " + openFileDialog2.FileName, "Exported");
                }
            //}
            /*catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }*/

                groupBoxGenerateGraphs.Enabled = true;
        }


        
        ///////////////////////////////////////////////////////////////////////////////////////////////
        //2nd STEP DONE////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        //PREPARATION BEFORE 3rd STEP//////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        //CONVERT ROW AND COLUMN TO CELL INDEX/////////////////////////////////////////////////////////
        private string cellIndex(int rowNumber, int columnNumber)
        {
            int column = columnNumber + 64;
            char columnIndex = (char)column;

            string cellIndex = columnIndex + rowNumber.ToString();

            return cellIndex;
        }

        //GET FORMULA FOR LIMIT////////////////////////////////////////////////////////////////////////
        private string limitFormula(int rowNumber, int maxColumn, Excel.Worksheet currentWorksheet)
        {
            string limitFormulaFrontPart = "=SERIES(,,{";
            string limitFormulaMiddlePart = "";
            string limitFormulaLastPart = "},2)";

            Excel.Range range = currentWorksheet.UsedRange;

            double limitValue;
            limitValue = (double)(range.Cells[rowNumber, 2] as Excel.Range).Value2;
            
            for (int i = 1; i <= maxColumn; i++)
            {
                limitFormulaMiddlePart = limitFormulaMiddlePart + limitValue + ",";
            }

            int lastIndexOfComma = limitFormulaMiddlePart.LastIndexOf(",");
            limitFormulaMiddlePart = limitFormulaMiddlePart.Substring(0, lastIndexOfComma);

            return limitFormulaFrontPart + limitFormulaMiddlePart + limitFormulaLastPart;
        }

        //POSITION OF LINE GRAPHS//////////////////////////////////////////////////////////////////////
        double positionLeft;
        double positionTop;
        int positionCounter;

        //SAVE FILE IN OTHER PLACE/////////////////////////////////////////////////////////////////////
        string savePath = "";

        //PREPARATION DONE/////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////////////////////////
        //3rd STEP: GENERATE GRAPH/////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        private void checkBoxSaveInOtherPlace_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxSaveInOtherPlace.Checked)
            {
                groupBoxSaveIn.Enabled = true;

            }
            else
            {
                groupBoxSaveIn.Enabled = false;
            }
        }

        private void buttonBrowseFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                savePath = folderBrowserDialog1.SelectedPath;
                textBoxSavePath.Text = savePath;
            }
        }

        private void buttonGenerateGraph_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook;

            if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

            }

            excelWorkbook = excelApp.Workbooks.Open(openFileDialog2.FileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            //worksheets///////////////////////////////////////////////////////////////////////////
            Excel.Worksheet ws_performanceSingle = excelWorkbook.Worksheets.get_Item(1);
            Excel.Worksheet ws_capacitySingle = excelWorkbook.Worksheets.get_Item(2);
            Excel.Worksheet ws_performance1000dwg = excelWorkbook.Worksheets.get_Item(3);
            Excel.Worksheet ws_capacity1000dwg = excelWorkbook.Worksheets.get_Item(4);
            Excel.Worksheet ws_performance100dwg = excelWorkbook.Worksheets.get_Item(5);
            Excel.Worksheet ws_capacity100dwg = excelWorkbook.Worksheets.get_Item(6);
            Excel.Worksheet ws_performance10dwg = excelWorkbook.Worksheets.get_Item(7);
            Excel.Worksheet ws_capacity10dwg = excelWorkbook.Worksheets.get_Item(8);
            Excel.Worksheet ws_performance5000dwg = excelWorkbook.Worksheets.get_Item(9);
            Excel.Worksheet ws_capacity5000dwg = excelWorkbook.Worksheets.get_Item(10);
            Excel.Worksheet ws_performance500dwg = excelWorkbook.Worksheets.get_Item(11);
            Excel.Worksheet ws_capacity500dwg = excelWorkbook.Worksheets.get_Item(12);
            Excel.Worksheet ws_performance7 = excelWorkbook.Worksheets.get_Item(13);
            Excel.Worksheet ws_capacity7 = excelWorkbook.Worksheets.get_Item(14);
            Excel.Worksheet ws_performance8 = excelWorkbook.Worksheets.get_Item(15);
            Excel.Worksheet ws_capacity8 = excelWorkbook.Worksheets.get_Item(16);
            Excel.Worksheet ws_performance9 = excelWorkbook.Worksheets.get_Item(17);
            Excel.Worksheet ws_capacity9 = excelWorkbook.Worksheets.get_Item(18);
            Excel.Worksheet ws_performance10 = excelWorkbook.Worksheets.get_Item(19);
            Excel.Worksheet ws_capacity10 = excelWorkbook.Worksheets.get_Item(20);

            Excel.Worksheet currentWorksheet;
            Excel.Worksheet[] allWorksheet = { ws_performanceSingle, ws_capacitySingle, ws_performance1000dwg, ws_capacity1000dwg, ws_performance100dwg, ws_capacity100dwg, ws_performance10dwg, ws_capacity10dwg, ws_performance5000dwg, ws_capacity5000dwg, ws_performance500dwg, ws_capacity500dwg, ws_performance7, ws_capacity7, ws_performance8, ws_capacity8, ws_performance9, ws_capacity9, ws_performance10, ws_capacity10 };

            int maxColumn;
            int maxRow;

            string newFileName = textBoxFileName.Text;

            for (int i = 0; i < 20; i++)
            {
                currentWorksheet = allWorksheet[i];

                //get max column and row
                maxColumn = findMaxColumn(currentWorksheet);
                maxRow = findMaxRow(currentWorksheet);

                positionLeft = 0;
                //int positionTopDefault = (maxRow + 5) * 20;
                int positionTopDefault = (maxRow) * 20;
                positionTop = positionTopDefault;
                positionCounter = 0;

                for (int j = 2; j <= maxRow; j++)
                {
                    //values for line graph
                    string valueIndex1 = cellIndex(j, 3);
                    string valueIndex2 = cellIndex(j, maxColumn);
                    Excel.Range chartValueRange = currentWorksheet.get_Range(valueIndex1, valueIndex2);

                    //x-axis for line graph
                    string xIndex1 = cellIndex(1, 3);
                    string xIndex2 = cellIndex(1, maxColumn);
                    Excel.Range xAxisRange = currentWorksheet.Range[xIndex1, xIndex2];

                    //create chart object
                    Excel.Range range = currentWorksheet.UsedRange;
                    string lineGraphTitle = (string)(range.Cells[j, 1] as Excel.Range).Value2;

                    //position line graphs
                    positionLeft = 820 * (j - 2) - (8200 * positionCounter);

                    if (positionLeft == 8200)
                    {
                        positionCounter = positionCounter + 1;
                        positionTop = positionTopDefault + (270 * positionCounter);
                        positionLeft = 0;
                    }

                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)currentWorksheet.ChartObjects();
                    Excel.ChartObject lineGraph = chartObjects.Add(positionLeft, positionTop - (maxRow * 2), 800, 250);
                    lineGraph.Chart.ChartType = Excel.XlChartType.xlLine;
                    lineGraph.Chart.HasTitle = true;
                    lineGraph.Chart.ChartTitle.Text = lineGraphTitle;
                    lineGraph.Chart.HasDataTable = true;

                    Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)lineGraph.Chart.SeriesCollection();
                    Excel.Series trend = seriesCollection.NewSeries();
                    
                    trend.Values = chartValueRange;
                    trend.XValues = xAxisRange;
                    trend.Name = "trend";

                    //get limit if any
                    try
                    {
                        if ((double)(range.Cells[j, 2] as Excel.Range).Cells.Value2 != null)
                        {
                            Excel.Series limit = seriesCollection.NewSeries();
                            limit.Formula = limitFormula(j, maxColumn - 2, currentWorksheet);
                            limit.Name = "limit";
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                    }

                    lineGraph.Chart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                }
            }

            //remove unused worksheets
            foreach ( Excel.Worksheet worksheet_in_check in allWorksheet)
            {
                if (worksheet_in_check.Name.Substring(0,2) == "un")
                {
                    excelApp.DisplayAlerts = false;
                    worksheet_in_check.Delete();
                    excelApp.DisplayAlerts = true;
                }
            }

            //Save the new Trend Report with graphs
            string messageBoxmessage = "Trend Report with Graphs \"" + newFileName + "\" is saved in: ";

            if (checkBoxSaveInOtherPlace.Checked)
            {
                excelWorkbook.SaveAs(savePath + @"\" + newFileName + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                messageBoxmessage = messageBoxmessage + "\"" + savePath + "\"";
            }
            else
            {
                excelWorkbook.SaveAs(newFileName + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                messageBoxmessage = messageBoxmessage + "\"My Documents\"";
            }

            excelApp.Quit();

            releaseObject(ws_performanceSingle);
            releaseObject(ws_capacitySingle);
            releaseObject(ws_performance10dwg);
            releaseObject(ws_capacity10dwg);
            releaseObject(ws_performance100dwg);
            releaseObject(ws_capacity100dwg);
            releaseObject(ws_performance500dwg);
            releaseObject(ws_capacity500dwg);
            releaseObject(ws_performance1000dwg);
            releaseObject(ws_capacity1000dwg);
            releaseObject(ws_performance5000dwg);
            releaseObject(ws_capacity5000dwg);
            releaseObject(ws_performance7);
            releaseObject(ws_capacity7);
            releaseObject(ws_performance8);
            releaseObject(ws_capacity8);
            releaseObject(ws_performance9);
            releaseObject(ws_capacity9);
            releaseObject(ws_performance10);
            releaseObject(ws_capacity10);
            releaseObject(excelWorkbook);
            releaseObject(excelApp);

            MessageBox.Show(messageBoxmessage);
        }
    }
}