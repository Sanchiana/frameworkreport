//Author:Sanchiana Carvalho

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
//using System.Drawing.Rectangle;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

//using Scripting;

namespace FrameworkChrome
{

    public class Class1
    {
        // public int strTestCaseID;
        public string resultpath;
        public string Photopath;
        public string pathstring1;
        public string destDirName3;
        public string strModule;
        public string strResultReportFile;
        public string test;
        public string test1;

        //public string strTestCaseName;

        public void Go()
        {

            ChromeDriverService service = ChromeDriverService.CreateDefaultService
                (@"C:\Users\CarvalhoS\Desktop");
            // properties on the service can be used to e.g. hide the command prompt

            // ChromeOptions options = new ChromeOptions
            //{
            //IgnoreZoomLevel = true
            // };

            IWebDriver driver = new ChromeDriver(service);
            driver.Navigate().GoToUrl("http://agi.abtassociates.com/");

        }
        public void ResultReport(string strTestCaseID1, string strTestCaseDesc2, string strTestCaseName)
        // public void ResultReport()
        {
            DateTime localDate = DateTime.Now;
            string result2 = localDate.ToString("yyyyMMddHHmmss");
            strResultReportFile = pathstring1 + "\\test1" + result2;
            //StreamWriteLiner sWrite = File.Open

            using (StreamWriter sWrite = new StreamWriter(strResultReportFile + ".txt"))
            //using (FileStream sWrite =File.OpenWriteLine(resultpath + ".html"))

            // FileStream fs = new FileStream(, FileMode.Create);


            {
                sWrite.WriteLine("<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'> ");

                sWrite.WriteLine("<style type=\"text/css\">");

                sWrite.WriteLine(".box1{");

                sWrite.WriteLine("font-family: 'Exo 2', sans-serif;");

                sWrite.WriteLine("font-size:22px;");

                sWrite.WriteLine("background-color: #FF0000;");

                sWrite.WriteLine("opacity: 0.8;");

                sWrite.WriteLine("color:#FFF;");

                sWrite.WriteLine("height:50px;");

                sWrite.WriteLine("font-weight:100;");

                sWrite.WriteLine("width:1024px;");

                sWrite.WriteLine("border-color:#FFF;");

                sWrite.WriteLine("border-width:2;");

                sWrite.WriteLine("border-top-left-radius: 4px;");

                sWrite.WriteLine("border-top-right-radius:4px;");

                sWrite.WriteLine("border-bottom-left-radius: 4px;");

                sWrite.WriteLine("border-bottom-right-radius:4px;");

                sWrite.WriteLine("box-shadow: 3px 3px #B4B4B4;}");


                sWrite.WriteLine(".gap{");

                sWrite.WriteLine("background-color:#FFF;");

                sWrite.WriteLine("height:10px;");

                sWrite.WriteLine("width:10px;}");
                sWrite.WriteLine(".box2{");

                sWrite.WriteLine("background-color:#C2C5CA;");

                sWrite.WriteLine("opacity: 0.6;");

                sWrite.WriteLine("width: 507px;");

                sWrite.WriteLine("height: auto;");

                sWrite.WriteLine("border-color:#000;");

                sWrite.WriteLine("border-width:2.5;");

                sWrite.WriteLine("border-top-left-radius: 4px;");

                sWrite.WriteLine("border-top-right-radius:4px;");

                sWrite.WriteLine("border-bottom-left-radius: 4px;");

                sWrite.WriteLine("border-bottom-right-radius:4px;");

                sWrite.WriteLine("box-shadow: 3px 3px #B4B4B4;}");


                sWrite.WriteLine(".box3{");

                sWrite.WriteLine("font-family: 'Andalus', sans-serif;");

                sWrite.WriteLine("font-size:16px;");

                sWrite.WriteLine("font-weight:100;");

                sWrite.WriteLine("color:#000;");

                sWrite.WriteLine("width: 490px;");

                sWrite.WriteLine("height: auto;");

                sWrite.WriteLine("vertical-align:top;");

                sWrite.WriteLine("padding-left: 15px;");

                sWrite.WriteLine("padding-top: 15px;");

                sWrite.WriteLine("padding-bottom: 15px;}");


                sWrite.WriteLine(".box4{");

                sWrite.WriteLine("background-color:#C2C5CA;");

                sWrite.WriteLine("opacity: 0.6;");

                sWrite.WriteLine("width: 507px;");

                sWrite.WriteLine("height: auto;");

                sWrite.WriteLine("border-color:#000;");

                sWrite.WriteLine("border-width:2;");

                sWrite.WriteLine("border-top-left-radius: 4px;");

                sWrite.WriteLine("border-top-right-radius:4px;");

                sWrite.WriteLine("border-bottom-left-radius: 4px;");

                sWrite.WriteLine("border-bottom-right-radius:4px;");

                sWrite.WriteLine("box-shadow: 3px 3px #B4B4B4;}");


                sWrite.WriteLine(".f3{");

                sWrite.WriteLine("font-size: 13px;");

                sWrite.WriteLine("font-family:Calibri;}");

                sWrite.WriteLine(".f3:HOVER{;");

                sWrite.WriteLine("background-color: white;}");


                sWrite.WriteLine(".tabledesign{");

                sWrite.WriteLine("border-collapse:collapse;");

                sWrite.WriteLine("border-color:#000;");

                sWrite.WriteLine("border-bottom-left-radius: 6px;");

                sWrite.WriteLine("box-shadow: 3px 3px #B4B4B4;}");


                sWrite.WriteLine(".tablestyle{");

                sWrite.WriteLine("font-family: 'Andalus', sans-serif;");

                sWrite.WriteLine("font-size:16px;");

                sWrite.WriteLine("font-weight:100;");

                sWrite.WriteLine("color:#000;");

                sWrite.WriteLine("padding-left: 8px;");

                sWrite.WriteLine("padding-top: 8px;");

                sWrite.WriteLine("padding-bottom: 8px;");

                sWrite.WriteLine("padding-right: 8px;}");


                sWrite.WriteLine(".tableheaderstyle{");

                sWrite.WriteLine("font-family: 'Andalus', sans-serif;");

                sWrite.WriteLine("font-size:16px;");

                sWrite.WriteLine("font-weight:100;");

                sWrite.WriteLine("color:#fff;");

                sWrite.WriteLine("padding-left: 8px;");

                sWrite.WriteLine("opacity:0.8;");

                sWrite.WriteLine("padding-top: 8px;");

                sWrite.WriteLine("padding-bottom: 8px;");

                sWrite.WriteLine("padding-right: 8px;");

                sWrite.WriteLine("background-color:#000;}");

                sWrite.WriteLine("</style>");

                //end of CSS

                sWrite.WriteLine("<!doctype html>");

                sWrite.WriteLine("<html>");
                sWrite.WriteLine("<head>");


                sWrite.WriteLine("<script type='text/javascript' src='./pages/js/jquery.min.js'></script>");

                sWrite.WriteLine("<script type='text/javascript' src='./pages/js/jquery.magnifier.js'></script>");


                sWrite.WriteLine("<meta charset='utf-8'>");

                sWrite.WriteLine("<title>Test Case Summary</title>");
                sWrite.WriteLine("</head>");


                sWrite.WriteLine("<body>");
                sWrite.WriteLine("<div align='center' style='height:auto'>");
                sWrite.WriteLine("<table width='1024' border='0'><tr>");
                sWrite.WriteLine("<img alt='ABT' src='./pages/images/ABT.png' width = '280' height='100' align = '<td>left'>");
                sWrite.WriteLine("</tr></table>");


                sWrite.WriteLine("<table width='1024' border='3'>");

                sWrite.WriteLine("<tr><th colspan='5' class='box1' style='border-color:'; scope='col' >TEST CASE SUMMARY</th></tr>");

                sWrite.WriteLine("<tr><th colspan='5' class='gap' scope='col'></th></tr>");

                sWrite.WriteLine("<tr>");

                sWrite.WriteLine("<td class='box2' valign='top'>");

                sWrite.WriteLine("<table border='0'  class='box3'>");

                sWrite.WriteLine("<tr><td colspan='3'>Summary</td></tr>");

                sWrite.WriteLine("<tr><td colspan='3' style='height:5px'>&nbsp;</td></tr>");

                sWrite.WriteLine("<tr><td valign='top'>Test Case Id</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify'>"+strTestCaseID1+"</td></tr>");

                sWrite.WriteLine("<tr><td valign='top'>Test Case Name</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify'>"+strTestCaseName+"</td></tr>");

                sWrite.WriteLine("<tr><td valign='top'>Description</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify'>"+strTestCaseDesc2+"</td></tr>");

                sWrite.WriteLine("<tr><td width='133' valign='top'>Status</td>");

                sWrite.WriteLine("<td width='4' valign='top'>:</td>");

                sWrite.WriteLine("<td width='356' valign='top' align='justify' style='color:green'>NA</td></tr>");

                sWrite.WriteLine("<tr><td valign='top' width = '250'>Verification points Passed</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify' style='color:green'>NA</td></tr>");

                sWrite.WriteLine("<tr><td valign='top' width = '250'>Verification points Failed</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify' style='color:red'>NA</td></tr>");

                sWrite.WriteLine("</table></td>");

                sWrite.WriteLine("<td class='gap'></td>");

                sWrite.WriteLine("<td class='box4' valign='top'>");

                sWrite.WriteLine("<table border='0' class='box3'>");

                sWrite.WriteLine("<tr><td colspan='3'>Execution/Platform Details</td></tr>");

                sWrite.WriteLine("<tr><td colspan='3' style='height:5px'>&nbsp;</td></tr>");

                sWrite.WriteLine("<tr><td width='156' valign='top'>OS</td>");

                sWrite.WriteLine("<td width='4' valign='top'>:</td>");

                sWrite.WriteLine("<td width='316' valign='top' align='justify'>Windows 7</td></tr>");

                sWrite.WriteLine("<tr><td valign='top'>Execution Date</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify'>" + localDate + "</td></tr>");

                sWrite.WriteLine("<tr><td valign='top'>Total Time Taken</td>");

                sWrite.WriteLine("<td valign='top'>:</td>");

                sWrite.WriteLine("<td valign='top' align='justify'>NA</td></tr>");

                sWrite.WriteLine("<tr class='gap'></tr>");

                sWrite.WriteLine("<tr class='gap'></tr>");

                sWrite.WriteLine("<tr class='gap'></tr>");

                sWrite.WriteLine("</table></td></tr>");

                sWrite.WriteLine("<tr><td class='gap'></td></tr>");

                sWrite.WriteLine("<tr><th colspan='5' class='box1' scope='col'>TEST CASE DETAILS</th></tr>");

                sWrite.WriteLine("<tr><td class='gap'></td></tr>");

                sWrite.WriteLine("<tr><td height='-1' colspan='5' valign='top'>");

                sWrite.WriteLine("<table width='1024'  border='1' class='tabledesign' cellspacing='0'>");

                sWrite.WriteLine("<tr valign='top'>");

                sWrite.WriteLine("<td width='48' class='tableheaderstyle'>Sr No</td>");

                sWrite.WriteLine("<td width='237' class='tableheaderstyle'>Step Description</td>");

                sWrite.WriteLine("<td width='131' class='tableheaderstyle'>Input Value</td>");

                sWrite.WriteLine("<td width='201' class='tableheaderstyle'>Expected Result</td>");

                sWrite.WriteLine("<td width='192' class='tableheaderstyle'>Actual Result</td>");

                sWrite.WriteLine("<td width='113' class='tableheaderstyle'>Time Taken in Seconds</td>");
                sWrite.WriteLine("<td width='72' class='tableheaderstyle'>Status</td>");
                sWrite.WriteLine("<td width='113' class='tableheaderstyle'>Screenshot</td>");
                
                //File.CreateText(resultpath + ".html").Flush();
                sWrite.Close();

            }








            // using (FileStream fs = new FileStream(resultpath + ".htm", FileMode.Create))
            // SendKeys.SendWait("%{F}{A}");


        }
        public string[] GetModulesToRun()
        {
            string[] GetModulesToRun = { "Test-Oracle" };
            return GetModulesToRun;
        }
        public void InvokeReporting()
        {


            // string dictTestCases;


            //Get Modules names to run

            string[] strModules = GetModulesToRun();

            //MessageBox.Show(strModules.Length.ToString());
            //arrModules = Split(strModules, ",");
            for (int iModules = 0; iModules < strModules.Length; iModules++)
            {
                strModule = strModules[iModules];
                // MessageBox.Show(strModule);

                string strTestCases = GetTestCaseToRun(strModule);
                //MessageBox.Show(strTestCases);
                string[] arrTestCaseWithSlNo = strTestCases.Split(',');
                //  Dictionary<string, string> dictResults = new Dictionary<string, string>();
                //  Dictionary<string[], string[]> dictSetResults = new Dictionary<string[], string[]>();
                Dictionary<string, string> openWith = new Dictionary<string, string>();

                int len = arrTestCaseWithSlNo.Length - 1;
                //MessageBox.Show(len.ToString());
                // MessageBox.Show(arrTestCaseWithSlNo.Length.ToString());
                for (int iTestCase = 0; iTestCase < len; iTestCase++)
                {
                    string temp = arrTestCaseWithSlNo[iTestCase];
                    // MessageBox.Show(temp);
                    string[] arrTestCase = temp.Split(';');
                    int lent = arrTestCase.Length;
                    //MessageBox.Show(lent.ToString());
                    string strTestCaseSlNo = arrTestCase[0];
                    // MessageBox.Show(strTestCaseSlNo);

                    string strTestCaseID = arrTestCase[1];
                    // MessageBox.Show(strTestCaseID);
                    string strTestCaseName = arrTestCase[2];
                    //MessageBox.Show(strTestCaseName);
                    string strTestCaseDesc = arrTestCase[3];
                    //MessageBox.Show(strTestCaseDesc);


                    string dictResults = ExecuteTestCase(strTestCaseSlNo, strTestCaseID, strTestCaseDesc, strTestCaseName);
                    string[] arrTest = dictResults.Split('|');
                    // MessageBox.Show(arrTest[0]);
                    int len1 = arrTest.Length - 1;
                    // MessageBox.Show(len1.ToString());
                    for (int iT = 0; iT < len1; iT++)

                    {
                        string temp1 = arrTest[iT];
                        // MessageBox.Show(temp1);
                        string[] arrParameters = temp1.Split('!');

                        string arrkeys = arrParameters[0];
                        // MessageBox.Show(arrkeys);
                        string arrItems = arrParameters[1];
                        //openWith.Add(arrkeys, arrItems);

                    }
                    //this.GetType().GetMethod(strTestCaseName).Invoke(this, null);
                    //Dictionary<string, string>.ValueCollection valueColl = openWith.Values;



                    //Task myFirstTask = Task.Factory.StartNew();
                    //Task.Run(() => +strTestCaseName+(openWith);

                    //MessageBox.Show(dictResults.ToString());
                    //string[] arrkeys = dictResults.Keys.ToArray();
                    // string[] arrItems = dictResults.Values.ToArray();
                    // dictSetResults.Add(arrkeys, arrItems);

                }


            }

        }
        public void AGI(Dictionary<string, string> DicArg)
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService
                (@"C:\Users\CarvalhoS\Desktop");
            // properties on the service can be used to e.g. hide the command prompt

            // ChromeOptions options = new ChromeOptions
            //{
            //IgnoreZoomLevel = true
            // };
            IWebDriver driver = new ChromeDriver(service);
            var time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.Navigate().GoToUrl(DicArg["URL"]);
            
            var time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            var t = time2 - time1;
            string p = t.ToString();
            writeResult("1", "opening the AGI", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("Tools & Resources")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("2", "Clicking on Tools & Resources", "", "", "Clicked", "pass", true, p);
            driver.FindElement(By.LinkText("Tools")).Click();
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("ATLAS (Abt Talent, Learning, and Support)")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("3", "Clicking on Abt Talent, Learning and Support (ATLAS)", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("AbtKnowledge")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("4", "Clicking on AbtKnowledge", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("AbtTravel Portal")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("5", "Clicking on AbtTravel Portal", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("Oracle")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("6", "Clicking on Oracle", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("ISMS")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("7", "Clicking on ISMS", "", "", "Clicked", "pass", true, p);
            driver.FindElement(By.LinkText("AGI")).Click();
            driver.FindElement(By.LinkText("Tools & Resources")).Click();
            driver.FindElement(By.LinkText("Tools")).Click();

            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("Outlook Web Access")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("8", "Clicking on Outlook Web Access", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
           
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("Reputational Capital Database")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("9", "Clicking on Reputational Capital Database", "", "", "Clicked", "pass", true, p);
            time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            driver.FindElement(By.LinkText("RepCap Planner")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("10", "Clicking on RepCap Planner", "", "", "Clicked", "pass", true, p);
            driver.FindElement(By.LinkText("Staff Directory")).Click();
            time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            t = time2 - time1;
            p = t.ToString();
            writeResult("11", "Clicking on Staff Directory", "", "", "Clicked", "pass", true, p);
            // time1 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            // driver.FindElement(By.LinkText("Reputational Capital Database")).Click();
            // time2 = DateTime.Now.Ticks / TimeSpan.TicksPerSecond;
            // t = time2 - time1;
            // p = t.ToString();
            // writeResult("1", "Clicking on More...", "", "", "Clicked", "pass", true, p);
        
        }
        public void Timesheet(Dictionary<string, string> DicArg)
        {
            string h = DicArg["From"];
            //MessageBox.Show(h);
            //MessageBox.Show(DicArg["To"]);
           // writeResult("1", "First Test Case1", "", "", "Clicked", "pass", true);
           // writeResult("2", "First Test Case1", "", "", "Clicked", "pass", true);
            //writeResult("3", "First Test Case1", "", "", "Clicked", "pass", true);
          //  writeResult("4", "First Test Case1", "", "", "Clicked", "pass", true);
           // writeResult("5", "First Test Case1", "", "", "Clicked", "pass", true);
            
        }
        public void ServiceNow(Dictionary<string, string> DicArg)
        {
            //MessageBox.Show(DicArg["Person"]);
           // writeResult("1", "First Test Case2", "", "", "Clicked", "pass", true);
           // writeResult("2", "First Test Case2", "", "", "Clicked", "pass", true);
          //  writeResult("3", "First Test Case2", "", "", "Clicked", "pass", true);
           // writeResult("4", "First Test Case2", "", "", "Clicked", "pass", true);
           // writeResult("5", "First Test Case2", "", "", "Clicked", "pass", true);
         
        }
        public void SSignOn(Dictionary<string, string> DicArg)
        {
            //  MessageBox.Show(DicArg["AbtKnowledge"]);
            writeResult("1", "First Test Case3", "", "", "Clicked", "pass", true,"");
            writeResult("2", "First Test Case3", "", "", "Clicked", "pass", true, "");
            writeResult("3", "First Test Case3", "", "", "Clicked", "pass", true, "");
            writeResult("4", "First Test Case3", "", "", "Clicked", "pass", true, "");
            writeResult("5", "First Test Case3", "", "", "Clicked", "pass", true, "");
            
        }

        //Get test cases list for the above modules and execute them

       
        public string GetTestCaseToRun(string strModule)
        {

            Excel.Application excelApp;
            Excel.Workbook excelWorkbook;
            // Excel.Sheets excelSheets;
            Excel.Worksheet excelWorksheet;
            Excel.Range range;
            //string currentSheet = "Sheet1";
            string Testcasepath = "C:\\Users\\CarvalhoS\\Project\\" + strModule + ".xlsx";
            excelApp = new Excel.Application();
            excelWorkbook = excelApp.Workbooks.Open(Testcasepath);// change it to datasheet sheet path later
            //excelSheets = excelWorkbook.Sheets;
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Module");
            //Excel.Range xlRange = excelWorksheet.UsedRange;
            //int col = xlRange.Columns.Count;
            // int row = xlRange.Rows.Count;
            range = excelWorksheet.UsedRange;
            //object misValue = System.Reflection.Missing.Value;
            //int   rCnt = range.Rows.Count;
            //int  cCnt = range.Columns.Count;
            Range excelRange = excelWorksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(
                       XlRangeValueDataType.xlRangeValueDefault);
            int iTotalColumns = excelWorksheet.UsedRange.Columns.Count;
            int iTotalRows = excelWorksheet.UsedRange.Rows.Count;
            string strTestCasesSet = null;
            string strSlNo = string.Empty;
            string strTestCaseID = string.Empty;
            string strTestCaseName = string.Empty;
            string strTestCaseDesc = string.Empty;
            //string strSlNo = string.Empty;

            for (int i = 2; i <= iTotalRows; i++)
            {
                string strExecute = valueArray[i, 5].ToString();



                if (strExecute == "YES")
                {

                    strSlNo = valueArray[i, 1].ToString();
                    //Debug.Print(strSlNo);

                    strTestCaseID = valueArray[i, 2].ToString();


                    strTestCaseName = valueArray[i, 3].ToString();


                    strTestCaseDesc = valueArray[i, 4].ToString();

                    strTestCasesSet = strTestCasesSet + strSlNo + ";" + strTestCaseID + ";" + strTestCaseName + ";" + strTestCaseDesc + ",";


                }


            }
            //MessageBox.Show(strTestCasesSet);

            excelWorkbook.Close();// change missValue to null
            excelApp.Quit();

            releaseObject(excelWorksheet);
            releaseObject(excelWorkbook);
            releaseObject(excelApp);
            var process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }

            return strTestCasesSet;
        }
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public string ExecuteTestCase(string strTestCaseSlNo, string strTestCaseID, string strTestCaseDesc, string strTestCaseName)
        {


            string Testcasepath = "C:\\Users\\CarvalhoS\\Project\\" + strModule + ".xlsx";
            //string strModule = "Module";
            // string strExcelSheetPath = Testcasepath + strModule + ".xlsx";
            Dictionary<string, string> dictParameters = new Dictionary<string, string>();
            Dictionary<string, string> dictResults = new Dictionary<string, string>();
            Dictionary<string, string> dictIterationResults = new Dictionary<string, string>();
            //dict.Add("one", 1);
            //int rCnt = 0;
            //int cCnt = 0;

            Excel.Application excelApp;
            Excel.Workbook excelWorkbook;
            // Excel.Sheets excelSheets;
            Excel.Worksheet excelWorksheet;
            Excel.Range range;
            //string Testcasepath = "C:\\Users\\CarvalhoS\\Documents\\theExample.xlsx";
            excelApp = new Excel.Application();
            excelWorkbook = excelApp.Workbooks.Open(Testcasepath);// change it to datasheet sheet path later
            //excelSheets = excelWorkbook.Sheets;
            //MessageBox.Show(strTestCaseName);
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(strTestCaseName);
            range = excelWorksheet.UsedRange;
            Range excelRange = excelWorksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(
                       XlRangeValueDataType.xlRangeValueDefault);
            //object misValue = System.Reflection.Missing.Value;
            // rCnt = range.Rows.Count;
            //cCnt = range.Columns.Count;
            int iTotalColumns = excelWorksheet.UsedRange.Columns.Count;
            //MessageBox.Show(iTotalColumns.ToString());
            int iTotalRows = excelWorksheet.UsedRange.Rows.Count;
            // MessageBox.Show(iTotalRows.ToString());
            // int  iIterations = GetIterations(excelWorksheet);
            int iIterations = 0;
            string setparaValue = null;
            for (int k = 2; k <= iTotalRows; k++)
            {
                string sExecute = valueArray[k, 1].ToString();
                // MessageBox.Show(sExecute);
                if (sExecute == "Y")
                {
                    iIterations = iIterations + 1;
                }

            }
            // MessageBox.Show(iIterations.ToString());
            //int iIterations = 1; // for now keep it one
            int iCurrentIteration = 0;


            // bool iFlag = true;
            DateTime iStartTime  = DateTime.Now;
            DateTime iTCStartTime = DateTime.Now;
           
            //MessageBox.Show(iTCStartTime.ToString());
            for (int i = 2; i <= iTotalRows; i++)
            {
                string mExecute = valueArray[i, 1].ToString();


                if (mExecute == "Y")
                {

                    iCurrentIteration = iCurrentIteration + 1;
                    if (iIterations == 1)
                    {
                        CreateResultStructure(0, strTestCaseName);
                    }
                    else
                    {
                        CreateResultStructure(iCurrentIteration, strTestCaseName);
                    }
                }


                

                // Environment.SetEnvironmentVariable."StepPassCount") = 0;
                Environment.SetEnvironmentVariable("StepPassCount", "0");
                Environment.SetEnvironmentVariable("StepFailCount", "0");
                ResultReport(strTestCaseID, strTestCaseDesc, strTestCaseName);

                dictParameters.Clear();
                // dictParameters.Add("Row", rCnt);
                for (int iCol = 2; iCol <= iTotalColumns; iCol++)
                {
                    string sParam = valueArray[1, iCol].ToString();
                    //MessageBox.Show(sParam);


                    string sValue = valueArray[i, iCol].ToString();
                    //  MessageBox.Show(sValue);

                    setparaValue = setparaValue + sParam + "!" + sValue + "|";
                    dictParameters.Add(sParam, sValue);

                    // MessageBox.Show(dictParameters.ToString());



                }
                string[] arrkeys = dictParameters.Keys.ToArray();
                string[] arrItems = dictParameters.Values.ToArray();
                // string ma = "Hie";
                // string p= strTestCaseName+"(ma)";
                // string[] parameters = { "Sam", "Perls" };
                //Dictionary<string, string>.ValueCollection valueColl = dictParameters.Values;
                //object[] args = new object[myArgs.Length];
                //this.GetType().GetMethod(strTestCaseName).Invoke(this,
                // new object[] { arrItems });
                this.GetType().GetMethod(strTestCaseName).Invoke(this,
                                         new object[] { dictParameters });
                // MessageBox.Show(setparaValue);
                
            }
            string value;
            string value1;
            value = Environment.GetEnvironmentVariable("StepPassCount");
            value1 = Environment.GetEnvironmentVariable("StepFailCount");
            int num1 = Convert.ToInt32(value);
            int num2 = Convert.ToInt32(value1);

            DateTime iTCStopTime = DateTime.Now;
            TimeSpan diff = iTCStopTime-iTCStartTime ;
            string diffInSeconds =  string.Format(
                       CultureInfo.CurrentCulture,
                       "{0} days, {1} hours, {2} minutes, {3} seconds",
                       diff.Days,
                       diff.Hours,
                       diff.Minutes,
                       diff.Seconds);


           // MessageBox.Show(diffInSeconds);
            //string diffInSeconds1 = diffInSeconds.ToString();
            CreateSummary(num1,num2,diffInSeconds);
            //MessageBox.Show(value);

            excelWorkbook.Close();// change missValue to null
            excelApp.Quit();

            releaseObject(excelWorksheet);
            releaseObject(excelWorkbook);
            releaseObject(excelApp);
            return setparaValue;
           

        }





        public void CreateResultStructure(int Iteration,string strTestCaseName)
        {
            DateTime dtTimeStamp = DateTime.Now;


            if (Iteration == 0)
            {
                strResultReportFile = resultpath + ".txt";
            }


            else
            {
                System.Threading.Thread.Sleep(1000);
                // MessageBox.Show(Iteration.ToString());
                strResultReportFile = resultpath + Iteration + ".txt";
            }



            CreateFolders(strTestCaseName);
        }
        public void CreateFolders(string strTestCaseName)

        {
           // string strTestCaseName = "Hello";//test case name for now
                                             //FileSystemObject fso = new FileSystemObject();
            DateTime dtTimeStamp = DateTime.Now;
            string month = dtTimeStamp.ToString("MMMM");
            string day = dtTimeStamp.ToString("yyyy-MM-dd");
            // System.IO.Directory.CreateDirectory(@"C:\Users\CarvalhoS\Desktop\Project\" + month);
            string sReportPath = @"C:\Users\CarvalhoS\Desktop\Results\" + month;
            //string day = String.Format("{D}", DateTime.Now);
            string des = @"C:\Users\CarvalhoS\Desktop\Results\pages\";
            if (!Directory.Exists(sReportPath))
            {
                Directory.CreateDirectory(sReportPath);
                copyDirectory(des, sReportPath);
                // System.IO.Directory.(des, sReportPath, true);

            }

            string pathstring = System.IO.Path.Combine(sReportPath, day);
            string destDirName1 = System.IO.Path.Combine(sReportPath, "pages");

            if (!Directory.Exists(pathstring))
            {
                Directory.CreateDirectory(pathstring);
                Directory.CreateDirectory(destDirName1);
                copyDirectory(des, destDirName1);
                //System.IO.File.Copy(des, pathstring, true);
            }
            pathstring1 = System.IO.Path.Combine(pathstring, strTestCaseName);
            string destDirName = System.IO.Path.Combine(pathstring, "pages");
            string destDirName2 = System.IO.Path.Combine(pathstring1, "pages");
            destDirName3 = System.IO.Path.Combine(pathstring1, "Screenshots");
            if (!Directory.Exists(pathstring1))
            {
                Directory.CreateDirectory(pathstring1);
                Directory.CreateDirectory(destDirName);
                Directory.CreateDirectory(destDirName2);
                Directory.CreateDirectory(destDirName3);
                copyDirectory(des, destDirName);
                copyDirectory(des, destDirName2);
                // System.IO.File.Copy(des, pathstring1, true);
            }

        }

        // Copy directory structure recursively
        public static void copyDirectory(string Src, string Dst)
        {
            String[] Files;

            if (Dst[Dst.Length - 1] != Path.DirectorySeparatorChar)
                Dst += Path.DirectorySeparatorChar;
            if (!Directory.Exists(Dst)) Directory.CreateDirectory(Dst);
            Files = Directory.GetFileSystemEntries(Src);
            foreach (string Element in Files)
            {
                // Sub directories
                if (Directory.Exists(Element))
                    copyDirectory(Element, Dst + Path.GetFileName(Element));
                // Files in directory
                else
                    File.Copy(Element, Dst + Path.GetFileName(Element), true);
            }
        }



        public void  writeResult(string iSrNo, string strStepDescription, string strTestData, string strExpResult, string strActResult, string strStatus, bool isScreenshot,string strTimeTaken)
        {
            string Photopath;
            //DateTime localDate1 = DateTime.Now;
            //string result2 = localDate1.ToString("yyyyMMddHH");
            // string resultpath = "C:\\Users\\CarvalhoS\\Desktop\\test1" + result2;
            using (StreamWriter sWrite = File.AppendText(strResultReportFile + ".txt"))
            {

                //bool isScreenshot = true;
                if (isScreenshot == true)
                {
                    DateTime localDate = DateTime.Now;
                    string sTimestamp = localDate.ToString("yyyyMMddHHmmss");

                    Bitmap captureBitmap = new Bitmap(1800, 1000, PixelFormat.Format32bppArgb);

                    System.Drawing.Rectangle captureRectangle = Screen.AllScreens[0].Bounds;

                    Graphics captureGraphics = Graphics.FromImage(captureBitmap);

                    captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size);
                    Photopath = destDirName3 + "\\Capture" + sTimestamp;

                    captureBitmap.Save(Photopath + ".png", ImageFormat.Png);

                    sWrite.WriteLine("<tr class='tablestyle' valign='top'>");


                    sWrite.WriteLine("<td >" + iSrNo + "</td>");

                    sWrite.WriteLine("<td >" + strStepDescription + "</td>");

                    sWrite.WriteLine("<td >" + strTestData + "</td>");

                    sWrite.WriteLine("<td >" + strExpResult + "</td>");

                    sWrite.WriteLine("<td >" + strActResult + "</td>");
                    sWrite.WriteLine("<td >" + strTimeTaken + " </td>");
                    //string strStatus = "pass";
                    //int StepPassCount = 0;
                    //int StepFailCount = 0;
                    if (strStatus == "pass")
                    {
                        sWrite.WriteLine("<td><font color='Green'>PASS</td>");

                        string value;

                        value = Environment.GetEnvironmentVariable("StepPassCount");
                        int numVal = Convert.ToInt32(value);
                        numVal = numVal + 1;
                        Environment.SetEnvironmentVariable("StepPassCount", numVal.ToString());
                       // StepPassCount = StepPassCount + 1;
                    }



                    else
                    {
                        sWrite.WriteLine("<td><font color='Red'>FAIL</td>");
                        string value;

                        value = Environment.GetEnvironmentVariable("StepFailCount");
                        int numVal = Convert.ToInt32(value);
                        numVal = numVal + 1;
                        Environment.SetEnvironmentVariable("StepFailCount", numVal.ToString());

                    }

                    





                    sWrite.WriteLine("<td >");

                    // bool isScreen = true;
                    if (isScreenshot == true)
                    {
                        //string strRelativePath = Photopath + ".png";

                        string strRelativePath = "./Screenshots/Capture" + sTimestamp + ".png";

                        sWrite.WriteLine("<img src='" + strRelativePath + "' border='1' style='width: 50px; height: 30px' />");

                        sWrite.WriteLine("<a href=#' rel= 'magnify[sc_" + sTimestamp + "]' >Zoom in</a>");
                        sWrite.WriteLine("</td></tr>");
                        sWrite.WriteLine("<img class='magnify' type = 'hidden' id='sc_" + sTimestamp + "' src='" + strRelativePath + "' border='1' data-magnifyby='20' style='width: 50px; height: 30px; display: none; ' />");
                        
                         


                    }


                    else
                    {
                        sWrite.WriteLine("No Screenshot available");
                       
                     
                    }
                   
                    sWrite.WriteLine("</td ></tr>");





                    if (isScreenshot == true)
                    {

                        //sWrite.WriteLine("<img class='magnify' type = 'hidden' id='sc_" + sTimestamp + " src='");

                        //sWrite.WriteLine(strRelativePath + " border='1' data-magnifyby='20' style='width: 50px; height: 30px; display: none; ' />");
                    }
                    // sWrite.WriteLine("</table></td>");

                    //sWrite.WriteLine("</tr>");

                    //sWrite.WriteLine("</table>");

                    //sWrite.WriteLine("</div></body>	</html>");






                    // sWrite.WriteLine("</tr>");


                 

                }

               
            }
          
        }
        public void CreateSummary(int iPassCount, int iFailCount, string timeElapsed)
        {
            //MessageBox.Show(iPassCount.ToString());
            //MessageBox.Show(iFailCount.ToString());
            //MessageBox.Show(timeElapsed);
            if (iFailCount > 0)

            {
                test = "<td width='356' valign='top' align='justify' style='color:green'>FAIL</td></tr>";

            }
              
            
            else if (iFailCount==0)
            {
                test = "<td width='356' valign='top' align='justify' style='color:green'>PASS</td></tr>";
            }
                
             
            
            string strVerificationPointsPassText = "<td valign='top' align='justify' style='color:green'>" + iPassCount + "</td></tr>";
           // MessageBox.Show(strVerificationPointsPassText);

            string strVerificationPointsFailText = "<td valign='top' align='justify' style='color:red'> "+ iFailCount+"</td></tr>";
           // MessageBox.Show(strVerificationPointsFailText);

            string TimeTakenText = "<td valign='top' align='justify'>" + timeElapsed + "  </td></tr>";

            int PercentageOfPassFail;
            PercentageOfPassFail = ((iPassCount * 100) / (iPassCount + iFailCount));
            string iPercentageOfPassFail = PercentageOfPassFail.ToString() + "%";

            // StreamReader reading = File.OpenText(strResultReportFile + ".txt");
          
            
                string text = File.ReadAllText(strResultReportFile +".txt");
            if (text.Contains("<td width='356' valign='top' align='justify' style='color:green'>NA</td></tr>"))
            {
               // MessageBox.Show("found");
                text = text.Replace("<td width='356' valign='top' align='justify' style='color:green'>NA</td></tr>", test);
                //File.WriteAllText(strResultReportFile + ".txt", strReportText);
            }
           
            if (text.Contains("<td valign='top' align='justify' style='color:red'>NA</td></tr>"))
             {
             // MessageBox.Show("found");
              text = text.Replace("<td valign='top' align='justify' style='color:red'>NA</td></tr>", "<td valign='top' align='justify' style='color:red'> " + iFailCount + "</td></tr>");
           //File.WriteAllText(strResultReportFile + ".txt", "<td valign='top' align='justify' style='color:green'>" + iFailCount + "</td></tr>");
             }
    
            if (text.Contains("<td valign='top' align='justify' style='color:green'>NA</td></tr>"))
            {
            // MessageBox.Show("found");
             text = text.Replace("<td valign='top' align='justify' style='color:green'>NA</td></tr>", "<td valign='top' align='justify' style='color:green'>" + iPassCount + "</td></tr>");
            //File.WriteAllText(strResultReportFile + ".txt", "<td valign='top' align='justify' style='color:green'>" + iPassCount + "</td></tr>");
            }
      
            if (text.Contains("<td valign='top' align='justify'>NA</td></tr>"))
            {
               //MessageBox.Show("found");
               text = text.Replace("<td valign='top' align='justify'>NA</td></tr>", TimeTakenText);
             
            }

            File.WriteAllText(strResultReportFile + ".txt", text);

            texttohtml();


        }
        public void texttohtml()
        {
            using (StreamWriter sWrite = File.AppendText(strResultReportFile + ".txt"))
            {
              //  MessageBox.Show("sECOND");
                sWrite.WriteLine("</table></td>");

                sWrite.WriteLine("</tr>");

                sWrite.WriteLine("</table>");

                sWrite.WriteLine("</div></body>	</html>");
            }
            File.Copy(strResultReportFile + ".txt", strResultReportFile + ".html");
            File.Delete(strResultReportFile + ".txt");

        }

    }
}
