using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using System.Xml.Linq;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;

namespace VSource
{
    public partial class Form1 : Form
    {

        public bool employeeIDRequest { get; set; }
        public bool lastFirstMIRequest { get; set; }
        public bool socialSecurityRequest { get; set; }
        public string internalCode { get; set; }
        public string externallCode { get; set; }
        public bool firstLastNameRequest { get; set; }
        
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                employeeIDRequest = true;
            }
            else
            {
                employeeIDRequest = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            internalCode = textBox1.Text;
            externallCode = textBox2.Text;
            
            // MessageBox.Show(socialSecurityRequest.ToString());
            try
            {
                string pasheet;

                var xpasheet = new XElement("eConnect",
                    new XElement("RQeConnectOutType",
                        new XElement("eConnectProcessInfo",
                          new XElement("Outgoing", "TRUE"),
                          new XElement("MessageID", "Employee")),
                        new XElement("eConnectOut",
                          new XElement("DOCTYPE", "Employee"),
                          new XElement("OUTPUTTYPE", "2"),
                          new XElement("FORLOAD", "0"),
                          new XElement("FORLIST", "1"),
                          new XElement("ACTION", "0"),
                          new XElement("ROWCOUNT", "4000"),
                          new XElement("REMOVE", "0"))));

                pasheet = xpasheet.ToString();

                // Create a connection string to specify the Microsoft Dynamics GP server and database
                // Change the data source and initial catalog to specify your server and database
                string sConnectionString = @"data source=VM-GPD01;initial catalog=TMI;integrated security=SSPI;persist security info=False;packet size=4096";

                // Create an eConnectMethods object
                eConnectMethods requester = new eConnectMethods();

                // Call the eConnect_Requester method of the eConnectMethods object to retrieve specified XML data
                string reqDoc = requester.eConnect_Requester(sConnectionString, EnumTypes.ConnectionStringType.SqlClient, pasheet);

                // Display the result of the eConnect_Requester method call
                //Console.Write(reqDoc);


                XElement blob = XElement.Parse(reqDoc);



                //Make Dictionary List of All Departments;

                var deptDictionary = new Dictionary<string, string>()
                {
                  {"ACCT","25-" + internalCode +"-10-110-01"},
                {"ADMIN","25-" + internalCode +"-10-140-01"},
                {"AMBO","25-" + internalCode +"-40-210-06"},
                {"AMID","25-" + internalCode +"-30-210-01"},
                {"AMSC","25-" + internalCode +"-20-210-01"},
                {"ARCMGT","25-" + externallCode +"-80-900-01"},
                {"CNTADM","25-" + internalCode +"-10-135-01"},
                {"CNTAVD","25-" + externallCode +"-20-800-01"},
                {"CNTRBO","25-" + externallCode +"-40-500-06"},
                {"CNTRCO","25-" + externallCode +"-20-500-08"},
                {"CNTRCT","25-" + externallCode +"-20-500-01"},
                {"CNTRID","25-" + externallCode +"-20-500-06"},
                {"CNTRMN","25-" + externallCode +"-20-500-07"},
                {"CNTRNC","25-" + externallCode +"-40-500-03"},
                {"CNTRSF","25-" + externallCode +"-20-500-05"},
                {"CNTRTX","25-" + externallCode +"-20-500-03"},
                {"DMBO","25-" + internalCode +"-20-215-06"},
                {"EXEC","25-" + internalCode +"-10-100-01"},
                {"EXECEN","25-" + internalCode +"-10-100-04"},
                {"FNANCE","25-" + internalCode +"-10-130-01"},
                {"HR","25-" + internalCode +"-10-120-01"},
                {"HR ID","25-" + internalCode +"-12-120-06"},
                {"INFRBO","25-" + externallCode +"-30-700-06"},
                {"INPRBO","25-" + internalCode +"-30-290-06"},
                {"INSRBO","25-" + externallCode +"-30-300-06"},
                {"INSRSC","25-" + externallCode +"-30-300-01"},
                {"MGDEXC","25-" + internalCode +"-30-280-01"},
                {"MIS","25-" + internalCode +"-10-160-01"},
                {"MKTSAL","25-" + internalCode +"-20-205-01"},
                {"MRKT","25-" + internalCode +"-10-170-01"},
                {"OCIOSC","25-" + externallCode +"-60-800-01"},
                {"OPEXEC","25-" + internalCode +"-10-139-01"},
                {"PRJMGT","25-" + internalCode +"-20-229-01"},
                {"RCRTBO","25-" + internalCode +"-10-220-06"},
                {"RCRTEN","25-" + internalCode +"-10-220-04"},
                {"RCRTM","25-" + internalCode +"-20-220-01"},
                {"RCRTME","25-" + internalCode +"-10-228-01"},
                {"RCRTU","25-" + internalCode +"-10-222-01"},
                {"RCRUIT","25-" + internalCode +"-10-220-01"},
                {"RMBO","25-" + internalCode +"-10-222-06"},
                {"SALEBO","25-" + internalCode +"-40-200-06"},
                {"SALEMG","25-" + internalCode +"-20-209-01"},
                {"SALES","25-" + internalCode +"-20-200-01"},
                {"SMSC","25-" + internalCode +"-20-235-01"},
                {"TRANBO","25-" + externallCode +"-30-310-06"},
                {"TSG","25-" + internalCode +"-20-230-01"},
                {"SALEEN","25-" + internalCode +"-20-234-01"},
                {"TSGMGS","25-" + externallCode +"-60-600-01"},
                {"CNTRNY","25-" + externallCode +"-20-500-02"},
                {"","MISSING"}

                };




                //cool stuff down here





                int queryCount = blob.Elements("eConnect").Elements("Employee").Count();


                string[,] employeeList = new string[queryCount, 10];

                int rowCount = 0;
                int columnCount = 0;
                //Console.WriteLine(blob.ToString());

                foreach (XElement employee in blob.Elements("eConnect").Elements("Employee"))
                {
                    if (employee.Element("INACTIVE").Value == "0")
                    {
                        

                       string mName;

                       if (employeeIDRequest == true)
                       {
                           
                           employeeList[rowCount, columnCount] = employee.Element("FRSTNAME").Value + " " + employee.Element("LASTNAME").Value;
                           columnCount += 1;
                       }
                       if (employeeIDRequest == true)
                       {
                          
                           employeeList[rowCount, columnCount] = employee.Element("FRSTNAME").Value + " " + employee.Element("LASTNAME").Value;
                           
                           columnCount += 1;
                       }
                       
                        
                        
                       if (lastFirstMIRequest == true)
                       {
                           if (employee.Element("MIDLNAME").Value == "")
                           { mName = ""; }
                           else
                           { mName = employee.Element("MIDLNAME").Value.Substring(0, 1) + "."; }

                           employeeList[rowCount, columnCount] = employee.Element("LASTNAME").Value + ", " +
                               employee.Element("FRSTNAME").Value + " " +
                                mName;

                           columnCount += 1;
                       }

                        employeeList[rowCount, columnCount] = employee.Element("DEPRTMNT").Value;
                        columnCount += 1;

                        if (socialSecurityRequest == true)
                        {
                            Func<string, string> socialize = s =>
                                {
                                    if (employee.Element("SOCSCNUM").Value == null)
                                    {
                                        s = "";
                                        return s;
                                    }
                                    else
                                    {
                                        s = employee.Element("SOCSCNUM").Value;
                                        s = s[0] + s[1] + s[2] + "-" + s[3] + s[4] + "-" + s[5] + s[6] + s[7] + s[8];
                                        return s;
                                    }

                                };
                            employeeList[rowCount, columnCount] = socialize(employee.Element("SOCSCNUM").Value);
                            columnCount += 1;
                        } 
                        
                        
                        
                        
                        try
                        {
                            employeeList[rowCount, columnCount] = deptDictionary[employee.Element("DEPRTMNT").Value];
                            columnCount += 1;
                        }
                        catch
                        {
                            employeeList[rowCount, columnCount] = "MISSING";
                            columnCount += 1;
                        }

                       
                        rowCount += 1;
                        columnCount = 0;
                    }
                }
               



                //doing excel stuff here
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;

                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "EMPLOYEE ID";
                oSheet.Cells[1, 2] = "THI FORMAT NAME";
                oSheet.Cells[1, 3] = "Department";
                oSheet.Cells[1, 4] = "Code";

                string endCornerSheet = "G" + (queryCount + 1).ToString();
                Excel.Range employeeRange = oSheet.Range["A2", endCornerSheet];

                employeeRange.Value2 = employeeList;

                //oRng = oSheet.get_Range("A2", endCornerSheet).Value2 = employeeRange;
                oRng = oSheet.get_Range("A2", endCornerSheet).Value2 = employeeList;

               


            }


            catch (Exception ex)
            {// Dislay any errors that occur to the console
               // MessageBox.Show(ex.ToString());
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                lastFirstMIRequest = true;
            }
            else
            {
                lastFirstMIRequest = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                socialSecurityRequest = true;
            }
            else
            {
                socialSecurityRequest = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                firstLastNameRequest = true;
            }
            else
            {
                firstLastNameRequest = false;
            }
        }
    }
}
