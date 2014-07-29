using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;




namespace SmokePPD
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class CodedUITest1
    {

        #region Declarations
        public static int methodId = 0;
        public static string pageName = string.Empty;
        public static string methodName = string.Empty;
        public static string link = string.Empty;
        public static string methodDescription = string.Empty;
        public static string methodResult = string.Empty;
        public static string errorMessage = string.Empty;
        public static string mailMessageBody = string.Empty;
        public static DateTime localDate = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.Now, TimeZoneInfo.Local.Id, "India Standard Time");
        
        #endregion

        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            methodId = 0;
            mailMessageBody += "<body>"; ;
            mailMessageBody = "<h3 style=\"font-family:Segoe UI\">PPD Smoke Test Automation as on " + localDate.ToString("MMM dd, yyyy") + "</h3>";
            mailMessageBody += "<table>";
            mailMessageBody += "<tr>";
            mailMessageBody += "<th>Test Case ID</th>";
            mailMessageBody += "<th>Application</th>";
            mailMessageBody += "<th>Test Method Name</th>";
            mailMessageBody += "<th>Description</th>";
            mailMessageBody += "<th>Project</th>";
            mailMessageBody += "<th>Status</th>";
            mailMessageBody += "<th>Error Message</th>";
            mailMessageBody += "</tr>";
        }

        [TestInitialize]
        public void MyTestInitialize()
        {
            //methodId = 0;
            pageName = string.Empty;
            pageName = string.Empty;
            methodDescription = string.Empty; 
            methodResult = string.Empty; 
            errorMessage = string.Empty;
            link = string.Empty;
        }

        [TestMethod]
        public void PPDDashboadPage_Staging_And_Production()
        {
            BrowserWindow browser = new BrowserWindow();
            bool isUIClicked = false;
            bool isStagingClicked = false;
            bool isProductionClicked = false;
            int uIClickCount = 0;


            browser = BrowserWindow.Launch("about:tabs");
            browser.Maximized = true;
            methodDescription = "This test method verifies whether the PPD Dashboard pages in staging and production environment exists, loads succesfully";



            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        if (!isStagingClicked)
                        {
                            #region For staging dashboard page

                            if (uIClickCount.Equals(0))
                            {
                                link = ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString();
                                browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString()));

                                Playback.Wait(15000);
                            }
                            else
                            {
                                browser.Refresh();
                                Playback.Wait(10000);
                            }

                            HtmlDiv itemGeographyFilter_staging = new HtmlDiv(browser);
                            itemGeographyFilter_staging.SearchProperties.Add(HtmlDiv.PropertyNames.Id, "cascadedGeographyFilter-Selected");

                            Mouse.Click(itemGeographyFilter_staging);

                            isStagingClicked = true;
                            #endregion 
                        }

                        #region For production dashboard page

                        if (!isProductionClicked)
                        {
                            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Production"].ToString()));

                            Playback.Wait(15000);
                        }
                        else
                        {
                            browser.Refresh();
                            Playback.Wait(10000);
                        }

                        HtmlDiv itemGeographyFilter_production = new HtmlDiv(browser);
                        itemGeographyFilter_production.SearchProperties.Add(HtmlDiv.PropertyNames.Id, "cascadedGeographyFilter-Selected");

                        Mouse.Click(itemGeographyFilter_production);
                        #endregion

                        isUIClicked = true;
                        isProductionClicked = true;
                    }
                    catch (Exception exception)
                    {                        
                        BrowserWindow.ClearCache();
                        browser.Refresh();

                        if(++uIClickCount > 5 )
                        {
                            errorMessage = exception.Message;
                            throw new Exception(exception.Message);
                        }                        
                    }

                }
            }
            catch (Exception exception)
            {

                throw new Exception(exception.Message);
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }
        }

        [TestCleanup]
        public void MyTestCleanUp()
        {            
            methodId++;
            methodName = this.testContextInstance.TestName;
            methodResult = this.testContextInstance.CurrentTestOutcome.ToString();
            mailMessageBody += "<tr>";
            mailMessageBody += "<td>" + methodId + "</td>";
            mailMessageBody += "<td><a href=\" " + link + "\">" + methodName + "</a></td>";
            mailMessageBody += "<td>" + methodName + "()</td>";
            mailMessageBody += "<td>" + methodDescription + "</td>";
            mailMessageBody += "<td>" + "PPD" + "</td>";
            mailMessageBody += "<td>" + methodResult + "</td>";
            mailMessageBody += "<td>" + errorMessage + "</td>";
            mailMessageBody += "</tr>";
            //SendMail(mailMessageBody, "PPD Smoke Test : Automation as on " + localDate.ToString("MMM dd, yyyy"));
        }

        [ClassCleanup]
        public static void MyClassCleanUp()
        {
            mailMessageBody += "</table>";
            mailMessageBody += "</body>";
            SendMail(mailMessageBody, "PPD Smoke Test : Automation as on " + localDate.ToString("MMM dd, yyyy"));
        }

        [TestMethod]
        public void ExportToExcelOnPPD()
        {
            BrowserWindow browser = new BrowserWindow();
            bool isUIClicked = false;
            bool isStagingClicked = false;
            bool isProductionClicked = false;
            int uIClickCount = 0;


            browser = BrowserWindow.Launch("about:tabs");
            browser.Maximized = true;
            methodDescription = "This test method verifies whether the Export to Excel Option on  PPD Dashboard pages in staging and production environment exists.";


            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        if (!isStagingClicked)
                        {
                            #region For staging dashboard page

                            if (uIClickCount.Equals(0))
                            {
                                link = ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString();
                                browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString()));

                                Playback.Wait(15000);
                            }
                            else
                            {
                                browser.Refresh();
                                Playback.Wait(10000);
                            }

                            HtmlSpan itemExportToSpan_staging = new HtmlSpan(browser);
                            itemExportToSpan_staging.SearchProperties.Add(HtmlSpan.PropertyNames.Id, "exportToExcel");

                            Mouse.Click(itemExportToSpan_staging.GetChildren()[0]);

                            Playback.Wait(2000);


                            isStagingClicked = true;
                            #endregion
                        }

                        #region For production dashboard page

                        if (!isProductionClicked)
                        {
                            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Production"].ToString()));

                            Playback.Wait(15000);
                        }
                        else
                        {
                            browser.Refresh();
                            Playback.Wait(10000);
                        }

                        HtmlSpan itemExportToSpan_production = new HtmlSpan(browser);
                        itemExportToSpan_production.SearchProperties.Add(HtmlSpan.PropertyNames.Id, "exportToExcel");

                        Mouse.Click(itemExportToSpan_production.GetChildren()[0]);

                        Playback.Wait(2000);
                        #endregion

                        isUIClicked = true;
                        isProductionClicked = true;
                    }
                    catch (Exception exception)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();

                        if (++uIClickCount > 5)
                        {
                            errorMessage = exception.Message;
                            throw new Exception(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {

                throw new Exception(exception.Message);
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }
        }

        [TestMethod]
        public void ExportToPPTOnPPD()
        {

            BrowserWindow browser = new BrowserWindow();
            bool isUIClicked = false;
            bool isStagingClicked = false;
            bool isProductionClicked = false;
            int uIClickCount = 0;


            browser = BrowserWindow.Launch("about:tabs");
            browser.Maximized = true;

            methodDescription = "This test method verifies whether the Export to PPT Option on  PPD Dashboard pages in staging and production environment exists.";

            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        if (!isStagingClicked)
                        {
                            #region For staging dashboard page

                            if (uIClickCount.Equals(0))
                            {
                                link = ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString();
                                browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Staging"].ToString()));

                                Playback.Wait(15000);
                            }
                            else
                            {
                                browser.Refresh();
                                Playback.Wait(10000);
                            }

                            HtmlSpan itemExportToSpan_staging = new HtmlSpan(browser);
                            itemExportToSpan_staging.SearchProperties.Add(HtmlSpan.PropertyNames.Id, "exportToPPT");

                            Mouse.Click(itemExportToSpan_staging.GetChildren()[0]);

                            Playback.Wait(2000);


                            isStagingClicked = true;
                            #endregion
                        }

                        #region For production dashboard page

                        if (!isProductionClicked)
                        {
                            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDDashboardPage_Production"].ToString()));

                            Playback.Wait(15000);
                        }
                        else
                        {
                            browser.Refresh();
                            Playback.Wait(10000);
                        }

                        HtmlSpan itemExportToSpan_production = new HtmlSpan(browser);
                        itemExportToSpan_production.SearchProperties.Add(HtmlSpan.PropertyNames.Id, "exportToPPT");

                        Mouse.Click(itemExportToSpan_production.GetChildren()[0]);

                        Playback.Wait(2000);
                        #endregion

                        isUIClicked = true;
                        isProductionClicked = true;
                    }
                    catch (Exception exception)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();

                        if (++uIClickCount > 5)
                        {
                            errorMessage = exception.Message;
                            throw new Exception(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {

                throw new Exception(exception.Message);
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }
        }

        [TestMethod]
        public void VerifyHeatMapPage()
        {
            bool isControlClicked = false;
            short controlClickCount = 0;

            BrowserWindow openedBrowser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["PPDHeatMapPage_Production"];
            openedBrowser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDHeatMapPage_Production"]));
            methodDescription = "This test method verifies whether the PPD Dashboard HeatMap page loads successfully.";
            Playback.Wait(25000);

            try
            {
                #region while loop

                while (isControlClicked.Equals(false) && controlClickCount < 5)
                {
                    try
                    {
                        // Waiting for 30 seconds
                        System.Threading.Thread.Sleep(25000);

                        #region Attempting to click on controls

                        // Trying to click on geography filter drop down

                        HtmlDiv geographyFilterDropDown = new HtmlDiv(openedBrowser);
                        geographyFilterDropDown.SearchProperties.Add(HtmlDiv.PropertyNames.Id, "geoDropDownLabel");

                        // Clicking geography drop down
                        Mouse.Click(geographyFilterDropDown);

                        // Resetting parameters
                        controlClickCount = 0;
                        isControlClicked = true;

                        #endregion
                    }
                    catch (Exception)
                    {
                        controlClickCount++;

                        BrowserWindow.ClearCache();
                        BrowserWindow.ClearCookies();
                        openedBrowser.Refresh();

                        if (controlClickCount >= 5)
                        {
                            throw;
                        }

                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            finally
            {
                openedBrowser.Dispose();
                openedBrowser.Close();
            }      
        }

        [TestMethod]
        public void ExportToExcelOnHeatMap()
        {
            bool isControlClicked = false;
            short controlClickCount = 0;

            BrowserWindow openedBrowser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["PPDHeatMapPage_Production"];
            openedBrowser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["PPDHeatMapPage_Production"]));

            methodDescription = "This test method verifies whether Export to Excel feature in the PPD Dashboard HeatMap page executes successfully.";

            Playback.Wait(25000);

            try
            {
                #region while loop

                while (isControlClicked.Equals(false) && controlClickCount < 5)
                {
                    try
                    {
                        // Waiting for 30 seconds
                        System.Threading.Thread.Sleep(25000);

                        #region Attempting to click on controls

                        // Trying to click on export to excel

                        HtmlSpan itemExportToSpan_staging = new HtmlSpan(openedBrowser);
                        itemExportToSpan_staging.SearchProperties.Add(HtmlSpan.PropertyNames.Id, "exportToExcel");

                        Mouse.Click(itemExportToSpan_staging.GetChildren()[0]);



                        // Resetting parameters
                        controlClickCount = 0;
                        isControlClicked = true;

                        #endregion
                    }
                    catch (Exception)
                    {
                        controlClickCount++;

                        BrowserWindow.ClearCache();
                        BrowserWindow.ClearCookies();
                        openedBrowser.Refresh();

                        if (controlClickCount >= 5)
                        {
                            throw;
                        }

                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            finally
            {
                openedBrowser.Dispose();
                openedBrowser.Close();
            }
        }

        [TestMethod]
        public void MetricUploadService()
        {
            short uIClickCount = 0;
            bool isUIClicked = false;


            BrowserWindow browser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["MetricUploadService"];
            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["MetricUploadService"]));

            methodDescription = "This test method verifies whether the PPD Metric Upload Service page loads successfully.";

            Playback.Wait(40000);
            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        HtmlComboBox metricComboBox = new HtmlComboBox(browser);
                        metricComboBox.SearchProperties.Add(HtmlComboBox.PropertyNames.Id, "ddlMetricTemplate");

                        Mouse.Click(metricComboBox);

                        isUIClicked = true;

                    }
                    catch (Exception ex)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();
                        Playback.Wait(40000);
                        if(++uIClickCount > 5)
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                throw;
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }   
        }

        [TestMethod]
        public void MetricApprovalService()
        {
            short uIClickCount = 0;
            bool isUIClicked = false;


            BrowserWindow browser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["MetricApprovalService"];
            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["MetricApprovalService"]));
            methodDescription = "This test method verifies whether the PPD Metric Approval Service page loads successfully.";

            Playback.Wait(15000);
            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        HtmlDiv approvedItem = new HtmlDiv(browser);
                        approvedItem.SearchProperties.Add(HtmlDiv.PropertyNames.Id, "_approved");

                        Mouse.Click(approvedItem);

                        isUIClicked = true;

                    }
                    catch (Exception ex)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();
                        Playback.Wait(15000);
                        if (++uIClickCount > 5)
                            throw;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                throw;
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }  
        }

        //[TestMethod]
        //[Timeout(TestTimeout.Infinite)]
        //public void MASCube()
        //{
        //    this.UIMap.OpenSSMS();
        //    Playback.Wait(15000);
        //    this.UIMap.ConnectToCube();
        //    Playback.Wait(5000);
        //    this.UIMap.OpenCubes();
        //    this.UIMap.AssertMethod1();
        //}

        //[TestMethod]
        //public void PPDCubeWithGMOTest1()
        //{

        //}

        //[TestMethod]
        //public void FY14_Expanded_Geo_Cube_with_GMO_Test1()
        //{

        //}

        //[TestMethod]
        //public void Process_PPD_Cube_Geo_Level_Changes()
        //{

        //}

        [TestMethod]
        public void PPDCubeFunctions()
        {
            try
            {

            }
            catch(Exception e)
            {
                
            }
        }

        //[TestMethod]
        //public void PPDMSSalesSecurity()
        //{

        //}

        [TestMethod]
        public void SecurityCheckReport()
        {
            short uIClickCount = 0;
            bool isUIClicked = false;


            BrowserWindow browser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["SecurityCheckReport"];
            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["SecurityCheckReport"]));

            methodDescription = "This test method verifies whether the PPD Security Check Report page loads successfully.";

            Playback.Wait(7000);
            try
            {
                
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {

                        HtmlEdit inputAlias = new HtmlEdit(browser);
                        inputAlias.SearchProperties.Add(HtmlEdit.PropertyNames.Id, "alias");

                        Mouse.Click(inputAlias);

                        isUIClicked = true;

                    }
                    catch (Exception ex)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();
                        Playback.Wait(7000);
                        if (++uIClickCount > 5)
                            throw;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                throw;
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }  
        }

        [TestMethod]
        public void VerifyMetricApprovalETL()
        {
            short uIClickCount = 0;
            bool isUIClicked = false;


            BrowserWindow browser = BrowserWindow.Launch("about:tabs");
            link = ConfigurationManager.AppSettings["MetricApprovalETL"];
            browser.NavigateToUrl(new System.Uri(ConfigurationManager.AppSettings["MetricApprovalETL"]));

            methodDescription = "This test method verifies whether the PPD Metric Approval Service ETL page loads successfully.";

            Playback.Wait(10000);
            try
            {
                while (uIClickCount < 5 && isUIClicked.Equals(false))
                {
                    try
                    {
                        HtmlComboBox dataSourceDropDown = new HtmlComboBox(browser);
                        dataSourceDropDown.SearchProperties.Add(HtmlComboBox.PropertyNames.Id, "ddlDataSource");

                        Mouse.Click(dataSourceDropDown);

                        isUIClicked = true;

                    }
                    catch (Exception ex)
                    {
                        BrowserWindow.ClearCache();
                        browser.Refresh();
                        Playback.Wait(15000);
                        if (++uIClickCount > 5)
                            throw;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                throw;
            }
            finally
            {
                browser.Dispose();
                browser.Close();
            }  
        }


        public static void SendMail(string messageBody, string subject)
        {
            try
            {
                  string FILEPATH = @"TableStyles.css";
                
                string sMailIToAuthenticate = "v-swsriv@microsoft.com"; //Variable to store the MailIToAuthenticate from Excel
        string sAuthencationPasswd = "July@2014";  //Variable to store the AuthencationPasswd from Excel
                string toAddress = "sritejkr@maqsoftware.net";
                System.Net.Mail.SmtpClient c = new System.Net.Mail.SmtpClient();
                c.Host = "smtphost.redmond.corp.microsoft.com";
                c.Port = 25;
                c.EnableSsl = true;
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                c.Credentials = new System.Net.NetworkCredential(sMailIToAuthenticate, sAuthencationPasswd);
                msg.To.Add(toAddress);
                //msg.CC.Add(sToAddress2);
                //msg.CC.Add(sToAddress3);
                msg.From = new System.Net.Mail.MailAddress(sMailIToAuthenticate, "PPD Smoke Test Automation");
                msg.Subject = subject;
                msg.IsBodyHtml = true;
                string messagestyle = "";
                string style = System.IO.File.ReadAllText("D:/DAILYFILES/2014/July/28July2014/SmokePPD/SmokePPD/TableStyles.css");
                messagestyle = ("<style>" + style + "</style>");

                string replaceWith = "";
                messagestyle = messagestyle.Replace("\r\n", replaceWith).Replace("\n", replaceWith).Replace("\r", replaceWith);

                messagestyle = messagestyle + messageBody;
                msg.Body = messagestyle;

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

                c.Send(msg);


            }
            catch (Exception)
            {
                throw;
            }

        }




        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
