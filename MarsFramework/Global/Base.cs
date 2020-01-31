using MarsFramework.Config;
using MarsFramework.Pages;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.PageObjects;
using RelevantCodes.ExtentReports;
using System;
using System.IO;
using System.Threading;
using TechTalk.SpecFlow;
using static MarsFramework.Global.GlobalDefinitions;

namespace MarsFramework.Global
{
    class Base
    {
        #region To access Path from resource file
        private static string solutionParentDirectory = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
       // public static String ExcelPath = Path.Combine(solutionParentDirectory, @"ExcelData\TestDataShareSkill.xlsx");
        public static int Browser = Int32.Parse(MarsResource.Browser);
       // public static String ExcelPath = @"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData";
        public static String ExcelPath = MarsResource.ExcelPath;
        public static string ScreenshotPath = MarsResource.ScreenShotPath;
        public static string ReportPath = MarsResource.ReportPath;
        //public static string ScreenshotPath = @"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\TestReports\Screenshots\";
       // public static string ReportPath = @"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\TestReports\Test.html";
        #endregion

        #region reports
        public static ExtentTest test;
        public static ExtentReports extent;
        #endregion
        public static void ExtentReports()
        {
            extent = new ExtentReports(ReportPath, true, DisplayOrder.NewestFirst);
            extent.LoadConfig(MarsResource.ReportXMLPath);
        }
        #region setup and tear down
        [SetUp]
        [Obsolete]
        public void Inititalize()
        {

            // advisasble to read this documentation before proceeding http://extentreports.relevantcodes.com/net/
            switch (Browser)
            {

                case 1:
                    GlobalDefinitions.Driver = new FirefoxDriver();
                    break;
                case 2:
                    GlobalDefinitions.Driver = new ChromeDriver();
                    GlobalDefinitions.Driver.Manage().Window.Maximize();
                    break;

            }

            #region Initialise Reports

            extent = new ExtentReports(ReportPath, false, DisplayOrder.NewestFirst);
            extent.LoadConfig(MarsResource.ReportXMLPath);

            #endregion
            //add the website
            Driver.Navigate().GoToUrl("http://192.168.99.100:5000");
            //Driver.Navigate().GoToUrl("http://www.skillswap.pro/");
            if (MarsResource.IsLogin == "true")
            {
                SignIn loginobj = new SignIn();
                loginobj.LoginSteps();
            }
            else
            {
                SignUp obj = new SignUp();
                obj.register();
            }

        }
        [TearDown]

       // [AfterScenario]
        public void TearDown()
        {
            // test = extent.StartTest("Add a Share Skill");
            Thread.Sleep(500);
            // Screenshot
            String img = SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Report");//AddScreenCapture(@"E:\Dropbox\VisualStudio\Projects\Beehive\TestReports\ScreenShots\");
           // test.Log(LogStatus.Info, "Image example: " + img);
            test.Log(LogStatus.Info, "Snapshot below: " + test.AddScreenCapture(img));
            // end test. (Reports)
            extent.EndTest(test);
            // calling Flush writes everything to the log file (Reports)
            extent.Flush();
            // Close the driver :)            
            GlobalDefinitions.Driver.Close();
            GlobalDefinitions.Driver.Quit();
        }
        #endregion

        [Test]
        [Obsolete]
        public void ShareSkill()
        {
           var shareskill = new Pages.ShareSkill();
            PageFactory.InitElements(Driver, shareskill);
            shareskill.EnterShareSkill();
            shareskill.viewtheskill();

        }
        [Test]
        [Obsolete]
        public void EditSkillManageListing()
        {
            var manageListings = new Pages.ManageListings();
            PageFactory.InitElements(Driver, manageListings);
            manageListings.ClickManageListing();
            manageListings.EditListing();
            manageListings.UpdatedSkill();
           
        }

        [Test]
        [Obsolete]
        public void ViewSkillManageListing()
        {
            var manageListings = new Pages.ManageListings();
            PageFactory.InitElements(Driver, manageListings);
            manageListings.ClickManageListing();
            manageListings.ClickView();
            
        }
        [Test]
        [Obsolete]
        public void DeleteSkillManageListing()
        {
            var manageListings = new Pages.ManageListings();
            PageFactory.InitElements(Driver, manageListings);
            manageListings.ClickManageListing();
            manageListings.ClickDelete();
            manageListings.YesDelete();
            
        }
    }
}