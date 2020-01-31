using MarsFramework.Global;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace MarsFramework.Pages
{
    class SignIn
    {
        [System.Obsolete]
        public SignIn()
        {
            PageFactory.InitElements(GlobalDefinitions.Driver, this);
        }

        #region  Initialize Web Elements 
        //Finding the Sign Link
        [FindsBy(How = How.XPath, Using = "//a[contains(text(),'Sign')]")]
        private IWebElement SignIntab { get; set; }

        // Finding the Email Field
        [FindsBy(How = How.Name, Using = "email")]
        private IWebElement Email { get; set; }

        //Finding the Password Field
        [FindsBy(How = How.Name, Using = "password")]
        private IWebElement Password { get; set; }

        //Finding the Login Button
        [FindsBy(How = How.XPath, Using = "//button[contains(text(),'Login')]")]
        private IWebElement LoginBtn { get; set; }

        #endregion

        internal void LoginSteps()
        {
            //Populate the excel data

            //Populate the excel data
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestData.xlsx", "SignIn");
            //Base.ExcelPath = @"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestData.xlsx";
            //Click on Signin Tab
            SignIntab.Click();
            //Enter email id
            Email.SendKeys(GlobalDefinitions.ExcelLib.ReadData(2,"Username"));
            //Enter password
            Password.SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Password"));
            //Click on Login button
            LoginBtn.Click();
        }
    }
}