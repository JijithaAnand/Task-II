using MarsFramework.Global;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using RelevantCodes.ExtentReports;
using System;
using System.Data;
using System.IO;
using System.Threading;
using static MarsFramework.Global.GlobalDefinitions;
using System.Runtime.InteropServices;
using AutoItX3Lib;
using static NUnit.Core.NUnitFramework;

namespace MarsFramework.Pages
{
    internal class ShareSkill
    {
        private const string StrSendText = @"C:\Users\JIJI\images.jpg";

        [System.Obsolete]
        public ShareSkill()
        {
            PageFactory.InitElements(GlobalDefinitions.Driver, this);
        }

        //Click on ShareSkill Button
        [FindsBy(How = How.LinkText, Using = "Share Skill")]
        public IWebElement ShareSkillButton { get; set; }

        //Enter the Title in textbox
        [FindsBy(How = How.Name, Using = "title")]
        private IWebElement Title { get; set; }

        //Enter the Description in textbox
        [FindsBy(How = How.Name, Using = "description")]
        private IWebElement Description { get; set; }

        //Click on Category Dropdown
        [FindsBy(How = How.Name, Using = "categoryId")]
        private IWebElement CategoryDropDown { get; set; }

        //Click on SubCategory Dropdown
        [FindsBy(How = How.Name, Using = "subcategoryId")]
        private IWebElement SubCategoryDropDown { get; set; }

        //Enter Tag names in textbox
        [FindsBy(How = How.XPath, Using = "//body/div/div/div[@id='service-listing-section']/div[contains(@class,'ui container')]/div[contains(@class,'listing')]/form[contains(@class,'ui form')]/div[contains(@class,'tooltip-target ui grid')]/div[contains(@class,'twelve wide column')]/div[contains(@class,'')]/div[contains(@class,'ReactTags__tags')]/div[contains(@class,'ReactTags__selected')]/div[contains(@class,'ReactTags__tagInput')]/input[1]")]
        private IWebElement Tags { get; set; }

        //Select the Service type
        [FindsBy(How = How.XPath, Using = "//form/div[5]/div[@class='twelve wide column']/div/div[@class='field']")]
        private IWebElement ServiceTypeOptions { get; set; }

        //Select the Location Type
        [FindsBy(How = How.XPath, Using = "//form/div[6]/div[@class='twelve wide column']/div/div[@class = 'field']")]
        private IWebElement LocationTypeOption { get; set; }

        //Click on Start Date dropdown
        [FindsBy(How = How.Name, Using = "k-sm-date-format")]
        private IWebElement StartDateDropDown { get; set; }

        //Click on End Date dropdown
        [FindsBy(How = How.Name, Using = "endDate")]
        private IWebElement EndDateDropDown { get; set; }

        //Storing the table of available days
        [FindsBy(How = How.XPath, Using = "//body/div/div/div[@id='service-listing-section']/div[@class='ui container']/div[@class='listing']/form[@class='ui form']/div[7]/div[2]/div[1]")]
        private IWebElement Days { get; set; }

        //Storing the starttime
        [FindsBy(How = How.XPath, Using = "//div[3]/div[2]/input[1]")]
        private IWebElement StartTime { get; set; }

        //Click on StartTime dropdown
        [FindsBy(How = How.XPath, Using = "//div[3]/div[2]/input[1]")]
        private IWebElement StartTimeDropDown { get; set; }

        //Click on EndTime dropdown
        [FindsBy(How = How.XPath, Using = "//div[3]/div[3]/input[1]")]
        private IWebElement EndTimeDropDown { get; set; }

        //Click on Skill Trade option
        [FindsBy(How = How.XPath, Using = "//form/div[8]/div[@class='twelve wide column']/div/div[@class = 'field']")]
        private IWebElement SkillTradeOption { get; set; }

        //Enter Skill Exchange
        [FindsBy(How = How.XPath, Using = "//div[@class='form-wrapper']//input[@placeholder='Add new tag']")]
        private IWebElement SkillExchange { get; set; }

        //Enter the amount for Credit
        [FindsBy(How = How.XPath, Using = "//input[@placeholder='Amount']")]
        private IWebElement CreditAmount { get; set; }

        //Click on Active/Hidden option
        [FindsBy(How = How.XPath, Using = "//form/div[10]/div[@class='twelve wide column']/div/div[@class = 'field']")]
        private IWebElement ActiveOption { get; set; }

        //Click on Save button
        [FindsBy(How = How.XPath, Using = "//input[@value='Save']")]
        private IWebElement Save { get; set; }
        public void UpdateSkill()
        {
            
        }
        internal void EnterShareSkill()
        {
            int rows;
           rows = GlobalDefinitions.ExcelLib.NumberofRows(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataShareSkill.xlsx", "ShareSkill");
            //Populate the excel data
             GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataShareSkill.xlsx", "ShareSkill");
                Thread.Sleep(1500);
            
            for (int i = 2; i <= rows+1; i++)
            {
                Thread.Sleep(1500);
                //Click on Share Skill
                ShareSkillButton.Click();
                //Enter Title
                Thread.Sleep(3500);
                Title.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Title"));
                //Enter Description
                Description.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Description"));
                //Enter Category
                CategoryDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Category"));
                //Enter Sub-Category
                SubCategoryDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "SubCategory"));
                //Enter Tags
                Tags.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Tags"));
                Tags.SendKeys(Keys.Enter);
                //Enter ServiceType
                String Service = GlobalDefinitions.ExcelLib.ReadData(i, "ServiceType");
                if (Service == "One-off service")
                {
                    ServiceTypeOptions.FindElement(By.XPath("/ html / body / div / div / div[1] / div[2] / div / form / div[5] / div[2] / div[1] / div[2] / div / input")).Click();
                }
                else
                {
                    ServiceTypeOptions.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[5]/div[2]/div[1]/div[1]/div/input")).Click();
                }
                //Enter LocationType
                String LocationType = GlobalDefinitions.ExcelLib.ReadData(i, "LocationType");
                if (LocationType == "On-site")
                {
                    LocationTypeOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[6]/div[2]/div/div[1]/div/input")).Click();
                }
                else
                {
                    LocationTypeOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[6]/div[2]/div/div[2]/div/input")).Click();
                }
                //Enter Startdate
                
                //StartDateDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Startdate").ToString());
                //Enter Enddate
                // EndDateDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Enddate"));
                //Enter Selectday
                // String Day=GlobalDefinitions.ExcelLib.ReadData(2, "Selectday");
                //if(Day== "Sun")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[2]/div[1]/div/input")).Click();
                //}
                //else if(Day=="Mon")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[3]/div[1]/div/input")).Click();
                // }
                //else if (Day == "Tue")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[4]/div[1]/div/input")).Click();
                //}
                //else if (Day == "Wed")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[5]/div[1]/div/input")).Click();
                //}
                //else if (Day == "Thu")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[6]/div[1]/div/input")).Click();
                //}
                //else if (Day == "Fri")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[7]/div[1]/div/input")).Click();
                //}
                //else if (Day == "Sat")
                //{
                //    Days.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[7]/div[2]/div[1]/div[8]/div[1]/div/input")).Click();
                //}

                //Enter Starttime
               // StartTimeDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Starttime"));
                //Enter Endtime
               // EndTimeDropDown.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Endtime"));
                //Enter SkillTrade
                string SkillTrade = GlobalDefinitions.ExcelLib.ReadData(i, "SkillTrade");
                if (SkillTrade == "Skill-Exchange")
                {
                    SkillTradeOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[2]/div/div[1]/div/label")).Click();
                }
                else
                {
                    SkillTradeOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[2]/div/div[2]/div/label")).Click();
                }
                //Enter Skill-Exchange
                SkillExchange.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Skill-Exchange"));
                SkillExchange.SendKeys(Keys.Enter);
                //Enter Credit
                if (SkillExchange.Text == "Credit")
                {
                    CreditAmount.SendKeys(GlobalDefinitions.ExcelLib.ReadData(i, "Credit"));
                }
                ////Work Sample
                Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[9]/div/div[2]/section/div/label/div/span/i")).Click();
                AutoItX3 autoit = new AutoItX3();
                autoit.WinActivate("Open");
                Thread.Sleep(1500);
                autoit.ControlSetText("Open", "", "Edit1", StrSendText);
                autoit.ControlClick("Open", "", "Button1");


                //Enter Active
                string Active = GlobalDefinitions.ExcelLib.ReadData(i, "Active");
                if (Active == "Active")
                {
                    ActiveOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[10]/div[2]/div/div[1]/div/input")).Click();
                }
                else
                {
                    ActiveOption.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[10]/div[2]/div/div[2]/div/input")).Click();
                }
                //Click on save
                Save.Click();
                Thread.Sleep(1500);
            }
        }
        internal void viewtheskill()
        {
            Global.Base bases = new Global.Base();
            Global.Base.ExtentReports();
            Thread.Sleep(1000);
            Global.Base.test = Global.Base.extent.StartTest("Share skill Added");
            int count = 0;
            count = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            while (true)
            {

                for (var i = 1; i <= count; i++)
                {
                    Thread.Sleep(1000);
                    IWebElement Category = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[2]"));
                    IWebElement Title = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                    IWebElement Description = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));


                    if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && Description.Text == ExcelLib.ReadData(2, "Description"))
                    {
                        //test = extent.StartTest("Add a Share Skill");
                        // Global.Base.test.Log(LogStatus.Info, "Snapshot below: " + Global.Base.test.AddScreenCapture(SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Report")));
                        Global.Base.test.Log(LogStatus.Pass, "Test Passed, Share Skill Added successfully");
                        SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Share skill updated successfully");
                        Assert.AreEqual(Title.Text, GlobalDefinitions.ExcelLib.ReadData(2, "Title"));
                        return;
                    }

                }
                return;
            }
        }
        
        
        
    }
}


