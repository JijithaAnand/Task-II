using AutoItX3Lib;
using MarsFramework.Global;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using RelevantCodes.ExtentReports;
using System;
using System.Threading;
using static MarsFramework.Global.GlobalDefinitions;
using static NUnit.Core.NUnitFramework;

namespace MarsFramework.Pages
{
    internal class ManageListings
    {
        private const string StrSendText = @"C:\Users\JIJI\images.jpg";
        [System.Obsolete]
        public ManageListings()
        {
            PageFactory.InitElements(GlobalDefinitions.Driver, this);
        }

        //Click on Manage Listings Link
        [FindsBy(How = How.LinkText, Using = "Manage Listings")]
        private IWebElement manageListingsLink { get; set; }

        //View the listing
        [FindsBy(How = How.XPath, Using = "/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[1]/td[8]/div/button[1]/i")]
        
        private IWebElement view { get; set; }

        //Delete the listing
        [FindsBy(How = How.XPath, Using = "/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr/td[8]/div/button[3]/i")]
        private IWebElement delete { get; set; }

        //Edit the listing
        [FindsBy(How = How.XPath, Using = "(//i[@class='outline write icon'])[1]")]
        private IWebElement edit { get; set; }

        //Click on Yes or No
        [FindsBy(How = How.XPath, Using = "//div[@class='actions']")]
        private IWebElement clickActionsButton { get; set; }
        private int check { get; set; }

        [System.Obsolete]
        internal void ClickManageListing()
        {
            Thread.Sleep(1500);
            Pages.ShareSkill shareSkill = new Pages.ShareSkill();
            Thread.Sleep(1500);
            //shareSkill.EnterShareSkill();
            manageListingsLink.Click();
        }
        internal void ClickEdit()
        {
            edit.Click();
        }
        internal void ClickView()
        {
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataManageListings.xlsx", "ViewSkill");
            Global.Base bases = new Global.Base();
            Global.Base.ExtentReports();
            Thread.Sleep(1000);
            Global.Base.test = Global.Base.extent.StartTest("View skill");
            int counts = 0;
            counts = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            check = 0;
            if (counts !=0)
            {
                if (counts == 1)
                {
                    Thread.Sleep(1000);
                    IWebElement Category = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr/td[2]"));
                    IWebElement Title = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr/td[3]"));
                    IWebElement Description = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr/td[4]"));


                    if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && Description.Text == ExcelLib.ReadData(2, "Description"))
                    {
                        check++;
                        if (check != 0)
                        {
                            view.Click();
                            viewListing();
                        }

                    }
                    return;
                }
                else
                {
                    
                        for (var i = 1; i <= counts; i++)
                        {
                            Thread.Sleep(1000);
                            IWebElement Category = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[2]"));
                            IWebElement Title = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                            IWebElement Description = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));


                            if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && Description.Text == ExcelLib.ReadData(2, "Description"))
                            {
                                check++;
                                if (check != 0)
                                {
                                Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr["+ i + "]/td[8]/div/button[1]/i")).Click();
                                    viewListing();
                                }
                                return;
                            }
                        

                    }
                }
                if (check == 0)
                {
                    check = 1;
                    Global.Base.test.Log(LogStatus.Pass, "Test Passed, Conserened Skill not found");
                    SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Conserned Skill not found");
                    Assert.AreEqual(check, 0);
                    return;
                }
            }
            else
            {
                Global.Base.test.Log(LogStatus.Pass, "Test passed, Skill List is Empty for view");
                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "List in Manage Listing is Empty to view");
                Assert.AreEqual(counts,0);
            }

        }
        internal void ClickDelete()
        {
            Global.Base bases = new Global.Base();
            Global.Base.ExtentReports();
            Thread.Sleep(1000);
            Global.Base.test = Global.Base.extent.StartTest("Delete skill");
            //Populate the Excel Sheet DeleteSkill
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataManageListings.xlsx", "DeleteSkill");
            int counts = 0;
            counts = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            int check = 0;
            if (counts !=0)
            {
                
                    for (var i = 1; i <= counts; i++)
                    {
                        Thread.Sleep(1000);
                        IWebElement Category = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[2]"));
                        IWebElement Title = Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                        IWebElement Description = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));


                        if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && Description.Text == ExcelLib.ReadData(2, "Description"))
                        {
                            delete.Click();
                        Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")).Click();
                        check++;
                        return;
                        }
                    }
            }
        }
        internal void EditListing()
        {


            Thread.Sleep(1000);
            //Populate the Excel Sheet ShareSkill
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataShareSkill.xlsx", "ShareSkill");
            int counts = 0;
            counts = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            int check = 0;
            if(counts!=0)
            {
                
                    for (var i = 1; i <= counts; i++)
                    {
                        IWebElement Title = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                        IWebElement desc = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));

                        if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && desc.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Description"))
                        {
                        Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[1]/td[8]/div/button[2]/i")).Click();
                        Thread.Sleep(1500);
                        //Populate the Excel Sheet
                        GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataManageListings.xlsx", "EditManageListings");
                        //Update Title
                        Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[1]/div/div[2]/div/div[1]/input")).SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Title"));
                        //Update Description
                        Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[2]/div/div[2]/div[1]/textarea")).SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Description"));
                        //Enter Category
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[3]/div[2]/div/div/select")).Click();
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[3]/div[2]/div/div/select/option[7]")).Click();
                        Driver.FindElement(By.XPath("/ html / body / div / div / div[1] / div[2] / div / form / div[3] / div[2] / div / div[2] / div[1] / select")).Click();
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[3]/div[2]/div/div[2]/div[1]/select/option[5]")).Click();
                        //Enter Tags
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[4]/div[2]/div/div/div/div/input")).SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Tags"));
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[4]/div[2]/div/div/div/div/input")).SendKeys(Keys.Enter);
                        //Enter ServiceType
                        String Service = GlobalDefinitions.ExcelLib.ReadData(2, "ServiceType");
                        if (Service == "One-off service")
                        {
                            Driver.FindElement(By.XPath("/ html / body / div / div / div[1] / div[2] / div / form / div[5] / div[2] / div[1] / div[2] / div / input")).Click();
                        }
                        else
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[5]/div[2]/div[1]/div[1]/div/input")).Click();
                        }
                        //Enter LocationType
                        String LocationType = GlobalDefinitions.ExcelLib.ReadData(2, "LocationType");
                        if (LocationType == "On-site")
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[6]/div[2]/div/div[1]/div/input")).Click();
                        }
                        else
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[6]/div[2]/div/div[2]/div/input")).Click();
                        }

                        //Enter SkillTrade
                        string SkillTrade = GlobalDefinitions.ExcelLib.ReadData(2, "SkillTrade");
                        if (SkillTrade == "Skill-Exchange")
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[2]/div/div[1]/div/label")).Click();
                        }
                        else
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[2]/div/div[2]/div/label")).Click();
                        }
                        //Enter Skill-Exchange
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[4]/div/div/div/div/div/input")).SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Skill-Exchange"));
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[4]/div/div/div/div/div/input")).SendKeys(Keys.Enter);
                        //Enter Credit
                        if (Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[4]/div/div/div/div/div/input")).Text == "Credit")
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[8]/div[4]/div/div/input")).SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Credit"));
                        }
                        ////Work Sample
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[9]/div/div[2]/section/div/label/div/span/i")).Click();
                        AutoItX3 autoit = new AutoItX3();
                        autoit.WinActivate("Open");
                        Thread.Sleep(1500);
                        autoit.ControlSetText("Open", "", "Edit1", StrSendText);
                        autoit.ControlClick("Open", "", "Button1");


                        //Enter Active
                        string Active = GlobalDefinitions.ExcelLib.ReadData(2, "Active");
                        if (Active == "Active")
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[10]/div[2]/div/div[1]/div/input")).Click();
                        }
                        else
                        {
                            Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[10]/div[2]/div/div[2]/div/input")).Click();
                        }
                        //Click on save
                        Driver.FindElement(By.XPath("/html/body/div/div/div[1]/div[2]/div/form/div[11]/div/input[1]")).Click();
                        Thread.Sleep(1500);
                        manageListingsLink.Click();

                        check++;
                            return;
                        
                        }
                    }
                return;
            }

        }




        internal void UpdatedSkill()
        {
            Global.Base bases = new Global.Base();
            Global.Base.ExtentReports();
            Thread.Sleep(1000);
            Global.Base.test = Global.Base.extent.StartTest("Update skill");
            int counts = 0;
            counts = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            //Populate the Excel Sheet
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataManageListings.xlsx", "EditManageListings");
            int check = 0;
            if (counts != 0)
            {
                for (var j = 2; j <= 2; j++)
                {
                    for (var i = 1; i <= counts; i++)
                    {
                        IWebElement Title = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                        IWebElement desc = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));

                        if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && desc.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Description"))
                        {
                            Global.Base.test.Log(LogStatus.Pass, "Test Passed, Manage Listing Updated successfully");
                            SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Updated successfully");
                            Assert.AreEqual(Title.Text, GlobalDefinitions.ExcelLib.ReadData(2, "Title"));
                            check++;
                            return;
                        }

                    }
                    Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[2]/button[" + j + "]")).Click();
                }
                if (check == 0 && counts == 0)
                {
                    Global.Base.test.Log(LogStatus.Pass, "Test Passed, Conserned Skill not found");
                    SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Conserned Skill not found");
                    Assert.AreEqual(check ,0);
                    return;
                }
            }
            else
            {
                Global.Base.test.Log(LogStatus.Pass, "Test passed, Skill List is Empty for Edit");
                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "List in Manage Listing is Empty to Edit");
                Assert.AreEqual(counts, 0);
            }
            
            
        }
        internal void viewListing()
        {
            //Global.Base bases = new Global.Base();
            //Global.Base.ExtentReports();
            Thread.Sleep(1000);
            //Global.Base.test = Global.Base.extent.StartTest("View skill")
                    //GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataShareSkill.xlsx", "ViewSkill");
                    //IWebElement Title = Global.GlobalDefinitions.Driver.FindElement(By.ClassName("skill-title"));
             ////       IWebElement desc = Global.GlobalDefinitions.Driver.FindElement(By.ClassName("description"));
             //       if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && desc.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Description"))
             //       {
                        Global.Base.test.Log(LogStatus.Pass, "Test Passed, Manage Listing Viewed successfully");
                        SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Viewed successfully");
             //   Assert.AreEqual(Title.Text, GlobalDefinitions.ExcelLib.ReadData(2, "Title"));
               // return;
              //      }
                
            //    elseC:\Users\JIJI\source\repos\marsframework-master\MarsFramework\Pages\Profile.cs
            //    {
            //        Global.Base.test.Log(LogStatus.Pass, "Test passed, The Consered Skill not found");
            //        SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Consered Skill Not Found In Manage Listing");
                
            //}
            
            
        }
        
        
        internal void YesDelete()
        {
            Global.Base bases = new Global.Base();
            Global.Base.ExtentReports();
            Thread.Sleep(1000);
            Global.Base.test = Global.Base.extent.StartTest("Delete skill");
            int counts = 0;
            counts = Global.GlobalDefinitions.Driver.FindElements(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr")).Count;
            //Populate the Excel Sheet
            GlobalDefinitions.ExcelLib.PopulateInCollection(@"C:\Users\JIJI\source\repos\marsframework-master\MarsFramework\ExcelData\TestDataManageListings.xlsx", "DeleteSkill");
            int check = 0;
            Thread.Sleep(1000);
           
           
                if (counts !=0)
                {
                    
                        for (var i = 1; i <= counts; i++)
                        {
                            IWebElement Title = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                            IWebElement desc = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));

                            if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && desc.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Description"))
                            {
                                Global.Base.test.Log(LogStatus.Fail, "Test Failed, Manage Listing Not Deleted");
                                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Not Deleted");
                        Assert.AreEqual(Title.Text, GlobalDefinitions.ExcelLib.ReadData(2, "Title"));
                        check++;       
                        return;
                            }
                            if(check==0)
                    {
                        Global.Base.test.Log(LogStatus.Pass, "Test passed, Manage Listing Deleted");
                        SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Deleted");
                        Assert.AreEqual(check,0);
                        return;
                    }
                        }
                        
                    

                }

else
            {
                Global.Base.test.Log(LogStatus.Pass, "Test passed, Skill List is Empty");
                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "List in Manage Listing is Empty");
                Assert.AreEqual(counts,0);
            }
            
        }
    
        internal void NoDelete()
        {
            Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[1]")).Click();
            Thread.Sleep(1000);
            try
            {
                while (true)
                {
                    for (var j = 2; j <= 100; j++)
                    {
                        for (var i = 1; i <= 10; i++)
                        {
                            IWebElement Title = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[3]"));
                            IWebElement desc = Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[1]/table/tbody/tr[" + i + "]/td[4]"));

                            if (Title.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Title") && desc.Text == GlobalDefinitions.ExcelLib.ReadData(2, "Description"))
                            {
                                Global.Base.test.Log(LogStatus.Pass, "Test Pass, Manage Listing Not Deleted");
                                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Not Deleted");
                                return;
                            }
                        }
                        Global.GlobalDefinitions.Driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[1]/div[2]/button[" + j + "]")).Click();
                    }

                }


            }
            catch
            {
                Global.Base.test.Log(LogStatus.Fail, "Test Failed, Manage Listing Not found");
                SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.Driver, "Manage Listing Not found");
            }
        }
        internal void Listings()
        {
            //Populate the Excel Sheet
            GlobalDefinitions.ExcelLib.PopulateInCollection(Base.ExcelPath, "ManageListings");


        }
    }
}
