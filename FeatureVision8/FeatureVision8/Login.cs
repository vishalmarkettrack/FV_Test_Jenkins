using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Configuration;
using System.Data;
using AventStack.ExtentReports;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace FeatureVision8
{
    public class Login
    {
        #region Private Variables
            private IWebDriver login;
            private ExtentTest test;
        #endregion

        //#region Public Methods

        //public Login(IWebDriver driver, ExtentTest testReturn)
        //{
        //    // TODO: Complete member initialization
        //    this.login = driver;
        //    test = testReturn;
        //}

        //public IWebDriver driver
        //{
        //    get { return this.login; }
        //    set { this.login = value; }
        //}

        ///// <summary>
        ///// Navigate to login page (Login URL get From the Login.xlsx Sheet)
        ///// </summary>
        ///// <returns></returns>
        //public Login navigateToLoginPage()
        //{
        //    driver.Navigate().GoToUrl(Common.ApplicationURL);
        //    Results.WriteStatus(test, "Pass", "Launched, URL <b>" + Common.ApplicationURL + "</b> successfully.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// To Verify Login Page
        ///// </summary>
        ///// <param name="Watermark">Verify Watermark</param>
        ///// <returns></returns>
        //public Login VerifyLoginPage(bool Watermark = true)
        //{
        //    //Assert.IsTrue(driver._waitForElement("id", "ctl00_imgMtLogo"), "Feature Vision Logo not found on Page.");
        //    Assert.AreEqual(true, driver._waitForElement("xpath", "//div[@class='logo']", 15), "Numerator Logo not found on Page.");

        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='Verlbl']"), "'Version Number' not Present on page.");
        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='button' and @value='LEARN MORE']"), "'LEARN MORE' Button not Present on Screen.");

        //    //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_WelComeTitle"), "'To access FeatureVision® enter your email address and password' Label not Found.");
        //    //Assert.AreEqual("To access FeatureVision® enter your email address and password", driver._getText("id", "ctl00_ContentPlaceHolder1_WelComeTitle"), "'To access FeatureVision® enter your email address and password' Label not Match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "EmailAddress"), "Email input area not found on Page.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "Password"), "Password input area not found on Page.");

        //    if (Watermark)
        //    {
        //        Assert.AreEqual(true, driver._getAttributeValue("id", "EmailAddress", "placeholder").Contains("Email"), "'Email Address' Watermark not Present.");
        //        Assert.AreEqual(true, driver._getAttributeValue("id", "Password", "placeholder").Contains("Password"), "'Password' Watermark not Present.");
        //    }

        //    //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_rememberme"), "'Remember' checkbox not Present");
        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@class='inputFieldLabel padding-left10']"), "'Remember my credentials on this computer' checkbox Label not Present.");
        //    //Assert.AreEqual("Remember my credentials on this computer", driver._getText("xpath", "//span[@class='inputFieldLabel padding-left10']"), "'Remember my credentials on this computer' checkbox Label not match");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[contains(text(), 'Log In')]"), "LOG IN Button not found on Page.");

        //    //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_reqDiv"), "Browser Requirements Icon not found on Page.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[text()='Reset Password']"), "Reset Password Link not found on Page.");
        //    //Assert.AreEqual("Reset Password", driver._getText("id", "ctl00_ContentPlaceHolder1_RetrivePasswordLink"), "'Reset Password' Link Label not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[text()='Request Access']"), "Request Access Link not found on Page.");
        //    //Assert.AreEqual("Request Access", driver._getText("xpath", "//a[@class='request-accessLink']"), "'Request Access' Link Label not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[text()='Privacy Notice']"), "'Privacy Policy' Label at Bottom not found.");
        //    //Assert.AreEqual("Privacy Notice", driver._getText("xpath", "//a[@class='privacy-policyLink']"), "'Privacy Notice' Label at Bottom not match.");

        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@class='loginheaderlink']"), "'Resolution' note at Bottom not found.");
        //    //Assert.AreEqual("FeatureVision® has been best optimized at 1280 x 1024 resolution.", driver._getText("xpath", "//span[@class='loginheaderlink']").Trim().Replace("\r\n", ""), "'FeatureVision® has been best optimized at 1280 x 1024 resolution.' Label at Bottom not match.");

        //    Results.WriteStatus(test, "Pass", "Verified, Login page.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Promo Detail on Screen
        ///// </summary>
        ///// <param name="sectionName">Section Name to Verify</param>
        ///// <returns></returns>
        //public Login verifyPromoDetailOnScreen(string sectionName)
        //{
        //    IWebElement body = driver._findElement("xpath", "//td[@class='Login-bg-color'][1]");
        //    bool avail = false;

        //    if (sectionName.Equals("Image"))
        //    {
        //        IList<IWebElement> imageOrVideo = body.FindElements(By.TagName("img"));
        //        for (int i = 0; i < imageOrVideo.Count; i++)
        //            if (imageOrVideo[i].GetAttribute("src").Contains(".png") || imageOrVideo[i].GetAttribute("src").Contains(".jpg") || imageOrVideo[i].GetAttribute("src").Contains(".jpeg") || imageOrVideo[i].GetAttribute("src").Contains(".png"))
        //            {
        //                avail = true;
        //                break;
        //            }
        //    }
        //    else
        //    {
        //        IList<IWebElement> content = body.FindElements(By.TagName("td"));
        //        string title = "";

        //        if (sectionName.Equals("Header"))
        //            title = "Amazon Primed: Are you coexisting with the Everything Store?";

        //        if (sectionName.Equals("Sub Header"))
        //            title = "Follow our ongoing Amazon coverage and view our most recent webinar";

        //        for (int i = 0; i < content.Count; i++)
        //            if (content[i].Text.Trim().Replace("\r\n", "").ToLower().Contains(title.ToLower()))
        //            {
        //                avail = true;
        //                break;
        //            }
        //    }

        //    Assert.AreEqual(true, avail, "Promo '" + sectionName + "' not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified, Promo " + sectionName + " Section on Screen.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click on Logo
        ///// </summary>
        ///// <returns></returns>
        //public Login clickOnLogo()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_imgMtLogo"), "MTG Logo not Present on Screen.");
        //    driver._click("id", "ctl00_imgMtLogo");
        //    Thread.Sleep(2000);
        //    Results.WriteStatus(test, "Pass", "Clicked On MTG Logo on Screen.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Navigate URL on Screen
        ///// </summary>
        ///// <param name="url">URL</param>
        ///// <returns></returns>
        //public Login verifyNavigateURLOnScreen(string url)
        //{
        //    Assert.AreEqual(true, driver.Url.Contains(url), "Navigate URL not Match. Url is : " + driver.Url);
        //    Results.WriteStatus(test, "Pass", "Verified, Navigate Url " + url + " on Screen.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Login using Valid Email Address & Password
        ///// </summary>
        ///// <param name="configurUser">Configure User Login Credentail Enter</param>
        ///// <returns></returns>
        //public Login loginUsingValidEmailAddressPassword(bool configureUser = false)
        //{
        //    #region Datasheet

        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
        //    string email, password = "";

        //    #endregion

        //    if (configureUser)
        //    {
        //        email = Email[1].ToString();
        //        password = Password[1].ToString();
        //    }
        //    else
        //    {
        //        email = Email[0].ToString();
        //        password = Password[0].ToString();
        //    }

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "Email Address Textarea not Present.");
        //    driver._type("id", "ctl00_ContentPlaceHolder1_EmailAddress", email);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_Password"), "Password Textarea not Present.");
        //    driver._type("id", "ctl00_ContentPlaceHolder1_Password", password);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Password : " + password);

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Login using Valid Email Id & Password
        ///// </summary>
        ///// <param name="column">Column Number to Find Data from the Excel Sheet</param>
        ///// <returns></returns>
        //public Login loginUsingValidEmailIdAndPassword(int column)
        //{
        //    #region Datasheet

        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
        //    string email, password = "";

        //    #endregion

        //    email = Email[column].ToString();
        //    password = Password[column].ToString();

        //    Assert.AreEqual(true, driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
        //    driver._type("id", "EmailAddress", email);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);

        //    Assert.AreEqual(true, driver._isElementPresent("id", "Password"), "Password Textarea not Present.");
        //    driver._type("id", "Password", password);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Password : " + password);

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Login using Invalid Email address and Password
        ///// </summary>
        ///// <returns></returns>
        //public Login loginUsingInvalidEmailAddressAndPassword()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "Email Address Textarea not Present.");
        //    driver._type("id", "ctl00_ContentPlaceHolder1_EmailAddress", driver._randomString(5) + "@test.com");
        //    Results.WriteStatus(test, "Pass", "Entered Invalid Email address");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_Password"), "Password Textarea not Present.");
        //    driver._type("id", "ctl00_ContentPlaceHolder1_Password", driver._randomString(6));
        //    Results.WriteStatus(test, "Pass", "Entered Invalid Password.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Login Using Invalid Email Address or Password
        ///// </summary>
        ///// <param name="password">Enter Invalid Password</param>
        ///// <returns></returns>
        //public Login loginUsingInvalidEmailAddressOrPassword(bool password)
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "Email Address Textarea not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_Password"), "Password Textarea not Present.");

        //    if (password)
        //    {
        //        driver._clearText("id", "ctl00_ContentPlaceHolder1_EmailAddress");
        //        driver._type("id", "ctl00_ContentPlaceHolder1_Password", driver._randomString(6));
        //        Results.WriteStatus(test, "Pass", "Entered Invalid Password.");
        //    }
        //    else
        //    {
        //        driver._clearText("id", "ctl00_ContentPlaceHolder1_Password");
        //        driver._type("id", "ctl00_ContentPlaceHolder1_EmailAddress", driver._randomString(5) + "@test.com");
        //        Results.WriteStatus(test, "Pass", "Entered Invalid Email address");
        //    }

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Select Remember Checkbox on Login Page
        ///// </summary>
        ///// <returns></returns>
        //public Login selectRememberCheckboxOnLoginPage()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_rememberme"), "Remember Checkbox not Present.");
        //    driver._click("id", "ctl00_ContentPlaceHolder1_rememberme");
        //    Results.WriteStatus(test, "Pass", "Selected Remember Checkbox on Login Page.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click Sign In Button
        ///// </summary>
        ///// <returns></returns>
        //public Login clickSignInButton()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[contains(text(), 'Log In')]"), "'Sign In' Button not Present.");
        //    driver._clickByJavaScriptExecutor("//button[contains(text(), 'Log In')]");
        //    Results.WriteStatus(test, "Pass", "Clicked, Sign In Button.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click Learn More Button on Screen
        ///// </summary>
        ///// <returns></returns>
        //public Login clickLearnMoreButtonOnScreen()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='button' and @value='LEARN MORE']"), "'LEARN MORE' Button not Present.");
        //    driver._clickByJavaScriptExecutor("//input[@type='button' and @value='LEARN MORE']");
        //    Results.WriteStatus(test, "Pass", "Clicked, Sign In Button.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click System Requirement Icon on Screen
        ///// </summary>
        ///// <returns></returns>
        //public Login clickSystemRequirementIconOnScreen()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_reqDiv"), "'System Requirement' Icon not Present.");
        //    driver._click("id", "ctl00_ContentPlaceHolder1_reqDiv");
        //    Results.WriteStatus(test, "Pass", "Clicked, System Requirement Icon on Screen.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click on Icon Button and Verify Video Screen
        ///// </summary>
        ///// <param name="iconName">Icon Name</param>
        ///// <returns></returns>
        //public Login clickOnIconButtonAndVerifyVideoScreen(string iconName)
        //{
        //    driver._selectFrameToDefaultContent();

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvVideo']/iframe"), "'Video' not Present.");
        //    driver._selectFrameWithinFrame("xpath", "//*[@id='dvVideo']/iframe");

        //    switch (iconName.ToLower())
        //    {
        //        case "play":
        //            if (driver._isElementPresent("xpath", "//button[@aria-label='Play' and @title='Play']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement playElement = driver._findElement("xpath", "//*[@id='player']/div[1]");
        //                    driver.MouseHoverByJavaScript(playElement);
        //                }
        //            }

        //            Assert.IsTrue(driver._waitForElement("xpath", "//button[@aria-label='Play' and @title='Play']", 15), "'Play' Button not Present.");
        //            driver._click("xpath", "//button[@aria-label='Play' and @title='Play']");

        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@class='play rounded-box state-playing' and @aria-label='Pause' and @title='Pause']"), "'Video not Playing.");
        //            Results.WriteStatus(test, "Pass", "Clicked, Play Icon and Verified Video Playing.");
        //            break;

        //        case "pause":
        //            if (driver._isElementPresent("xpath", "//button[@aria-label='Pause' and @title='Pause']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement pauseElement = driver._findElement("xpath", "//*[@id='player']");
        //                    driver.MouseHoverByJavaScript(pauseElement);
        //                }
        //            }

        //            Assert.IsTrue(driver._waitForElement("xpath", "//button[@aria-label='Pause' and @title='Pause']", 15), "'Pause' Button not Present.");
        //            driver._click("xpath", "//button[@aria-label='Pause' and @title='Pause']");

        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@class='play rounded-box state-paused' and @aria-label='Play' and @title='Play']"), "Video not Paused.");
        //            Results.WriteStatus(test, "Pass", "Clicked, Pause Icon and Verified Video Paused.");
        //            break;

        //        case "fullscreen":
        //            if (driver._isElementPresent("xpath", "//button[@class='fullscreen']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement fullscreenElement = driver._findElement("xpath", "//*[@id='player']/div[1]");
        //                    driver.MouseHoverByJavaScript(fullscreenElement);
        //                }
        //            }

        //            if (driver._getAttributeValue("xpath", "//button[@class='fullscreen']", "title").Contains("Enter full screen"))
        //                driver._clickByJavaScriptExecutor("//button[@class='fullscreen']/div");

        //            Thread.Sleep(1000);
        //            //Assert.AreEqual("Exit full screen", driver._getAttributeValue("xpath", "//button[@class='fullscreen']", "title"), "'Full Scrren' not enabled.");
        //            Results.WriteStatus(test, "Pass", "Clicked, Full Scrren Icon and Verified Screen.");
        //            break;

        //        case "exitscreen":
        //            if (driver._isElementPresent("xpath", "//button[@class='fullscreen']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement exitscreenElement = driver._findElement("xpath", "//*[@id='player']/div[1]");
        //                    driver.MouseHoverByJavaScript(exitscreenElement);
        //                }
        //            }

        //            if (driver._getAttributeValue("xpath", "//button[@class='fullscreen']", "title") == "Exit full screen")
        //                driver._clickByJavaScriptExecutor("//button[@class='fullscreen']/div");

        //            Thread.Sleep(1000);
        //            //Assert.AreEqual("Enter full screen", driver._getAttributeValue("xpath", "//button[@class='fullscreen']", "title"), "'Full Scrren' not Disabled.");
        //            Results.WriteStatus(test, "Pass", "Clicked, Exit Full Scrren Icon and Verified Screen.");
        //            break;

        //        case "hd":
        //            if (driver._isElementPresent("xpath", "//button[@class='hd' and @title='Select video quality']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement hdElement = driver._findElement("xpath", "//*[@id='player']/div[1]");
        //                    driver.MouseHoverByJavaScript(hdElement);
        //                }
        //            }

        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@class='hd' and @title='Select video quality']"));
        //            driver._click("xpath", "//button[@class='hd' and @title='Select video quality']"); Thread.Sleep(1000);

        //            Results.WriteStatus(test, "Pass", "Clicked, HD Icon.");
        //            break;

        //        case "volume":
        //            if (driver._isElementPresent("xpath", "//div[@class='volume' and @role='slider']") == false)
        //            {
        //                while (driver._getAttributeValue("xpath", "//*[@id='player']", "tabindex") != null)
        //                {
        //                    IWebElement volumeElement = driver._findElement("xpath", "//*[@id='player']/div[1]");
        //                    driver.MouseHoverByJavaScript(volumeElement);
        //                }
        //            }

        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@class='hd' and @title='Select video quality']"));
        //            driver._click("xpath", "//button[@class='hd' and @title='Select video quality']"); Thread.Sleep(1000);

        //            Results.WriteStatus(test, "Pass", "Clicked, HD Icon.");
        //            break;

        //        default:
        //            break;
        //    }
        //    driver._selectFrameToDefaultContent();

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// to Verify Error Message on Login Page
        ///// </summary>
        ///// <param name="message">Error Message Content</param>
        ///// <returns></returns>
        //public Login verifyErrorMessageOnLoginPage(string message)
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_ErrorMsg", 10), "Error Message not Present.");
        //    Assert.AreEqual(message, driver._getText("id", "ctl00_ContentPlaceHolder1_ErrorMsg"), "Error Message not Match.");
        //    Results.WriteStatus(test, "Pass", "Verified, Error Message on Login Page.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Email and Password Field on Login Page
        ///// </summary>
        ///// <param name="BlankValue">Verify Email Address & Password Field Blank</param>
        ///// <returns></returns>
        //public Login verifyEmailAndPasswordFieldOnLoginPage(bool BlankValue = false)
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");

        //    if (BlankValue)
        //    {
        //        Assert.AreEqual("", driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Email Address' Field not Blank.");
        //        Assert.AreEqual("", driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Password' Field not Blank.");

        //        Results.WriteStatus(test, "Pass", "Verified, Email Address and Password Field on Login Page.");
        //    }
        //    else
        //    {
        //        Assert.AreEqual(Email[0].ToString(), driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Email Address' Field is Blank.");
        //        Assert.AreEqual(true, driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress") != "", "'Password' Field is Blank.");

        //        Results.WriteStatus(test, "Pass", "Verified, Email Address and Password Field on Login Page.");
        //    }

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Remember Checkbox on Login Page
        ///// </summary>
        ///// <param name="isChecked">Checked box Checked verify</param>
        ///// <returns></returns>
        //public Login verifyRememberCheckboxOnLoginPage(bool unChecked = false)
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_rememberme"), "Remember Checkbox not Present.");

        //    if (unChecked)
        //    {
        //        Assert.AreEqual("x-grid3-row-checker", driver._getAttributeValue("id", "ctl00_ContentPlaceHolder1_rememberme", "class"), "Remember Checkbox not Checked.");
        //        Results.WriteStatus(test, "Pass", "Verified, Remember Checkbox not Checked on Login Page.");
        //    }
        //    else
        //    {
        //        Assert.AreEqual("x-grid3-row-checker-active", driver._getAttributeValue("id", "ctl00_ContentPlaceHolder1_rememberme", "class"), "Remember Checkbox not Checked.");
        //        Results.WriteStatus(test, "Pass", "Verified, Remember Checkbox Checked on Login Page.");
        //    }

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click Links on Login Page
        ///// </summary>
        ///// <param name="linkName">Link Name to Click</param>
        ///// <returns></returns>
        //public Login clickLinksOnLoginPage(string linkName)
        //{
        //    bool link = false;
        //    if (linkName.Equals("Reset Password"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_RetrivePasswordLink"), "'Reset Password' Link not Present.");
        //        driver._click("id", "ctl00_ContentPlaceHolder1_RetrivePasswordLink");
        //        driver._waitForElement("id", "ctl00_ContentPlaceHolder1_tblUserEmail", 20);
        //        link = true;
        //    }

        //    if (linkName.Equals("Request Access"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[@class='request-accessLink']"), "'Request Access' Link not Present.");
        //        driver._clickByJavaScriptExecutor("//a[@class='request-accessLink']");
        //        driver._waitForElement("id", "popupDiv1", 20);
        //        link = true;
        //    }

        //    Assert.AreEqual(true, link);
        //    Results.WriteStatus(test, "Pass", "Clicked, " + linkName + " Link on Login Page .");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Navigate URL to Verify Login Page, Enter Credential and Verify Home Page with Client Name & Database
        ///// </summary>
        ///// <param name="clientName">Select Client Name from List</param>
        ///// <param name="database">Select Database from List</param>
        ///// <param name="columnNo">Column Number to Find Data from the Excel Sheet</param>
        ///// <returns></returns>
        //public Login loginAndVerifyHomePageWithClientAndDatabase(string clientName = "Procter & Gamble", string database = "Detail Data (PEP)", int columnNo = 0)
        //{
        //    navigateToLoginPage().VerifyLoginPage();
        //    loginUsingValidEmailIdAndPassword(columnNo).clickSignInButton();

        //    Home homePage = new Home(driver, test);

        //    //homePage.verifyHomePage();
        //    //if (clientName != "")
        //    //    homePage.selectClientOrDatabaseFromUserProfileMenu("Client Name", clientName);
        //    //if (database != "")
        //    //    homePage.selectClientOrDatabaseFromUserProfileMenu("Database", database);

        //    return new Login(driver, test);
        //}

        //#region Browser Requirement

        ///// <summary>
        ///// To Verify Browser Requirement Popup Window
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyBrowserRequirementPopupWindow()
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "popupDiv1", 15), "Browser Requirement Popup Window not found.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "dvTitle1"), "'Alert' Title not Found.");
        //    Assert.AreEqual("Alert", driver._getText("id", "dvTitle1").Trim().Replace("\r\n", ""), "'Alert' Title not Match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "closepopup1"), "'Close' Icon not Present on Winodw.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']"), "Content Not Present on Popup Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[2]"), "'Minimum Requirements' Column not Present.");
        //    Assert.AreEqual("Minimum Requirements", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[2]").Trim().Replace("\r\n", ""), "'Minimum Requirements' Column Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[3]"), "'Your System' Column not Present.");
        //    Assert.AreEqual("Your System", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[3]").Trim().Replace("\r\n", ""), "'Your System' Column Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[4]"), "'Status' Column not Present.");
        //    Assert.AreEqual("Status", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[1]/td[4]").Trim().Replace("\r\n", ""), "'Status' Column Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[2]/td[1]"), "'Browser' Row not Present.");
        //    Assert.AreEqual("Browser", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[2]/td[1]").Trim().Replace("\r\n", ""), "'Browser' Row Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[3]/td[1]"), "'JavaScript' Row not Present.");
        //    Assert.AreEqual("JavaScript", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[3]/td[1]").Trim().Replace("\r\n", ""), "'JavaScript' Row Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[4]/td[1]"), "'Cookies' Row not Present.");
        //    Assert.AreEqual("Cookies", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[4]/td[1]").Trim().Replace("\r\n", ""), "'Cookies' Row Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[5]/td[1]"), "'Popups' Row not Present.");
        //    Assert.AreEqual("Popups", driver._getText("xpath", "//table[@class='messagepopup']/tbody/tr/td/div/table/tbody/tr[5]/td[1]").Trim().Replace("\r\n", ""), "'Popups' Row Name not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "btn1no0"), "Ok Button not Present.");

        //    Results.WriteStatus(test, "Pass", "Verified, Access Request Popup Window.");
        //    return new Login(driver, test);
        //}

        //#endregion

        //#region Reset Password

        ///// <summary>
        ///// To Verify Reset Password Page
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyResetPasswordPage()
        //{
        //    Assert.AreEqual(true, driver._waitForElement("xpath", "//img[contains(@src,'/ImagesPreLogin/Logo.png')]", 15), "Numerator Logo not found on Page.");
        //    Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_tdInstruction", 15), "Numerator Reset Password Page not found.");

        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='headercontainer']"), "'LEARN MORE ABOUT MARKET TRACK' Label not Found.");
        //    //Assert.AreEqual("LEARN MORE ABOUTMARKET TRACK", driver._getText("xpath", "//td[@class='headercontainer']").Trim().Replace("\r\n", ""), "'LEARN MORE ABOUT MARKET TRACK' Label not Match.");

        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvVideo']"), "Video not Present.");

        //    Assert.AreEqual("Reset Numerator Password", driver._getText("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "'Reset Numerator Password' Label not Match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_tdInstruction2"), "Second Instruction not found.");
        //    Assert.AreEqual("If you have forgotten, misplaced or want to change your Numerator password, enter your email address below and click the submit button", driver._getText("id", "ctl00_ContentPlaceHolder1_tdInstruction2"), "'Second Instruction' message not Match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='inputFieldLabel' and contains(text(), 'Please send Password Reset link to my email address:')]"), "'Please send Password Reset link to my email address:' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_userEmail"), "Email Address input area not found on Page.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_submitEmail"), "Submit Button not found on Page.");
        //    Assert.AreEqual("Submit", driver._getValue("id", "ctl00_ContentPlaceHolder1_submitEmail"), "'Submit' Button Label not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ButtonLoginBack"), "Back to Login Button not found on Page.");
        //    Assert.AreEqual("Back to Login", driver._getValue("id", "ButtonLoginBack"), "'Back to Login' Button Label not match.");

        //    //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not found.");
        //    //Assert.AreEqual("View Privacy Policy", driver._getText("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not match.");

        //    Results.WriteStatus(test, "Pass", "Verified, Reset Password Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Enter Email Address on Reset Password Page
        ///// </summary>
        ///// <param name="validEmail">Valid Email Address</param>
        ///// <returns></returns>
        //public Login enterEmailAddressOnResetPasswordPage(bool validEmail = false)
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_userEmail"), "Email Address textarea not Pesent.");

        //    if (validEmail)
        //    {
        //        driver._type("id", "ctl00_ContentPlaceHolder1_userEmail", Email[0].ToString());
        //        Results.WriteStatus(test, "Pass", "Entered, Valid Email Address on Reset Password Page.");
        //    }
        //    else
        //    {
        //        driver._type("id", "ctl00_ContentPlaceHolder1_userEmail", driver._randomString(5) + "@test.com");
        //        Results.WriteStatus(test, "Pass", "Entered, Invalid Email Address on Reset Password Page.");
        //    }
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click Button on Reset Password Page
        ///// </summary>
        ///// <param name="buttonName">Button Name to Click</param>
        ///// <returns></returns>
        //public Login clickButtonOnResetPasswordPage(string buttonName)
        //{
        //    bool button = false;

        //    if (buttonName.Equals("Submit"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_submitEmail"), "'Submit' Button not Present.");
        //        driver._click("id", "ctl00_ContentPlaceHolder1_submitEmail");
        //        button = true;
        //    }

        //    if (buttonName.Equals("Back to Login"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "ButtonLoginBack"), "'Back to Login' Button not Present.");
        //        driver._click("id", "ButtonLoginBack");
        //        button = true;
        //    }

        //    Assert.AreEqual(true, button);
        //    Results.WriteStatus(test, "Pass", "Clicked, " + buttonName + " Button on Reset Password Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Warning Message and Login link on Reset Password Page
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyWarningMessageAndLoginLinkOnResetPasswordPage()
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_messageWorning", 15), "Warning Message not Present.");
        //    Assert.AreEqual("Password Reset link has been sent to your email address.", driver._getText("id", "ctl00_ContentPlaceHolder1_messageWorning"), "'Password Reset link has been sent to your email address.' Warning message not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ButtonLoginBack"), "'Login' Button not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified, Watning Message & Login Link on Reset Password Page.");

        //    return new Login(driver, test);
        //}

        //#endregion

        //#region Access Request Popup Window

        ///// <summary>
        ///// To Verify Access Request Popup Window
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyAccessRequestPopupWindow()
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "popupDiv1", 15), "Access Request Popup Window not found.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "dvTitle1"), "'Access Request' Title not Found.");
        //    Assert.AreEqual("Access Request", driver._getText("id", "dvTitle1").Trim().Replace("\r\n", ""), "'Access Request' Title not Match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "closepopup1"), "'Close' Icon not Present on Winodw.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='headerMsg']"), "'Access Request Form' Information not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[@class='aLinkColor']"), "'Support Email' address not Present.");
        //    Assert.AreEqual("promohelp.numerator.com", driver._getText("xpath", "//a[@class='aLinkColor']").Trim().Replace("\r,\n", ""), "'Support Email' address not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Company or Organization :']"), "'Company or Organization' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtCompany"), "'Company or Organization' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Country:']"), "'Country' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ddlCountry"), "'Country' Dropdown not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* First Name:']"), "'First Name' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtFirstName"), "'First Name' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Last Name:']"), "'Last Name' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtLastName"), "'Last Name' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Email Address:']"), "'Email Address' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtEmailAddress"), "'Email Address' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Phone:']"), "'Phone' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtPhone"), "'Phone' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = '* Title:']"), "'Title' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtTitle"), "'Title' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='request-accessLabel' and text() = 'Additional Information/Comments:']"), "'Additional Information/Comments' Label not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "txtComments"), "'Additional Information/Comments' textarea not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "btnSave"), "Save Button not found on Page.");
        //    Assert.AreEqual("Save", driver._getValue("id", "btnSave"), "'Save' Button Label not match.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "btnCancel"), "Cancel Button not found on Page.");
        //    Assert.AreEqual("Cancel", driver._getValue("id", "btnCancel"), "'Cancel' Button Label not match.");

        //    Results.WriteStatus(test, "Pass", "Verified, Access Request Popup Window.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Enter Data on Access Request Window
        ///// </summary>
        ///// <returns></returns>
        //public String enterDataOnAccessRequestWindow()
        //{
        //    driver._type("id", "txtCompany", "AutoComp" + driver._randomString(4, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Company Name on Access Request Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='ddlCountry']/option"), "Country List not Present on Screen.");
        //    IList<IWebElement> countryCollection = driver._findElements("xpath", "//*[@id='ddlCountry']/option");
        //    Random select = new Random();
        //    int x = select.Next(0, countryCollection.Count);
        //    countryCollection[x].Click();
        //    Results.WriteStatus(test, "Pass", "Selected, Country on Access Request Window.");

        //    string firstName = "AutoF" + driver._randomString(4, true);
        //    driver._type("id", "txtFirstName", firstName);
        //    Results.WriteStatus(test, "Pass", "Entered, First Name on Access Request Window.");

        //    driver._type("id", "txtLastName", "AutoL" + driver._randomString(4, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Last Name on Access Request Window.");

        //    driver._type("id", "txtEmailAddress", "Test" + driver._randomString(4, true) + "@test.com");
        //    Results.WriteStatus(test, "Pass", "Entered, Email Address on Access Request Window.");

        //    driver._type("id", "txtPhone", driver._randomString(8, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Phone on Access Request Window.");

        //    driver._type("id", "txtTitle", "AutoT" + driver._randomString(5, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Title on Access Request Window.");

        //    driver._type("id", "txtComments", "Comment" + driver._randomString(6, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Comments on Access Request Window.");

        //    return firstName;
        //}

        ///// <summary>
        ///// Click Button on Access Request Popup Window
        ///// </summary>
        ///// <param name="buttonName">Button Name to Click</param>
        ///// <returns></returns>
        //public Login clickButtonOnAccessRequestPopupWindow(string buttonName)
        //{
        //    bool button = false;

        //    if (buttonName.Equals("Save"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "btnSave"), "'Save' Button not Present.");
        //        //driver._click("id", "btnSave");
        //        button = true;
        //    }

        //    if (buttonName.Equals("Cancel"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "btnCancel"), "'Cancel' Button not Present.");
        //        driver._click("id", "btnCancel");
        //        button = true;
        //    }

        //    Assert.AreEqual(true, button);
        //    Results.WriteStatus(test, "Pass", "Clicked, " + buttonName + " Button on Access Request Popup Window.");
        //    return new Login(driver, test);
        //}

        //#endregion

        //#region Privacy Policy

        ///// <summary>
        ///// Verify Image Logo on Login Page
        ///// </summary>
        ///// <param name="logo">Logo Title to Verify</param>
        ///// <returns></returns>
        //public Login verifyImageLogoOnLoginPage(string logo)
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_imgMtLogo"), "'Logo' not Present on Login Page.");
        //    Assert.AreEqual(true, driver._getAttributeValue("id", "ctl00_imgMtLogo", "title").Contains(logo), "'" + logo + "' Logo not Match.");
        //    Results.WriteStatus(test, "Pass", "Verified, '" + logo + "' Image Logo on Login Page.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Get Image Logo Location on Home Page
        ///// </summary>
        ///// <returns>Image Title</returns>
        //public String getImageLogoLocationOnLoginPage()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_imgMtLogo"), "'Image Logo' not Present on Login Page.");
        //    string location = driver._getAttributeValue("id", "ctl00_imgMtLogo", "src");
        //    Results.WriteStatus(test, "Pass", "Get '" + location + "' Image Logo Location on Login Page.");

        //    return location;
        //}

        ///// <summary>
        ///// Click 'View Privacy Link' on Login Page
        ///// </summary>
        ///// <returns></returns>
        //public Login clickViewPrivacyLinkOnLoginPage()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not found.");
        //    Assert.AreEqual("View Privacy Policy", driver._getText("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not match.");
        //    driver._clickByJavaScriptExecutor("//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a");
        //    Results.WriteStatus(test, "Pass", "Clicked, 'View Privacy Policy' Link on Login Page.");

        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Privacy Policy Popup Window on Page
        ///// </summary>
        ///// <param name="logoLocation">Verify Logo</param>
        ///// <param name="supportMailAddress">Verify Support Mail Address</param>
        ///// <returns></returns>
        //public Login verifyPrivacyPolicyPopupWindowOnPage(string logoLocation = "", string supportMailAddress = "")
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "popupDiv1", 15), "'Privacy Policy' Popup Window not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "dvTitle1"), "'Privacy Policy' Label not Present on Window.");
        //    Assert.AreEqual("Privacy Policy", driver._getText("id", "dvTitle1"), "'Privacy Policy' Label not match on Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "privacypolicy"), "'Privacy Policy' Content not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "imgPrivacyLogo"), "'Numerator' Logo not Present on Window.");

        //    if (logoLocation != "")
        //        Assert.AreEqual(logoLocation, driver._getAttributeValue("id", "imgPrivacyLogo", "src"), "'Logo' not verified on Popup Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "btn1no0"), "'Close' Button not Present on Window.");
        //    Assert.AreEqual("Close", driver._getValue("id", "btn1no0"), "'Close' Button Label not match on Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "closepopup1"), "'Close' Icon not Present on Window.");

        //    if (supportMailAddress != "")
        //    {
        //        IWebElement elements = driver._findElement("id", "privacypolicy");
        //        IList<IWebElement> content = elements.FindElements(By.TagName("a"));

        //        for (int i = 0; i < content.Count; i++)
        //            Assert.AreEqual(supportMailAddress, content[i].Text, "'" + content[i].Text + "' Support Email not Match.");
        //    }

        //    Results.WriteStatus(test, "Pass", "Verified, Privacy Policy Popup Window on Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Click Close Button On Privacy Policy Window
        ///// </summary>
        ///// <returns></returns>
        //public Login clickCloseButtonOnPrivacyPolicyWindow()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "btn1no0"), "'Close' Button not Present on Window.");
        //    driver._click("id", "btn1no0");
        //    Thread.Sleep(1000);
        //    Results.WriteStatus(test, "Pass", "Clicked, Close Button on Privacy Policy Window.");

        //    return new Login(driver, test);
        //}

        //#endregion

        //#region Outlook Mails Methods

        ///// <summary>
        ///// Verify Outlook Login Screen to Enter Credential and Click SignIn Button
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyOutlookLoginScreenToEnterCredentialAndClickSignInButton()
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";

        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Outlook");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Outlook");

        //    if (driver._isElementPresent("id", "cred_userid_inputtext"))
        //    {

        //        Assert.IsTrue(driver._waitForElement("id", "cred_userid_inputtext", 20), "Email Address Textarea not Present.");
        //        Assert.AreEqual(true, driver._isElementPresent("id", "cred_password_inputtext"), "Password Textarea not Present.");
        //        Assert.AreEqual(true, driver._isElementPresent("id", "cred_sign_in_button"), "SignIn Button not Present.");
        //        Results.WriteStatus(test, "Pass", "Verified, Outlook Login Screen.");

        //        driver._type("id", "cred_userid_inputtext", Email[0].ToString());
        //        Thread.Sleep(1000);
        //        driver._type("id", "cred_password_inputtext", Password[0].ToString());
        //        Thread.Sleep(3000);

        //        driver._click("id", "cred_sign_in_button");
        //        Thread.Sleep(5000);
        //        Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
        //    }
        //    else
        //    {
        //        Assert.IsTrue(driver._waitForElement("xpath", "//input[@type='email' and @name='loginfmt']", 20), "Email Address Textarea not Present.");
        //        driver._type("xpath", "//input[@type='email' and @name='loginfmt']", Email[0].ToString());
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Next']"), "Next Button not Present.");
        //        driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Next']");
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@name='passwd' and @type='password']"), "Password Textarea not Present.");
        //        driver._type("xpath", "//input[@name='passwd' and @type='password']", Password[0].ToString());
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Sign in']"), "Sign in Button not Present.");
        //        driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Sign in']");
        //        Thread.Sleep(1000);

        //        if (driver._isElementPresent("xpath", "//input[@type='button' and @value='No']"))
        //            driver._clickByJavaScriptExecutor("//input[@type='button' and @value='No']");
        //        Thread.Sleep(1000);
        //    }

        //    Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Outlook Home Page 
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyOutlookHomePage()
        //{
        //    Assert.IsTrue(driver._waitForElement("xpath", "//span[@title='Inbox' and text() = 'Inbox']", 15), "Inbox Folder not Present.");
        //    Assert.IsTrue(driver._waitForElement("xpath", "//div[@role='option' and @aria-haspopup='true']", 15), "Emails List not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified, Outlook Home Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Select Reset Password Mail to Open Reset Link
        ///// </summary>
        ///// <returns></returns>
        //public Login selectResetPasswordMailToOpenResetLink()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']"), "Mails Subject not Present.");
        //    IList<IWebElement> mailSubjects = driver._findElements("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']");
        //    bool avail = false;

        //    for (int m = 0; m < mailSubjects.Count; m++)
        //    {
        //        if (mailSubjects[m].Text.Contains("Your Numerator Password Reset Request"))
        //        {
        //            mailSubjects[m].Click();
        //            avail = true;
        //            break;
        //        }
        //    }
        //    Assert.AreEqual(true, avail, "'Your Numerator Password Reset Request' Mail not Present.");
        //    Results.WriteStatus(test, "Pass", "Selected, Reset Password Email.");

        //    Assert.IsTrue(driver._waitForElement("xpath", "//div[@aria-label='Message Contents']", 15), "Message Content not Present.");
        //    IWebElement body = driver._findElement("xpath", "//div[@aria-label='Message Contents']");
        //    IList<IWebElement> content = body.FindElements(By.TagName("a"));
        //    bool resetLink = false;

        //    for (int i = 0; i < content.Count(); i++)
        //    {
        //        if (content[i].GetAttribute("href").Contains("ResetPassword.aspx"))
        //        {
        //            content[i].Click();
        //            resetLink = true;
        //            Thread.Sleep(5000);
        //            break;
        //        }
        //    }
        //    Assert.AreEqual(true, resetLink, "'Reset Password' Link not Present on Content.");
        //    Results.WriteStatus(test, "Pass", "Clicked, Reset Password Link from Email.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// select Summary Report Requested Data Mail To Download Report
        ///// </summary>
        ///// <returns></returns>
        //public Login selectSummaryReportRequestedDataMailToDownloadReport(string mailTitle = "Summary Report - Requested Data")
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']"), "Mails Subject not Present.");
        //    Thread.Sleep(8000);
        //    IList<IWebElement> mailSubjects = driver._findElements("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']");
        //    bool avail = false;

        //    for (int m = 0; m < mailSubjects.Count; m++)
        //    {
        //        if (mailSubjects[m].Text.Contains(mailTitle))
        //        {
        //            mailSubjects[m].Click();
        //            avail = true;
        //            break;
        //        }
        //    }
        //    Assert.AreEqual(true, avail, "'" + mailTitle + "' Mail not Present.");
        //    Results.WriteStatus(test, "Pass", "Selected, '" + mailTitle + "' Email.");

        //    Assert.IsTrue(driver._waitForElement("xpath", "//span[@class='_fc_4 o365buttonLabel' and contains(@id,'_ariaId_') and text() = 'Download']", 15), "Download Link not Present.");
        //    driver._clickByJavaScriptExecutor("//span[@class='_fc_4 o365buttonLabel' and contains(@id,'_ariaId_') and text() = 'Download']");
        //    Thread.Sleep(8000);
        //    Results.WriteStatus(test, "Pass", "Clicked, Download File Link on Email.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Reset Screen to Enter Password
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyResetScreenToEnterPassword()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_imgMtLogo"), "Image Logo Not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "'Reset Numerator Password' Instruction not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_Text1"), "Email Address not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_password"), "New Password Textarea not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_RePassword"), "Re-type New Password Textarea not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_submitPassword"), "Reset Password Button not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified, Reset Password Screen.");

        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");

        //    driver._type("id", "ctl00_ContentPlaceHolder1_password", Password[0].ToString());
        //    Thread.Sleep(1000);
        //    driver._type("id", "ctl00_ContentPlaceHolder1_RePassword", Password[0].ToString());
        //    Thread.Sleep(1000);

        //    driver._click("id", "ctl00_ContentPlaceHolder1_submitPassword");
        //    Thread.Sleep(5000);
        //    Results.WriteStatus(test, "Pass", "Entered, Password and Clicked Reset Password Button.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Password Updated Label and Click Login Link
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyPasswordUpdatedAndClickLoginLink()
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_messageWorning", 10), "'Password updated successfully.' Label not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='ctl00_ContentPlaceHolder1_messageWorning']/a"), "'Login' link not Present.");
        //    driver._clickByJavaScriptExecutor("//*[@id='ctl00_ContentPlaceHolder1_messageWorning']/a");
        //    Thread.Sleep(3000);
        //    Results.WriteStatus(test, "Pass", "Verified, Password Updated Label and Clicked Login Link.");
        //    return new Login(driver, test);
        //}

        //#endregion

        //#region Gmail Methods

        ///// <summary>
        ///// Verify Gmail Login Page and Entered Credential
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyGmailLoginPageAndEnterCredential()
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Gmail");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Gmail");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "initialView"), "Login Page not present.");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='identifierId' and @type='email' and @name='identifier']"), "Email or phone Textarea not present.");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@type='button' and text()='Forgot email?']"), "'Forgot email?' Link not present.");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='identifierNext' and @role='button']"), "'Next' Button not present.");

        //    driver._click("xpath", "//input[@id='identifierId' and @type='email' and @name='identifier']"); Thread.Sleep(1000);
        //    driver._type("xpath", "//input[@id='identifierId' and @type='email' and @name='identifier']", "" + Email[0].ToString() + "");
        //    driver._clickByJavaScriptExecutor("//div[@id='identifierNext' and @role='button']");
        //    Results.WriteStatus(test, "Pass", "Entered '" + Email[0].ToString() + "' Email & Clicked Next Button.");
        //    Thread.Sleep(5000);

        //    if (driver._isElementPresent("xpath", "//input[@type='email']"))
        //    {
        //        driver._type("xpath", "//input[@type='email']", "" + Email[0].ToString() + ""); Thread.Sleep(2000);
        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Next']"), "'Next' Button not Present.");
        //        driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Next']"); Thread.Sleep(5000);
        //        Results.WriteStatus(test, "Pass", "Entered '" + Email[0].ToString() + "' Email & Clicked Next Button for MicroSoft Page with Numerator Security.");

        //        if (driver._isElementPresent("xpath", "//div[@id='displayName' and @class='identity']"))
        //        {
        //            Assert.AreEqual("" + Email[0].ToString() + "", driver._getText("xpath", "//div[@id='displayName' and @class='identity']"), "'" + Email[0].ToString() + "' Display Email Address not present.");
        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='password']"));

        //            driver._type("xpath", "//input[@type='password']", "" + Password[0].ToString() + ""); Thread.Sleep(2000);
        //            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Sign in']"), "'Sign in' Button not Present.");
        //            driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Sign in']"); Thread.Sleep(5000);
        //            Results.WriteStatus(test, "Pass", "Entered '" + Password[0].ToString() + "' Password & Clicked SignIn Button for MicroSoft Page with Numerator Security.");

        //            if (driver._isElementPresent("xpath", "//input[@type='submit' and @value='Yes']"))
        //            {
        //                driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Yes']"); Thread.Sleep(3000);
        //                Results.WriteStatus(test, "Pass", "Clicked YES Option from popup Window.");
        //            }
        //        }
        //    }

        //    if (driver._isElementPresent("xpath", "//div[@role='button']//.//span[text()='Continue']"))
        //    {
        //        driver._clickByJavaScriptExecutor("//div[@role='button']//.//span[text()='Continue']"); Thread.Sleep(5000);
        //        Results.WriteStatus(test, "Pass", "Clicked Continue Button to Login into Gmail Account.");
        //    }

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[@aria-label='Google apps' and @role='button']"), "'Google' App Button not Present.");
        //    driver._clickByJavaScriptExecutor("//a[@aria-label='Google apps' and @role='button']");
        //    driver._waitForElement("xpath", "//a[@aria-label='Google apps' and @role='button' and @aria-expanded='true']", 10);

        //    IWebElement Gmail = driver.FindElement(By.XPath("//div[@aria-label='Google apps' and @role='region' and @aria-hidden='false']/ul/li/a[contains(@href,'mail.google.com')]"));
        //    Gmail.Click(); Thread.Sleep(7000);
        //    Results.WriteStatus(test, "Pass", "Clicked 'Google App' Option and Select 'Gmail' on it.");

        //    driver.SwitchTo().Window(driver.WindowHandles.Last());
        //    Thread.Sleep(5000);
        //    driver._waitForElement("xpath", "//a[contains(@href,'#inbox') and @title='Inbox']", 15);
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[contains(@href,'#inbox') and @title='Inbox']"), "'Inbox' Title not present.");

        //    driver._clickByJavaScriptExecutor("//a[contains(@href,'#inbox') and @title='Inbox']");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//tbody/tr[@draggable='true']"));
        //    Results.WriteStatus(test, "Pass", "Verify Gmail Screen and Clicked Inbox Folder.");
        //    Results.WriteStatus(test, "Pass", "Verified Gamil Login Page and Entered Credential.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Subject Line of Email and Click Email
        ///// </summary>
        ///// <param name="subject">Subject Line to Verify</param>
        ///// <returns></returns>
        //public Login verifySubjectLineOfEmailAndClickEmail(string common_Subject, string subjectLine)
        //{
        //    string EmailSubjectLine = "Numerator Promotions Intel " + common_Subject;
        //    if (subjectLine != "")
        //        EmailSubjectLine = EmailSubjectLine + " - " + subjectLine;

        //    for (int i = 0; i < 50; i++)
        //    {
        //        if (driver._isElementPresent("xpath", "//table[@class='F cf zt']/tbody/tr//.//div[@role='link']//.//*[contains(text(),'" + EmailSubjectLine + "')]"))
        //            break;
        //        else
        //            Thread.Sleep(3000);
        //    }

        //    driver._waitForElement("xpath", "//table[@class='F cf zt']/tbody/tr//.//div[@role='link']//.//*[contains(text(),'" + EmailSubjectLine + "')]", 20);
        //    while (driver._isElementPresent("xpath", "//table[@class='F cf zt']/tbody/tr//.//div[@role='link']//.//*[contains(text(),'" + EmailSubjectLine + "')]") == true)
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//table[@class='F cf zt']/tbody/tr//.//div[@role='link']//.//*[contains(text(),'" + EmailSubjectLine + "')]"));
        //        driver._click("xpath", "//table[@class='F cf zt']/tbody/tr//.//div[@role='link']//.//*[contains(text(),'" + EmailSubjectLine + "')]"); Thread.Sleep(5000);
        //    }

        //    Assert.AreEqual(true, driver._waitForElement("xpath", "//div[@class='adn ads']"));
        //    Results.WriteStatus(test, "Pass", "Verified Subject Title and Clicked Mail.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify Attached File and Click to Download
        ///// </summary>
        ///// <param name="fileName">FileName to Verify</param>
        ///// <returns></returns>
        //public String verifyAttachedFileAndClickToDownload(string checkFileName = "")
        //{
        //    string fileName = "";
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[contains(@class,'ie') and contains(text(),'" + fileName + "')]"));
        //    driver._scrollintoViewElement("xpath", "//span[contains(@class,'ie') and contains(text(),'" + fileName + "')]");

        //    IWebElement element = driver._findElement("xpath", "//span[contains(@class,'ie') and contains(text(),'" + fileName + "')]");
        //    fileName = driver._getAttributeValue("xpath", "//div[contains(@aria-label,'Download attachment " + fileName + "') and @role='button']", "aria-label").Replace("Download attachment ", "").Replace(".xlsx", "").Replace(".pdf", "").Replace(".zip", "").Replace(".pptx", "");
        //    driver.MouseHoverByJavaScript(element);
        //    driver._clickByJavaScriptExecutor("//div[contains(@aria-label,'Download attachment " + fileName + "') and @role='button']"); Thread.Sleep(5000);

        //    if (checkFileName != "")
        //    {
        //        Assert.AreEqual(true, fileName.Contains(checkFileName), "'" + checkFileName + "' Report File Name not match");
        //        Results.WriteStatus(test, "Pass", "Verified Body of Email and Verifeid Download File Name with Report File name.");
        //    }
        //    else
        //        Results.WriteStatus(test, "Pass", "Verified Body of Email and Download File.");
        //    return fileName;
        //}

        ///// <summary>
        ///// Verify Tab Values of Downloaded File From Email
        ///// </summary>
        ///// <param name="chartName"></param>
        ///// <param name="tabName"></param>
        ///// <returns></returns>
        //public Login verifyTabValuesOfDownloadedFileFromEmail(string downloadedFileName, string[] tabName)
        //{
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;
        //    string FilePath = "";

        //    string sourceDir = ExtentManager.ResultsDir + "\\";
        //    string[] fileEntries = Directory.GetFiles(sourceDir);

        //    foreach (string fileName in fileEntries)
        //    {
        //        if (fileName.Contains(downloadedFileName))
        //        {
        //            FilePath = fileName;
        //            break;
        //        }
        //    }

        //    xlApp = new Excel.Application();
        //    xlWorkBook = xlApp.Workbooks.Open(FilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
        //    bool avail = false;
        //    int num = xlWorkBook.Sheets.Count;

        //    for (int i = 0; i < tabName.Length; i++)
        //    {
        //        for (int s = 1; s <= num; s++)
        //        {
        //            avail = false;
        //            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(s);
        //            if (xlWorkSheet.Name.Equals(tabName[i]))
        //            {
        //                avail = true;
        //                break;
        //            }
        //        }
        //        Assert.AreEqual(true, avail, "'" + xlWorkSheet.Name + "' Tab Not Present.");
        //    }

        //    Assert.AreEqual(true, avail, "" + tabName + " Tab Not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified " + tabName + " Tab from Excel File.");
        //    return new Login(driver, test);
        //}

        ///// <summary>
        ///// Verify File Downloaded Or Not from Email
        ///// </summary>
        ///// <param name="downloadedFileName">Downloaded File Name</param>
        ///// <returns></returns>
        //public Login verifyFileDownloadedOrNotFromEmail(string downloadedFileName)
        //{
        //    string FilePath = "";
        //    bool avail = false;
        //    string sourceDir = ExtentManager.ResultsDir + "\\";
        //    string[] fileEntries = Directory.GetFiles(sourceDir);

        //    foreach (string fileName in fileEntries)
        //    {
        //        if (fileName.Contains(downloadedFileName))
        //        {
        //            FilePath = fileName;
        //            avail = true;
        //            break;
        //        }
        //    }

        //    Assert.AreEqual(true, avail, "'" + downloadedFileName + "' Downloaded File Not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified '" + downloadedFileName + "' File from Email.");
        //    return new Login(driver, test);
        //}

        //#endregion

        //#endregion

        #region Public Methods
        public Login(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.login = driver;
            test = testReturn;
        }

        public IWebDriver driver
        {
            get { return this.login; }
            set { this.login = value; }
        }

        /// <summary> Navigate to login page (Login URL get From the Login.xlsx Sheet) </summary>
        /// <returns></returns>
        public Login navigateToLoginPage()
        {
            driver.Navigate().GoToUrl(Common.ApplicationURL);
            Results.WriteStatus(test, "Pass", "Launched, URL <b>" + Common.ApplicationURL + "</b> successfully.");
            return new Login(driver, test);
        }

        /// <summary> To Verify Login Page </summary>
        /// <param name="Watermark">Verify Watermark</param>
        /// <returns></returns>
        public Login VerifyLoginPage(bool Watermark = true)
        {
            //verifyNavigateURLOnScreen("https://promotionsintel.numerator.com/login.aspx");

            //Assert.AreEqual(true, driver._waitForElement("xpath", "//img[contains(@src,'/ImagesPreLogin/Logo.png')]"), "Numerator Logo not found on Page.");
            Assert.AreEqual(true, driver._waitForElement("class", "logo"), "Numerator Logo not found on Page.");

            Assert.AreEqual(true, driver._isElementPresent("id", "EmailAddress"), "Email input textbox not found on Page.");
            Assert.AreEqual(true, driver._isElementPresent("id", "Password"), "Password input textbox not found on Page.");

            if (Watermark)
            {
                Assert.AreEqual(true, driver._getAttributeValue("id", "EmailAddress", "placeholder").Contains("Email"), "'Email' Watermark not Present.");
                Assert.AreEqual(true, driver._getAttributeValue("id", "Password", "placeholder").Contains("Password"), "'Password' Watermark not Present.");
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'verifyLogin();')]"), "'Log In' button not found on Page.");
            Assert.AreEqual("Log In", driver._findElement("xpath", "//*[contains(@ng-click,'verifyLogin();')]").Text, "'Log In' button text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'openResetPasssword(true);')]"), "'Reset Password' link not found on Page.");
            Assert.AreEqual("Reset Password", driver._findElement("xpath", "//*[contains(@ng-click,'openResetPasssword(true);')]").Text, "'Reset Password' link text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'requestAccess()')]"), "'Request Access' link not found on Page.");
            Assert.AreEqual("Request Access", driver._findElement("xpath", "//*[contains(@ng-click,'requestAccess()')]").Text, "'Request Access' link text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'openPrivacyPolicy()')]"), "'Privacy Notice' link not found on Page.");
            Assert.AreEqual("Privacy Notice", driver._findElement("xpath", "//*[contains(@ng-click,'openPrivacyPolicy()')]").Text, "'Privacy Notice' link text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'ssoLogin()')]"), "'Log in with Numerator Single Sign On' button not found on Page.");
            Assert.AreEqual("Log in with Numerator Single Sign On", driver._findElement("xpath", "//*[contains(@ng-click,'ssoLogin()')]").Text, "'Log in with Numerator Single Sign On' button text not match.");

            Results.WriteStatus(test, "Pass", "Verified, Login page.");
            return new Login(driver, test);
        }

        /// <summary> Verify Navigate URL on Screen </summary>
        /// <param name="url">URL</param>
        /// <returns></returns>
        public Login verifyNavigateURLOnScreen(string url)
        {
            Assert.AreEqual(true, driver.Url.Contains(url), "Navigate URL not Match. \nUrl is : " + driver.Url);
            Results.WriteStatus(test, "Pass", "Verified, Navigate Url " + url + " on Screen.");
            return new Login(driver, test);
        }

        /// <summary> Login using Valid Email Address</summary>
        /// <param name="configurUser">Configure User Login Credentail Enter</param>
        /// <returns></returns>
        public Login loginUsingValidEmailAddress(bool configureUser = false)
        {
            #region Datasheet
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
            string email;
            #endregion

            if (configureUser)
            {
                email = Email[2].ToString();
            }
            else
            {
                email = Email[0].ToString();
            }

            Assert.IsTrue(driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
            driver._type("id", "EmailAddress", email);
            Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);

            return new Login(driver, test);
        }

        ///// <summary> Login using Valid Email Address & Password </summary>
        ///// <param name="configurUser">Configure User Login Credentail Enter</param>
        ///// <returns></returns>
        //public Login loginUsingValidEmailAddressPassword(bool configureUser = false)
        //{
        //    #region Datasheet
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
        //    string email, password = "";
        //    #endregion

        //    if (configureUser)
        //    {
        //        email = Email[2].ToString();
        //        password = Password[2].ToString();
        //    }
        //    else
        //    {
        //        email = Email[4].ToString();
        //        password = Password[4].ToString();
        //    }

        //    Assert.IsTrue(driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
        //    driver._type("id", "EmailAddress", email);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);
        //    Assert.IsTrue(driver._isElementPresent("id", "Password"), "Password Textarea not Present.");
        //    driver._type("id", "Password", password);
        //    Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Password : " + password);

        //    return new Login(driver, test);
        //}

        /// <summary> Login using Valid Email Address & Password </summary>
        /// <param name="configurUser">Configure User Login Credentail Enter</param>
        /// <returns></returns>
        public Login loginUsingValidEmailAddressPassword(bool configureUser = false, string eid = "EmailAddress", string pid = "Password")
        {
            #region Datasheet
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
            string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
            string email, password = "";
            #endregion

            if (configureUser)
            {
                email = Email[2].ToString();
                password = Password[2].ToString();
            }
            else
            {
                email = Email[4].ToString();
                password = Password[4].ToString();
            }

            //Assert.IsTrue(driver._isElementPresent("id", eid), "Email Address Textarea not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id," + eid + ")]"), "'Email' textbox not Present.");
            driver._type("id", eid, email);
            //driver._findElement("xpath", "//*[contains(@id," + eid + ")]").SendKeys(email);
            //Thread.Sleep(1000);
            Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);
            //Assert.IsTrue(driver._isElementPresent("id", pid), "Password Textarea not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id," + pid + ")]"), "'Password' textbox not Present.");
            driver._type("id", pid, password);
            Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Password : " + password);

            return new Login(driver, test);
        }

        /// <summary> Login using Valid Email Id & Password </summary>
        /// <param name="column">Column Number to Find Data from the Excel Sheet</param>
        /// <returns></returns>
        public Login loginUsingValidEmailIdAndPassword(int column)
        {
            #region Datasheet
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
            string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
            string email, password = "";
            #endregion

            email = Email[column].ToString();
            password = Password[column].ToString();

            Assert.IsTrue(driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
            driver._type("id", "EmailAddress", email);
            Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Email Address : " + email);
            Assert.IsTrue(driver._isElementPresent("id", "Password"), "Password Textarea not Present.");
            driver._type("id", "Password", password);
            Results.WriteStatus(test, "Pass", "Information Inputed successfully.<b> Password : " + password);

            return new Login(driver, test);
        }

        /// <summary> Login using Invalid Email address and Password </summary>
        /// <returns></returns>
        public Login loginUsingInvalidEmailAddressAndPassword(bool password, string eid = "@test.com")
        {
            Assert.IsTrue(driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
            driver._type("id", "EmailAddress", driver._randomString(5) + eid);
            Results.WriteStatus(test, "Pass", "Entered Invalid Email address");

            if (password)
            {
                Assert.IsTrue(driver._isElementPresent("id", "Password"), "Password Textarea not Present.");
                driver._type("id", "Password", driver._randomString(6));
                Results.WriteStatus(test, "Pass", "Entered Invalid Password.");
            }
            return new Login(driver, test);
        }

        /// <summary> Login Using Invalid Email Address or Password </summary>
        /// <param name="password">Enter Invalid Password</param>
        /// <returns></returns>
        public Login loginUsingInvalidEmailAddressOrPassword(bool email, string eid = "@test.com")
        {
            if (email)
            {
                Assert.IsTrue(driver._isElementPresent("id", "EmailAddress"), "Email Address Textarea not Present.");
                //driver._clearText("id", "Password");
                driver._type("id", "EmailAddress", driver._randomString(5) + eid);
                Results.WriteStatus(test, "Pass", "Entered Invalid Email address");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("id", "Password"), "Password Textarea not Present.");
                driver._clearText("id", "EmailAddress");
                driver._type("id", "Password", driver._randomString(6));
                Results.WriteStatus(test, "Pass", "Entered Invalid Password.");
            }

            return new Login(driver, test);
        }

        /// <summary> Click Sign In Button </summary>
        /// <returns></returns>
        public Login clickSignInButton()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[contains(text(), 'Log In')]"), "'Sign In' Button not Present.");
            driver._clickByJavaScriptExecutor("//button[contains(text(), 'Log In')]");
            if(driver._isElementPresent("xpath", "//button[contains(text(), 'Log In')]"))
                driver._clickByJavaScriptExecutor("//button[contains(text(), 'Log In')]");
            Results.WriteStatus(test, "Pass", "Clicked, Sign In Button.");

            return new Login(driver, test);
        }

        /// <summary> Click on any Button </summary>
        /// <returns></returns>
        public Login clickOnButton(string clickID)
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'" + clickID + "')]"), "'Button not found on Page.");
            driver._clickByJavaScriptExecutor("//*[contains(@ng-click,'" + clickID + "')]");
            Results.WriteStatus(test, "Pass", "Clicked, on Button.");

            return new Login(driver, test);
        }

        /// <summary> to Verify Error Message on Login Page </summary>
        /// <param name="message">Error Message Content</param>
        /// <returns></returns>
        public Login verifyErrorMessageOnLoginPage(string message)
        {
            Assert.IsTrue(driver._waitForElement("class", "ng-binding", 10), "Error Message not Present.");
            Assert.AreEqual(message, driver._getText("xpath", "//*[@ng-bind-html='objLoginData.errorMessage' and @class='ng-binding']"), "Error Message not Match.");
            //Console.WriteLine("message " + message);
            //Console.WriteLine("on page error " + driver._getText("xpath", "//*[@ng-bind-html='objLoginData.errorMessage' and @class='ng-binding']"));
            //Console.WriteLine(driver._findElement("xpath", "//*[@ng-bind-html='objLoginData.errorMessage' and @class='ng-binding']").Text);
            Results.WriteStatus(test, "Pass", "Verified, Error Message on Login Page.");

            return new Login(driver, test);
        }

        /// <summary> to Verify Success Message on Page </summary>
        /// <param name="message">Success Message Content</param>
        /// <returns></returns>
        public Login verifySuccessMessageOnPage(string message)
        {
            Assert.IsTrue(driver._waitForElement("class", "ng-binding", 10), "Success Message not Present.");
            Assert.AreEqual(message, driver._getText("xpath", "//*[@ng-if='objLoginData.showSuccessMessage']/div[1]"), "Success Message not Match.");
            //Console.WriteLine("message : " + message);
            //Console.WriteLine("on page : " + driver._findElement("xpath", "//*[@ng-if='objLoginData.showSuccessMessage']").Text);
            Results.WriteStatus(test, "Pass", "Verified, Success Message on Page.");

            return new Login(driver, test);
        }

        /// <summary> to Verify Website Maintenance on Page </summary>
        /// <param name="message">Message Content</param>
        /// <returns></returns>
        public Login verifyWebsiteMaintenanceMessageOnPage(string message)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//*[@ng-bind-html='objServerMaintenance.alertText']", 10), "Website Maintenance Message not Present.");
                //Console.WriteLine("message : " + message);
                //Console.WriteLine("on page : " + driver._findElement("xpath", "//*[@ng-bind-html='objServerMaintenance.alertText']").Text);
                Assert.AreEqual(message, driver._getText("xpath", "//*[@ng-bind-html='objServerMaintenance.alertText']"), "Website Maintenance: Message not Match.");

                Results.WriteStatus(test, "Pass", "Verified, Website Maintenance Message on Page.");

            return new Login(driver, test);
        }

        ///// <summary> Verify Email and Password Field on Login Page </summary>
        ///// <param name="BlankValue">Verify Email Address & Password Field Blank</param>
        ///// <returns></returns>
        //public Login verifyEmailAndPasswordFieldOnLoginPage(bool BlankValue = false)
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");

        //    if (BlankValue)
        //    {
        //        Assert.AreEqual("", driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Email Address' Field not Blank.");
        //        Assert.AreEqual("", driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Password' Field not Blank.");

        //        Results.WriteStatus(test, "Pass", "Verified, Email Address and Password Field on Login Page.");
        //    }
        //    else
        //    {
        //        Assert.AreEqual(Email[0].ToString(), driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress"), "'Email Address' Field is Blank.");
        //        Assert.AreEqual(true, driver._getValue("id", "ctl00_ContentPlaceHolder1_EmailAddress") != "", "'Password' Field is Blank.");

        //        Results.WriteStatus(test, "Pass", "Verified, Email Address and Password Field on Login Page.");
        //    }
        //    return new Login(driver, test);
        //}

        /// <summary> Click Links on Login Page </summary>
        /// <param name="linkName">Link Name to Click</param>
        /// <returns></returns>
        public Login clickLinksOnLoginPage(string linkName)
        {
            //bool link = false;
            if (linkName.Equals("Reset Password"))
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'openResetPasssword(true);')]"), "'Reset Password' Link not Present.");
                driver._click("xpath", "//*[contains(@ng-click,'openResetPasssword(true);')]");
                //driver._waitForElement("id", "EmailAddress", 20);
                //link = true;
            }

            if (linkName.Equals("Request Access"))
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'requestAccess()')]"), "'Request Access' Link not Present.");
                //driver._clickByJavaScriptExecutor("//a[@class='request-accessLink']");
                driver._click("xpath", "//*[contains(@ng-click,'requestAccess()')]");
                driver._waitForElement("class", "row popup ng-scope", 20);
                //link = true;
            }

            if (linkName.Equals("Privacy Notice"))
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'openPrivacyPolicy()')]"), "'Privacy Notice' Link not Present.");
                //driver._click("xpath", "//*[contains(@ng-click,'openPrivacyPolicy()')]");
                driver._clickByJavaScriptExecutor("//*[contains(@ng-click,'openPrivacyPolicy()')]");
                //link = true;
            }

            if (linkName.Equals("Log in with Numerator Single Sign On"))
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'ssoLogin();')]"), "'Log in with Numerator Single Sign On' Link not Present.");
                driver._click("xpath", "//*[contains(@ng-click,'ssoLogin();')]");
                //driver._clickByJavaScriptExecutor("//*[contains(@ng-click,'openPrivacyPolicy()')]");
                //link = true;
            }
            //Assert.AreEqual(true, link);
            Results.WriteStatus(test, "Pass", "Clicked, " + linkName + " Link on Login Page .");
            return new Login(driver, test);
        }

        ///// <summary> Navigate URL to Verify Login Page, Enter Credential and Verify Home Page with Client Name & Database </summary>
        ///// <param name="clientName">Select Client Name from List</param>
        ///// <param name="database">Select Database from List</param>
        ///// <param name="columnNo">Column Number to Find Data from the Excel Sheet</param>
        ///// <returns></returns>
        //public Login loginAndVerifyHomePageWithClientAndDatabase(string clientName = "Procter & Gamble", string database = "Detail Data (PEP)", int columnNo = 0)
        //{
        //    navigateToLoginPage().verifyLoginPage();
        //    loginUsingValidEmailIdAndPassword(columnNo).clickSignInButton("verifyLogin();");

        //    //Home homePage = new Home(driver, test);

        //    //homePage.verifyHomePage();
        //    //if (clientName != "")
        //        //  homePage.selectClientOrDatabaseFromUserProfileMenu("Client Name", clientName);
        //        //if (database != "")
        //          //homePage.selectClientOrDatabaseFromUserProfileMenu("Database", database);

        //            //return new Login(driver, test);
        //}

        #region SSO

        /// <summary> To Verify SSO Login Page </summary>
        /// <param name="Watermark">Verify Watermark</param>
        /// <returns></returns>
        public Login verifySSOLoginPage(bool Watermark = true, string ScreenType = "Desktop")
        {
            string screenclass;
            if (ScreenType == "Desktop")
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-lg";
            }
            else
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-sm";
            }

            Assert.AreEqual(true, driver._waitForElement("xpath", "//*[@class='" + screenclass + "']/div/div/div/center/img"), "Numerator Logo not found on Page.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@class='textDescription-customizable ']"), "'Sign in' text not found on Page.");
            Assert.AreEqual("Sign in with your email and password", driver._findElement("xpath", "//*[@class='" + screenclass + "']//*[@class='textDescription-customizable ']").Text, "'Sign in' text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@for='signInFormUsername']"), "'Email' text not found on Page.");
            Assert.AreEqual("Email", driver._findElement("xpath", "//*[@class='" + screenclass + "']//*[@for='signInFormUsername']").Text, "'Email' text not match.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormUsername']"), "Email input textbox not found on Page.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@for='signInFormPassword']"), "'Password' text not found on Page.");
            Assert.AreEqual("Password", driver._findElement("xpath", "//*[@class='" + screenclass + "']//*[@for='signInFormPassword']").Text, "'Password' text not match.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormPassword']"), "Password input textbox not found on Page.");

            if (Watermark)
            {
                Assert.AreEqual(true, driver._getAttributeValue("id", "signInFormUsername", "placeholder").Contains("Email"), "'Email' Watermark not Present.");
                Assert.AreEqual(true, driver._getAttributeValue("id", "signInFormPassword", "placeholder").Contains("Password"), "'Password' Watermark not Present.");
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@class='redirect-customizable']"), "'Forgot your password?' link not found on Page.");
            Assert.AreEqual("Forgot your password?", driver._findElement("xpath", "//*[@class='" + screenclass + "']//*[@class='redirect-customizable']").Text, "'Forgot your password?' link text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']"), "'Sign In' button not found on Page.");
            Assert.AreEqual(true, driver._getAttributeValue("xpath", "//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']", "value").Contains("Sign in"), "'Password' Watermark not Present.");

            Results.WriteStatus(test, "Pass", "Verified, Login page.");
            return new Login(driver, test);
        }

        /// <summary> Click on SSO SignIn Button </summary>
        /// <returns></returns>
        public Login clickSSOSignInButton(bool configureUser = false, string ScreenType = "Desktop")
        {
            #region Datasheet
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
            string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");
            string email, password = "";
            #endregion

            if (configureUser)
            {
                email = Email[2].ToString();
                password = Password[2].ToString();
            }
            else
            {
                email = driver._randomString(5) + "@numerator.com";
                password = driver._randomString(6);
            }

            string screenclass;
            if (ScreenType == "Desktop")
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-lg";
            }
            else
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-sm";
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']"), "'Button not found on Page.");
            driver._clickByJavaScriptExecutor("//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']");
            Results.WriteStatus(test, "Pass", "Clicked, on Sign In Button.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormUsername']"), "Email Textarea not Present.");
            driver._type("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormUsername']", email);
            Results.WriteStatus(test, "Pass", "Entered Invalid Email address");

                Assert.IsTrue(driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormPassword']"), "Password Textarea not Present.");
                driver._type("xpath", "//*[@class='" + screenclass + "']//*[@id='signInFormPassword']", password);
                Results.WriteStatus(test, "Pass", "Entered Invalid Password.");

            driver._clickByJavaScriptExecutor("//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']");

            if (configureUser)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//*[@class='" + screenclass + "']//*[@id='loginErrorMessage']"), "Error Message not Present.");
                Assert.AreEqual("The username or password you entered is invalid", driver._getText("xpath", "//*[@class='" + screenclass + "']//*[@id='loginErrorMessage']"), "Error Message not Match.");
                Results.WriteStatus(test, "Pass", "Verified, Error Message on Login Page.");
            }
            return new Login(driver, test);
        }

        /// <summary> Verify Forgot your password? link </summary>
        /// <returns></returns>
        public Login verifyForgotYourPasswordLink(bool Watermark = true, string ScreenType = "Desktop")
        {
            string screenclass;
            if (ScreenType == "Desktop")
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-lg";
            }
            else
            {
                screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-sm";
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@class='redirect-customizable']"), "'Forgot your password?' link not found on Page.");
            Assert.AreEqual("Forgot your password?", driver._findElement("xpath", "//*[@class='" + screenclass + "']//*[@class='redirect-customizable']").Text, "'Forgot your password?' link text not match.");
            driver._clickByJavaScriptExecutor("//*[@class='" + screenclass + "']//*[@class='redirect-customizable']");

            Assert.AreEqual(true, driver._waitForElement("class", "logo-customizable"), "Numerator Logo not found on Page.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='modal-body']/div/div/h1"), "'Forgot your password?' text not found on Page.");
            Assert.AreEqual("Forgot your password?", driver._findElement("xpath", "//*[@class='modal-body']/div/div/h1").Text, "'Forgot your Password' text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='modal-body']/div/div/span"), "'Forgot your password' info text not found on Page.");
            Assert.AreEqual("Enter your Email below and we will send a message to reset your password", driver._findElement("xpath", "//*[@class='modal-body']/div/div/span").Text, "'Forgot your Password' info text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='username']"), "Email input textbox not found on Page.");

            if (Watermark)
            {
                Assert.AreEqual(true, driver._getAttributeValue("id", "username", "placeholder").Contains("Email"), "'Email' Watermark not Present.");
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@name='reset_my_password']"), "'Reset my password' button not found on Page.");
            Assert.AreEqual("Reset my password", driver._findElement("xpath", "//*[@name='reset_my_password']").Text, "'Reset my password' button text not match.");

            Results.WriteStatus(test, "Pass", "Verified, Reset my password page.");
            return new Login(driver, test);
        }

        /// <summary> Login using Valid Email </summary>
        /// <returns></returns>
        public Login loginUsingInvalidEmailinSSO(bool configureUser = false, string ScreenType = "Desktop")
        {
            #region Datasheet
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");
            string email = "";
            #endregion

            if (configureUser)
            {
                email = Email[2].ToString();
            }
            else
            {
                email = driver._randomString(5) + "@numerator.com";
            }

            //string screenclass;
            //if (ScreenType == "Desktop")
            //{
            //    screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-lg";
            //}
            //else
            //{
            //    screenclass = "modal-content background-customizable modal-content-mobile visible-md visible-sm";
            //}

            driver._clickByJavaScriptExecutor("//*[@name='reset_my_password']");
            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']"), "'Button not found on Page.");
            //driver._clickByJavaScriptExecutor("//*[@class='" + screenclass + "']//*[@name='signInSubmitButton']");
            Results.WriteStatus(test, "Pass", "Clicked, on Reset my password Button.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//*[@id='username']"), "Email Textarea not Present.");
            driver._type("xpath", "//*[@id='username']", email);
            Results.WriteStatus(test, "Pass", "Entered Invalid Email address");

            driver._clickByJavaScriptExecutor("//*[@name='reset_my_password']");

            if (configureUser==false)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//*[@id='errorMessage']"), "Message not Present.");
                Assert.AreEqual("Could not reset password for the account, please contact support or try again", driver._getText("xpath", "//*[@id='errorMessage']"), "Error Message not Match.");
                Results.WriteStatus(test, "Pass", "Verified, Error Message on Reset my password Page.");
            }
            return new Login(driver, test);
        }
        
        #endregion

        #region Reset Password

        /// <summary> To Verify Reset Password Page </summary>
        /// <returns></returns>
        public Login verifyResetPasswordPage(bool Watermark = true)
        {
            //Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "Numerator Reset Password Page not found.");

            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='headercontainer']"), "'LEARN MORE ABOUT MARKET TRACK' Label not Found.");
            //Assert.AreEqual("LEARN MORE ABOUTMARKET TRACK", driver._getText("xpath", "//td[@class='headercontainer']").Trim().Replace("\r\n", ""), "'LEARN MORE ABOUT MARKET TRACK' Label not Match.");

            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvVideo']"), "Video not Present.");

            //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "'Reset Numerator Password' Label not Found.");
            //Assert.AreEqual("Reset Numerator Password", driver._getText("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "'Reset Numerator Password' Label not Match.");

            //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_tdInstruction2"), "Second Instruction not found.");
            //Assert.AreEqual("If you have forgotten, misplaced or want to change your Numerator password,enter your email address below and click the submit button", driver._getText("id", "ctl00_ContentPlaceHolder1_tdInstruction2"), "'Second Instruction' message not Match.");

            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//td[@class='inputFieldLabel' and contains(text(), 'Email Address:')]"), "'Email Address:' Label not found.");
            //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_userEmail"), "Email Address input area not found on Page.");

            Assert.AreEqual(true, driver._waitForElement("class", "logo"), "Numerator Logo not found on Page.");

            Assert.AreEqual(true, driver._isElementPresent("id", "EmailAddress"), "Email input area not found on Page.");
            if (Watermark)
            {
                Assert.AreEqual(true, driver._getAttributeValue("id", "EmailAddress", "placeholder").Contains("Email"), "'Email' Watermark not Present.");
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'openResetPasssword(false);')]"), "'Cancel' button not found on Page.");
            Assert.AreEqual("Cancel", driver._findElement("xpath", "//*[contains(@ng-click,'openResetPasssword(false);')]").Text, "'Cancel' button text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@ng-click,'sendResetEmail();')]"), "'Reset Password' button not found on Page.");
            Assert.AreEqual("Request Password Reset", driver._findElement("xpath", "//*[contains(@ng-click,'sendResetEmail();')]").GetAttribute("value"), "'Reset Password' button text not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='tblMain']/div[1]/div/div[2]/div[3]"), "Reset Instruction not found.");
            Assert.AreEqual("Enter your email address above and click on Request Password Reset button to get password reset link.", driver._getText("xpath", "//*[@id='tblMain']/div[1]/div/div[2]/div[3]/span"), "'Reset Instruction' message not Match.");

            Results.WriteStatus(test, "Pass", "Verified, Reset Password Page.");
            return new Login(driver, test);
        }

        ///// <summary> Enter Email Address on Reset Password Page </summary>
        ///// <param name="validEmail">Valid Email Address</param>
        ///// <returns></returns>
        //public Login enterEmailAddressOnResetPasswordPage(bool validEmail = false)
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");

        //    Assert.IsTrue(driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_userEmail"), "Email Address textarea not Pesent.");

        //    if (validEmail)
        //    {
        //        driver._type("id", "ctl00_ContentPlaceHolder1_userEmail", Email[0].ToString());
        //        Results.WriteStatus(test, "Pass", "Entered, Valid Email Address on Reset Password Page.");
        //    }
        //    else
        //    {
        //        driver._type("id", "ctl00_ContentPlaceHolder1_userEmail", driver._randomString(5) + "@test.com");
        //        Results.WriteStatus(test, "Pass", "Entered, Invalid Email Address on Reset Password Page.");
        //    }
        //    return new Login(driver, test);
        //}

        ///// <summary> Click Button on Reset Password Page </summary>
        ///// <param name="buttonName">Button Name to Click</param>
        ///// <returns></returns>
        //public Login clickButtonOnResetPasswordPage(string buttonName)
        //{
        //    bool button = false;

        //    if (buttonName.Equals("Submit"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_submitEmail"), "'Submit' Button not Present.");
        //        driver._click("id", "ctl00_ContentPlaceHolder1_submitEmail");
        //        button = true;
        //    }

        //    if (buttonName.Equals("Back to Login"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "ButtonLoginBack"), "'Back to Login' Button not Present.");
        //        driver._click("id", "ButtonLoginBack");
        //        button = true;
        //    }

        //    Assert.AreEqual(true, button);
        //    Results.WriteStatus(test, "Pass", "Clicked, " + buttonName + " Button on Reset Password Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary> Verify Warning Message and Login link on Reset Password Page </summary>
        ///// <returns></returns>
        //public Login verifyWarningMessageAndLoginLinkOnResetPasswordPage()
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "ctl00_ContentPlaceHolder1_messageWorning"), "Warning Message not Present.");
        //    Assert.AreEqual("Password Reset link has been sent to your email address.", driver._getText("xpath", "//*[@id='ctl00_ContentPlaceHolder1_messageWorning']/table/tbody/tr[1]/td/span"), "'Password Reset link has been sent to your email address.' Warning message not found.");
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[@class='loginheaderlink' and text() = 'Login']"), "'Login' Link not Present.");
        //    Results.WriteStatus(test, "Pass", "Verified, Watning Message & Login Link on Reset Password Page.");

        //    return new Login(driver, test);
        //}

        #endregion

        #region Access Request Popup Window

        /// <summary> To Verify Access Request Popup Window </summary>
        /// <returns></returns>
        public Login verifyAccessRequestPopupWindow()
        {
            driver._waitForElement("class", "row popup ng-scope", 20);
            Assert.IsTrue(driver._waitForElement("xpath", "//*[@ng-controller='requestAccessController']", 20), "Access Request Popup Window not found.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'-aria-labelledby')]"), "'Access Request' Title not Found.");
            Assert.AreEqual("Access Request", driver._getText("xpath", "//*[contains(@id,'-aria-labelledby')]").Trim().Replace("\r\n", ""), "'Access Request' Title not Match.");

            //Assert.AreEqual(true, driver._isElementPresent("id", "btnClose"), "'Close' Icon not Present on Winodw.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-click='closeThisDialog();']"), "Cancel Button not found on Page.");
            Assert.AreEqual("×", driver._findElement("xpath", "//*[@ng-click='closeThisDialog();']").Text, "'Cancel' Button Label not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'-aria-describedby')]"), "'Access Request' information not Present.");
            //Console.WriteLine(driver._findElement("xpath", "//*[contains(@id,'-aria-describedby')]").Text);

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//a[@class='aLinkColor ng-binding']"), "'Support Email' address not Present.");
            Assert.AreEqual("promointelsupport@numerator.com", driver._getText("xpath", "//a[@class='aLinkColor ng-binding']").Trim().Replace("\r,\n", ""), "'Support Email' address not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[1]/td[1]"), "'Company or Organization' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.company']"), "'Company or Organization' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[2]/td[1]"), "'Country' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.country']"), "'Country' DDL not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[3]/td[1]"), "'First Name' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.firstName']"), "'First Name' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[4]/td[1]"), "'Last Name' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.lastName']"), "'Last Name' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[5]/td[1]"), "'Email Address' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.emailId']"), "'Email Address' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[6]/td[1]"), "'Phone' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.phone']"), "'Phone' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[7]/td[1]"), "'Title' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.title']"), "'Title' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[8]/td[1]"), "'Additional Information/Comments' label not found.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.comments']"), "'Additional Information/Comments' textarea not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'ngdialog')]/div[2]/div/div/div/div[2]/div/div/div/table[2]/tbody/tr[9]/td[1]"), "'Required Information' label not found.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-click='saveRequestAccess();']"), "Save Button not found on Page.");
            Assert.AreEqual("Save", driver._getText("xpath", "//*[@ng-click='saveRequestAccess();']"), "'Save' Button Label not match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='btn btn-default' and @ng-click='closeThisDialog();']"), "Cancel Button not found on Page.");
            Assert.AreEqual("Cancel", driver._findElement("xpath", "//*[@class='btn btn-default' and @ng-click='closeThisDialog();']").Text, "'Cancel' Button Label not match.");

            Results.WriteStatus(test, "Pass", "Verified, Access Request Popup Window.");
            return new Login(driver, test);
        }

        /// <summary> To Verify Alert Access Request Popup </summary>
        /// <returns></returns>
        public Login verifyAlertAccessRequestPopup()
        {
            driver._waitForElement("id", "myModal", 20);
            Assert.IsTrue(driver._waitForElement("xpath", "//*[@class='modal-content']", 20), "Alert Access Request Popup not found.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[contains(@id,'-aria-labelledby')]"), "'Alert' Title not Found.");
            if (driver._getText("xpath", "//*[@class='modal-title ng-binding' and contains(@id,'-aria-labelledby')]") == "Alert")
            {
                Assert.AreEqual("Alert", driver._getText("xpath", "//*[@class='modal-title ng-binding' and contains(@id,'-aria-labelledby')]"), "'Alert' Title not Match.");
            }
            else
            {
                Assert.AreEqual("Oops, we hit a snag", driver._getText("xpath", "//*[@class='modal-title ng-binding' and contains(@id,'-aria-labelledby')]"), "'Alert' Title not Match.");
            }

            //Assert.AreEqual(true, driver._isElementPresent("id", "btnClose"), "'Close' Icon not Present on Winodw.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-click='closeThisDialog()']"), "Cancel Button not found on Page.");
            Assert.AreEqual("×", driver._findElement("xpath", "//*[@ng-click='closeThisDialog()']").Text, "'Cancel' Button Label not match.");

            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='myModal']/div/div/div[2]/div/div/div/p/span[1]"), "'Alert' information not Present.");
            //Console.WriteLine(driver._findElement("xpath", "//*[@id='myModal']/div/div/div[2]/div/div/div/p/span[1]").Text);

            //Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='myModal']/div/div/div[2]/div/div/div/p/span[2]"), "'Alert' fileds not Present.");
            //Assert.AreEqual("promohelp.numerator.com", driver._getText("xpath", "//a[@class='aLinkColor ng-binding']").Trim().Replace("\r,\n", ""), "'Support Email' address not match.");
            //Console.WriteLine(driver._findElement("xpath", "//*[@id='myModal']/div/div/div[2]/div/div/div/p/span[2]").Text);

            //string AlertInfo = "Alert " + driver._findElement("xpath", "//*[@id='myModal']/div/div/div[2]/div/div/div/p/span[2]").Text;
            //Assert.AreEqual(FieldID,AlertInfo, "'Alert' info not Match.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='btn btn-default ng-binding' and @ng-click='closeThisDialog()']"), "Ok Button not found on Page.");
            Assert.AreEqual("Ok", driver._findElement("xpath", "//*[@class='btn btn-default ng-binding' and @ng-click='closeThisDialog()']").Text, "'Ok' Button Label not match.");
            driver._clickByJavaScriptExecutor("//*[@class='btn btn-default ng-binding' and @ng-click='closeThisDialog()']");
            Thread.Sleep(1000);

            Results.WriteStatus(test, "Pass", "Verified, Alert Access Request Popup.");
            return new Login(driver, test);
        }

        ///// <summary> Enter Data on Access Request Window </summary>
        ///// <returns></returns>
        //public String enterDataOnAccessRequestWindow()
        //{
        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.company']", "AutoComp" + driver._randomString(4, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Company Name on Access Request Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess.country']"), "Country List not Present on Screen.");
        //    IList<IWebElement> countryCollection = driver._findElements("xpath", "//*[@ng-model='objRequestAccess.country']/option");
        //    Random select = new Random();
        //    int x = select.Next(0, countryCollection.Count);
        //    countryCollection[x].Click();
        //    Results.WriteStatus(test, "Pass", "Selected, Country on Access Request Window.");

        //    string firstName = "AutoF" + driver._randomString(4, true);
        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.firstName']", firstName);
        //    Results.WriteStatus(test, "Pass", "Entered, First Name on Access Request Window.");

        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.lastName']", "AutoL" + driver._randomString(4, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Last Name on Access Request Window.");

        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.emailId']", "Test" + driver._randomString(4, true) + "@test.com");
        //    Results.WriteStatus(test, "Pass", "Entered, Email Address on Access Request Window.");

        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.phone']", driver._randomString(8, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Phone on Access Request Window.");

        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.title']", "AutoT" + driver._randomString(5, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Title on Access Request Window.");

        //    driver._type("xpath", "//*[@ng-model='objRequestAccess.comments']", "Comment" + driver._randomString(6, true));
        //    Results.WriteStatus(test, "Pass", "Entered, Comments on Access Request Window.");

        //    return firstName;
        //}

        // <summary> Insert Single filed Data in Access Request Window </summary>
        // <returns></returns>
        //public String DataOnAccessRequestPopup(string FieldID)
        public Login DataOnAccessRequestPopup(string FieldID, string Name = "")
        {
            if (FieldID == "country")
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@ng-model='objRequestAccess." + FieldID + "']"), FieldID + " field not Present on Screen.");
                IList<IWebElement> countryCollection = driver._findElements("xpath", "//*[@ng-model='objRequestAccess." + FieldID + "']/option");
                Random select = new Random();
                int x = select.Next(0, countryCollection.Count);
                countryCollection[x].Click();
                Results.WriteStatus(test, "Pass", "Selected, " + FieldID + " on Access Request Window.");
            }
            else
            {
                driver._type("xpath", "//*[@ng-model='objRequestAccess." + FieldID + "']", driver._randomString(4, true) + Name);
                Results.WriteStatus(test, "Pass", "Entered, " + FieldID + " on Access Request Window.");
            }

            clickOnButton("saveRequestAccess();");
            verifyAlertAccessRequestPopup();

            return new Login(driver, test);
        }

        ///// <summary> Click Button on Access Request Popup Window </summary>
        ///// <param name="buttonName">Button Name to Click</param>
        ///// <returns></returns>
        //public Login clickButtonOnAccessRequestPopupWindow(string buttonName)
        //{
        //    bool button = false;

        //    if (buttonName.Equals("Save"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "btnSave"), "'Save' Button not Present.");
        //        //driver._click("id", "btnSave");
        //        button = true;
        //    }

        //    if (buttonName.Equals("Cancel"))
        //    {
        //        Assert.AreEqual(true, driver._isElementPresent("id", "btnCancel"), "'Cancel' Button not Present.");
        //        driver._click("id", "btnCancel");
        //        button = true;
        //    }

        //    Assert.AreEqual(true, button);
        //    Results.WriteStatus(test, "Pass", "Clicked, " + buttonName + " Button on Access Request Popup Window.");
        //    return new Login(driver, test);
        //}

        #endregion

        //#region Privacy Policy

        ///// <summary> Verify Image Logo on Login Page </summary>
        ///// <param name="logo">Logo Title to Verify</param>
        ///// <returns></returns>
        //public Login verifyImageLogoOnLoginPage(string logo)
        //{
        //    Assert.IsTrue(driver._isElementPresent("id", "ctl00_imgMtLogo"), "'Logo' not Present on Login Page.");
        //    Assert.AreEqual(true, driver._getAttributeValue("id", "ctl00_imgMtLogo", "title").Contains(logo), "'" + logo + "' Logo not Match.");
        //    Results.WriteStatus(test, "Pass", "Verified, '" + logo + "' Image Logo on Login Page.");

        //    return new Login(driver, test);
        //}

        /// <summary> Get Image Logo Location on Home Page </summary>
        /// <returns>Image Title</returns>
        public String getImageLogoLocationOnLoginPage()
        {
            Assert.AreEqual(true, driver._isElementPresent("class", "site-logo"), "'Image Logo' not Present on Login Page.");
            string location = driver._getAttributeValue("class", "site-logo", "src");
            Results.WriteStatus(test, "Pass", "Get '" + location + "' Image Logo Location on Login Page.");

            return location;
        }

        /// <summary> Click 'View Privacy Link' on Login Page </summary>
        /// <returns></returns>
        public Login clickViewPrivacyLinkOnLoginPage()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not found.");
            Assert.AreEqual("View Privacy Policy", driver._getText("xpath", "//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a"), "'View Privacy Policy' Label at Bottom not match.");
            driver._clickByJavaScriptExecutor("//*[@id='tblFooterLinks']/tbody/tr/td/table/tbody/tr/td[4]/a");
            Results.WriteStatus(test, "Pass", "Clicked, 'View Privacy Policy' Link on Login Page.");

            return new Login(driver, test);
        }

        ///// <summary> Verify Privacy Policy Popup Window on Page </summary>
        ///// <param name="logoLocation">Verify Logo</param>
        ///// <param name="supportMailAddress">Verify Support Mail Address</param>
        ///// <returns></returns>
        //public Login verifyPrivacyPolicyPopupWindowOnPage(string logoLocation = "", string supportMailAddress = "")
        //{
        //    Assert.IsTrue(driver._waitForElement("id", "popupDiv1"), "'Privacy Policy' Popup Window not Present.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "dvTitle1"), "'Privacy Policy' Label not Present on Window.");
        //    Assert.AreEqual("Privacy Policy", driver._getText("id", "dvTitle1"), "'Privacy Policy' Label not match on Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "privacypolicy"), "'Privacy Policy' Content not Present.");
        //    Assert.AreEqual(true, driver._isElementPresent("id", "imgPrivacyLogo"), "'Numerator' Logo not Present on Window.");

        //    if (logoLocation != "")
        //        Assert.AreEqual(logoLocation, driver._getAttributeValue("id", "imgPrivacyLogo", "src"), "'Logo' not verified on Popup Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "btn1no0"), "'Close' Button not Present on Window.");
        //    Assert.AreEqual("Close", driver._getValue("id", "btn1no0"), "'Close' Button Label not match on Window.");

        //    Assert.AreEqual(true, driver._isElementPresent("id", "closepopup1"), "'Close' Icon not Present on Window.");

        //    if (supportMailAddress != "")
        //    {
        //        IWebElement elements = driver._findElement("id", "privacypolicy");
        //        IList<IWebElement> content = elements.FindElements(By.TagName("a"));

        //        for (int i = 0; i < content.Count; i++)
        //        {
        //            Assert.AreEqual(supportMailAddress, content[i].Text, "'" + content[i].Text + "' Support Email not Match.");
        //        }
        //    }

        //    Results.WriteStatus(test, "Pass", "Verified, Privacy Policy Popup Window on Page.");
        //    return new Login(driver, test);
        //}

        ///// <summary> Click Close Button On Privacy Policy Window </summary>
        ///// <returns></returns>
        //public Login clickCloseButtonOnPrivacyPolicyWindow()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("id", "btn1no0"), "'Close' Button not Present on Window.");
        //    driver._click("id", "btn1no0");
        //    Thread.Sleep(1000);
        //    Results.WriteStatus(test, "Pass", "Clicked, Close Button on Privacy Policy Window.");

        //    return new Login(driver, test);
        //}

        //#endregion

        #region GMail Methods

        ///// <summary> Verify Outlook Login Screen to Enter Credential and Click SignIn Button </summary>
        ///// <returns></returns>
        //public Login verifyOutlookLoginScreenToEnterCredentialAndClickSignInButton()
        //{
        //    string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";

        //    string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Gmail");
        //    string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Outlook");

        //    if (driver._isElementPresent("id", "cred_userid_inputtext"))
        //    {

        //        Assert.IsTrue(driver._waitForElement("id", "cred_userid_inputtext", 20), "Email Address Textarea not Present.");
        //        Assert.AreEqual(true, driver._isElementPresent("id", "cred_password_inputtext"), "Password Textarea not Present.");
        //        Assert.AreEqual(true, driver._isElementPresent("id", "cred_sign_in_button"), "SignIn Button not Present.");
        //        Results.WriteStatus(test, "Pass", "Verified, Outlook Login Screen.");

        //        driver._type("id", "cred_userid_inputtext", Email[0].ToString());
        //        Thread.Sleep(1000);
        //        driver._type("id", "cred_password_inputtext", Password[0].ToString());
        //        Thread.Sleep(3000);

        //        driver._click("id", "cred_sign_in_button");
        //        Thread.Sleep(5000);
        //        Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
        //    }
        //    else
        //    {
        //        Assert.IsTrue(driver._waitForElement("xpath", "//input[@type='email' and @name='loginfmt']", 20), "Email Address Textarea not Present.");
        //        driver._type("xpath", "//input[@type='email' and @name='loginfmt']", Email[0].ToString());
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Next']"), "Next Button not Present.");
        //        driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Next']");
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@name='passwd' and @type='password']"), "Password Textarea not Present.");
        //        driver._type("xpath", "//input[@name='passwd' and @type='password']", Password[0].ToString());
        //        Thread.Sleep(1000);

        //        Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Sign in']"), "Sign in Button not Present.");
        //        driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Sign in']");
        //        Thread.Sleep(1000);

        //        if (driver._isElementPresent("xpath", "//input[@type='button' and @value='No']"))
        //            driver._clickByJavaScriptExecutor("//input[@type='button' and @value='No']");
        //        Thread.Sleep(1000);
        //    }

        //    Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
        //    return new Login(driver, test);
        //}

        /// <summary> Verify Gmail Login Screen to Enter Credential and Click SignIn Button </summary>
        /// <returns></returns>
        public Login verifyGmailLoginScreenToEnterCredentialAndClickNextButton()
        {
            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";

            string[] Email = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Email", "Valid");

            Assert.IsTrue(driver._waitForElement("id", "identifierId", 20), "Email Textarea not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@class='RveJvd snByac']"), "Next Button not Present.");
            Results.WriteStatus(test, "Pass", "Verified, Gmail Login Screen.");

            driver._type("id", "identifierId", Email[2].ToString());
            Thread.Sleep(5000);
            //driver._click("class", "RveJvd snByac");
            driver._clickByJavaScriptExecutor("//*[@class='RveJvd snByac']");
            Thread.Sleep(10000);
            driver._clickByJavaScriptExecutor("//*[@class='RveJvd snByac']");
            Thread.Sleep(10000);
            Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Process.");

            //if (driver._isElementPresent("id", "cred_userid_inputtext"))
            //{

            //    Assert.IsTrue(driver._waitForElement("id", "cred_userid_inputtext", 20), "Email Address Textarea not Present.");
            //    Assert.AreEqual(true, driver._isElementPresent("id", "cred_password_inputtext"), "Password Textarea not Present.");
            //    Assert.AreEqual(true, driver._isElementPresent("id", "cred_sign_in_button"), "SignIn Button not Present.");
            //    Results.WriteStatus(test, "Pass", "Verified, Outlook Login Screen.");

            //    driver._type("id", "cred_userid_inputtext", Email[0].ToString());
            //    Thread.Sleep(1000);
            //    driver._type("id", "cred_password_inputtext", Password[0].ToString());
            //    Thread.Sleep(3000);

            //    driver._click("id", "cred_sign_in_button");
            //    Thread.Sleep(5000);
            //    Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
            //}
            //else
            //{
            //    Assert.IsTrue(driver._waitForElement("xpath", "//input[@type='email' and @name='loginfmt']", 20), "Email Address Textarea not Present.");
            //    driver._type("xpath", "//input[@type='email' and @name='loginfmt']", Email[0].ToString());
            //    Thread.Sleep(1000);

            //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Next']"), "Next Button not Present.");
            //    driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Next']");
            //    Thread.Sleep(1000);

            //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@name='passwd' and @type='password']"), "Password Textarea not Present.");
            //    driver._type("xpath", "//input[@name='passwd' and @type='password']", Password[0].ToString());
            //    Thread.Sleep(1000);

            //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@type='submit' and @value='Sign in']"), "Sign in Button not Present.");
            //    driver._clickByJavaScriptExecutor("//input[@type='submit' and @value='Sign in']");
            //    Thread.Sleep(1000);

            //    if (driver._isElementPresent("xpath", "//input[@type='button' and @value='No']"))
            //        driver._clickByJavaScriptExecutor("//input[@type='button' and @value='No']");
            //    Thread.Sleep(1000);
            //}

            Results.WriteStatus(test, "Pass", "Entered, Credential and Clicked SignIn Button.");
            return new Login(driver, test);
        }

        /// <summary> Verify Gmail Home Page </summary>
        /// <returns></returns>
        public Login verifyGmailHomePage()
        {
            Thread.Sleep(10000);
            Assert.IsTrue(driver._waitForElement("xpath", "//*[@class='J-Ke n0' and @title='Inbox']"), "Inbox Folder not Present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//*[@class='F cf zt' and @role='grid']"), "Emails List not Present.");
            Thread.Sleep(10000);
            Results.WriteStatus(test, "Pass", "Verified, Gmail Home Page.");
            return new Login(driver, test);
        }

        /// <summary> Select Reset Password Mail to Open Reset Link </summary>
        /// <returns></returns>
        public Login selectResetPasswordMailToOpenResetLink()
        {
            Thread.Sleep(10000);
            Assert.AreEqual(true, driver._isElementPresent("id", ":27"), "Mails Subject not Present.");
            IList<IWebElement> mailSubjects = driver._findElements("xpath", "//span[contains(@data-thread-id,'#thread-f:')]");
            bool avail = false;

            for (int m = 0; m < mailSubjects.Count; m++)
            {
                //Console.WriteLine("/n " + mailSubjects[m].Text);
                if (mailSubjects[m].Text.Contains("Your Numerator Promotions Intel Password Reset Request"))
                {
                    mailSubjects[m].Click();
                    avail = true;
                    break;
                }
            }
            Assert.AreEqual(true, avail, "'Your Numerator Promotions Intel Password Reset Request' Mail not Present.");
            Results.WriteStatus(test, "Pass", "Selected, Reset Password Email.");

            Assert.IsTrue(driver._waitForElement("xpath", "//span[@style='color:#336e74']"), "Message Content not Present.");
            IWebElement body = driver._findElement("xpath", "//span[@style='color:#336e74']");
            IList<IWebElement> content = body.FindElements(By.TagName("a"));
            bool resetLink = false;

            for (int i = 0; i < content.Count(); i++)
            {
                if (content[i].GetAttribute("href").Contains("Login.aspx"))
                {
                    content[i].Click();
                    resetLink = true;
                    Thread.Sleep(5000);
                    break;
                }
            }
            Assert.AreEqual(true, resetLink, "'Reset Password' Link not Present on Content.");
            Results.WriteStatus(test, "Pass", "Clicked, Reset Password Link from Email.");
            return new Login(driver, test);
        }

        /// <summary> Verify Reset Screen to Enter Password </summary>
        /// <returns></returns>
        public Login verifyResetScreenToEnterPassword()
        {
            Assert.AreEqual(true, driver._isElementPresent("class", "logo"), "Image Logo Not Present.");
            //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_tdInstruction"), "'Reset Numerator Password' Instruction not Present.");
            //Assert.AreEqual(true, driver._isElementPresent("id", "ctl00_ContentPlaceHolder1_Text1"), "Email Address not Present.");
            Assert.AreEqual(true, driver._isElementPresent("id", "NewPassword"), "New Password Textarea not Present.");
            Assert.AreEqual(true, driver._isElementPresent("id", "ConfirmNewPassword"), "Confirm New Password Textarea not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//button[@ng-click='setNewPassword();']"), "Set New Password Button not Present.");
            Results.WriteStatus(test, "Pass", "Verified, Reset Password Screen.");

            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Login.xlsx";
            string[] Password = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "Password", "Valid");

            driver._type("id", "NewPassword", Password[0].ToString());
            Thread.Sleep(1000);
            driver._type("id", "ConfirmNewPassword", Password[0].ToString());
            Thread.Sleep(1000);
            //clickSignInButton("setNewPassword();");
            driver._click("xpath", "//button[@ng-click='setNewPassword();']");
            Thread.Sleep(5000);
            Results.WriteStatus(test, "Pass", "Entered, Password and Clicked Reset Password Button.");
            return new Login(driver, test);
        }

        /// <summary> Verify Password Updated Label and Click Login Link </summary>
        /// <returns></returns>
        public Login verifyPasswordUpdatedAndClickLoginLink()
        {
            Assert.AreEqual(true, driver._isElementPresent("class", "logo"), "Image Logo Not Present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//span[@ng-bind-html='objPasswordStatus.message']", 10), "'Password has been updated successfully.' Label not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@value='Log In']"), "'Login' button not Present.");
            driver._clickByJavaScriptExecutor("//*[@ng-click='openResetPasssword(false);']");
            Thread.Sleep(3000);
            Results.WriteStatus(test, "Pass", "Verified, Password Updated Label and Clicked Login Link.");
            return new Login(driver, test);
        }

        ///// <summary>
        ///// select Summary Report Requested Data Mail To Download Report
        ///// </summary>
        ///// <returns></returns>
        //public Login verifyGmailLoginPageAndEnterCredential()
        //{
        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']"), "Mails Subject not Present.");
        //    Thread.Sleep(8000);
        //    IList<IWebElement> mailSubjects = driver._findElements("xpath", "//span[@class='lvHighlightAllClass lvHighlightSubjectClass']");
        //    bool avail = false;

        //    for (int m = 0; m < mailSubjects.Count; m++)
        //    {
        //        if (mailSubjects[m].Text.Contains(mailTitle))
        //        {
        //            mailSubjects[m].Click();
        //            avail = true;
        //            break;
        //        }
        //    }
        //    Assert.AreEqual(true, avail, "'" + mailTitle + "' Mail not Present.");
        //    Results.WriteStatus(test, "Pass", "Selected, '" + mailTitle + "' Email.");

        //    Assert.IsTrue(driver._waitForElement("xpath", "//span[@class='_fc_4 o365buttonLabel' and contains(@id,'_ariaId_') and text() = 'Download']"), "Download Link not Present.");
        //    driver._clickByJavaScriptExecutor("//span[@class='_fc_4 o365buttonLabel' and contains(@id,'_ariaId_') and text() = 'Download']");
        //    Thread.Sleep(8000);
        //    Results.WriteStatus(test, "Pass", "Clicked, Download File Link on Email.");
        //    return new Login(driver, test);
        //}
        #endregion

        public Login loginAndVerifyHomePageWithClient(string clientName = "Procter & Gamble", int columnNo = 0)
        {
            navigateToLoginPage().VerifyLoginPage();
            loginUsingValidEmailIdAndPassword(columnNo).clickSignInButton();

            Home homePage = new Home(driver, test);

            homePage.VerifyHomePage();
            homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
            if (clientName != "")
                homePage.VerifyClientAndChangeIfItDoesNotMatch(clientName);

            return new Login(driver, test);
        }

        #endregion
    }
}
