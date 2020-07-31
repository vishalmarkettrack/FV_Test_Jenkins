using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FeatureVision8.FeatureVision8
{
    [TestFixture]
    [Parallelizable(ParallelScope.Fixtures)]
    public class TestSuite001_Login : Base
    {

        #region Private Variables
        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        #endregion

        #region Public Fixture Methods
        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite001_Login).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite001_Login).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);

            return driver;
        }

        [TearDown]
        public void TestFixtureTearDown()
        {
            extent.Flush();
            driver.Quit();
        }
        #endregion

        #region Test Methods
        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC001_VerifyLoginPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC001-Verify Login Page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifymessagewhenuserInsertInvalidEmailOrPasswordAndClickLogInbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify message when user Insert invalid Email or Password and click Log In button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickOnButton("verifyLogin();");
                loginPage.verifyErrorMessageOnLoginPage("Your e-mail or password were incorrect.");

                loginPage.loginUsingInvalidEmailAddressOrPassword(true).clickOnButton("verifyLogin();");
                loginPage.verifyErrorMessageOnLoginPage("Your e-mail or password were incorrect.");

                loginPage.loginUsingInvalidEmailAddressOrPassword(false).clickOnButton("verifyLogin();");
                loginPage.verifyErrorMessageOnLoginPage("Your e-mail or password were incorrect.");

                loginPage.loginUsingInvalidEmailAddressAndPassword(true).clickOnButton("verifyLogin();");
                loginPage.verifyErrorMessageOnLoginPage("Your e-mail or password were incorrect.");

                loginPage.loginUsingInvalidEmailAddressAndPassword(false, "").clickOnButton("verifyLogin();");
                loginPage.verifyErrorMessageOnLoginPage("Invalid email format.");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyPagewhenuserInsertValidEmailAndPasswordAndClickLogInbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Page when user Insert valid Email and Password and click Log In button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailAddressPassword(true).clickOnButton("verifyLogin();");
                Thread.Sleep(20000);
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyResetPasswordLink(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Reset Password link.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password");
                loginPage.verifyResetPasswordPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifymessagewhenuserInsertInvalidEmailAndClickRequestPasswordResetbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify message when user Insert invalid Email and click Request Password Reset button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password").verifyResetPasswordPage();

                loginPage.clickOnButton("sendResetEmail();");
                loginPage.verifyErrorMessageOnLoginPage("Please enter your e-mail address to reset your password.");

                loginPage.loginUsingInvalidEmailAddressAndPassword(false, "").clickOnButton("sendResetEmail();");
                loginPage.verifyErrorMessageOnLoginPage("Invalid email format.");

                loginPage.loginUsingInvalidEmailAddressOrPassword(true).clickOnButton("sendResetEmail();");
                loginPage.verifyErrorMessageOnLoginPage("Your e-mail was not recognized by our system. Is it possible you entered it incorrectly?");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifymessagewhenuserInsertValidEmailAndClickRequestPasswordResetbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify message when user Insert valid Email and click Request Password Reset button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password").verifyResetPasswordPage();

                loginPage.loginUsingValidEmailAddress(true).clickOnButton("sendResetEmail();");
                loginPage.verifySuccessMessageOnPage("Success!");
                Thread.Sleep(3000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifywhenuserClickCancelbuttonOnResetPasswordPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify when user click Cancel button on Reset Password page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password").verifyResetPasswordPage();

                loginPage.clickOnButton("openResetPasssword(false);");
                Thread.Sleep(3000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifymessagewhenuserInsertInvalidPasswordAndClickSetNewPasswordbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify message when user Insert invalid Password and click Set New Password button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password").verifyResetPasswordPage();
                loginPage.loginUsingValidEmailAddress(true).clickOnButton("sendResetEmail();");

                driver.Navigate().GoToUrl("https://mail.google.com");
                loginPage.verifyGmailLoginScreenToEnterCredentialAndClickNextButton();
                //loginPage.verifyGmailHomePage();
                loginPage.selectResetPasswordMailToOpenResetLink();

                driver.SwitchTo().Window(driver.WindowHandles.Last());
                loginPage.verifyResetScreenToEnterPassword();
                loginPage.verifyPasswordUpdatedAndClickLoginLink();

                //loginPage.loginUsingValidEmailAddressPassword(true).clickSignInButton("verifyLogin();");
                Thread.Sleep(20000);
                //homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifymessagewhenuserInsertValidPasswordAndClickSetNewPasswordbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify message when user Insert valid Password and click Set New Password button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Reset Password").verifyResetPasswordPage();
                loginPage.loginUsingValidEmailAddress(true).clickOnButton("sendResetEmail();");

                driver.Navigate().GoToUrl("https://mail.google.com");
                loginPage.verifyGmailLoginScreenToEnterCredentialAndClickNextButton();
                //loginPage.verifyGmailHomePage();
                loginPage.selectResetPasswordMailToOpenResetLink();

                driver.SwitchTo().Window(driver.WindowHandles.Last());
                loginPage.verifyResetScreenToEnterPassword();
                loginPage.verifyPasswordUpdatedAndClickLoginLink();

                loginPage.loginUsingValidEmailAddressPassword(true).clickOnButton("verifyLogin();");
                Thread.Sleep(20000);
                //homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyRequestAccessLink(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Request Access link.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Request Access");
                loginPage.verifyAccessRequestPopupWindow();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyAlertwhenusernotInsertValueAndClickSavebutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Alert when user not Insert value and click Save button In RequestAccessPopup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Request Access").verifyAccessRequestPopupWindow();

                loginPage.clickOnButton("saveRequestAccess();").verifyAlertAccessRequestPopup();
                //loginPage.clickSignInButton("closeThisDialog()");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyAlertwhenuserInsertValueAndClickSavebutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Alert when user Insert value and click Save button In RequestAccessPopup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Request Access").verifyAccessRequestPopupWindow();

                //loginPage.enterDataOnAccessRequestWindow();

                loginPage.DataOnAccessRequestPopup("company");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("country");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("firstName");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("lastName");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("emailId");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("phone");
                Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("title");
                Thread.Sleep(1000);

                //loginPage.DataOnAccessRequestPopup("comments");
                //Thread.Sleep(1000);

                loginPage.DataOnAccessRequestPopup("emailId", "@numerator.com");
                Thread.Sleep(1000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifywhenuserclickCancelbuttonOnAccessRequestPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify when user click Cancel button on Access Request popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Request Access").verifyAccessRequestPopupWindow();

                loginPage.clickOnButton("closeThisDialog();");
                Thread.Sleep(3000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyPrivacyNoticeLink(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify Privacy Notice link.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Privacy Notice");
                //loginPage.verifyNavigateURLOnScreen("https://www.numerator.com/privacy-notice");
                //Thread.Sleep(2000);
                //string imageLocation = loginPage.getImageLogoLocationOnLoginPage();
                //loginPage.clickViewPrivacyLinkOnLoginPage();
                //loginPage.verifyPrivacyPolicyPopupWindowOnPage(imageLocation, "help.markettrack.com");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyWebsiteMaintenancemessage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify Website Maintenance message.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.verifyWebsiteMaintenanceMessageOnPage("Website Maintenance: Numerator Promotions Intel website will be offline on 02/11/2020 between 06:30 AM and 11:30 AM GMT for scheduled maintenance. Thank you for your patience while we upgrade the site.");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyLoginwithNumeratorSingleSignOn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify Log in with Numerator Single Sign On button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On");
                loginPage.verifySSOLoginPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifymessagewhenuserInsertInvalidEmailOrPasswordAndClickSignInbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify message when user Insert invalid Email or Password and click Sign In button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On").verifySSOLoginPage();
                loginPage.clickSSOSignInButton();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifyPagewhenuserInsertValidEmailOrPasswordAndClickSignInbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify Page when user Insert valid Email or Password and click Sign In button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On").verifySSOLoginPage();
                loginPage.clickSSOSignInButton(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_VerifyForgotYourPasswordLink(String Bname)
        {
            TestFixtureSetUp(Bname, "TC019-Verify Forgot your password? link.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On").verifySSOLoginPage();
                loginPage.verifyForgotYourPasswordLink();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC019");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifymessagewhenuserInsertInvalidEmailAndClickResetmypasswordbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify message when user Insert invalid Email and click Reset my password button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On").verifySSOLoginPage();
                loginPage.verifyForgotYourPasswordLink();
                loginPage.loginUsingInvalidEmailinSSO();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifyPagewhenuserInsertValidEmailAndClickResetmypasswordbutton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify Page when user Insert valid Email and click Reset my password button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.clickLinksOnLoginPage("Log in with Numerator Single Sign On").verifySSOLoginPage();
                loginPage.verifyForgotYourPasswordLink();
                loginPage.loginUsingInvalidEmailinSSO(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite001_Login_TC021");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
