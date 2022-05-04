using System;
using System.IO;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Download
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Customer number:");
            string customerNumber = Console.ReadLine();
            string customerEmail = "andrevo@gmail.com";

            Console.WriteLine("Password:");
            string password = Console.ReadLine();

            IWebDriver driver = new ChromeDriver(Directory.GetCurrentDirectory());

            //DownloadManoBustasInvoice(driver, customerEmail, customerNumber, password);
            DownloadVstInvoice(driver, customerNumber, password);
            DownloadVvInvoice(driver, customerNumber, password);

            //Šitas turi būti pabaigoje, nes jis aktualus visiems action'ams, kadangi reikia sulaukti kol jie baigs download'int
            Thread.Sleep(5000);
            
            driver.Quit();
        }

        private static void DownloadManoBustasInvoice(IWebDriver driver, string customerEmail, string customerNumber, string password)
        {
            driver.Navigate().GoToUrl("https://www.ebustas.lt");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            IWebElement usernameElement = driver.FindElement(By.Id("username"));
            usernameElement.SendKeys(customerEmail);

            IWebElement passwordElement = driver.FindElement(By.Id("password"));
            passwordElement.SendKeys(password);

            IWebElement cookiesElement = driver.FindElement(By.Id("CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll"));
            cookiesElement.Click();

            driver.SwitchTo().Frame(3);
            Thread.Sleep(3000);
            IWebElement chatElement = driver.FindElement(By.Id("minimizeChat"));
            chatElement.Click();

            driver.SwitchTo().DefaultContent();
            IWebElement loginButton = driver.FindElement(By.ClassName("login--button"));
            loginButton.Click();
            Thread.Sleep(3000);

            ChooseRightCustomer(driver, customerNumber);

            IWebElement downloadIcon = driver.FindElement(By.ClassName("bill-info--summary--icon"));
            downloadIcon.Click();
        }

        private static void ChooseRightCustomer(IWebDriver driver, string customerNumber)
        {
            IWebElement chooseUser = driver.FindElement(By.ClassName("header-user"));
            chooseUser.Click();

            IWebElement chosenUser = null;
            if (customerNumber == "0446181")
            {
                chosenUser = driver.FindElement(By.PartialLinkText("11-169"));
            }
            else
            {
                chosenUser = driver.FindElement(By.PartialLinkText("3-40"));
            }

            chosenUser.Click();
            Thread.Sleep(2000);
        }

        private static void DownloadVstInvoice(IWebDriver driver, string customerNumber, string password)
        {
            driver.Navigate().GoToUrl("https://savitarna.chc.lt/saskaitos/");

            IWebElement usernameElement = driver.FindElement(By.Id("txt1"));
            usernameElement.SendKeys(customerNumber);

            IWebElement passwordElement = driver.FindElement(By.Id("txt2"));
            passwordElement.SendKeys(password);

            IWebElement loginButton = driver.FindElement(By.Id("btn"));
            loginButton.Click();

            try
            {
                IWebElement closeButton = driver.FindElement(By.Id("btnClose"));
                closeButton.Click();
            }
            catch
            {

            }

            IWebElement downloadIcon = driver.FindElement(By.Id("btnShow"));
            downloadIcon.Click();
        }

        private static void DownloadVvInvoice(IWebDriver driver, string customerNumber, string password)
        {
            customerNumber = customerNumber.TrimStart('0');

            driver.Navigate().GoToUrl("https://savitarna.vv.lt/login.html#FORM:HomeLoginBlock:V:type=F:H764542779");

            //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
            //return wait.Until(drv => drv.FindElement(by));
            Thread.Sleep(1000);

            IWebElement usingPassword = driver.FindElement(By.ClassName("white_button_text"));
            usingPassword.Click();

            IWebElement usernameElement = driver.FindElement(By.CssSelector("input.gwt-TextBox"));
            usernameElement.SendKeys(customerNumber);

            IWebElement passwordElement = driver.FindElement(By.Id("inputPass"));
            passwordElement.SendKeys(password);

            IWebElement loginButton = driver.FindElement(By.ClassName("green_button"));
            loginButton.Click();

            driver.Manage().Window.Maximize();

            Thread.Sleep(2000);

            IWebElement paymentsButton = driver.FindElement(By.Id("payment_text"));
            paymentsButton.Click();

            Thread.Sleep(5000);
            IWebElement downloadIcon = driver.FindElement(By.ClassName("icon-file"));
            downloadIcon.Click();
        }
    }
}