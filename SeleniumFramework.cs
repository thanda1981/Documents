using AutoIt;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using Oracle.DataAccess.Client;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Framework
{
    public class SeleniumFramework
    {
        public static IWebDriver driver;
        public static string folderPath;
        //public static object ConfigurationManager { get; private set; }

        public static IWebDriver Init(string browsercode, string browserpath)
        {
            //driver = new ChromeDriver();
            if (browsercode == "IE")
            {
                if (browserpath != "")
                {
                    var ieOptions = new InternetExplorerOptions()
                    {
                        IntroduceInstabilityByIgnoringProtectedModeSettings = true,
                        IgnoreZoomLevel = true,
                        EnableNativeEvents = false
                    };
                    driver = new InternetExplorerDriver(browserpath, ieOptions);
                }
                else
                {
                    var ieOptions1 = new InternetExplorerOptions()
                    {
                        IntroduceInstabilityByIgnoringProtectedModeSettings = true,
                        IgnoreZoomLevel = true,
                        EnableNativeEvents = false
                    };
                    driver = new InternetExplorerDriver(ieOptions1);
                }

            }
            else if (browsercode == "EDGE")
            {
                //driver = new EdgeDriver();
                if (browserpath != "")
                {
                    driver = new EdgeDriver(browserpath);
                }
            }
            else if (browsercode == "FF")
            {
                //driver = new FirefoxDriver();
                if (browserpath != "")
                {
                    driver = new FirefoxDriver(browserpath);
                }
            }
            else if (browsercode == "CHROME")
            {
                if (browserpath != "")
                {
                    Console.WriteLine("browser path 1");
                    var chromeOptionsforPath = new ChromeOptions();
                    chromeOptionsforPath.AddUserProfilePreference("download.default_directory", @"D:\Downloads");
                    chromeOptionsforPath.AddUserProfilePreference("download.prompt_for_download", false);
                    chromeOptionsforPath.AddUserProfilePreference("download.directory_upgrade", true);
                    chromeOptionsforPath.AddUserProfilePreference("safebrowsing.enabled", true);
                    chromeOptionsforPath.Proxy = null;
                    chromeOptionsforPath.AddUserProfilePreference("plugins.plugins_disabled", new string[] { "Chrome PDF Viewer" });
                    chromeOptionsforPath.AddUserProfilePreference("credentials_enable_service", true);
                    chromeOptionsforPath.AddUserProfilePreference("profile.password_manager_enabled", true);
                    chromeOptionsforPath.AddUserProfilePreference("disable-popup-blocking", true);
                    chromeOptionsforPath.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    driver = new ChromeDriver(@browserpath, chromeOptionsforPath);
                    
                }
                else
                {
                    Console.WriteLine("browser path 2");
                    var chromeOptions = new ChromeOptions();
                    chromeOptions.AddUserProfilePreference("download.default_directory", @"D:\Downloads");
                    chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                    chromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
                    chromeOptions.AddUserProfilePreference("safebrowsing.enabled", true);
                    // chromeOptions.AddUserProfilePreference("disable-popup-blocking", true);
                    chromeOptions.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    chromeOptions.Proxy = null;
                    driver = new ChromeDriver(chromeOptions);
                }
            }
            return driver;
        }



        public static void SetClickButton(string Locator,string Element)
        {
            switch(Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(Element), 20);
                    // screenshot(folderPath);
                    IWebElement IdElement = driver.FindElement(By.Id(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", IdElement);
                    break;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(Element), 10);
                    // screenshot(folderPath);
                    IWebElement XpathElement = driver.FindElement(By.XPath(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", XpathElement);
                    break;

                case "PARTIALLINKTEXT":
                    waitForPageUntilElementVisible(By.PartialLinkText(Element), 10);
                    // screenshot(folderPath);
                    IWebElement PartialLinkElement = driver.FindElement(By.PartialLinkText(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", PartialLinkElement);
                    break;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(Element), 10);
                    // screenshot(folderPath);
                    IWebElement CssElement = driver.FindElement(By.CssSelector(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", CssElement);
                    break;

                case "LINKTEXT":
                    waitForPageUntilElementVisible(By.LinkText(Element), 10);
                    // screenshot(folderPath);
                    IWebElement LinkTextElement = driver.FindElement(By.LinkText(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", LinkTextElement);
                    break;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(Element), 10);
                    // screenshot(folderPath);
                    IWebElement ClassNameElement = driver.FindElement(By.ClassName(Element));
                    driver.ExecuteJavaScript<object>("arguments[0].click();", ClassNameElement);
                    break;

            }

        }

        public static void setClickandClearSendValue(string Locator, string Element, string value)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    SeleniumFramework.waitForPageUntilElementVisible(By.Id(Element), 50);
                    IWebElement id = driver.FindElement(By.Id(Element));
                    id.Click();
                    id.Clear();
                    id.SendKeys(value);
                    break;

                case "XPATH":
                    SeleniumFramework.waitForPageUntilElementVisible(By.XPath(Element), 50);
                    IWebElement XPath = driver.FindElement(By.XPath(Element));
                    XPath.Click();
                    XPath.Clear();
                    XPath.SendKeys(value);
                    break;

                case "CSS":
                    SeleniumFramework.waitForPageUntilElementVisible(By.CssSelector(Element), 50);
                    IWebElement CssSelector = driver.FindElement(By.CssSelector(Element));
                    CssSelector.Click();
                    CssSelector.Clear();
                    CssSelector.SendKeys(value);
                    break;

                case "PARTIAL":
                    SeleniumFramework.waitForPageUntilElementVisible(By.PartialLinkText(Element), 50);
                    IWebElement PartialLinkText = driver.FindElement(By.PartialLinkText(Element));
                    PartialLinkText.Click();
                    PartialLinkText.Clear();
                    PartialLinkText.SendKeys(value);
                    break;

                case "LINKTEXT":
                    SeleniumFramework.waitForPageUntilElementVisible(By.LinkText(Element), 50);
                    IWebElement LinkText = driver.FindElement(By.LinkText(Element));
                    LinkText.Click();
                    LinkText.Clear();
                    LinkText.SendKeys(value);
                    break;

                case "CLASSNAME":
                    SeleniumFramework.waitForPageUntilElementVisible(By.ClassName(Element), 50);
                    IWebElement ClassName = driver.FindElement(By.ClassName(Element));
                    ClassName.Click();
                    ClassName.Clear();
                    ClassName.SendKeys(value);
                    break;

                case "TAGNAME":
                    SeleniumFramework.waitForPageUntilElementVisible(By.TagName(Element), 50);
                    IWebElement TagName = driver.FindElement(By.TagName(Element));
                    TagName.Click();
                    TagName.Clear();
                    TagName.SendKeys(value);
                    break;
            }
        }

        public static void setTextBoxSendValueandTab(string Locator,string Element,string Value)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(Element), 10);
                    // screenshot(folderPath);
                    IWebElement IdElement = driver.FindElement(By.Id(Element));
                    IdElement.SendKeys(Value);
                    IdElement.SendKeys(Keys.Tab);
                    break;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(Element), 10);
                    // screenshot(folderPath);
                    IWebElement XpathElement = driver.FindElement(By.XPath(Element));
                    XpathElement.SendKeys(Value);
                    XpathElement.SendKeys(Keys.Tab);
                    break;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(Element), 10);
                    // screenshot(folderPath);
                    IWebElement CssElement = driver.FindElement(By.CssSelector(Element));
                    CssElement.SendKeys(Value);
                    CssElement.SendKeys(Keys.Tab);
                    break;

                case "PARTIAL":
                    waitForPageUntilElementVisible(By.PartialLinkText(Element), 10);
                    // screenshot(folderPath);
                    IWebElement PartialElement = driver.FindElement(By.PartialLinkText(Element));
                    PartialElement.SendKeys(Value);
                    PartialElement.SendKeys(Keys.Tab);
                    break;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(Element), 10);
                    // screenshot(folderPath);
                    IWebElement ClassNameElement = driver.FindElement(By.ClassName(Element));
                    ClassNameElement.SendKeys(Value);
                    ClassNameElement.SendKeys(Keys.Tab);
                    break;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(Element), 10);
                    // screenshot(folderPath);
                    IWebElement NameElement = driver.FindElement(By.Name(Element));
                    NameElement.SendKeys(Value);
                    NameElement.SendKeys(Keys.Tab);
                    break;
            }

        }

        public static void setTextBoxSendValueandEnter(string Locator, string Element, string Value)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(Element), 10);
                    // screenshot(folderPath);
                    IWebElement IdElement = driver.FindElement(By.Id(Element));
                    IdElement.SendKeys(Value);
                    IdElement.SendKeys(Keys.Enter);
                    break;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(Element), 10);
                    // screenshot(folderPath);
                    IWebElement XpathElement = driver.FindElement(By.XPath(Element));
                    XpathElement.SendKeys(Value);
                    XpathElement.SendKeys(Keys.Enter);
                    break;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(Element), 10);
                    // screenshot(folderPath);
                    IWebElement CssElement = driver.FindElement(By.CssSelector(Element));
                    CssElement.SendKeys(Value);
                    CssElement.SendKeys(Keys.Enter);
                    break;

                case "PARTIAL":
                    waitForPageUntilElementVisible(By.PartialLinkText(Element), 10);
                    // screenshot(folderPath);
                    IWebElement PartialElement = driver.FindElement(By.PartialLinkText(Element));
                    PartialElement.SendKeys(Value);
                    PartialElement.SendKeys(Keys.Enter);
                    break;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(Element), 10);
                    // screenshot(folderPath);
                    IWebElement ClassNameElement = driver.FindElement(By.ClassName(Element));
                    ClassNameElement.SendKeys(Value);
                    ClassNameElement.SendKeys(Keys.Enter);
                    break;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(Element), 10);
                    // screenshot(folderPath);
                    IWebElement NameElement = driver.FindElement(By.Name(Element));
                    NameElement.SendKeys(Value);
                    NameElement.SendKeys(Keys.Enter);
                    break;
            }

        }

        public static void setTextBoxSendValue(string Locator, string Element, string Value)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(Element), 10);
                    // screenshot(folderPath);
                    IWebElement IdElement = driver.FindElement(By.Id(Element));
                    IdElement.SendKeys(Value);
                    break;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(Element), 10);
                    // screenshot(folderPath);
                    IWebElement XpathElement = driver.FindElement(By.XPath(Element));
                    XpathElement.SendKeys(Value);
                    break;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(Element), 10);
                    // screenshot(folderPath);
                    IWebElement CssElement = driver.FindElement(By.CssSelector(Element));
                    CssElement.SendKeys(Value);
                    break;

                case "PARTIAL":
                    waitForPageUntilElementVisible(By.PartialLinkText(Element), 10);
                    // screenshot(folderPath);
                    IWebElement PartialElement = driver.FindElement(By.PartialLinkText(Element));
                    PartialElement.SendKeys(Value);
                    break;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(Element), 10);
                    // screenshot(folderPath);
                    IWebElement ClassNameElement = driver.FindElement(By.ClassName(Element));
                    ClassNameElement.SendKeys(Value);
                    break;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(Element), 10);
                    // screenshot(folderPath);
                    IWebElement NameElement = driver.FindElement(By.Name(Element));
                    NameElement.SendKeys(Value);
                    break;
            }

        }

        public static void setDropDown(string Locator, IWebElement Element, string Value)
        {
            switch (Locator.ToUpper())
            {
                case "INDEX":
                    SelectElement BasedOnIndex = new SelectElement(Element);
                    BasedOnIndex.SelectByIndex(int.Parse(Value));
                    break;
                case "VALUE":
                    SelectElement BasedOnValue = new SelectElement(Element);
                    BasedOnValue.SelectByValue(Value);
                    break;
                case "TEXT":
                    SelectElement BasedOnText = new SelectElement(Element);
                    BasedOnText.SelectByText(Value);
                    break;
            }
        }

        public static void SelectCheckBox(string Locator,string element)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    IWebElement SelectChecBoxID = driver.FindElement(By.Id(element));
                    SelectChecBoxID.Click();
                    break;
                case "XPATH":
                    IWebElement SelectChecBoxXpath = driver.FindElement(By.XPath(element));
                    SelectChecBoxXpath.Click();
                    break;
                case "CLASSNAME":
                    IWebElement SelectChecBoxClass = driver.FindElement(By.ClassName(element));
                    SelectChecBoxClass.Click();
                    break;
                case "NAME":
                    IWebElement SelectChecBoxName = driver.FindElement(By.Name(element));
                    SelectChecBoxName.Click();
                    break;
                case "CSS":
                    IWebElement SelectChecBoxCss = driver.FindElement(By.CssSelector(element));
                    SelectChecBoxCss.Click();
                    break;
            }

        }
        


        public static void SelectMultiTables(string table1, string table2, string buttonname, string bookno)
        {

            new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(30)).Until(
            d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            IList<IWebElement> mtrlist = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(table1)).FindElements(By.TagName("tr")));
            IList<IWebElement> mtdlist;
            int mtdcnt = 0;
            int trline = 0;
            int tdline = 0;
            foreach (var t in mtrlist)
            {
                trline = trline + 1;
                mtdlist = t.FindElements(By.TagName("td"));
                mtdcnt = mtdlist.Count;
                for (int k = 0; k < mtdcnt; k++)
                {
                    tdline = tdline + 1;
                    if (mtdlist[k].Text != "")
                    {
                        string bookingno = mtdlist[k].Text;
                        if (bookingno.Equals(bookno))
                        {
                            Multitableclickbutton(table2, buttonname, trline);
                            goto endloop;
                        }
                    }
                }

                tdline = 0;
            }
            endloop:;
        }

        public static void Multitableclickbutton(string tableXpath, string buttonname, int trline)
        {

            IList<IWebElement> countoftr = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(tableXpath)).FindElements(By.TagName("tr")));
            IList<IWebElement> tdbutton = countoftr[trline - 1].FindElements(By.TagName("button"));
            for (int j = 0; j < tdbutton.Count; j++)
            {
                string s = tdbutton[j].GetAttribute("title");
                if (s.Equals(buttonname))
                {
                    SeleniumFramework.HighlightElement(tdbutton[j]);
                    tdbutton[j].Click();
                    // screenshot(folderPath);
                    break;
                }
            }

        }


        public static string isAlertPresent(string value)
        {
            string noAlertPresent = null;
            try
            {
                string alertValueToPass = AlertHandling(value);
                driver.SwitchTo().Alert().Accept();
                return alertValueToPass;
            }
            catch (NoAlertPresentException)
            {
                return noAlertPresent;
            }
        }

        public static string AlertHandling(string value)
        {
            string alertText = driver.SwitchTo().Alert().Text;

            var a = Regex.Replace(alertText, "[^0-9,.]+", string.Empty);
            string g = a.Substring(0, 3);
            float typeValue = float.Parse(g);
            float floatvalue = float.Parse(value);
            if (floatvalue >= typeValue)
            {
                Console.WriteLine("Given value Accepted" + value + Keys.Enter);
                return floatvalue.ToString();
            }
            else
            {
                return (typeValue + 1).ToString();
            }
        }

        public static string GetAttribute(string Locator,string element,string AttributeName)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    IWebElement FindElementId = driver.FindElement(By.Id(element));
                    string idValue = FindElementId.GetAttribute(AttributeName);
                    return idValue;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    IWebElement FindElementXpath = driver.FindElement(By.XPath(element));
                    string XpathValue = FindElementXpath.GetAttribute(AttributeName);
                    return XpathValue;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(element), 10);
                    IWebElement FindElementClass = driver.FindElement(By.ClassName(element));
                    string ClassValue = FindElementClass.GetAttribute(AttributeName);
                    return ClassValue;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(element), 10);
                    IWebElement FindElementName = driver.FindElement(By.Name(element));
                    string NameValue = FindElementName.GetAttribute(AttributeName);
                    return NameValue;

                case "CSS":
                    IWebElement SelectChecBoxCss = driver.FindElement(By.CssSelector(AttributeName));
                    waitForPageUntilElementVisible(By.CssSelector(element), 10);
                    IWebElement FindElementCssSelector  = driver.FindElement(By.CssSelector(element));
                    string CssSelectorValue = FindElementCssSelector.GetAttribute(AttributeName);
                    return CssSelectorValue;

            }
            return null;
        }

        public static string GetText(string Locator, string element)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    string TextValueId = driver.FindElement(By.Id(element)).Text;
                    return TextValueId;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    string textValueXpath = driver.FindElement(By.XPath(element)).Text;
                    return textValueXpath;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(element), 10);
                    string TextValueClassName = driver.FindElement(By.ClassName(element)).Text;
                    return TextValueClassName;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(element), 10);
                    string TextValueName = driver.FindElement(By.Name(element)).Text;
                    return TextValueName;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(element), 10);
                    string TextValueCssSelector = driver.FindElement(By.CssSelector(element)).Text;
                    return TextValueCssSelector;

            }
            return null;
        }

        public static bool VerifyElement(string Locator,string element)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    bool FindElementId = driver.FindElement(By.Id(element)).Displayed;
                    return FindElementId;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    bool FindElementXPath = driver.FindElement(By.Id(element)).Displayed;
                    return FindElementXPath;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(element), 10);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    waitForPageUntilElementVisible(By.ClassName(element), 10);
                    bool FindElementClassName = driver.FindElement(By.Id(element)).Displayed;
                    return FindElementClassName;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(element), 10);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    waitForPageUntilElementVisible(By.Name(element), 10);
                    bool FindElementName = driver.FindElement(By.Id(element)).Displayed;
                    return FindElementName;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(element), 10);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    waitForPageUntilElementVisible(By.CssSelector(element), 10);
                    bool FindElementCssSelector = driver.FindElement(By.Id(element)).Displayed;
                    return FindElementCssSelector;

            }
            return false;

        }

        public static void MouseHoverAction(string Locator, string element)
        {
            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    // screenshot(folderPath);
                    IWebElement IdElement = SeleniumFramework.driver.FindElement(By.Id(element));
                    Actions actionId = new Actions(driver);
                    actionId.MoveToElement(IdElement).Click().Perform();
                    break;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    // screenshot(folderPath);
                    IWebElement XpathElement = SeleniumFramework.driver.FindElement(By.XPath(element));
                    Actions actionXpath = new Actions(driver);
                    actionXpath.MoveToElement(XpathElement).Click().Perform();
                    break;

                case "DOUBLECLICKID":
                    waitForPageUntilElementVisible(By.Id(element), 20);
                    // screenshot(folderPath);
                    IWebElement IdDCElement = SeleniumFramework.driver.FindElement(By.Id(element));
                    Actions actionDCId = new Actions(driver);
                    actionDCId.MoveToElement(IdDCElement).DoubleClick().Build().Perform();
                    break;

                case "DOUBLECLICKXPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 20);
                    // screenshot(folderPath);
                    IWebElement XpathDCElement = SeleniumFramework.driver.FindElement(By.XPath(element));
                    Actions actionDCXpath = new Actions(driver);
                    actionDCXpath.MoveToElement(XpathDCElement).DoubleClick().Build().Perform();
                    break;


            }

        }
        public static string GetDateFormatValue(string Locator, string element)
        {

            switch (Locator.ToUpper())
            {
                case "ID":
                    waitForPageUntilElementVisible(By.Id(element), 10);
                    IWebElement FindElementId = driver.FindElement(By.Id(element));
                    string dateFormatValueid = FindElementId.GetAttribute("data-format");
                    return dateFormatValueid;

                case "XPATH":
                    waitForPageUntilElementVisible(By.XPath(element), 10);
                    // screenshot(folderPath);
                    IWebElement FindElementXpath = driver.FindElement(By.XPath(element));
                    string dateFormatValuexpath = FindElementXpath.GetAttribute("data-format");
                    return dateFormatValuexpath;

                case "CSS":
                    waitForPageUntilElementVisible(By.CssSelector(element), 10);
                    // screenshot(folderPath);
                    IWebElement CssElement = driver.FindElement(By.CssSelector(element));
                    string dateFormatValuecss = CssElement.GetAttribute("data-format");
                    return dateFormatValuecss;

                case "PARTIAL":
                    waitForPageUntilElementVisible(By.PartialLinkText(element), 10);
                    // screenshot(folderPath);
                    IWebElement PartialElement = driver.FindElement(By.PartialLinkText(element));
                    string dateFormatValuepartial = PartialElement.GetAttribute("data-format");
                    return dateFormatValuepartial;

                case "CLASSNAME":
                    waitForPageUntilElementVisible(By.ClassName(element), 10);
                    // screenshot(folderPath);
                    IWebElement ClassNameElement = driver.FindElement(By.ClassName(element));
                    string dateFormatValueclass = ClassNameElement.GetAttribute("data-format");
                    return dateFormatValueclass;

                case "NAME":
                    waitForPageUntilElementVisible(By.Name(element), 10);
                    // screenshot(folderPath);
                    IWebElement NameElement = driver.FindElement(By.Name(element));
                    string dateFormatValuename = NameElement.GetAttribute("data-format");
                    return dateFormatValuename;
            }
            return null;
        }

        public static string GetDateFormatValueBasedonDaysId(string ID, int days)
        {
            waitForPageUntilElementVisible(By.Id(ID), 10);
            IWebElement FindElementId = driver.FindElement(By.Id(ID));
            string dateFormatValue = FindElementId.GetAttribute("data-format");
            string dateTime = DateTime.Today.AddDays(days).ToString(dateFormatValue, new CultureInfo("en-US"));
            return dateTime;
        }
        public static string GetDateFormatValueBasedonDaysXpath(string Xpath, int days)
        {
            waitForPageUntilElementVisible(By.XPath(Xpath), 10);
            IWebElement FindElementId = driver.FindElement(By.XPath(Xpath));
            string dateFormatValue = FindElementId.GetAttribute("data-format");
            string dateTime = DateTime.Today.AddDays(days).ToString(dateFormatValue, new CultureInfo("en-US"));
            return dateTime;
        }


        public static void SpanClassdropdown(string arrowButton, string textElement, string value)
        {
            driver.FindElement(By.XPath(arrowButton)).Click();
            driver.FindElement(By.XPath(textElement)).SendKeys(value);
            System.Threading.Thread.Sleep(1000);
            Actions action = new Actions(driver);
            action.SendKeys(Keys.Enter);
        }
        public static void HighlightElement(IWebElement highLightelement)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)SeleniumFramework.driver;
            js.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", highLightelement, " border: 3px solid red;");
        }

       		
        public static void SelectTablevalue(string Xpath, string valuetoPass, int val)
        {
            // screenshot(folderPath);
            rerun:;
            IWebElement table = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            IList<IWebElement> tableValues = new List<IWebElement>(table.FindElements(By.TagName("tr")));
            IList<IWebElement> rowValues;
            int rowCount = tableValues.Count;
            foreach (var row in tableValues)
            {
                try
                {
                    rowValues = row.FindElements(By.TagName("td"));
                    string PassingValue = rowValues[val].Text;
                    if (valuetoPass.Equals(rowValues[val].Text))
                    {
                        bool rowvalue = rowValues[val].Displayed;
                        int tableCount = 0;
                        while (rowvalue == true)
                        {
                            if (tableCount > 3)
                            {
                                goto end;
                            }
                            WebDriverWait wait = new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(10));
                            wait.Until(ExpectedConditions.TextToBePresentInElement(rowValues[val], valuetoPass));
                            // screenshot(folderPath);
                            driver.ExecuteJavaScript<object>("arguments[0].click();", rowValues[val]);
                            // screenshot(folderPath);
                            rowvalue = rowValues[val].Displayed;
                            tableCount++;
                        }
                        goto end;
                    }
                }
                catch (StaleElementReferenceException ex)
                {
                    Console.WriteLine(ex);
                    goto rerun;
                }

            }//***Using specific cell value   
            end:;
        }

        public static void SelectTableWithClickablePoint(string Xpath, string valuetoPass, int val, string ClickablePointText)
        {
            // screenshot(folderPath);
            rerun:;
            IWebElement table = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            IList<IWebElement> tableValues = new List<IWebElement>(table.FindElements(By.TagName("tr")));
            IList<IWebElement> rowValues;
            int rowCount = tableValues.Count;
            foreach (var row in tableValues)
            {
                try
                {
                    rowValues = row.FindElements(By.TagName("td"));
                    string PassingValue = rowValues[val].Text;
                    if (valuetoPass.Equals(rowValues[val].Text))
                    {
                        bool rowvalue = rowValues[val].Displayed;
                        int tableCount = 0;
                        while (rowvalue == true)
                        {
                            if (tableCount > 3)
                            {
                                goto end;
                            }
                            WebDriverWait wait = new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(10));
                            wait.Until(ExpectedConditions.TextToBePresentInElement(rowValues[val], valuetoPass));
                            // screenshot(folderPath);
                            IWebElement clicakblePoint = SeleniumFramework.driver.FindElement(By.XPath("//*[contains(@onclick," + "\"" + ClickablePointText + "\")]"));
                            clicakblePoint.GetAttribute("onclick");
                            IWebElement exactclicablePoint = clicakblePoint.FindElement(By.TagName("i"));
                            SeleniumFramework.HighlightElement(exactclicablePoint);
                            Actions action = new Actions(driver);
                            action.MoveToElement(exactclicablePoint).Perform();
                            exactclicablePoint.Click();
                            // screenshot(folderPath);
                            rowvalue = rowValues[val].Displayed;
                            tableCount++;
                        }
                        goto end;
                    }
                }
                catch (StaleElementReferenceException ex)
                {
                    Console.WriteLine(ex);
                    goto rerun;
                }

            }//***Using specific cell value   
            end:;
        }

        public static void SelectTablevalueWithClickablePoint(string Xpath, string valuetoPass, int val)
        {
            // screenshot(folderPath);
            rerun:;
            IWebElement table = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            IList<IWebElement> tableValues = new List<IWebElement>(table.FindElements(By.TagName("tr")));
            IList<IWebElement> rowValues;
            int rowCount = tableValues.Count;
            foreach (var row in tableValues)
            {
                try
                {
                    rowValues = row.FindElements(By.TagName("td"));
                    string PassingValue = rowValues[val].Text;
                    if (valuetoPass.Equals(rowValues[val].Text))
                    {
                        bool rowvalue = rowValues[val].Displayed;
                        int tableCount = 0;
                        while (rowvalue == true)
                        {
                            if (tableCount > 3)
                            {
                                goto end;
                            }
                            WebDriverWait wait = new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(10));
                            wait.Until(ExpectedConditions.TextToBePresentInElement(rowValues[val], valuetoPass));
                            // screenshot(folderPath);
                            driver.ExecuteJavaScript<object>("arguments[0].click();", rowValues[val].FindElement(By.TagName("button")));
                            // screenshot(folderPath);
                            rowvalue = rowValues[val].Displayed;
                            tableCount++;
                        }
                        goto end;
                    }
                }
                catch (StaleElementReferenceException ex)
                {
                    Console.WriteLine(ex);
                    goto rerun;
                }

            }//***Using specific cell value   
            end:;
        }

        public static void ScrollingOption(IWebElement element)
        {
            ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].scrollIntoView(true);", element);

        }

        public static void SelectTablesButtonClick(string TableXpath, string ValueToCompare, int Val, string GetattributeButton, string ButtonToCompare)
        {
            IList<IWebElement> mtrlist = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(TableXpath)).FindElements(By.TagName("tr")));
            IList<IWebElement> mtdlist;
            IList<IWebElement> rmtdlist;
            IList<IWebElement> butlist;
            int buttoncount = 0;
            int mtrchk = 0;
            int mtdchk = 0;
            int butchk = 0;
            foreach (var t in mtrlist) //tr list of table
            {
                mtrchk = mtrchk + 1;
                mtdlist = t.FindElements(By.TagName("td"));

                if (mtdlist[Val].Text.Equals(ValueToCompare))
                {
                    rmtdlist = t.FindElements(By.TagName("td"));

                    foreach (var r in rmtdlist) //td list of table
                    {
                        mtdchk = mtdchk + 1;
                        butlist = r.FindElements(By.TagName("button"));
                        if (butlist.Count > 0)
                        {
                            foreach (var b in butlist)
                            {
                                butchk = butchk + 1;
                                string butval = butlist[buttoncount].GetAttribute(GetattributeButton);
                                buttoncount = buttoncount + 1;
                                if (GetattributeButton == "onclick")
                                {
                                    if (butval.Contains(ButtonToCompare))
                                    {
                                        ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", b);
                                        goto endloop;
                                    }
                                }
                                else if (GetattributeButton == "title")
                                {
                                    if (butval.Equals(ButtonToCompare))
                                    {
                                        ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", b);
                                        goto endloop;
                                    }
                                }
                            }

                        }
                    }


                }

            }
            endloop:;
            Console.WriteLine("count :{0} and {1} and {2}", mtrchk, mtdchk, butchk);
        }

        public static bool IsTestElementPresent(string Xpath)
        {
            try
            {
                bool verifyElement = driver.FindElement(By.XPath(Xpath)).Displayed == true;
                // HighlightElement(verifyElement);
                if (verifyElement == false)
                {
                    return false;

                }
                return true;

            }
            catch (Exception)
            {
                return false;
            }
        }


        public static void SelectTablevalueMultiple(string Xpath, string valuetoPass1, int val, string valuetoPass2, int val2)
        {
            IWebElement table = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            IList<IWebElement> tableValues = new List<IWebElement>(table.FindElements(By.TagName("tr")));
            IList<IWebElement> rowValues;
            int rowCount = tableValues.Count;
            foreach (var row in tableValues)
            {
                rowValues = row.FindElements(By.TagName("td"));
                int rowCounts = rowValues.Count;
                if (rowValues[val].Text.Contains(valuetoPass1) && rowValues[val2].Text.Contains(valuetoPass2))
                {
                    Actions action = new Actions(SeleniumFramework.driver);
                    action.MoveToElement(rowValues[val]).Build().Perform();
                    Thread.Sleep(1000);
                    rowValues[val].Click();
                    // screenshot(folderPath);
                    break;
                }//***Using specific cell value
                rowCount++;
            }
        }



        public static List<string> ParamValues(string SQLEnvironment, string query, string procedure)
        {
            OracleConnection DbConnection = new OracleConnection();
            DbConnection.ConnectionString = ConfigurationManager.ConnectionStrings[SQLEnvironment].ToString();
            //OracleConnection DbConnection = new OracleConnection("Data Source=MSCV5DEV; User ID=DEPOT_TEST_AUTOMATION; Password=DEPOT_TEST_AUTOMATION");
            DbConnection.Open();
            OracleCommand DBCommand = new OracleCommand(query, DbConnection);
            //  DBCommand.CommandType = CommandType.Text;

            string getOutputParamNames;

            /*if (SQLEnvironment.Equals("Depot_Test"))
            {
                getOutputParamNames = "SELECT ARGUMENT_NAME FROM ALL_ARGUMENTS  WHERE OWNER = 'DEPOT_TEST_AUTOMATION' AND PACKAGE_NAME|| '.' || OBJECT_NAME = upper('" + procedure + "')  AND IN_OUT = 'OUT' ORDER BY POSITION ";
            }

            else
            {
                getOutputParamNames = "SELECT ARGUMENT_NAME FROM ALL_ARGUMENTS  WHERE OWNER = 'OVDEPOT' AND PACKAGE_NAME|| '.' || OBJECT_NAME = upper('" + procedure + "')  AND IN_OUT = 'OUT' ORDER BY POSITION ";
            }*/


            getOutputParamNames = "SELECT ARGUMENT_NAME FROM ALL_ARGUMENTS  WHERE OWNER = (SELECT user from dual) AND PACKAGE_NAME|| '.' || OBJECT_NAME = upper('" + procedure + "')  AND IN_OUT = 'OUT' ORDER BY POSITION ";

            List<Tuple<string, string>> outputvalues = ReturnMultipleRows(SQLEnvironment, getOutputParamNames);
            foreach (var data in outputvalues)
            {
                DBCommand.Parameters.Add(data.Item2, OracleDbType.Varchar2, 32767, data.Item2, ParameterDirection.Output);
            }

            DBCommand.ExecuteNonQuery();
            List<string> result = new List<string>();


            foreach (var data in outputvalues)
            {
                string b = DBCommand.Parameters[data.Item2].Value.ToString();
                result.Add(b);
            }

            DbConnection.Close();
            return result;
        }

        public static List<Tuple<string, string>> ReturnMultipleRows(string SQLEnvironment, string sql)
        {
            List<Tuple<string, string>> row = new List<Tuple<string, string>>();
            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = ConfigurationManager.ConnectionStrings[SQLEnvironment].ToString();
            //conn.ConnectionString = ConfigurationManager.ConnectionStrings["Depot"].ToString();
            //OracleConnection conn = new OracleConnection("Data Source=MSCV5DEV; User ID=DEPOT_TEST_AUTOMATION; Password=DEPOT_TEST_AUTOMATION");

            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            conn.Open();
            cmd.CommandText = sql;
            cmd.CommandType = CommandType.Text;
            OracleDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {

                //result.Add(row);
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string columnname = reader.GetName(i);
                    string columnvalue = reader.GetValue(i).ToString();
                    row.Add(new Tuple<string, string>(columnname, columnvalue));
                }
            }
            reader.Close();
            conn.Close();
            return row;
        }

        public static Dictionary<string, string> RowValues(string Environment, string sql)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = ConfigurationManager.ConnectionStrings[Environment].ToString();
            //conn.ConnectionString = "User ID=MSC_DEPOT;Password=MSC_DEPOT;Data Source=MSCV5DEV";
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            conn.Open();
            cmd.CommandText = sql;
            cmd.CommandType = CommandType.Text;
            OracleDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string columnname = reader.GetName(i);
                    string columnvalue = reader.GetValue(i).ToString();
                    result.Add(columnname, columnvalue);
                }
                break;
            }
            reader.Close();
            conn.Close();
            return result;
        }

        public static Dictionary<string, string> ReturnSingleRow(string Env, string sql)
        {

            Dictionary<string, string> result = new Dictionary<string, string>();

            //var test = ConfigurationManager.ConnectionStrings[Env].ToString();

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[Env].ToString());

            SqlCommand command = new SqlCommand();
            SqlDataReader reader;
            conn.Open();

            command.Connection = conn;

            command.CommandText = sql;
            command.CommandTimeout = 800;  //IN SECS
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string columnname = reader.GetName(i);
                    string columnvalue = reader.GetValue(i).ToString();

                    result.Add(columnname, columnvalue);
                }

                break;
            }
            reader.Close();
            conn.Close();

            return result;
        }

        public static Int32 GetRandomNumber(int NoOfDigit)
        {
            string outstart = null;
            string outend = null;
            int start = 1;
            int end   = 9;
            for (int i = 1; i <= NoOfDigit; i++)
            {
                if (outstart == null)
                {
                    outstart = start.ToString();
                    outend   = end.ToString();
                }
                else
                {
                    outstart = string.Concat(outstart + start);
                    outend   = string.Concat(outend + end);
                }
            }
            Random randomss = new Random();
            Int32 randomDigit = randomss.Next(int.Parse(outstart), int.Parse(outend));
            return randomDigit;
        }

        public static Int32 GetRandomNumber(int NoOfDigit,int startpoint,int endpoint)
        {
            string outstart = null;
            string outend = null;
            int start = startpoint;
            int end = endpoint;
            for (int i = 1; i <= NoOfDigit; i++)
            {
                if (outstart == null)
                {
                    outstart = start.ToString();
                    outend = end.ToString();
                }
                else
                {
                    outstart = string.Concat(outstart + start);
                    outend = string.Concat(outend + end);
                }
            }
            Random randomss = new Random();
            Int32 randomDigit = randomss.Next(int.Parse(outstart), int.Parse(outend));
            return randomDigit;
        }

      
        public static void ClickHiddenElementByXpath(string Xpath)
        {
            // screenshot(folderPath);
            waitForPageUntilElementVisible(By.XPath(Xpath), 50);
            ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", SeleniumFramework.driver.FindElement(By.XPath(Xpath)));
            // screenshot(folderPath);
        }

        public static void ClickHiddenElementById(string Id)
        {
            // screenshot(folderPath);
            waitForPageUntilElementVisible(By.Id(Id), 500);
            ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", SeleniumFramework.driver.FindElement(By.Id(Id)));
            // screenshot(folderPath);
        }
        public static void MouseHoverJavaScriptById(string Id)
        {
            // screenshot(folderPath);
            string mouseOverScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover',true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');}";
            ((IJavaScriptExecutor)driver).ExecuteScript(mouseOverScript,
                    SeleniumFramework.driver.FindElement(By.Id(Id)));
        }


        public static void UntilFindElementByXpath(string Xpath, int timeoutInSeconds)
        {
            // screenshot(folderPath);
            var wait = new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(timeoutInSeconds));
            wait.Until(drv => drv.FindElement(By.XPath(Xpath)));
            // screenshot(folderPath);
        }

        public static void Screenshot(string imageLocation)
        {
            string fileName = SeleniumFramework.GetTimestamp(DateTime.Now);
            ITakesScreenshot  Screenshotdriver = driver as ITakesScreenshot;
            Screenshot  Screenshot =  Screenshotdriver.GetScreenshot();
            string saveLocation = @imageLocation + fileName + ".png";
            Screenshot image = ((ITakesScreenshot)driver).GetScreenshot();
            image.SaveAsFile(saveLocation,  ScreenshotImageFormat.Png);
        }

        public static string GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }

        public static void listboxselectiontrail(string Xpath, string ScrolldownXpath, string id, string value)
        {

            top:;
            IWebElement Textfield = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            Textfield.Clear();
            Textfield.Click();
            Textfield.SendKeys(value);
            string flag = "No";
            IWebElement items = SeleniumFramework.driver.FindElement(By.Id(id));
            IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.TagName("li")));
            try
            {
                for (int i = 0; i < liValues.Count; i++)
                {
                    string indexValue = (liValues[i].GetAttribute("data-offset-index"));

                    string listText = SeleniumFramework.driver.FindElement(By.XPath("//li[@data-offset-index='" + indexValue + "']")).Text;
                    if (listText.Equals(value))
                    {
                        IWebElement dropdownbutton = SeleniumFramework.driver.FindElement(By.XPath(ScrolldownXpath));
                        dropdownbutton.Click();
                        dropdownbutton.Click();
                        // screenshot(folderPath);
                        IWebElement Uilink = SeleniumFramework.driver.FindElement(By.XPath("//li[@data-offset-index='" + indexValue + "']"));
                        Actions action = new Actions(SeleniumFramework.driver);
                        action.MoveToElement(Uilink).Click().Perform();
                        // screenshot(folderPath);
                        action.SendKeys(Keys.Tab);

                        flag = "Yes";
                        goto endloop;
                    }
                }
            }
            catch (StaleElementReferenceException)
            {
                goto top;
            }
            if (flag == "No")
            {
                goto top;
            }
            endloop:;
        }

        public static void ListBoxSelectionByliID(string Xpath, string id, string value, int seconds)
        {

            waitForPageUntilElementVisible(By.XPath(Xpath), 50);
            // screenshot(folderPath);
            IWebElement Textfield = SeleniumFramework.driver.FindElement(By.XPath(Xpath));
            Textfield.Clear();
            //            Textfield.Click();
            ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Textfield);
            int count = 0;
            foreach (char val in value)
            {
                count = count + 1;
                Textfield.SendKeys(val.ToString());
                // screenshot(folderPath);
                Task.Delay(seconds).Wait();
                if (id != "")
                {
                    if (count > 4)
                    {
                        IWebElement items = SeleniumFramework.driver.FindElement(By.Id(id));
                        IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.TagName("li")));
                        Console.WriteLine("list if2 :{0} and value: {1}", liValues.Count, value);
                        if (liValues.Count == 1)
                        {
                            Textfield.SendKeys(Keys.Enter);
                            // screenshot(folderPath);
                            Textfield.SendKeys(Keys.Tab);
                            // screenshot(folderPath);
                            goto endstat;
                        }
                    }
                }
            }
            Textfield.SendKeys(Keys.Enter);
            // screenshot(folderPath);
            Textfield.SendKeys(Keys.Tab);
            // screenshot(folderPath);
            endstat:;
        }

        public static void ListBoxSelectionXpathXpath(string ButtonXpath, string xpath, string xpathSendKeys, string valueToSearch)
        {
            rerun:;
            waitForPageUntilElementVisible(By.XPath(ButtonXpath), 10);
            // screenshot(folderPath);
            IWebElement cli = SeleniumFramework.driver.FindElement(By.XPath(ButtonXpath));
            Actions action = new Actions(driver);
            action.MoveToElement(cli).Click().Perform();
            // screenshot(folderPath);
            SeleniumFramework.driver.FindElement(By.XPath(xpathSendKeys)).SendKeys(valueToSearch);
            // screenshot(folderPath);
            waitForPageUntilElementVisible(By.Id(xpath), 10);
            IWebElement items = SeleniumFramework.driver.FindElement(By.Id(xpath));
            IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.TagName("li")));
            for (int i = 1; i < liValues.Count; i++)
            {
                try
                {
                    string indexValue = (liValues[i].GetAttribute("data-offset-index"));
                    string listText = SeleniumFramework.driver.FindElement(By.XPath("//li[@data-offset-index='" + indexValue + "']")).Text;
                    if (listText.Equals(valueToSearch))
                    {
                        driver.ExecuteJavaScript<object>("arguments[0].click();", liValues[i]);
                        // screenshot(folderPath);
                    }
                }
                catch (StaleElementReferenceException)
                {
                    goto rerun;
                }

            }
        }

        public static void ListBoxSelectionXpathID(string ButtonXpath, string Id, string xpathSendKeys, string valueToSearch)
        {
            rerun:;
            waitForPageUntilElementVisible(By.XPath(ButtonXpath), 10);
            // screenshot(folderPath);
            IWebElement cli = SeleniumFramework.driver.FindElement(By.XPath(ButtonXpath));
            Actions action = new Actions(driver);
            action.MoveToElement(cli).Click().Perform();
            SeleniumFramework.driver.FindElement(By.XPath(xpathSendKeys)).SendKeys(valueToSearch);
            // screenshot(folderPath);
            waitForPageUntilElementVisible(By.Id(Id), 10);
            IWebElement items = SeleniumFramework.driver.FindElement(By.Id(Id));
            IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.TagName("li")));
            for (int i = 1; i < liValues.Count; i++)
            {
                try
                {
                    string indexValue = (liValues[i].GetAttribute("data-offset-index"));
                    string listText = SeleniumFramework.driver.FindElement(By.XPath("//li[@data-offset-index='" + indexValue + "']")).Text;
                    if (listText.Equals(valueToSearch))
                    {
                        driver.ExecuteJavaScript<object>("arguments[0].click();", liValues[i]);
                    }
                }
                catch (StaleElementReferenceException)
                {
                    goto rerun;
                }

            }
        }

        public static void ListBoxSelectionXpathIDWithOutSendKeys(string ButtonXpath, string Id, string valueToSearch)
        {
            SeleniumFramework.waitForPageUntilElementVisible(By.XPath(ButtonXpath), 10);
            // screenshot(folderPath);
            rerun:;
            IWebElement cli = SeleniumFramework.driver.FindElement(By.XPath(ButtonXpath));
            try
            {
                SeleniumFramework.driver.ExecuteJavaScript<object>("arguments[0].click();", cli);
                // screenshot(folderPath);
                SeleniumFramework.waitForPageUntilElementVisible((By.TagName("li")), 50);
            }
            catch (StaleElementReferenceException)
            {
                goto rerun;
            }
            //IWebElement items = MyMscControls.driver.FindElement(By.Id(Id));
            IWebElement items = driver.FindElement(By.XPath("//*[@id='" + Id + "']"));
            IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.TagName("li")));
            string Xpathlocator;
            for (int i = 0; i < liValues.Count; i++)
            {
                string className = liValues[i].GetAttribute("innerText");

                if (className.Equals(valueToSearch))
                {
                    try
                    {
                        //IWebElement clicakblePoint = MyMscControls.driver.FindElement(By.XPath("//*[contains(text()," + "\"" + valueToSearch + "\")]"));
                        Xpathlocator = "//*[@id='" + Id + "']/child::li[text()=" + "\"" + valueToSearch + "\"]";
                        IWebElement clicakblePoint = driver.FindElement(By.XPath(Xpathlocator));
                        bool chceck = clicakblePoint.Displayed;
                        if (chceck == false)
                        {
                            driver.ExecuteJavaScript<object>("arguments[0].click();", liValues[i]);
                            // liValues[i].Click();
                            SeleniumFramework.HighlightElement(liValues[i]);
                            // screenshot(folderPath);
                            goto end;
                        }
                        else
                        {
                            SeleniumFramework.HighlightElement(clicakblePoint);
                            driver.ExecuteJavaScript<object>("arguments[0].click();", clicakblePoint);
                            // screenshot(folderPath);
                            goto end;
                        }
                    }
                    catch (StaleElementReferenceException)
                    {
                        //IWebElement clicakblePoint = MyMscControls.driver.FindElement(By.XPath("//*[contains(text(),'" + valueToSearch + "')]"));
                        IWebElement clicakblePoint = driver.FindElement(By.XPath("//*[@id='" + Id + "']/child::li[text()='" + valueToSearch + "']"));
                        SeleniumFramework.HighlightElement(clicakblePoint);
                        driver.ExecuteJavaScript<object>("arguments[0].click();", clicakblePoint);
                        // screenshot(folderPath);
                        goto end;
                    }
                }
            }
            end:;
        }

       

        public static void SelectListWithOutSendKeys(string ButtonXpath, string Id, string valueToSearch)
        {
            SeleniumFramework.waitForPageUntilElementVisible(By.XPath(ButtonXpath), 10);
            rerun:;
            IWebElement cli = SeleniumFramework.driver.FindElement(By.XPath(ButtonXpath));
            try
            {
                SeleniumFramework.driver.ExecuteJavaScript<object>("arguments[0].click();", cli);
                SeleniumFramework.waitForPageUntilElementVisible((By.XPath("//*[@id='" + Id + "']/child::li")), 50);
            }
            catch (StaleElementReferenceException)
            {
                goto rerun;
            }
            IWebElement items = SeleniumFramework.driver.FindElement(By.Id(Id));
            IList<IWebElement> liValues = new List<IWebElement>(items.FindElements(By.XPath("//*[@id='" + Id + "']/child::li")));

            for (int i = 0; i < liValues.Count; i++)
            {
                string className = liValues[i].GetAttribute("innerText");

                if (className.Equals(valueToSearch))
                {
                    try
                    {
                        IWebElement clicakblePoint = SeleniumFramework.driver.FindElement(By.XPath("//*[contains(text()," + "\"" + valueToSearch + "\")]"));
                        bool chceck = clicakblePoint.Displayed;
                        if (chceck == false)
                        {
                            SeleniumFramework.driver.ExecuteJavaScript<object>("arguments[0].click();", liValues[i]);
                            // liValues[i].Click();
                            SeleniumFramework.HighlightElement(liValues[i]);
                            goto end;
                        }
                        else
                        {
                            i++;
                            IWebElement childElemenetClick = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + Id + "']/child::li[" + i + "]"));
                            SeleniumFramework.HighlightElement(childElemenetClick);
                            SeleniumFramework.driver.ExecuteJavaScript<object>("arguments[0].click();", childElemenetClick);
                            goto end;
                        }
                    }
                    catch (StaleElementReferenceException)
                    {
                        IWebElement clicakblePoint = SeleniumFramework.driver.FindElement(By.XPath("//*[contains(text(),'" + valueToSearch + "')]"));
                        SeleniumFramework.HighlightElement(clicakblePoint);
                        SeleniumFramework.driver.ExecuteJavaScript<object>("arguments[0].click();", clicakblePoint);
                        goto end;
                    }
                }
            }
            end:;
        }


        public static void ListBoxSelectionXpathIDByIndex(string ButtonXpath, string Id, string xpath)
        {
            driver.FindElement(By.XPath(ButtonXpath)).Click();
            Thread.Sleep(1000);
            IWebElement listBox = SeleniumFramework.driver.FindElement(By.Id(Id));
            Thread.Sleep(1000);
            driver.FindElement(By.XPath(xpath)).Click();
            // screenshot(folderPath);
        }

        public static string GenerateAlphabets(int length)
        {
            Random random = new Random();
            string characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            StringBuilder sb = new StringBuilder(length);
            for (int i = 0; i < length; i++)
            {
                sb.Append(characters[random.Next(characters.Length)]);
            }
            return sb.ToString();
        }

        public static string FolderCreation(string path,string screenName)
        {
            string foldername = SeleniumFramework.GetTimestamp(DateTime.Now);
            //folderPath = @"D:\ScreenStatus\" + screenName + "\\";
            string filepath = @path + screenName + "\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            return folderPath;
        }

        public static void SetMultipleComboboxbyXpath(string XPath, string valuetoserach)
        {
            waitForPageUntilElementVisible(By.XPath(XPath), 10);
            IWebElement TextboxName = SeleniumFramework.driver.FindElement(By.XPath(XPath));
            TextboxName.Click();
            Thread.Sleep(2000);
            TextboxName.SendKeys(valuetoserach);
            Thread.Sleep(2000);
            TextboxName.SendKeys(Keys.ArrowDown);
            Thread.Sleep(2000);
            TextboxName.SendKeys(Keys.Enter);
        }

        
        public static void ClickDateUsingCalendar(string Xpath, string caldisplayxpathid, string targetdate)
        {
            SeleniumFramework.waitForPageUntilElementVisible(By.XPath(Xpath), 50);
            SetClickButton("xpath",Xpath);

            //splited Target Date "dd-MM-YYYY"
            string[] expectdatesplit = targetdate.Split('-');
            string expecteddate = expectdatesplit[0];
            string expectedmonth = expectdatesplit[1];
            string expectedyear = expectdatesplit[2];

            Console.WriteLine("Expect date :{0} and {1} and {2}", expecteddate, expectedmonth, expectedyear);
            string nextMonth = DateTime.Now.AddMonths(0).ToString("MMMM yyyy");

            //To find Current Year of Lable
            SeleniumFramework.waitForPageUntilElementVisible(By.XPath("//*[@id='" + caldisplayxpathid + "']//child::div"), 50);
            string DynamicGetid = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + caldisplayxpathid + "']//child::div")).GetAttribute("id");
            Console.WriteLine("OUTPUT :{0}", DynamicGetid);


            SeleniumFramework.HighlightElement(SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + DynamicGetid + "']//child::a[@aria-live='assertive']")));
            IWebElement lable = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + DynamicGetid + "']//child::a[@aria-live='assertive']"));
            string lableval = lable.GetAttribute("innerText");
            Console.WriteLine("lableval :{0}", lableval);

            DateTime datevalue = (Convert.ToDateTime(lableval.ToString()));

            string currmonth = datevalue.Month.ToString();
            string curryear = datevalue.Year.ToString();

            //splited lable month and year Date "MM/yyyy"
            /*string lablevalue = DateTime.Parse(lableval).ToString("MM/yyyy");
              string[] lablesplit = lablevalue.Split('/');
              string currmonth = lablesplit[0];
              string curryear = lablesplit[1];*/

            Console.WriteLine("curr month :{0}", currmonth);
            //Year Click
            YearClick(curryear, expectedyear, DynamicGetid, lable);
            //Month Click
            MonthClick(expectedmonth, expectedyear);
            //Date Click
            DateClick(expecteddate, expectedmonth, expectedyear);

        }

        public static void YearClick(string CurrYear, string ExpectedYear, string DynamicGetid, IWebElement Findyear)
        {

            if (int.Parse(CurrYear) == int.Parse(ExpectedYear))
            {
                Console.WriteLine("year equal");
                ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Findyear);
            }
            else
            {
                ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Findyear);
                ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Findyear);

            reprocess:;
                string lableval1 = Findyear.GetAttribute("innerText");
                Console.WriteLine("lable 1 :{0}", lableval1);

                string[] lablesplit1 = lableval1.Split('-');
                string startyear = lablesplit1[0];
                string endyear = lablesplit1[1];

                if (int.Parse(ExpectedYear) >= int.Parse(startyear) && int.Parse(ExpectedYear) <= int.Parse(endyear))
                {
                    IWebElement yeareleidentify = SeleniumFramework.driver.FindElement(By.XPath("//a[@data-value='" + ExpectedYear + "/0/1']"));
                    string yeareleidentifytext = yeareleidentify.Text;

                    if (int.Parse(yeareleidentifytext).Equals(int.Parse(ExpectedYear)))
                    {
                        yeareleidentify.Click();
                    }

                }
                else
                {
                    if (int.Parse(CurrYear) < int.Parse(ExpectedYear))
                    {
                        IWebElement nextarrow = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + DynamicGetid + "']//child::a[@aria-label='Next']"));
                        ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", nextarrow);

                    }
                    else
                    {
                        IWebElement nextarrow = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + DynamicGetid + "']//child::a[@aria-label='Previous']"));
                        ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", nextarrow);
                    }
                    goto reprocess;

                }

            }

        }


        public static void MonthClick(string ExpMonth, string ExpYear)
        {
            if (ExpMonth.Substring(0, 1) == "0")
            {
                string TrimExpMonth = ExpMonth.Replace("0", "");
                int ReplaceMonth = int.Parse(TrimExpMonth) - 1;
                IWebElement Montheleidentify = SeleniumFramework.driver.FindElement(By.XPath("//a[@data-value='" + ExpYear + "/" + ReplaceMonth + "/1']"));
                string Montheleidentifytext = Montheleidentify.Text;
                //Montheleidentify.Click();
                ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Montheleidentify);

            }
            else
            {
                int ReplaceMonth = int.Parse(ExpMonth) - 1;
                IWebElement Montheleidentify = SeleniumFramework.driver.FindElement(By.XPath("//a[@data-value='" + ExpYear + "/" + ReplaceMonth + "/1']"));
                string Montheleidentifytext = Montheleidentify.Text;
                //Montheleidentify.Click();
                ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Montheleidentify);
            }

        }

        public static void DateClick(string ExpDate, string ExpMonth, string ExpYear)
        {
            string ClickableDate;
            if (ExpDate.Substring(0, 1) == "0")
            {
                string TrimExpdate = ExpDate.Replace("0", "");
                int ReplaceMonth = int.Parse(ExpMonth) - 1;
                Console.WriteLine("final date click : {0} / {1} / {2}", ExpYear, ReplaceMonth, TrimExpdate);
                ClickableDate = ExpYear + "/" + ReplaceMonth + "/" + TrimExpdate;
            }
            else
            {
                string TrimExpdate = ExpDate.Replace("0", "");
                int ReplaceMonth = int.Parse(ExpMonth) - 1;
                Console.WriteLine("final date click : {0} / {1} / {2}", ExpYear, ReplaceMonth, ExpDate);
                ClickableDate = ExpYear + "/" + ReplaceMonth + "/" + ExpDate;
            }
            Console.WriteLine("final date : {0}", ClickableDate);
            IWebElement Dateeleidentify = SeleniumFramework.driver.FindElement(By.XPath("//a[@data-value='" + ClickableDate + "']"));
            string Dateeleidentifytext = Dateeleidentify.Text;
            ((IJavaScriptExecutor)SeleniumFramework.driver).ExecuteScript("arguments[0].click();", Dateeleidentify);
        }


        public static void GetSpecificListBoxSelectionXpathIDWithSendKeys(string TextFieldiD, string Id, string valueToSearch)
        {
            IWebElement cli = SeleniumFramework.driver.FindElement(By.Id(TextFieldiD));
            int charcount = 0;
            foreach (char val in valueToSearch)
            {
                cli.SendKeys(val.ToString());
                charcount++;

                if (charcount > 3)
                {
                    IWebElement elementVisible = SeleniumFramework.waitForPageUntilElementVisible((By.TagName("li")), 80);
                    break;
                }
            }
            SeleniumFramework.waitForPageUntilElementVisible(By.XPath("//*[@id='"+ Id + "']/child::li"), 80);
            List<IWebElement> POLList = new List<IWebElement>(SeleniumFramework.driver.FindElements(By.XPath("//*[@id='" + Id + "']/child::li")));
            int count = 1;
            foreach (var BookingList in POLList)
            {
                string innertext = BookingList.GetAttribute("innerText");
                int posA = innertext.IndexOf("-");

                string specificvalue = innertext.Substring(0, posA);
                Console.WriteLine(specificvalue);
                if (specificvalue.Equals(valueToSearch))
                {
                    SeleniumFramework.waitForPageUntilElementVisible(By.XPath("//*[@id='" + Id + "']/child::li[" + count + "]"), 50);
                    IWebElement POL = SeleniumFramework.driver.FindElement(By.XPath("//*[@id='" + Id + "']/child::li[" + count + "]"));
                    SeleniumFramework.ClickHiddenElementByXpath("//*[@id='" + Id + "']/child::li[" + count + "]");
                    break;
                }
                count++;
            }

        }




        public IWebElement SelectLocator(string Locator, string LocatorValue)
        {
            switch (Locator)
            {
                case "id":
                    IWebElement id = driver.FindElement(By.Id(LocatorValue));
                    return id;

                case "XPath":
                    IWebElement XPath = driver.FindElement(By.XPath(LocatorValue));
                    return XPath;

                case "CssSelector":
                    IWebElement CssSelector = driver.FindElement(By.CssSelector(LocatorValue));
                    return CssSelector;

                case "PartialLinkText":
                    IWebElement PartialLinkText = driver.FindElement(By.PartialLinkText(LocatorValue));
                    return PartialLinkText;

                case "LinkText":
                    IWebElement LinkText = driver.FindElement(By.LinkText(LocatorValue));
                    return LinkText;

                case "ClassName":
                    IWebElement ClassName = driver.FindElement(By.ClassName(LocatorValue));
                    return ClassName;

                case "TagName":
                    IWebElement TagName = driver.FindElement(By.TagName(LocatorValue));
                    return TagName;

                default:
                    return null;
            }
        }

        public static string ReadText(string fileName,string SearchText)
        {
            var lastLine = File.ReadLines(fileName).Count();
            StreamReader sr = new StreamReader(fileName);
            string readTxt=null;
            for (int i = 1; i < lastLine; i++)
            {
                 readTxt= sr.ReadLine();
                if (readTxt.Contains("NAD"))
                {
                    switch (SearchText)
                    {
                        case "Forwarder":
                            if (readTxt.Contains("FW"))
                            {
                                return readTxt;
                            }
                            break;
                        case "Notify":
                            if (readTxt.Contains("NI"))
                            {
                                return readTxt;
                            }
                            break;

                        case "Second Notify":
                            if (readTxt.Contains("N1"))
                            {
                                return readTxt;
                            }
                            break;
                        case "Invoicing Company":
                            if (readTxt.Contains("FP"))
                            {
                                return readTxt;
                            }
                            break;
                        case "Consignee":
                            if (readTxt.Contains("CN"))
                            {
                                return readTxt;
                            }
                            break;
                        case "Booking Party":
                            if (readTxt.Contains("CZ"))
                            {
                                return readTxt;
                            }
                            break;
                        default:
                            Console.WriteLine("Parties Not Avaialbale");
                            break;


                    }
                }
            }
            return null;
        }

        public static void FileDeleteFromFolder(string Path,string Filename)
        {
            string[] FilePath = Directory.GetFiles(Path).Where(P => P.Contains(Filename)).ToArray();

            if (FilePath.Count() > 0)
            {
                for (int i = 0; i < FilePath.Count(); i++)
                {
                    if ((System.IO.File.Exists(FilePath[i])))
                    {
                        System.IO.File.Delete(FilePath[i]);
                    }
                }
            }
        }

        public static string PDFreadandgettext(string PDFFilePath, string Findingtext)
        {
            PDDocument doc = PDDocument.load(PDFFilePath);
            PDFTextStripper stripper = new PDFTextStripper();
            string pdfTExt = stripper.getText(doc);

            int initialstarts = pdfTExt.IndexOf(Findingtext);
            //int initialend = pdfTExt.IndexOf(EndingName, initialstarts);
            //string src = pdfTExt.Substring(initialstarts, (initialend - initialstarts));
            string Pdftextvalue = pdfTExt.Substring(initialstarts, Findingtext.Length);
            /*string[] textSplit = src.Split(':');

            string s = textSplit[1].Replace("\r\n", "");*/
            return Pdftextvalue;
        }

        public static void UploadFile(string Identificate/*ex : XPath, Id, LinkText,..*/, string Locater,string FilePath)
        {
            switch (Identificate.ToLower())
            {
                case "id":
                    SeleniumFramework.driver.FindElement(By.Id(Locater)).Click();
                    break;

                case "xpath":
                    SeleniumFramework.driver.FindElement(By.XPath(Locater)).Click();
                    break;

                case "cssselector":
                    SeleniumFramework.driver.FindElement(By.CssSelector(Locater)).Click();
                    break;

                case "partiallinktext":
                    SeleniumFramework.driver.FindElement(By.PartialLinkText(Locater)).Click();
                    break;

                case "linktext":
                    SeleniumFramework.driver.FindElement(By.LinkText(Locater)).Click();
                    break;

                case "classname":
                    SeleniumFramework.driver.FindElement(By.ClassName(Locater)).Click();
                    break;

                case "tagname":
                    SeleniumFramework.driver.FindElement(By.TagName(Locater)).Click();
                    break;

            }

            System.Threading.Thread.Sleep(2000);
            SeleniumFramework.SetFileUploadTextBox("Open", "Edit1", FilePath);
            SeleniumFramework.SetFileUploadButton("Open", "Button1");

        }




        public static int GetTablesRowCount(string Xpath)
        {
            int Count = SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.XPath("tr")).Count();
            return Count;
        }

        public static int GetTablesRowIndex(string Xpath, string ContainsValue)
        {

            int RowIndex = -1;
            int NoOfPages = 0;
            int PageCount = 0;
            string CheckMorePage = "";

            IList<IWebElement> TablesTrElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr"))).Where(p => !string.IsNullOrEmpty(p.Text)).ToList();

            RowIndex = TablesTrElement.IndexOf(TablesTrElement.FirstOrDefault(p => p.Text.Contains(ContainsValue)));

        Rerun:;
            if (RowIndex < 0)
            {
                NoOfPages = SeleniumFramework.driver.FindElement(By.XPath(Xpath + "/following::ul[contains(@class,'page')]")).FindElements(By.TagName("li")).Count;
                PageCount = NoOfPages - 1;
                CheckMorePage = SeleniumFramework.driver.FindElement(By.XPath(Xpath + "/following::ul[contains(@class,'page')]/li[" + NoOfPages + "]/a")).GetAttribute("title");
            }

            if (PageCount > 1)
            {
                for (int j = 1; j < PageCount; j++)
                {
                    SeleniumFramework.MouseHoverAction("Xpath",Xpath + "/following::a[@aria-label='Go to the next page']");
                    TablesTrElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr")));
                    RowIndex = TablesTrElement.IndexOf(TablesTrElement.FirstOrDefault(p => p.Text.Contains(ContainsValue)));
                    if (RowIndex >= 0)
                    {
                        break;
                    }
                }
            }

            if (RowIndex < 0)
            {
                if (CheckMorePage.Equals("More pages"))
                {
                    goto Rerun;
                }

            }

            return RowIndex;

        }

        public static int GetTablesRowIndex_OLD(string Xpath, string ContainsValue)
        {
            /*int RowIndex = -1;

            IList<IWebElement> TablesTrElement = new List<IWebElement>(MyMscControls.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr")));

            RowIndex = TablesTrElement.IndexOf(TablesTrElement.FirstOrDefault(p => p.Text.Contains(ContainsValue)));

            return RowIndex;*/

            int RowIndex = -1;
            int PageCount = 0;
            //IList<IWebElement> TablesTrElement = new List<IWebElement>(MyMscControls.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr")));
            IList<IWebElement> TablesTrElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr"))).Where(p => !string.IsNullOrEmpty(p.Text)).ToList();

            RowIndex = TablesTrElement.IndexOf(TablesTrElement.FirstOrDefault(p => p.Text.Contains(ContainsValue)));

            if (RowIndex < 0)
            {
                PageCount = (SeleniumFramework.driver.FindElement(By.XPath(Xpath+ "/following::ul[contains(@class,'page')]")).FindElements(By.TagName("li")).Count) - 1;
            }

            if (PageCount > 1 )
            {
                for (int j = 1; j < PageCount; j++)
                {
                    SeleniumFramework.MouseHoverAction("Xpath",Xpath + "/following::a[@aria-label='Go to the next page']");
                    TablesTrElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr")));
                    RowIndex = TablesTrElement.IndexOf(TablesTrElement.FirstOrDefault(p => p.Text.Contains(ContainsValue)));
                    if (RowIndex >= 0)
                    {
                        break;
                    }
                }
            }

            return RowIndex;

        }

        public static int GetTablesCellIndex(string Xpath, int RowIndex, string Text)
        {
            int CellIndex = -1;

            IList<IWebElement> TablesTrElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("tr")));

            IList<IWebElement> TablesTdElement = new List<IWebElement>(TablesTrElement[RowIndex].FindElements(By.TagName("td")));

            CellIndex = TablesTdElement.IndexOf(TablesTdElement.FirstOrDefault(s => s.Text.Equals(Text)));

            return CellIndex;
        }

        public static void ClickTablesButton(string Xpath, string tagname, string Getattributename, string ButtonName)
        {
           IList<IWebElement> TablesTdElement = new List<IWebElement>(SeleniumFramework.driver.FindElement(By.XPath(Xpath)).FindElements(By.TagName("td")));

            IList<IWebElement> ButtonsElement = TablesTdElement.Where(t => t.FindElements(By.TagName(tagname)).Count > 0).ToList();
            
            try
            {
                int buttonindex = ButtonsElement.IndexOf(ButtonsElement.FirstOrDefault(y => y.FindElement(By.TagName(tagname)).GetAttribute(Getattributename).Equals(ButtonName)));

                IWebElement ButtonElement = ButtonsElement.FirstOrDefault(y => y.FindElement(By.TagName(tagname)).GetAttribute(Getattributename).Equals(ButtonName));
                ButtonElement.Click();
            }
            catch(NullReferenceException)
            {
                int buttonindex = ButtonsElement.IndexOf(ButtonsElement.FirstOrDefault(y => y.FindElement(By.TagName(tagname)).GetAttribute(Getattributename).Contains(ButtonName)));

                IWebElement ButtonElement = ButtonsElement.FirstOrDefault(y => y.FindElement(By.TagName(tagname)).GetAttribute(Getattributename).Contains(ButtonName));
                ButtonElement.Click();
                //((IJavaScriptExecutor)MyMscControls.driver).ExecuteScript("arguments[0].click();", ButtonElement);
            }

        }

        public static void WaitTillDialogWindowClosed(string id)
        {
            bool dialogpage = true;
            while (dialogpage == true)
            {
                dialogpage = SeleniumFramework.driver.FindElement(By.Id(id)).Displayed;
            }
        }

        public static void SetFileUploadTextBox(string title, string classInstance, string filepath)
        {
            AutoItX.ControlFocus(title, "", classInstance);
            AutoItX.ControlSetText(title, "", classInstance, @filepath);
        }
        public static void SetFileUploadButton(string title, string classInstance)
        {
            AutoItX.ControlClick(title, "", classInstance);
        }
        public static void SetTextboxValuebyByXpathWithOutTab(string XpathValue, string value)
        {
            waitForPageUntilElementVisible(By.XPath(XpathValue), 10);
            // screenshot(folderPath);
            IWebElement textBoxName = driver.FindElement(By.XPath(XpathValue));
            textBoxName.Clear();
            textBoxName.SendKeys(value);
            textBoxName.Click();
            string alertValue = isAlertPresent(value);
            if (alertValue != null)
            {
                textBoxName.Clear();
                textBoxName.SendKeys(alertValue + Keys.Tab);
            }

        }
        public static IWebElement waitForPageUntilElementVisible(By Locator, int maxseconds)
        {
            new WebDriverWait(SeleniumFramework.driver, TimeSpan.FromSeconds(50)).Until(
            d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
        rerun:;
            try
            {
                return new WebDriverWait(driver, TimeSpan.FromSeconds(maxseconds)).Until(ExpectedConditions.ElementToBeClickable(Locator));
            }
            catch (StaleElementReferenceException)
            {
                goto rerun;
            }

        }

    }
}

