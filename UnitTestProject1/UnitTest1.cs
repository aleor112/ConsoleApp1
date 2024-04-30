using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using System;
using System.Diagnostics;
using OpenQA.Selenium;
using NUnit.Framework;

public class Tests
{
    IWebDriver driver;

    [SetUp]
    public void Setup()
    {
        var chromeOptions = new ChromeOptions();
        chromeOptions.PageLoadStrategy = PageLoadStrategy.Default;

        driver = new ChromeDriver(chromeOptions);

        Trace.Listeners.Add(new ConsoleTraceListener());
    }

    [Test]
    public void Test1()
    {
        driver.Url = "https://trumpf.service-now.com/now/nav/ui/classic/params/target/task_list.do%3Fsysparm_query%3DstateNOT%2520IN-5%252C3%252C4%252C7%252C8%252C10%252C9%252C6%252C103%252C104%252C106%252C107%255EORstateIN-5%252C3%252C4%252C7%252C-4%252C-2%252C0%252C10%255Eassignment_groupLIKEMSP%2520-%2520adesso%2FInternalApplications%255EORassignment_groupLIKEMSP%2520-%2520adesso%2FLTS%255Eclosed_atISEMPTY%255Eshort_descriptionNOT%2520LIKEITD%3A%255EORDERBYpriority%255EORDERBYsys_created_on%26sysparm_first_row%3D1%26sysparm_view%3D";

        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

        var macroponent = driver.FindElement(By.TagName("macroponent-f51912f4c700201072b211d4d8c26010"));
        var macroponentRoot = macroponent.GetShadowRoot();
        var snCanvasAppshellRoot = macroponentRoot.FindElement(By.TagName("sn-canvas-appshell-root"));

        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(macroponent.GetAttribute("app-id"), "a84adaf4c700201072b211d4d8c260b7");
    }

    [TearDown]
    public void CloseBrowser()
    {
        driver.Quit();
        Trace.Flush();
    }
}