using Aspose.Cells;
using OpenQA.Selenium.DevTools.V118.Browser;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace TestProject1
{
    public class Tests
    {
        IWebDriver driver;
        bool shouldAddAssignmentTime = false;
        bool shouldAddResolvedTime = false;
        bool shouldAddReactionTime = false;
        [SetUp]
        public void Setup()
        {
            var edgeOptions = new EdgeOptions();
            edgeOptions.DebuggerAddress = "127.0.0.1:9222";
            //string driverPath = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe";

            // Create an instance of the ChromeDriver by providing the path to the existing executable
            driver = new EdgeDriver(edgeOptions);

            Trace.Listeners.Add(new ConsoleTraceListener());
        }

        [Test]
        public void Test1()
        {
            var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=state%20IN-5%2C3%2C4%2C7%2C8%2C10%2C9%2C6%2C103%2C104%2C106%2C107%5EORstateIN-5%2C3%2C4%2C7%2C-4%2C-2%2C0%2C10%5Eclosed_atISEMPTY%5Eshort_descriptionNOT%20LIKEITD:%5EORDERBYpriority%5EORDERBYsys_created_on&sysparm_first_row=1&sysparm_view=";
            driver.Navigate().GoToUrl(url);


            //IWebElement usernameField = driver.FindElement(By.Id("i0116"));
            //usernameField.SendKeys("aleksandar.stefanov.external@trumpf.com");

            //IWebElement loginButton = driver.FindElement(By.Id("idSIButton9"));
            //loginButton.Click();


            //IWebElement passwordField = driver.FindElement(By.Id("i0118"));
            //passwordField.SendKeys("TrumpIsTheBest!123");

            //loginButton = driver.FindElement(By.Id("idSIButton9"));
            //loginButton.Click();

            //IWebElement? shadowHost = driver.FindElement(By.CssSelector("#shadow_host"));
            //var shadowRoot = shadowHost.GetShadowRoot();

            //retrieve links
            List<string> urlsToTaksDetails = new List<string>();

            var nextButton = driver.FindElements(By.Name("vcr_next")).Last();

            do
            {

                var macroponent = driver.FindElement(By.Id("task_table"));
                var tbody = macroponent.FindElement(By.TagName("tbody"));
                var trs = tbody.FindElements(By.TagName("tr"));
                GetTaskItemsUrls(trs, urlsToTaksDetails);

                nextButton.Click();
            }
            while (nextButton.Enabled);

            List<Record> records = new List<Record>();

            // generate records
            foreach (var taskUrl in urlsToTaksDetails)
            {
                driver.Navigate().GoToUrl(taskUrl);
                var taskIdElements = driver.FindElement(By.CssSelector("#sys_readonly\\.incident\\.number, #sys_readonly\\.sc_req_item\\.number"));
                var ticketId = taskIdElements.GetDomAttribute("value");

                var tbody = driver.FindElement(By.ClassName("activities-form"));
                var listItems = tbody.FindElements(By.TagName("li")).Reverse();
                Dictionary<string, string> importantInfoPairs = new Dictionary<string, string>()
                {
                    {Const.AssignedTo,"" },
                    {Const.AssignmentGroup,"" },
                    {Const.AssignmentTime,"" },
                    {Const.PreviousAssignmentGroup,"" },
                    {Const.Priority,"" },
                    {Const.ReactionTime,"" },
                    {Const.ResolutionTime,"" },
                    {Const.Category,"" },
                    {Const.TicketId, ticketId }
                };
                foreach (var li in listItems)
                {
                    var text = li.Text;
                    if (!text.Contains("Field changes")) continue;

                    CheckForImportantInformationAndExtract(text, importantInfoPairs);

                    Record record = TryCreateRecord(importantInfoPairs);
                    if(record != null) records.Add(record);
                }
            }

            //generate excel from records
            GenerateExcel(records);


            //var macroponentRoot = macroponent.GetShadowRoot();
            //var snCanvasAppshellRoot = macroponentRoot.FindElement(By.TagName("sn-canvas-appshell-root"));

            //Assert.AreEqual(macroponent.GetAttribute("app-id"), "a84adaf4c700201072b211d4d8c260b7");
        }

        private void GenerateExcel(List<Record> records)
        {
            // Instantiate a new Workbook
            Workbook book = new Workbook();
            // Obtaining the reference of the worksheet
            Worksheet sheet = book.Worksheets[0];

            sheet.Cells.ImportCustomObjects((System.Collections.ICollection)records,
                new string[] { "TicketId", "ColleagueName" , "AssignmentTime", "ReactionTime", "ResolutionTime", "Priority", "Category", "AssignmentGroup" }, // propertyNames
                true, // isPropertyNameShown
                0, // firstRow
                0, // firstColumn
                records.Count, // Number of objects to be exported
                true, // insertRows
                null, // dateFormatString
                false); // convertStringToNumber

            // Save the Excel file
            book.Save("ExportedCustomObjects.xlsx");
        }

        private Record TryCreateRecord(Dictionary<string, string> importantInfoPairs)
        {
            if (importantInfoPairs[Const.AssignedTo] == "[Empty]" || !importantInfoPairs[Const.AssignmentGroup].Contains("adesso"))
            {
                return null;
            }
            var record = new Record()
            {
                AssignmentGroup = importantInfoPairs[Const.AssignmentGroup],
                AssignmentTime = importantInfoPairs[Const.AssignmentTime],
                Category = importantInfoPairs[Const.Category],
                Priority = importantInfoPairs[Const.Priority],
                ColleagueName = importantInfoPairs[Const.AssignedTo],
                ReactionTime = importantInfoPairs[Const.ReactionTime],
                ResolutionTime = importantInfoPairs[Const.ResolutionTime],
                TicketId = importantInfoPairs[Const.TicketId]
            };

            if(String.IsNullOrWhiteSpace(record.AssignmentGroup) || String.IsNullOrWhiteSpace(record.AssignmentTime) ||
            String.IsNullOrWhiteSpace(record.Category) || String.IsNullOrWhiteSpace(record.Priority) ||
            String.IsNullOrWhiteSpace(record.ColleagueName) || String.IsNullOrWhiteSpace(record.ReactionTime) ||
            String.IsNullOrWhiteSpace(record.ResolutionTime))
            {
                return null;
            }

            return record;
        }

        private void CheckForImportantInformationAndExtract(string text, Dictionary<string, string> dict)
        {
            KeyValuePair<string, string> lastpair = new KeyValuePair<string, string>(null, null);
            if(dict.Count > 0)
            {
                lastpair = dict.Last();
            }

            if (text.Contains("Field changes"))
            {
                Regex regex = new Regex(@"\b\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b");
                var time = regex.Match(text).Value;
                if (shouldAddAssignmentTime)
                {
                    dict[Const.AssignmentTime] = time;
                    shouldAddAssignmentTime = false;
                }
                if (shouldAddResolvedTime)
                {
                    dict[Const.ResolutionTime] = time;
                    shouldAddResolvedTime = false;
                }

                if (shouldAddReactionTime)
                {
                    dict[Const.ReactionTime] = time;
                    shouldAddReactionTime = false;
                }
            }
            else if (text.Contains(Const.Priority))
            {
                var priorityText = text.Substring(text.LastIndexOf("-") + 1).Trim();
                dict[Const.Priority] = priorityText;
            }
            else if (text.Contains(Const.AssignmentGroup))
            {
                var assignmentGroupText = text.Substring(text.IndexOf("\r\n") + 1).Trim();
                var previousGroup = dict[Const.AssignmentGroup];
                if (previousGroup.Contains("adesso"))
                {
                    shouldAddResolvedTime = true;
                }
                if (assignmentGroupText.Contains("adesso"))
                {
                    shouldAddAssignmentTime = true;
                }
                dict[Const.AssignmentGroup] = assignmentGroupText;
                dict[Const.PreviousAssignmentGroup] = previousGroup;
            }
            else if (text.Contains("Assigned to"))
            {
                var assignedTo = text.Substring(text.IndexOf("Assigned to") + 11).Trim();
                if (dict[Const.AssignmentGroup].Contains("adesso"))
                {
                    shouldAddReactionTime = true;
                }
                dict[Const.AssignedTo] = assignedTo;
            };
        }

        private void GetTaskItemsUrls(System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> trs,List<string> urlList)
        {
            foreach (var tr in trs)
            {
                string txt = string.Empty;
                try
                {
                    txt = tr.Text;
                    var firstTd = tr.FindElements(By.TagName("td"))[2];
                    var ss = firstTd.Text;
                    var link = firstTd.FindElement(By.TagName("a"));
                    var taskUrl = link.GetAttribute("href");
                    urlList.Add(taskUrl);
                }
                catch (Exception ex) 
                {
                    Console.WriteLine("Didn't retrieve url for row whit text " + txt + "Exception " + ex);
                }
                //open Task details in new tab
                //Actions actions = new Actions(driver);
                //actions.KeyDown(Keys.Control).Click(link).KeyUp(Keys.Control).Perform();

                //string originalWindowHandle = driver.CurrentWindowHandle;
                //foreach (string windowHandle in driver.WindowHandles)
                //{
                //    if (windowHandle != originalWindowHandle)
                //    {
                //        driver.SwitchTo().Window(windowHandle);
                //        break;
                //    }
                //}
            }
        }

        [TearDown]
        public void CloseBrowser()
        {
            driver.Quit();
            Trace.Flush();
        }
    }
}