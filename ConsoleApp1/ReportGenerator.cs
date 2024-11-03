using Aspose.Cells;
using Aspose.Cells.Timelines;
using log4net;
using log4net.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TestProject1;
using static System.Net.Mime.MediaTypeNames;

namespace ConsoleApp1
{
    public class ReportGenerator
    {
        ChromeDriver driver;
        bool shouldAddAssignmentTime = false;
        bool shouldAddResolvedTime = false;
        bool shouldAddReactionTime = false;
        bool shouldUpdateReacord = false;
        bool shouldAddClosedTime = false;
        bool hasUserChanged = false;
        bool shouldAddForwadedTime = false;
        string urlWithPlaceholder = "";
        int ticketsProcessed = 0;
        List<Record> _records;
        List<string> processedTicketNumbers = new List<string>();
        ILog _log;
        List<string> exsitingIds = new List<string>();
        public ReportGenerator(ILog log)
        {
            Setup();
            _log = log;
        }

        public void GenerateReport()
        {
            //var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC%5EORnumberSTARTSWITHRITM&sysparm_first_row=1&sysparm_view=";
            //var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC231110-0613384%5EORnumberSTARTSWITHINC230824-0578405&sysparm_first_row=1&sysparm_view=";
            //var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC231110-0613384%5EORnumberSTARTSWITHINC230824-0578405%5ENQnumberSTARTSWITHRITM231117-0565004&sysparm_first_row=1&sysparm_view=";
            //var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC%5EORnumberSTARTSWITHINC%5Esys_created_onONToday@javascript:gs.beginningOfToday()@javascript:gs.endOfToday()&sysparm_first_row=1&sysparm_view=";
            //var url = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC240126-0642703&sysparm_first_row=1&sysparm_view=";
            //urlWithPlaceholder = "https://trumpf.service-now.com/task_list.do?sysparm_query=stateNOT%20IN3%2C4%2C7%2C8%2C10%2C9%2C6%2C103%2C104%2C106%2C107%5Eassignment_groupLIKEMSP%20-%20adesso/InternalApplications%5Eshort_descriptionNOT%20LIKEITD%20&sysparm_first_row={0}&sysparm_view=";
            //one year ago urlWithPlaceholder = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC%5EORnumberSTARTSWITHRITM%5Esys_created_onONOne%20year%20ago@javascript:gs.beginningOfOneYearAgo()@javascript:gs.endOfOneYearAgo()&sysparm_first_row={0}&sysparm_view=%27";
            //last week urlWithPlaceholder = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC%5EORnumberSTARTSWITHRITM%5Esys_created_onONLast%20week@javascript:gs.beginningOfLastWeek()@javascript:gs.endOfLastWeek()&sysparm_first_row={0}&sysparm_view=";
            var startFromPosition = StateManager.ReadApplicationState();
            _log.Info("Retrieved row state " + startFromPosition);
            //Const.urlPlaceholder = "https://trumpf.service-now.com/task_list.do?sysparm_query=numberSTARTSWITHINC%5EORnumberSTARTSWITHRITM%5Esys_created_onONLast%20week@javascript:gs.beginningOfLastWeek()@javascript:gs.endOfLastWeek()&sysparm_first_row={0}&sysparm_view=";
            var url = "";
            if(startFromPosition != -1)
            {
                url = String.Format(Const.urlPlaceholder, startFromPosition);
            }
            else
            {
                url = String.Format(Const.urlPlaceholder, 1);
            }
            driver.Navigate().GoToUrl("https://trumpf.service-now.com/");

            Console.WriteLine("Log in with 2FA and press enter in console after that:");
            Console.ReadLine();
            //exsitingIds = ExcelHelper.GetIdsFromExcel("ticketReport_time_20240623_173045_fromRow_1.xlsx");
            //var exsitingIds2 = ExcelHelper.GetIdsFromExcel("ticketReport_time_20240623_212530_fromRow_1316.xlsx");
            //exsitingIds.AddRange(exsitingIds2);
            //OpenTabs(driver);
            //((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
            //List<string> tabs = new List<string>(driver.WindowHandles);
            //driver.SwitchTo().Window(tabs[1]);
            driver.Navigate().GoToUrl(url);
            
            //retrieve links
            List<SearchResultData> ticketUrlsList = new List<SearchResultData>();
            var ticketsCount = GetTicketUrls(ticketUrlsList);

            _log.Info("Last ticket row retrieved" + ticketsCount);

            GenerateRecords(ticketUrlsList);
            //generate excel from records
            ExcelHelper.GenerateExcel(_records, ticketsCount.ToString());
            StateManager.SaveApplicationState(ticketsCount);

        }

        public void OnExitProgram()
        {
            try
            {
                var startPosition = StateManager.ReadApplicationState();
                ExcelHelper.GenerateExcel(_records, (startPosition + ticketsProcessed).ToString());
                StateManager.SaveApplicationState(startPosition + ticketsProcessed);
            }
            catch(Exception ex)
            {
                _log.Error("Error while saving state on exit " + ex);
            }
        }

        private List<Record> GenerateRecords(List<SearchResultData> ticketUrlsList)
        {
            _records = new List<Record>();
            var counter = 0;

            // generate records
            do
            {
                try
                {
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //var nextUrls = urlsForTicketDetails.Skip((counter + 1) * Const.numberOfTabs).Take(Const.numberOfTabs).ToList();
                    var urls = ticketUrlsList
                        .Skip(counter * Const.numberOfTabs)
                        .Take(Const.numberOfTabs)
                        .Where(url => !processedTicketNumbers.Contains(url.TicketNumber))
                        .ToList();
                    LoadUrlsIntoTabs(urls);
                    for (var i = 0; i < urls.Count; i++)
                    {
                        var url = urls[i];
                        driver.SwitchTo().Window(driver.WindowHandles[urls.Count - i]);
                        try
                        {
                            GenerateRecordsFromUrl(url, _records);
                            ticketsProcessed++;
                        }
                        catch (Exception ex)
                        {
                            _log.Error($"Task url {url} error: " + ex);
                        }
                        driver.Close();
                    }
                }
                catch (Exception e)
                {
                    _log.Error($"Error while generating records {e}");
                }
                counter++;
            }
            while (counter * Const.numberOfTabs < ticketUrlsList.Count);


            driver.SwitchTo().Window(driver.WindowHandles[0]);
            return _records;
        }

        private void LoadUrlsIntoTabs(List<SearchResultData> urls)
        {
            for(int i = 0;i < urls.Count;i++)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript($"window.open('{urls[i].Url}', '_blank')");
            }
        }

        private void RedirectTab(string url)
        {
            ((IJavaScriptExecutor)driver).ExecuteAsyncScript($"window.location = '{url}'");
        }

        private void CloseTabs(List<string> urls)
        {
            for (int i = 1; i <= urls.Count; i++)
            {
                driver.SwitchTo().Window(driver.WindowHandles[i]);
                driver.Close();
            }
        }

        private int GetTicketUrls(List<SearchResultData> ticketUrlsList)
        {
            var totalRowsElement = driver.FindElement(By.CssSelector(".list_nav_bottom "));
            var totalRowsText = totalRowsElement.Text;
            var t = new string(totalRowsElement.Text.Trim().Split("\r\n")[0]
                .Where(e => char.IsDigit(e) || e == ' ')
                .ToArray())
                .Split(" ")
                .Where(e => !String.IsNullOrWhiteSpace(e))
                .ToList();


            _log.Info("Retrieve urls from list of tickets");
            int errorCounter = 0;
            int ticketRowsProcessed = int.Parse(t[0]);
            var totalTicketRows = int.Parse(t[2]);
            _log.Info("Total rows: " + totalTicketRows + " Start from: " + ticketRowsProcessed);
            var prevTotalRowsText = String.Empty;
            do
            {
                try
                {
                    totalRowsText = driver.FindElement(By.CssSelector(".list_nav_bottom ")).Text;

                    //unequal when the new page has loaded
                    if (prevTotalRowsText != totalRowsText)
                    {
                        prevTotalRowsText = totalRowsText;
                        var table = driver.FindElement(By.Id("metric_instance_table"));
                        var tableRowsHtml = table.FindElements(By.TagName("tr")).Skip(2);
                        GetTicketUrls(tableRowsHtml, ticketUrlsList);
                        if(ticketUrlsList.Count >= Const.BatchSizeTicketGenerate) 
                        {
                            var records = GenerateRecords(ticketUrlsList);
                            ExcelHelper.GenerateExcel(records, ticketRowsProcessed.ToString());
                            ticketRowsProcessed += ticketUrlsList.Count;
                            ticketUrlsList.Clear();

                            //var url = String.Format(Const.urlPlaceholder, ticketRowsProcessed);
                            //driver.Navigate().GoToUrl(url);
                        }
                        var nextButtonClick = driver.FindElements(By.CssSelector("button[name*='vcr_next']")).Last();
                        if (nextButtonClick.GetAttribute("disabled") == "true")
                        {
                            _log.Info("Next button is disabled. Breaking loop for retrieving ticket urls");
                            break;
                        }

                        nextButtonClick.Click();
                        Thread.Sleep(4000);
                        errorCounter = 0;
                    }
                    else
                    {
                        //the page hasn't reloaded yet, sleep
                        Thread.Sleep(2000);
                    }
                }
                catch (Exception e)
                {
                    errorCounter++;
                    _log.Error("Error while getting ticket urls from ServiceNow search: " + e + " Error count: " + errorCounter);
                    Thread.Sleep(3000);
                }
            }
            while (errorCounter < 5 && ticketRowsProcessed < totalTicketRows);
            return ticketRowsProcessed;
        }

        private void GenerateRecordsFromUrl(SearchResultData ticketUrl, List<Record> records)
        {
            var ticketIdElements = driver.FindElement(By.CssSelector("#sys_readonly\\.incident\\.number, #sys_readonly\\.sc_req_item\\.number"));
            var ticketId = ticketIdElements.GetDomAttribute("value");
            var serviceElement = driver.TryGetElement(By.CssSelector("#sys_display\\.sc_req_item\\.business_service, #sys_display\\.incident\\.business_service"));
            var service = serviceElement?.GetDomAttribute("value");
            var serviceOfferingElement = driver.TryGetElement(By.CssSelector("#sys_display\\.sc_req_item\\.service_offering, #sys_display\\.incident\\.service_offering"));
            var serviceOffering = serviceOfferingElement?.GetDomAttribute("value");
            var tbody = driver.FindElement(By.ClassName("activities-form"));
            var listItems = tbody.FindElements(By.TagName("li")).Reverse();
            Dictionary<string, string> importantInfoDict = new Dictionary<string, string>()
                {
                    {Const.AssignedTo,"" },
                    {Const.AssignmentGroup,"" },
                    {Const.AssignmentTime,"" },
                    {Const.PreviousAssignmentTime,"" },
                    {Const.PreviousAssignmentGroup,"" },
                    {Const.Priority,"" },
                    {Const.ReactionTime,"" },
                    {Const.ResolutionTime,"" },
                    {Const.Category,"" },
                    {Const.ForwardedAt, "" },
                    {Const.TicketId, ticketId },
                    {Const.Description, "" },
                    {Const.Application, "" },
                    {Const.TicketOpenedAt, "" },
                    {Const.TicketClosedAt, "" },
                    {Const.State, "" },
                    {Const.Service, service },
                    {Const.ServiceOffering, serviceOffering },
                    {Const.HasAdessoGroup, "" },
                    {Const.HasCreatedRecord, "" }
                };

            if (ticketId.StartsWith("RITM"))
            {
                try
                {
                    var descriptionElements = driver.TryGetElement(By.Id("sys_readonly.sc_req_item.description"));
                    importantInfoDict[Const.Description] = descriptionElements != null ? descriptionElements.Text : "";
                    importantInfoDict[Const.Category] = "Request";
                }
                catch (Exception e)
                {
                    _log.Error($"Error while getting description for {ticketId}" + e);
                }
            }
            if (ticketId.StartsWith("INC"))
            {
                var shortDescriptionElement = driver.TryGetElement(By.CssSelector("#incident\\.short_description"));
                var shortDescription = shortDescriptionElement?.GetDomAttribute("value");
                importantInfoDict[Const.Category] = "Incident";
                if (shortDescription != null && shortDescription.Contains("ITD:"))
                {
                    importantInfoDict[Const.Category] = "ITD";
                }
            }

            var record = new Record() 
            { 
                AllConnectedAdessoSupporters = "",
                TicketClosedAt = ticketUrl.EndTime, 
                TicketForwardedAt = ticketUrl.StartTime, 
                TicketOpenedAt = ticketUrl.CreatedAt,
                TicketNumber = ticketUrl.TicketNumber,
                TicketId = ticketUrl.TicketId
            };
            //go through list items to get necessary information to create a record
            foreach (var li in listItems)
            {
                try
                {
                    var text = li.Text;
                    if (text.Contains("Translate")) continue;

                    ExtractImportantInformationFromTicketFieldChanges(text, importantInfoDict);
                }
                catch (Exception ex)
                {
                    _log.Error("Error while getting information from a list item inside a ticket page: Error " + ex);
                }

                if (shouldUpdateReacord)
                {
                    UpdateRecordObject(importantInfoDict, record);
                    if (!String.IsNullOrWhiteSpace(record.TicketNumber))
                    {
                        importantInfoDict[Const.HasCreatedRecord] = "true";
                    }
                }
            }

            if(!String.IsNullOrWhiteSpace(record.TicketNumber))
            {
                processedTicketNumbers.Add(record.TicketNumber);
                records.Add(record);
            }

            if (!String.IsNullOrWhiteSpace(importantInfoDict[Const.HasAdessoGroup]) && String.IsNullOrWhiteSpace(importantInfoDict[Const.HasCreatedRecord]))
            {
                _log.Error($"Warning! Adesso group detected but record wasn't created! Ticket {ticketId}");
            }
        }


        private void UpdateRecordObject(Dictionary<string, string> importantInfoPairs, Record record)
        {
            shouldUpdateReacord = false;
            var currentAssignmentGroup = importantInfoPairs[Const.AssignmentGroup];
            var previousAssignmentGroup = importantInfoPairs[Const.PreviousAssignmentGroup];

            if (!Helper.IsGroupFromAdesso(currentAssignmentGroup) &&
                !Helper.IsGroupFromAdesso(previousAssignmentGroup))
            {
                return;
            }
            if (!String.IsNullOrWhiteSpace(record.TicketNumber))
            {
                record.AssignmentTime = importantInfoPairs[Const.AssignmentTime];
                record.Priority = importantInfoPairs[Const.Priority];
                record.ColleagueName = importantInfoPairs[Const.AssignedTo];
                record.TicketNumber = importantInfoPairs[Const.TicketId];
                record.Description = importantInfoPairs[Const.Description];
                record.TicketType = importantInfoPairs[Const.Category];
               // record.TicketOpenedAt = importantInfoPairs[Const.TicketOpenedAt];
                record.Service = importantInfoPairs[Const.Service];
                record.ServiceOffering  = importantInfoPairs[Const.ServiceOffering];
            }

            //if (String.IsNullOrWhiteSpace(record.TicketForwardedAt))
            //{
            //    record.TicketForwardedAt = importantInfoPairs[Const.ForwardedAt];
            //}

            if (String.IsNullOrWhiteSpace(record.ReactionTime))
            {
                record.ReactionTime = importantInfoPairs[Const.ReactionTime];
            }
            if (String.IsNullOrWhiteSpace(record.ResolutionTime))
            {
                record.ResolutionTime = importantInfoPairs[Const.ResolutionTime];
            }
            record.Status = importantInfoPairs[Const.State];
            //record.TicketClosedAt = importantInfoPairs[Const.TicketClosedAt];
            record.Application = importantInfoPairs[Const.Application];
            record.ColleagueName = importantInfoPairs[Const.AssignedTo];
            var index = record.ColleagueName.IndexOf("was");
            if (index != -1)
            {
                record.ColleagueName = record.ColleagueName.Substring(0, index);
            }
            record.CurrentAssignmentGroup = currentAssignmentGroup;
            if(record.ColleagueName != "[Empty]")
            {
                record.AddColleagueToConnectedSupportersList();
            }
            record.SetAssignmentGroupFlagsToTrueIfNecessary();

            //if (!importantInfoPairs[Const.State].ToLower().Contains("resolved") && !importantInfoPairs[Const.State].ToLower().Contains("complete"))
            //{
            //    record.CurrentAssignmentGroup = previousAssignmentGroup;
            //    record.AssignmentTime = importantInfoPairs[Const.PreviousAssignmentTime];
            //    record.ReactionTime = importantInfoPairs[Const.PreviousReactionTime];
            //}
            //if (String.IsNullOrWhiteSpace(record.CurrentAssignmentGroup) || String.IsNullOrWhiteSpace(record.AssignmentTime) ||
            //String.IsNullOrWhiteSpace(record.TicketType) || String.IsNullOrWhiteSpace(record.Priority) ||
            //String.IsNullOrWhiteSpace(record.ColleagueName) || String.IsNullOrWhiteSpace(record.ReactionTime) ||
            //String.IsNullOrWhiteSpace(record.ResolutionTime))
            //{
            //    return null;
            //}
            //return record;
        }

        private void ExtractImportantInformationFromTicketFieldChanges(string field, Dictionary<string, string> importantInfoDict)
        {

            if (field.Contains("Field changes"))
            {
                Regex regex = new Regex(@"\b\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b");
                var time = regex.Match(field).Value;
                if (String.IsNullOrEmpty(importantInfoDict[Const.TicketOpenedAt]))
                {
                    importantInfoDict[Const.TicketOpenedAt] = time;
                }
                if (shouldAddClosedTime)
                {
                    importantInfoDict[Const.TicketClosedAt] = time;
                    shouldAddClosedTime = false;
                }
                if (shouldAddAssignmentTime)
                {
                    importantInfoDict[Const.PreviousAssignmentTime] = importantInfoDict[Const.AssignmentTime];
                    importantInfoDict[Const.AssignmentTime] = time;
                    shouldAddAssignmentTime = false;
                }
                if (shouldAddResolvedTime)
                {
                    importantInfoDict[Const.ResolutionTime] = time;
                    shouldAddResolvedTime = false;
                    shouldUpdateReacord = true;
                }

                if (shouldAddForwadedTime)
                {
                    importantInfoDict[Const.ForwardedAt] = time;
                    shouldAddForwadedTime = false;
                }

                if (shouldAddReactionTime)
                {
                    if (String.IsNullOrWhiteSpace(importantInfoDict[Const.ReactionTime]))
                    {
                       importantInfoDict[Const.ReactionTime] = time;
                    }
                    shouldAddReactionTime = false;
                    shouldUpdateReacord = true;
                }
            }
            else if (field.Contains(Const.Priority))
            {
                var priorityText = field.Substring(field.LastIndexOf("-") + 1).Trim();
                importantInfoDict[Const.Priority] = priorityText;
            }
            else if (field.Contains(Const.Description))
            {
                importantInfoDict[Const.Description] = field;
                importantInfoDict[Const.Application] = GetApplicationKeywordFromDescription(field);
            }
            else if(field.Contains("State"))
            {
                var state = Helper.ClearRedundantPartOfString(field);
                importantInfoDict[Const.State] = state;
                if(state.Contains("Closed") || state.Contains("Resolved"))
                {
                    if (state.Contains("Closed"))
                    {
                        shouldAddClosedTime = true;
                    }
                    shouldAddResolvedTime = true;
                }
            }
            else if (field.Contains("Assignment group"))
            {
                var currentAssignmentGroup = Helper.ClearRedundantPartOfString(field);
                var previousAssignmentGroup = importantInfoDict[Const.AssignmentGroup];
                if (!Helper.IsGroupFromAdesso(previousAssignmentGroup))
                {
                    if (Helper.IsGroupFromAdesso(currentAssignmentGroup))
                    {
                        shouldAddForwadedTime = true;
                    }
                    shouldAddResolvedTime = true;
                }

                if (Helper.IsGroupFromAdesso(currentAssignmentGroup))
                {
                    importantInfoDict[Const.HasAdessoGroup] = "true";
                    shouldAddReactionTime = true;
                    shouldAddAssignmentTime = true;
                }

                importantInfoDict[Const.AssignmentGroup] = currentAssignmentGroup;
                importantInfoDict[Const.PreviousAssignmentGroup] = previousAssignmentGroup;
            }
            else if (field.Contains("Assigned to"))
            {
                var assignedTo = field.Substring(field.IndexOf("Assigned to") + 11).Trim();
                if (Helper.IsGroupFromAdesso(importantInfoDict[Const.AssignmentGroup])&& !field.StartsWith("[Empty]"))
                {
                    shouldAddReactionTime = true;
                    importantInfoDict[Const.AssignedTo] = assignedTo;
                }
            };
        }

        private string GetApplicationKeywordFromDescription(string description)
        {
            var descriptionLower = description.ToLower();
            foreach(var applicationName in Const.applicationWords)
            {
                if (descriptionLower.Contains(applicationName))
                {
                    return applicationName;
                }
            }

            return "";
        }

        private void GetTicketUrls(IEnumerable<IWebElement> trs, List<SearchResultData> urlList)
        {
            var t = String.Empty;
            foreach (var tr in trs)
            {
                try
                {
                    t = tr.Text;
                    var urlTd = tr.FindElements(By.TagName("td"))[4]; 
                    var incidentId = urlTd.Text.Replace("Incident:","").Trim();
                    if (!exsitingIds.Contains(incidentId))
                    {
                        File.AppendAllLines("unexisting.txt", new List<string>(){ incidentId });
                    }
                    else
                    {
                        continue;
                    }
                    var link = urlTd.FindElement(By.TagName("a"));
                    var taskUrl = link.GetAttribute("href");
                    var startDate = tr.FindElements(By.TagName("td"))[6].Text;
                    var endDate = tr.FindElements(By.TagName("td"))[7].Text;
                    var createdAt = tr.FindElements(By.TagName("td"))[2].Text;
                    urlList.Add(new SearchResultData() {TicketNumber = incidentId, Url = taskUrl, StartTime = startDate, EndTime = endDate, CreatedAt = createdAt});
                }
                catch (Exception ex)
                {
                    _log.Error("Didn't retrieve url. Exception " + ex + "Row text: " + t);
                }
            }
        }
        private void Setup()
        {
            var edgeOptions = new ChromeOptions();
            //string profileDirectory = @"C:\Users\astefanov\AppData\Local\Microsoft\Edge\User Data";

            //edgeOptions.AddArguments($"--user-data-dir={profileDirectory}");
            //edgeOptions.AddArguments($"--headless");
            edgeOptions.AddArguments("--profile-directory=Default");
            //edgeOptions.AddArgument("user-data-dir=" + profileDirectory);

            //var edgeOptions2 = new EdgeOptions();
            //edgeOptions.AddArgument("--user-data-dir=C:\\ChromeData");
            //edgeOptions.AddArgument("--remote-debugging-port=9222");
            //edgeOptions.DebuggerAddress = "127.0.0.1:5555";

            //  driver = new EdgeDriver(edgeOptions);
            driver = new ChromeDriver(edgeOptions);

            Trace.Listeners.Add(new ConsoleTraceListener());
        }

    }
}
