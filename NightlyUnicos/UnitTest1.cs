using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace NightlyUnicos
{
    [TestClass]
    public class ReportGeneration
    {
        IWebDriver driverObject;

        [TestMethod]
        public void MasterBranchTest()
        {

            KillMethod();
            string date = System.DateTime.Today.ToString();
            string daysName = System.DateTime.Today.DayOfWeek.ToString();
            string[] todaysDate = date.Split(" ");
            ExcelClass excelObject = new ExcelClass();

            driverObject = new ChromeDriver();
            string stableBuildEntryInExcel = "Stable";
            string emptyValue = " ";
            int rowNumberOfFirstBuildInfoHyperLinkInExcel = 52;
            int columnNumberOfBuildInfoHyperLinkInExcel = 6;
            int rowNumberOfFirstStableBuildInfoToBeWritten = 54;
            int columnNumberForStableBuildWrite = 4;
            int rowNumberofFirstBuildForNumberOfFailuresWrite = 53;
            int columnNumberForNumberOfFailuresWrite = 4;
            string failureResultToBeWrittenInExcel;
            excelObject.OpenExcelMethod("C:\\Users\\INPRSRI4\\Desktop\\New folder (15)\\MasterBranchTestReport");
            Thread.Sleep(200);
            excelObject.SelectSheetNumber(1);
            Thread.Sleep(200);
            excelObject.RunMacro("Macro1");
            excelObject.ExcelWriteCell(51, 2, "Jenkins Stable Run Report -" + todaysDate[0] + "(" + (daysName) + ")");
            Thread.Sleep(200);
            excelObject.FindExcelHyperLink();
            IWebElement FindBuildIsStableTestResultLink;
            IWebElement driverToFindNumberOfFailures;
            IWebElement driverToFindNumberOfTestCasesExecuted;
            IWebElement FindFailuresName;
            List<string> ListOfFailuresName = new List<string>();
            List<string> NewListOfFailuresName = new List<string>();
            List<int> CountOfEveryUniqueFailures = new List<int>();
            string[] FindFWBuildNumberSecondSplit;
            string actualFWBuildNumberWebScrapped;
            string numberOfFailuresInBuild;

            excelObject.SelectSheetNumber(2);
            Thread.Sleep(500);
            string totalNumberOfBuildsString = excelObject.ExcelReadCell(23, 1).ToString();
            string[] totalNumberOfBuildsInString = totalNumberOfBuildsString.Split(" ");
            int totalNumberOfBuilds = int.Parse(totalNumberOfBuildsInString[1]);
            excelObject.SelectSheetNumber(3);
            Thread.Sleep(200);
            string userName = excelObject.ExcelReadCell(29, 7);
            string passWord = excelObject.ExcelReadCell(31, 7);
            excelObject.SelectSheetNumber(2);
            Thread.Sleep(200);
            for (int i = 1; i < 11; i++)
            {
                excelObject.SelectSheetNumber(2);
                Thread.Sleep(200);
                string urlOfNightlySetupBuildReadFromExcel = excelObject.ExcelReadCell(i, 1);
                Thread.Sleep(200);
                driverObject.Navigate().GoToUrl(urlOfNightlySetupBuildReadFromExcel);
               
                if (i == 1)
                {
                    IWebElement enterUserName = driverObject.FindElement(By.Name("j_username"));
                    enterUserName.SendKeys(userName);
                    IWebElement enterPassWord = driverObject.FindElement(By.Name("j_password"));
                    enterPassWord.SendKeys(passWord);
                    Thread.Sleep(500);
                    IWebElement ClickSubmit = driverObject.FindElement(By.Name("Submit"));
                    ClickSubmit.Click();
                }
                
                IWebElement webElementToFindBuildNumberFromUrl = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/h1[1]/span[2]"));
                string buildInformationTextReadFromWebElement = webElementToFindBuildNumberFromUrl.Text;
                string[] firstSplitOfTextOfBuildInformationReadFromWebElement = buildInformationTextReadFromWebElement.Split("#");
                string[] secondSplitOfTextOfBuildInformationReadFromWebElement = firstSplitOfTextOfBuildInformationReadFromWebElement[1].Split(" ");
                string lastSuccessfulBuildNumberInText = secondSplitOfTextOfBuildInformationReadFromWebElement[0];
                int lastSuccessfulBuildNumberInInteger = int.Parse(lastSuccessfulBuildNumberInText);

                string[] setupUrlToUseFromExcel = urlOfNightlySetupBuildReadFromExcel.Split("lastSuccessfulBuild");
                string currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInText;
                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                string urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger + 1).ToString();
                driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                Thread.Sleep(2000);
            loopback:
                var error404PresentCheckWebElement = driverObject.FindElements(By.XPath("/html/body/h2"));

                if (error404PresentCheckWebElement.Count == 0)
                {
                    lastSuccessfulBuildNumberInInteger++;
                    urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger + 1).ToString();
                    driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                    Thread.Sleep(500);
                    goto loopback;
                }

                else if (error404PresentCheckWebElement.Count > 0)
                {
                    urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger).ToString();
                    driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                    Thread.Sleep(1000);
                }

                currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInInteger.ToString();
                excelObject.SelectSheetNumber(1);
                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

            LoopBack1:
                IWebElement FindFWBuildNumberWebElement = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[2]/div[1]/div[1]"));

                string FindFWBuildNumberFirstSplit = FindFWBuildNumberWebElement.Text;
                if ((FindFWBuildNumberFirstSplit == "") || (FindFWBuildNumberFirstSplit=="Firmware Loading Error -aborting test"))
                {
                    actualFWBuildNumberWebScrapped = " ";
                }

                else
                {
                    FindFWBuildNumberSecondSplit = FindFWBuildNumberFirstSplit.Split("/");
                    string[] FindFWBuildNumberThirdSplit = FindFWBuildNumberSecondSplit[1].Split("\r");
                    actualFWBuildNumberWebScrapped = FindFWBuildNumberThirdSplit[0];
                }

                string[] hyperLinkName = currentUrlToBeChecked.Split("job/");
                hyperLinkName = hyperLinkName[4].Split("/");
                if (actualFWBuildNumberWebScrapped == "UNICOS_Turel_FW_Feature")
                {
                    lastSuccessfulBuildNumberInInteger = lastSuccessfulBuildNumberInInteger - 1;
                    currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInInteger.ToString();
                    driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                    goto LoopBack1;
                }
                string actualNameOfSetupHyperLinkToBeWrittenInExcel = hyperLinkName[0] + "-Build#" + lastSuccessfulBuildNumberInInteger;
                Thread.Sleep(1000);
                excelObject.EditExcelHyperLink((i), currentUrlToBeChecked, actualNameOfSetupHyperLinkToBeWrittenInExcel, totalNumberOfBuilds);

                if (FindFWBuildNumberFirstSplit != "")
                {
                    FindBuildIsStableTestResultLink = driverObject.FindElement(By.LinkText("Test Result"));
                    FindBuildIsStableTestResultLink.Click();
                    FindBuildIsStableTestResultLink = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[1]"));
                    numberOfFailuresInBuild = FindBuildIsStableTestResultLink.Text;
                    string[] listofbuildFailureBeforeStart = numberOfFailuresInBuild.Split(" ");
                    numberOfFailuresInBuild = listofbuildFailureBeforeStart[0];
                }

                else
                {
                    numberOfFailuresInBuild = " ";
                }

                if (numberOfFailuresInBuild.Contains("0"))
                {
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, "");
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, stableBuildEntryInExcel);
                }

                else
                {
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, emptyValue);
                }

                if (FindFWBuildNumberFirstSplit != "")
                {
                    string testReportUrl = currentUrlToBeChecked + "/testReport/";
                    driverObject.Navigate().GoToUrl(testReportUrl);
                    driverToFindNumberOfFailures = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[1]"));
                    string totalNumberOfFailures = driverToFindNumberOfFailures.Text;
                    string[] listTotalNumberOfFailures = totalNumberOfFailures.Split(" ");
                    driverToFindNumberOfTestCasesExecuted = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[3]"));
                    string totalNumberOfTestCases = driverToFindNumberOfTestCasesExecuted.Text;
                    string[] listTotalNumberOfTestCases = totalNumberOfTestCases.Split(" ");
                    failureResultToBeWrittenInExcel = ("Total failures- " + listTotalNumberOfFailures[0] + " failures out of " + listTotalNumberOfTestCases[0] + " tests");
                    excelObject.ExcelWriteCell(rowNumberofFirstBuildForNumberOfFailuresWrite, columnNumberForNumberOfFailuresWrite, failureResultToBeWrittenInExcel);
                    excelObject.ExcelWriteCell(rowNumberOfFirstBuildInfoHyperLinkInExcel, columnNumberOfBuildInfoHyperLinkInExcel, actualFWBuildNumberWebScrapped);
                    driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                    Thread.Sleep(200);
                    if ((int.Parse(listTotalNumberOfFailures[0]) > 0 && (int.Parse(listTotalNumberOfFailures[0])) < 1000))
                    {
                        testReportUrl = currentUrlToBeChecked + "/testReport/";
                        driverObject.Navigate().GoToUrl(testReportUrl);
                        string path = " ";
                        ListOfFailuresName.Clear();
                        for (int numberOfFailures = 0; numberOfFailures < int.Parse(listTotalNumberOfFailures[0]); numberOfFailures++)
                        {
                            path = "/html[1]/body[1]/div[3]/div[2]/table[2]/tbody[1]/" + "tr[" + (numberOfFailures + 1) + "]/td[1]/a[3]";
                            FindFailuresName = driverObject.FindElement(By.XPath(path));
                            string FailureNameText = FindFailuresName.Text;
                            if (FailureNameText.Contains("("))
                            {
                                string[] FailureNameChange = FailureNameText.Split("(");
                                FailureNameText = FailureNameChange[0];
                            }
                            ListOfFailuresName.Add(FailureNameText);

                        }

                        NewListOfFailuresName.Clear();
                        foreach (string e in ListOfFailuresName)
                        {
                            if (!(NewListOfFailuresName.Contains(e)))
                            {
                                NewListOfFailuresName.Add(e);
                            }
                        }

                        for (int jr = 0; jr < NewListOfFailuresName.Count; jr++)
                        {
                            string damt = NewListOfFailuresName[jr];
                            int count1 = 0;
                            for (int sr = 0; sr < ListOfFailuresName.Count; sr++)
                            {
                                if (ListOfFailuresName[sr] == damt)
                                {
                                    count1++;
                                }
                            }
                            CountOfEveryUniqueFailures.Add(count1);
                        }

                        driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

                        for (int failedCount = 0; failedCount < NewListOfFailuresName.Count; failedCount++)
                        {
                            excelObject.SelectSheetNumber(3);
                            Thread.Sleep(500);
                            excelObject.CopyRowFormat();
                            excelObject.SelectSheetNumber(1);
                            Thread.Sleep(500);
                            excelObject.InsertExcelRow(rowNumberOfFirstStableBuildInfoToBeWritten + 1);
                            excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, 2, NewListOfFailuresName[failedCount]);
                            excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, 3, (CountOfEveryUniqueFailures[failedCount]).ToString());
                            rowNumberOfFirstStableBuildInfoToBeWritten++;
                            rowNumberOfFirstBuildInfoHyperLinkInExcel++;
                            rowNumberofFirstBuildForNumberOfFailuresWrite++;
                        }
                    }

                }

                else
                {
                    excelObject.ExcelWriteCell(rowNumberofFirstBuildForNumberOfFailuresWrite, columnNumberForNumberOfFailuresWrite, "Total Failures:No Test Report found/Build is still running.So do manual analysis");
                    excelObject.ExcelWriteCell(rowNumberOfFirstBuildInfoHyperLinkInExcel, columnNumberOfBuildInfoHyperLinkInExcel, "No firmware build was found: check manually");
                }

                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

                if (i == 1)
                {
                    excelObject.ExcelWriteCell(50, 2, "UNICOS_Turel/" + actualFWBuildNumberWebScrapped);
                }
                rowNumberOfFirstBuildInfoHyperLinkInExcel = rowNumberOfFirstBuildInfoHyperLinkInExcel + 4;
                rowNumberOfFirstStableBuildInfoToBeWritten = rowNumberOfFirstStableBuildInfoToBeWritten + 4;
                rowNumberofFirstBuildForNumberOfFailuresWrite = rowNumberofFirstBuildForNumberOfFailuresWrite + 4;
            }
            driverObject.Close();
            driverObject.Quit();
            excelObject.SaveAsExcelFile("C:\\Users\\INPRSRI4\\Desktop\\New folder (15)\\MasterBranchTestReport");
            excelObject.CloseExcelMethod();
            KillMethod();
        }

        [TestMethod]
        public void FeatureBranchTest()
        {

            KillMethod();
            string date = System.DateTime.Today.ToString();
            string daysName = System.DateTime.Today.DayOfWeek.ToString();
            string[] todaysDate = date.Split(" ");
            ExcelClass excelObject = new ExcelClass();

            driverObject = new ChromeDriver();
            string stableBuildEntryInExcel = "Stable";
            string emptyValue = " ";
            int rowNumberOfFirstBuildInfoHyperLinkInExcel = 20;
            int columnNumberOfBuildInfoHyperLinkInExcel = 6;
            int rowNumberOfFirstStableBuildInfoToBeWritten = 22;
            int columnNumberForStableBuildWrite = 4;
            int rowNumberofFirstBuildForNumberOfFailuresWrite = 21;
            int columnNumberForNumberOfFailuresWrite = 4;
            string failureResultToBeWrittenInExcel;
            excelObject.OpenExcelMethod("C:\\Users\\INPRSRI4\\Desktop\\New folder (15)\\FeatureBranchTestReport");
            Thread.Sleep(200);
            excelObject.SelectSheetNumber(1);
            Thread.Sleep(200);
            excelObject.RunMacro("Macro1");
            excelObject.ExcelWriteCell(19, 2, "Jenkins Stable Run Report -" + todaysDate[0] + "(" + (daysName) + ")");
            Thread.Sleep(200);
            excelObject.FindExcelHyperLink();
            IWebElement FindBuildIsStableTestResultLink;
            IWebElement driverToFindNumberOfFailures;
            IWebElement driverToFindNumberOfTestCasesExecuted;
            IWebElement FindFailuresName;
            List<string> ListOfFailuresName = new List<string>();
            List<string> NewListOfFailuresName = new List<string>();
            List<int> CountOfEveryUniqueFailures = new List<int>();
            string[] FindFWBuildNumberSecondSplit;
            string actualFWBuildNumberWebScrapped;
            string numberOfFailuresInBuild;

            excelObject.SelectSheetNumber(2);
            Thread.Sleep(500);
            string totalNumberOfBuildsString = excelObject.ExcelReadCell(23, 1).ToString();
            string[] totalNumberOfBuildsInString = totalNumberOfBuildsString.Split(" ");
            int totalNumberOfBuilds = int.Parse(totalNumberOfBuildsInString[1]);
            excelObject.SelectSheetNumber(3);
            Thread.Sleep(200);
            string userName = excelObject.ExcelReadCell(29, 7);
            string passWord = excelObject.ExcelReadCell(31, 7);
            excelObject.SelectSheetNumber(2);
            Thread.Sleep(200);
            for (int i = 1; i < 3; i++)
            {
                excelObject.SelectSheetNumber(2);
                Thread.Sleep(200);
                string urlOfNightlySetupBuildReadFromExcel = excelObject.ExcelReadCell(i, 1);
                Thread.Sleep(200);
                driverObject.Navigate().GoToUrl(urlOfNightlySetupBuildReadFromExcel);

                if (i == 1)
                {
                    IWebElement enterUserName = driverObject.FindElement(By.Name("j_username"));
                    enterUserName.SendKeys(userName);
                    IWebElement enterPassWord = driverObject.FindElement(By.Name("j_password"));
                    enterPassWord.SendKeys(passWord);
                    Thread.Sleep(500);
                    IWebElement ClickSubmit = driverObject.FindElement(By.Name("Submit"));
                    ClickSubmit.Click();
                }

                IWebElement webElementToFindBuildNumberFromUrl = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/h1[1]/span[2]"));
                string buildInformationTextReadFromWebElement = webElementToFindBuildNumberFromUrl.Text;
                string[] firstSplitOfTextOfBuildInformationReadFromWebElement = buildInformationTextReadFromWebElement.Split("#");
                string[] secondSplitOfTextOfBuildInformationReadFromWebElement = firstSplitOfTextOfBuildInformationReadFromWebElement[1].Split(" ");
                string lastSuccessfulBuildNumberInText = secondSplitOfTextOfBuildInformationReadFromWebElement[0];
                int lastSuccessfulBuildNumberInInteger = int.Parse(lastSuccessfulBuildNumberInText);

                string[] setupUrlToUseFromExcel = urlOfNightlySetupBuildReadFromExcel.Split("lastSuccessfulBuild");
                string currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInText;
                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                string urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger + 1).ToString();
                driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                Thread.Sleep(2000);
            loopback:
                var error404PresentCheckWebElement = driverObject.FindElements(By.XPath("/html/body/h2"));

                if (error404PresentCheckWebElement.Count == 0)
                {
                    lastSuccessfulBuildNumberInInteger++;
                    urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger + 1).ToString();
                    driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                    Thread.Sleep(500);
                    goto loopback;
                }

                else if (error404PresentCheckWebElement.Count > 0)
                {
                    urlOfBuildPlusOneToBeChecked = setupUrlToUseFromExcel[0] + (lastSuccessfulBuildNumberInInteger).ToString();
                    driverObject.Navigate().GoToUrl(urlOfBuildPlusOneToBeChecked);
                    Thread.Sleep(1000);
                }

                currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInInteger.ToString();
                excelObject.SelectSheetNumber(1);
                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

            LoopBack1:
                IWebElement FindFWBuildNumberWebElement = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[2]/div[1]/div[1]"));

                string FindFWBuildNumberFirstSplit = FindFWBuildNumberWebElement.Text;
                if ((FindFWBuildNumberFirstSplit == "") || (FindFWBuildNumberFirstSplit == "Firmware Loading Error -aborting test"))
                {
                    actualFWBuildNumberWebScrapped = " ";
                }

                else
                {
                    FindFWBuildNumberSecondSplit = FindFWBuildNumberFirstSplit.Split("/");
                    string[] FindFWBuildNumberThirdSplit = FindFWBuildNumberSecondSplit[2].Split("\r");
                    actualFWBuildNumberWebScrapped = FindFWBuildNumberThirdSplit[0];
                }

                string[] hyperLinkName = currentUrlToBeChecked.Split("job/");
                hyperLinkName = hyperLinkName[4].Split("/");
                if (actualFWBuildNumberWebScrapped == "UNICOS_Turel_FW_Master")
                {
                    lastSuccessfulBuildNumberInInteger = lastSuccessfulBuildNumberInInteger + 1;
                    currentUrlToBeChecked = setupUrlToUseFromExcel[0] + lastSuccessfulBuildNumberInInteger.ToString();
                    driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                    goto LoopBack1;
                }
                string actualNameOfSetupHyperLinkToBeWrittenInExcel = hyperLinkName[0] + "-Build#" + lastSuccessfulBuildNumberInInteger;
                Thread.Sleep(1000);
                excelObject.EditExcelHyperLink((i), currentUrlToBeChecked, actualNameOfSetupHyperLinkToBeWrittenInExcel, totalNumberOfBuilds);

                if (FindFWBuildNumberFirstSplit != "")
                {
                    FindBuildIsStableTestResultLink = driverObject.FindElement(By.LinkText("Test Result"));
                    FindBuildIsStableTestResultLink.Click();
                    FindBuildIsStableTestResultLink = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[1]"));
                    numberOfFailuresInBuild = FindBuildIsStableTestResultLink.Text;
                    string[] listofbuildFailureBeforeStart = numberOfFailuresInBuild.Split(" ");
                    numberOfFailuresInBuild = listofbuildFailureBeforeStart[0];
                }

                else
                {
                    numberOfFailuresInBuild = " ";
                }

                if (numberOfFailuresInBuild.Contains("0"))
                {
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, "");
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, stableBuildEntryInExcel);
                }

                else
                {
                    excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, columnNumberForStableBuildWrite, emptyValue);
                }

                if (FindFWBuildNumberFirstSplit != "")
                {
                    string testReportUrl = currentUrlToBeChecked + "/testReport/";
                    driverObject.Navigate().GoToUrl(testReportUrl);
                    driverToFindNumberOfFailures = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[1]"));
                    string totalNumberOfFailures = driverToFindNumberOfFailures.Text;
                    string[] listTotalNumberOfFailures = totalNumberOfFailures.Split(" ");
                    driverToFindNumberOfTestCasesExecuted = driverObject.FindElement(By.XPath("/html[1]/body[1]/div[3]/div[2]/div[1]/div[3]"));
                    string totalNumberOfTestCases = driverToFindNumberOfTestCasesExecuted.Text;
                    string[] listTotalNumberOfTestCases = totalNumberOfTestCases.Split(" ");
                    failureResultToBeWrittenInExcel = ("Total failures- " + listTotalNumberOfFailures[0] + " failures out of " + listTotalNumberOfTestCases[0] + " tests");
                    excelObject.ExcelWriteCell(rowNumberofFirstBuildForNumberOfFailuresWrite, columnNumberForNumberOfFailuresWrite, failureResultToBeWrittenInExcel);
                    excelObject.ExcelWriteCell(rowNumberOfFirstBuildInfoHyperLinkInExcel, columnNumberOfBuildInfoHyperLinkInExcel, actualFWBuildNumberWebScrapped);
                    driverObject.Navigate().GoToUrl(currentUrlToBeChecked);
                    Thread.Sleep(200);
                    if ((int.Parse(listTotalNumberOfFailures[0]) > 0 && (int.Parse(listTotalNumberOfFailures[0])) < 1000))
                    {
                        testReportUrl = currentUrlToBeChecked + "/testReport/";
                        driverObject.Navigate().GoToUrl(testReportUrl);
                        string path = " ";
                        ListOfFailuresName.Clear();
                        for (int numberOfFailures = 0; numberOfFailures < int.Parse(listTotalNumberOfFailures[0]); numberOfFailures++)
                        {
                            path = "/html[1]/body[1]/div[3]/div[2]/table[2]/tbody[1]/" + "tr[" + (numberOfFailures + 1) + "]/td[1]/a[3]";
                            FindFailuresName = driverObject.FindElement(By.XPath(path));
                            string FailureNameText = FindFailuresName.Text;
                            if (FailureNameText.Contains("("))
                            {
                                string[] FailureNameChange = FailureNameText.Split("(");
                                FailureNameText = FailureNameChange[0];
                            }
                            ListOfFailuresName.Add(FailureNameText);

                        }

                        NewListOfFailuresName.Clear();
                        foreach (string e in ListOfFailuresName)
                        {
                            if (!(NewListOfFailuresName.Contains(e)))
                            {
                                NewListOfFailuresName.Add(e);
                            }
                        }

                        for (int jr = 0; jr < NewListOfFailuresName.Count; jr++)
                        {
                            string damt = NewListOfFailuresName[jr];
                            int count1 = 0;
                            for (int sr = 0; sr < ListOfFailuresName.Count; sr++)
                            {
                                if (ListOfFailuresName[sr] == damt)
                                {
                                    count1++;
                                }
                            }
                            CountOfEveryUniqueFailures.Add(count1);
                        }

                        driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

                        for (int failedCount = 0; failedCount < NewListOfFailuresName.Count; failedCount++)
                        {
                            excelObject.SelectSheetNumber(3);
                            Thread.Sleep(500);
                            excelObject.CopyRowFormat();
                            excelObject.SelectSheetNumber(1);
                            Thread.Sleep(500);
                            excelObject.InsertExcelRow(rowNumberOfFirstStableBuildInfoToBeWritten + 1);
                            excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, 2, NewListOfFailuresName[failedCount]);
                            excelObject.ExcelWriteCell(rowNumberOfFirstStableBuildInfoToBeWritten, 3, (CountOfEveryUniqueFailures[failedCount]).ToString());
                            rowNumberOfFirstStableBuildInfoToBeWritten++;
                            rowNumberOfFirstBuildInfoHyperLinkInExcel++;
                            rowNumberofFirstBuildForNumberOfFailuresWrite++;
                        }
                    }

                }

                else
                {
                    excelObject.ExcelWriteCell(rowNumberofFirstBuildForNumberOfFailuresWrite, columnNumberForNumberOfFailuresWrite, "Total Failures:No Test Report found/Build is still running.So do manual analysis");
                    excelObject.ExcelWriteCell(rowNumberOfFirstBuildInfoHyperLinkInExcel, columnNumberOfBuildInfoHyperLinkInExcel, "No firmware build was found: check manually");
                }

                driverObject.Navigate().GoToUrl(currentUrlToBeChecked);

                if (i == 1)
                {
                    excelObject.ExcelWriteCell(18, 2, "UNICOS_Turel/UNICOS_Turel_FW_Feature/" + actualFWBuildNumberWebScrapped);
                }
                rowNumberOfFirstBuildInfoHyperLinkInExcel = rowNumberOfFirstBuildInfoHyperLinkInExcel + 4;
                rowNumberOfFirstStableBuildInfoToBeWritten = rowNumberOfFirstStableBuildInfoToBeWritten + 4;
                rowNumberofFirstBuildForNumberOfFailuresWrite = rowNumberofFirstBuildForNumberOfFailuresWrite + 4;
            }
            driverObject.Close();
            driverObject.Quit();
            excelObject.SaveAsExcelFile("C:\\Users\\INPRSRI4\\Desktop\\New folder (15)\\FeatureBranchTestReport");
            excelObject.CloseExcelMethod();
            KillMethod();
        }

        public void KillMethod()
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))
                    clsProcess.Kill();
            Process[] chromeInstances = Process.GetProcessesByName("chrome");
            foreach (Process p in chromeInstances)
                p.Kill();
        }

    }
}
