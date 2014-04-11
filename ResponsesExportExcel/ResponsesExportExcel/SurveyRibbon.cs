using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Checkbox.Wcf.Services.Proxies;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using System.Xml;
using System.Configuration;

namespace ResponsesExportExcel
{  
    public partial class SurveyRibbon
    {        
        AuthenticationServiceClient authClient;
        SurveyManagementServiceClient surveyClient;
        ResponseDataServiceClient responseDataClient;
        ServiceOperationResultOfstring authTicket;

        int rowCount = 0;
        Dictionary<string, string> questions = new Dictionary<string, string>();
        Excel.Worksheet surveySheet;

        private void SampleRibbon_Load(object sender, RibbonUIEventArgs e)
        {           
            surveySheet = Globals.ThisAddIn.Application.ActiveSheet;
        }

        private void setServices()
        {
            surveyClient = new SurveyManagementServiceClient("BasicHttpBinding_ISurveyManagementService");
            responseDataClient = new ResponseDataServiceClient("BasicHttpBinding_IResponseDataService");
            authClient = new AuthenticationServiceClient("BasicHttpBinding_IAuthenticationService");
        }

        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            FormLogin formLogin = new FormLogin();
            formLogin.ShowDialog();
            if (formLogin.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                ClientInfo clientInfo = formLogin.getLoginInfo();

                setXmlFile(clientInfo.name);
                setServices();
                authTicket = authClient.Login(clientInfo.name, clientInfo.password);
                if (authTicket == null || string.IsNullOrWhiteSpace(authTicket.ResultData))
                {
                    System.Windows.Forms.MessageBox.Show("Wrong login or password.");
                    authTicket = null;
                }
                var listSurveys = surveyClient.ListAvailableSurveys(authTicket.ResultData, 0, 30, "", true, "", "");
                cbSurvey.Items.Clear();
                for (int i = 0; i < listSurveys.ResultData.TotalItemCount; i++)
                {
                    RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                    item.Label = listSurveys.ResultData.ResultPage[i].Name;
                    cbSurvey.Items.Add(item);
                }
                System.Windows.Forms.MessageBox.Show("You are logged in as " + clientInfo.name + ".");
            }
            
        }

        private void setXmlFile(String name)
        {
            String userName = name.Substring(0, name.IndexOf('@'));
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            XmlNodeList endpoints = xmlDoc.GetElementsByTagName("endpoint");
            foreach (XmlNode endpoint in endpoints)
            {
                endpoint.Attributes["address"].Value = endpoint.Attributes["address"].Value.Replace("dev.checkbox".ToString(), userName + ".checkboxonline");
            }

            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            ConfigurationManager.RefreshSection("endpoint");
        }

        private void btnSurvey_Click(object sender, RibbonControlEventArgs e)
        {
            if (authTicket == null)
            {
                System.Windows.Forms.MessageBox.Show("You must log in first.");
                return;
            }
            if (string.IsNullOrEmpty(cbSurvey.Label))
            {
                System.Windows.Forms.MessageBox.Show("You must select the survey first.");
                return;
            }
            getAnswersBySurveyName(cbSurvey.SelectedItem.ToString());
            cbQuestion1.Items.Clear();
            cbQuestion2.Items.Clear();
            for (int i = 0; i < cbQuestion.Items.Count; i++)
            {
                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                item.Label = cbQuestion.Items[i].Label;
                cbQuestion1.Items.Add(item);
                item = this.Factory.CreateRibbonDropDownItem();
                item.Label = cbQuestion.Items[i].Label;
                cbQuestion2.Items.Add(item);
            }
            surveySheet = Globals.ThisAddIn.Application.ActiveSheet;
        }

        private string getNextColumn(string name)
        {
            string result = name;
            if (name[name.Length - 1] < 'Z')
            {
                char last = (char)(name[name.Length - 1] + 1);
                return result.Remove(name.Length - 1).Insert(name.Length - 1, last.ToString());
            }
            int i = name.Length - 1;
            while ((i >= 0) && (name[i] == 'Z'))
            {
                i--;
            }
            if (i >= 0)
            {
                char last = (char)(name[i] + 1);
                result = result.Remove(i).Insert(i, last.ToString());
                for (; i < name.Length; i++)
                {
                    result = result.Remove(i).Insert(i, "A");                 
                }
                return result.ToString();
            }
            result = "";
            for (i = 0; i < name.Length + 1; i++)
            {
                result += "A";
            }
            return result;
        }

        

        private void getAnswersBySurveyName(string surveyName)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            string column = "A";
            int row = 1;

            var surveyInfo = surveyClient.GetSurveyInfoByName(authTicket.ResultData, surveyName);
            if (surveyInfo == null)
            {
                System.Windows.Forms.MessageBox.Show("Survey is not available.");
                return;
            }

            var page = responseDataClient.ListSurveyResponses(authTicket.ResultData, surveyInfo.ResultData.Id, 0, 1, "", "", "", true);
            if (page == null)
            {
                System.Windows.Forms.MessageBox.Show("Survey responses are not available.");
                return;
            }
            var responses = page.ResultData.ResultPage;

            List<String> answers = new List<String>();

            activeSheet.Range[column + row.ToString()].Value2 = "Id";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "Guid";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "Started";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "CompletionDate";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "Time";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "LastEditDate";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "RespondentIP";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "ResponseLanguage";           
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "UserIdentifier";
            column = getNextColumn(column);
            activeSheet.Range[column + row.ToString()].Value2 = "Invitee";
            column = getNextColumn(column);

            List<int> itemIds = new List<int>();

            cbQuestion.Items.Clear();
            for (int i = 1; i < surveyInfo.ResultData.ItemIds.Length; i++)
            {
                var itemInfo = surveyClient.GetSurveyItemData(authTicket.ResultData, surveyInfo.ResultData.Id, surveyInfo.ResultData.ItemIds[i]);
                if (itemInfo.ResultData.IsAnswerable)
                {
                    var question = itemInfo.ResultData.TextData[0].TextValues.NameValueList[0].Value;
                    question = Regex.Replace(question, @"<[^>]*>", String.Empty);
                    if (itemInfo.ResultData.TypeName.Equals("Matrix"))
                    {
                        for (int k = 0; k < itemInfo.ResultData.ChildItemIds.Length; k++)
                        {
                            var childItem = surveyClient.GetSurveyItemData(authTicket.ResultData, surveyInfo.ResultData.Id,
                                itemInfo.ResultData.ChildItemIds[k]);
                            if (childItem.ResultData.TypeName.Equals("Message"))
                            {
                                var option = childItem.ResultData.TextData[0].TextValues.NameValueList[0].Value;
                                activeSheet.Range[column + row.ToString()].Value2 = question + "_" + option;

                                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                                item.Label = question + "_" + option;
                                cbQuestion.Items.Add(item);
                                questions.Add(item.Label, column);

                                column = getNextColumn(column);
                            }
                            else
                            {
                                itemIds.Add(itemInfo.ResultData.ChildItemIds[k]);
                            }
                        }
                    }
                    else
                    {
                        itemIds.Add(surveyInfo.ResultData.ItemIds[i]);
                        if (itemInfo.ResultData.TypeName.Equals("Checkboxes") || (itemInfo.ResultData.TypeName.Equals("RankOrder")))
                        {
                            for (int j = 0; j < itemInfo.ResultData.Options.Length; j++)
                            {
                                var option = itemInfo.ResultData.Options[j].TextData[0].TextValues.NameValueList[0].Value;
                                activeSheet.Range[column + row.ToString()].Value2 = question + "_" + option;

                                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                                item.Label = question + "_" + option;
                                cbQuestion.Items.Add(item);
                                questions.Add(item.Label, column);

                                column = getNextColumn(column);
                            }
                        }
                        else
                        {
                             activeSheet.Range[column + row.ToString()].Value2 = question;

                             RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                             item.Label = question;
                             cbQuestion.Items.Add(item);
                             questions.Add(item.Label, column);

                             column = getNextColumn(column);
                        }
                    }

                }
                else
                {
                    itemIds.Add(surveyInfo.ResultData.ItemIds[i]);
                }
            }

            row++;
            column = "A";

            for (int j = 0; j < responses.Length; j++)
            {
                var responseGuid = responses[j].Guid;
                var responseItems = responseDataClient.GetAnswersForResponseByGuid(authTicket.ResultData, surveyInfo.ResultData.Id, "en-US",
                                                                 responseGuid);
                
                var items = responseItems.ResultData;
                int temp;

                int[] itemNum = new int[items.Length];
                for (int k = 0; k < items.Length; k++)
                {
                    itemNum[k] = k;
                }
                for (int i = 0; i < items.Length - 1; i++)
                {
                    for (int ii = 0; ii < items.Length - i - 1; ii++)
                    {
                        if (items[itemNum[ii]].ItemId > items[itemNum[ii+1]].ItemId)
                        {
                            temp = itemNum[ii];
                            itemNum[ii] = itemNum[ii + 1];
                            itemNum[ii + 1] = temp;
                        }
                    }
                }
                
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].Id.ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].Guid.ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].Started.ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].CompletionDate.ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = (responses[j].CompletionDate - responses[j].Started).ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].LastEditDate.ToString();
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].RespondentIp;
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].ResponseLanguage;           
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].UserIdentifier;
                column = getNextColumn(column);
                activeSheet.Range[column + row.ToString()].Value2 = responses[j].Invitee;
                column = getNextColumn(column);
                
                items = responseItems.ResultData;

                int itemIdNum = 0;
                for (int i = 0; i < items.Length; i++)
                {
                    var item = items[itemNum[i]];

                    if (item.ItemId != itemIds[itemIdNum])
                    {
                        while (itemIds[itemIdNum] != item.ItemId)
                        {
                            column = getNextColumn(column);
                            itemIdNum++;
                        }
                    }
      
                        itemIdNum++;
                        var itemAnswers = item.Answers;
                        var itemInfo = surveyClient.GetSurveyItemData(authTicket.ResultData, surveyInfo.ResultData.Id, item.ItemId);
                        if (itemInfo.ResultData.TypeName.Equals("Checkboxes"))
                        {
                            for (int k = 0; k < itemInfo.ResultData.Options.Length; k++)
                            {
                                int optNum = 0;
                                string optText = itemInfo.ResultData.Options[k].TextData[0].TextValues.NameValueList[0].Value;
                                
                                while ((optNum < itemAnswers.Length) && (!itemAnswers[optNum].OptionText.Equals(optText)))
                                {
                                    optNum++;
                                }
                                if (optText.Equals("Other:"))
                                {
                                    if (optNum == itemAnswers.Length)
                                    {
                                        activeSheet.Range[column + row.ToString()].Value2 = "0";
                                        column = getNextColumn(column);
                                    }
                                    else
                                    {                                      
                                        activeSheet.Range[column + row.ToString()].Value2 = itemAnswers[optNum].AnswerText;
                                        column = getNextColumn(column);
                                    }
                                }
                                else
                                {
                                    if (optNum == itemAnswers.Length)
                                    {
                                        activeSheet.Range[column + row.ToString()].Value2 = "0";
                                        column = getNextColumn(column);
                                    }
                                    else
                                    {
                                        activeSheet.Range[column + row.ToString()].Value2 = "1";
                                        column = getNextColumn(column);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (itemInfo.ResultData.TypeName.Equals("RankOrder"))
                            {
                                for (int k = 0; k < itemInfo.ResultData.Options.Length; k++)
                                {
                                    int optNum = 0;
                                    string optText = itemInfo.ResultData.Options[k].TextData[0].TextValues.NameValueList[0].Value;
                                    while ((optNum < itemAnswers.Length) && (!itemAnswers[optNum].OptionText.Equals(optText)))
                                    {
                                        optNum++;
                                    }
                                    optNum++;
                                    activeSheet.Range[column + row.ToString()].Value2 = optNum.ToString();
                                    column = getNextColumn(column);
                                }
                            }
                            else
                            {
                                int[] ansNum = new int[itemAnswers.Length];
                                for (int k = 0; k < itemAnswers.Length; k++)
                                {
                                    ansNum[k] = k;
                                }
                                for (int q = 0; q < itemAnswers.Length - 1; q++)
                                {
                                    for (int ii = 0; ii < itemAnswers.Length - i - 1; ii++)
                                    {
                                        if (itemAnswers[ansNum[ii]].AnswerId > itemAnswers[ansNum[ii + 1]].AnswerId)
                                        {
                                            temp = ansNum[ii];
                                            ansNum[ii] = ansNum[ii + 1];
                                            ansNum[ii + 1] = temp;
                                        }
                                    }
                                }
                                for (int k = 0; k < itemAnswers.Length; k++)
                                {
                                    var answer = itemAnswers[ansNum[k]];
                                    if (answer.OptionId.HasValue && !string.IsNullOrWhiteSpace(answer.AnswerText))
                                    {
                                        activeSheet.Range[column + row.ToString()].Value2 = "Other: " + itemAnswers[0].AnswerText;
                                        column = getNextColumn(column);
                                    }
                                    else
                                    {
                                        if (answer.OptionId.HasValue)
                                        {
                                            activeSheet.Range[column + row.ToString()].Value2 = answer.OptionText;
                                            column = getNextColumn(column);
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrWhiteSpace(answer.AnswerText))
                                            {
                                                activeSheet.Range[column + row.ToString()].Value2 = answer.AnswerText;
                                                column = getNextColumn(column);
                                            }
                                        }
                                    }
                                }
                            }
                        }                                                             
                }

                row++;
                column = "A";
            }

            rowCount = row;
            surveySheet = activeSheet;
        }

        private void btnChart_Click(object sender, RibbonControlEventArgs e)
        {      
            getAnswersChart(cbQuestion.SelectedItem.ToString());
        }

        private void getAnswersChart(string ans)
        {
            string numColumn = questions[ans];
            Dictionary<string, int> answers = new Dictionary<string, int>();
            
            string title = surveySheet.Range[numColumn + "1"].Value2.ToString();
            for (int i = 2; i <= rowCount; i++)
            {
                string key;
                if (surveySheet.Range[numColumn + i.ToString()].Value2 == null)
                    key = "";
                else
                    key = surveySheet.Range[numColumn + i.ToString()].Value2.ToString();
                if (key != "")
                {
                    if (answers.ContainsKey(key))
                    {
                        answers[key]++;
                    }
                    else
                    {
                        answers.Add(key, 1);
                    }
                }
            }

            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.Sheets.Add();

            int k = 2;
            activeSheet.Range["A1"].Value2 = ans;
            activeSheet.Range["B1"].Value2 = "Count";

            foreach (var answer in answers)
            {
                activeSheet.Range["A" + (k).ToString()].Value2 =
                    '"'.ToString() + answer.Key.ToString() + '"'.ToString();
                activeSheet.Range["B" + (k).ToString()].Value2 =
                    answer.Value;
                k++;
            }           

            Excel.Range data = Globals.ThisAddIn.Application.ActiveSheet.Range["A2", "B" + (answers.Count + 1).ToString()];
            Excel.Chart myNewChart =
            Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddChart(XlChartType: Excel.XlChartType.xl3DPie,
                    Left: Type.Missing, Top: Type.Missing, Width: Type.Missing,
                    Height: Type.Missing).Select();
            myNewChart = Globals.ThisAddIn.Application.ActiveChart;
            myNewChart.SetSourceData(data, Excel.XlRowCol.xlColumns);
            myNewChart.SetElement(Office.MsoChartElementType.msoElementChartTitleAboveChart);
            myNewChart.ChartTitle.Text = title;
        }

        private void btnCrTab_Click(object sender, RibbonControlEventArgs e)
        {
            string numCol1 = questions[cbQuestion1.SelectedItem.Label];
            string numCol2 = questions[cbQuestion2.SelectedItem.Label];
            getAnswersCrossTable(numCol1, numCol2);
        }

        private void getAnswersCrossTable(string numCol1, string numCol2)
        {
            Dictionary<string, Dictionary<string, int>> answers = new Dictionary<string, Dictionary<string, int>>();

            string title = surveySheet.Range[numCol1 + "1"].Value2.ToString() + " " +
                surveySheet.Range[numCol2 + "1"].Value2.ToString();
            for (int i = 2; i <= rowCount; i++)
            {
                string key;
                if (surveySheet.Range[numCol1 + i.ToString()].Value2 == null)
                    key = "";
                else
                    key = surveySheet.Range[numCol1 + i.ToString()].Value2.ToString();
                if (key != "")
                {
                    string key2;
                    if (surveySheet.Range[numCol2 + i.ToString()].Value2 == null)
                        key2 = "";
                    else
                        key2 = surveySheet.Range[numCol2 + i.ToString()].Value2.ToString();

                    if (answers.ContainsKey(key))
                    {
                        if (key2 != "")
                        {
                            if (answers[key].ContainsKey(key2))
                            {
                                answers[key][key2]++;
                            }
                            else
                            {
                                answers[key].Add(key2, 1);
                            }
                        }
                    }
                    else
                    {
                        Dictionary<string, int> val = new Dictionary<string, int>();
                        if (key2 != "" )
                            val.Add(key2, 1);
                        answers.Add(key, val);
                    }
                }
            }

            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.Sheets.Add();

            int k = 2;
            Dictionary<string, string> ans_col = new Dictionary<string, string>();
            string freeCol = "B";
            foreach (var answer in answers)
            {
                activeSheet.Range["A" + (k).ToString()].Value2 =
                        '"'.ToString() + answer.Key.ToString() + '"'.ToString();
                foreach (var ans in answer.Value)
                {
                    string col = "B";
                    while (!col.Equals(freeCol))
                    {
                        activeSheet.Range[col + (k).ToString()].Value2 =
                            0;
                        col = getNextColumn(col);
                    }
                    if (ans_col.ContainsKey(ans.Key))
                    {                    
                        activeSheet.Range[ans_col[ans.Key] + (k).ToString()].Value2 =
                            ans.Value;
                    }
                    else
                    {
                        ans_col.Add(ans.Key, freeCol);
                        if (k > 2)
                        {                            
                            for (int i = 2; i < k; i++)
                            {
                                activeSheet.Range[ans_col[ans.Key] + (i).ToString()].Value2 =
                                    0;
                            }
                        }
                        freeCol = getNextColumn(freeCol);
                        activeSheet.Range[ans_col[ans.Key] + "1"].Value2 =
                            '"'.ToString() + ans.Key.ToString() + '"'.ToString();
                        activeSheet.Range[ans_col[ans.Key] + (k).ToString()].Value2 =
                            ans.Value;
                        
                    }
                }
                k++;
            }
        }

    }

    public class ClientInfo
    {
        public string name;
        public string password;
    }
}
