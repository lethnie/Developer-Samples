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

namespace ResponsesExportExcel
{  
    public partial class SurveyRibbon
    {
        ClientInfo clientInfo;
        AuthenticationServiceClient authClient;
        SurveyManagementServiceClient surveyClient;
        ResponseDataServiceClient responseDataClient;
        ServiceOperationResultOfstring authTicket;

        int rowCount = 0;
        Dictionary<string, string> questions = new Dictionary<string, string>();
        Excel.Worksheet surveySheet;

        private void SampleRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            surveyClient = new SurveyManagementServiceClient("BasicHttpBinding_ISurveyManagementService");
            responseDataClient = new ResponseDataServiceClient("BasicHttpBinding_IResponseDataService");
            authClient = new AuthenticationServiceClient("BasicHttpBinding_IAuthenticationService");

            surveySheet = Globals.ThisAddIn.Application.ActiveSheet;
        }

        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            FormLogin formLogin = new FormLogin();
            formLogin.ShowDialog();
            if (formLogin.DialogResult == System.Windows.Forms.DialogResult.OK)
            {              
                clientInfo = formLogin.getLoginInfo();
                
                authTicket = authClient.Login(clientInfo.name, clientInfo.password);
                if (authTicket == null || string.IsNullOrWhiteSpace(authTicket.ResultData))
                {
                    System.Windows.Forms.MessageBox.Show("Wrong login or password.");
                    authTicket = null;
                }
                var listSurveys = surveyClient.ListAvailableSurveys(authTicket.ResultData, 0, 10, "", true, "", "");
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
            getAnswersBySurveyName(cbSurvey.Text);
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
            //firstColumn = column;

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
            getAnswersChart(questions[cbQuestion.Text]);
        }

        private void getAnswersChart(string numColumn)
        {
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

            int k = 1;
            foreach (var answer in answers)
            {
                activeSheet.Range["A" + (k).ToString()].Value2 =
                    '"'.ToString() + answer.Key.ToString() + '"'.ToString();
                activeSheet.Range["B" + (k).ToString()].Value2 =
                    answer.Value;
                k++;
            }           

            Excel.Range data = Globals.ThisAddIn.Application.ActiveSheet.Range["A1", "B" + (answers.Count).ToString()];
            Excel.Chart myNewChart =
            Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddChart(XlChartType: Excel.XlChartType.xl3DPie,
                    Left: Type.Missing, Top: Type.Missing, Width: Type.Missing,
                    Height: Type.Missing).Select();
            myNewChart = Globals.ThisAddIn.Application.ActiveChart;
            myNewChart.SetSourceData(data, Excel.XlRowCol.xlColumns);
            myNewChart.SetElement(Office.MsoChartElementType.msoElementChartTitleAboveChart);
            myNewChart.ChartTitle.Text = title;
        }

    }

    public class ClientInfo
    {
        public string name;
        public string password;
    }
}
