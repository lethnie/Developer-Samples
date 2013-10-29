using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace Responses
{
    class Program
    {
        static void Main(string[] args)
        {
            string surveyName, authName, authPwd;
            authName = args[0];
            authPwd = args[1];
            surveyName = args[2];

            var surveyClient = new SurveyManagementServiceClient("BasicHttpBinding_ISurveyManagementService");
            var responseDataClient = new ResponseDataServiceClient("BasicHttpBinding_IResponseDataService");
            var authClient = new AuthenticationServiceClient("BasicHttpBinding_IAuthenticationService");
            var surveyEditor = new SurveyEditorServiceClient("BasicHttpBinding_ISurveyEditorService");

            var authTicket = authClient.Login(authName, authPwd);
            if (string.IsNullOrWhiteSpace(authTicket.ResultData))
            {
                Environment.Exit(1);
            }

            var surveyInfo = surveyClient.GetSurveyInfoByName(authTicket.ResultData, surveyName);
            if (surveyInfo == null)
            {
                Environment.Exit(1);
            }

            var page = responseDataClient.ListSurveyResponses(authTicket.ResultData, surveyInfo.ResultData.Id, 0, 1, "", "", "", true);
            if (page == null)
            {
                Environment.Exit(1);
            }
            var responses = page.ResultData.ResultPage;

            List<String> answers = new List<String>();
            var itemAnswer = new StringBuilder();
            var itemQuestion = new StringBuilder(); 
            
            if (responses == null)
            {
                Environment.Exit(1);
            }

            itemQuestion.Append('"' + "Id" + '"' + ',');
            itemQuestion.Append('"' + "Guid" + '"' + ',');
            itemQuestion.Append('"' + "Started" + '"' + ',');
            itemQuestion.Append('"' + "CompletionDate" + '"' + ',');
            itemQuestion.Append('"' + "Time" + '"' + ',');
            itemQuestion.Append('"' + "LastEditDate" + '"' + ',');
            itemQuestion.Append('"' + "RespondentIP" + '"' + ',');
            itemQuestion.Append('"' + "ResponseLanguage" + '"' + ',');
            itemQuestion.Append('"' + "UserIdentifier" + '"' + ',');
            itemQuestion.Append('"' + "Invitee" + '"' + ',');

            List<int> itemIds = new List<int>();
            List<String> questions = new List<String>();

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
                                if (k < itemInfo.ResultData.ChildItemIds.Length - 1)
                                {
                                    if (surveyClient.GetSurveyItemData(authTicket.ResultData, surveyInfo.ResultData.Id,
                                itemInfo.ResultData.ChildItemIds[k + 1]).ResultData.TypeName.Equals("Message"))
                                        questions.Add('"' + question + "_" + option + '"' + ",");
                                    else
                                        questions.Add('"' + question + "_" + option + '"');
                                }
                                else
                                    questions.Add('"' + question + "_" + option + '"');
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
                                if (j < itemInfo.ResultData.Options.Length - 1)
                                    questions.Add('"' + question + "_" + option + '"' + ",");
                                else
                                    questions.Add('"' + question + "_" + option + '"');

                            }
                        }
                        else
                            questions.Add('"' + question + '"');
                    }
                    if (i < surveyInfo.ResultData.ItemIds.Length - 1)
                        questions.Add(",");
                }
                else
                {
                    itemIds.Add(surveyInfo.ResultData.ItemIds[i]);
                }
            }

            for (int i = 0; i < questions.Count; i++)
            {
                itemQuestion.Append(questions[i]);
            }

            answers.Add(itemQuestion.ToString());

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
                
                itemAnswer = new StringBuilder();
                itemAnswer.Append('"' + responses[j].Id.ToString() + '"' + ',');
                itemAnswer.Append('"' + responses[j].Guid.ToString() + '"' + ',');
                itemAnswer.Append('"' + responses[j].Started.ToString() + '"' + ',');
                itemAnswer.Append('"' + responses[j].CompletionDate.ToString() + '"' + ',');
                itemAnswer.Append('"' + (responses[j].CompletionDate - responses[j].Started).ToString() + '"' + ',');
                itemAnswer.Append('"' + responses[j].LastEditDate.ToString() + '"' + ',');
                itemAnswer.Append('"' + responses[j].RespondentIp + '"' + ',');
                itemAnswer.Append('"' + responses[j].ResponseLanguage + '"' + ',');
                itemAnswer.Append('"' + responses[j].UserIdentifier + '"' + ',');
                itemAnswer.Append('"' + responses[j].Invitee + '"' + ',');
                items = responseItems.ResultData;

                int itemIdNum = 0;
                for (int i = 0; i < items.Length; i++)
                {
                    var item = items[itemNum[i]];

                    if (item.ItemId != itemIds[itemIdNum])
                    {
                        while (itemIds[itemIdNum] != item.ItemId)
                        {
                            if (itemIdNum < itemIds.Count - 1)
                                itemAnswer.Append('"'.ToString() + '"'.ToString() + ",");
                            else
                                itemAnswer.Append('"'.ToString() + '"'.ToString());
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
                                        itemAnswer.Append('"' + "0" + '"');
                                    else
                                        itemAnswer.Append('"' + itemAnswers[optNum].AnswerText + '"');
                                }
                                else
                                {
                                    if (optNum == itemAnswers.Length)
                                        itemAnswer.Append('"' + "0" + '"');
                                    else
                                        itemAnswer.Append('"' + "1" + '"');
                                }
                                if (k < itemInfo.ResultData.Options.Length - 1)
                                    itemAnswer.Append(",");
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
                                    itemAnswer.Append('"' + optNum.ToString() + '"');
                                    if (k < itemInfo.ResultData.Options.Length - 1)
                                        itemAnswer.Append(",");
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
                                        itemAnswer.Append('"' + "Other: " + itemAnswers[0].AnswerText + '"');
                                    }
                                    else
                                    {
                                        if (answer.OptionId.HasValue)
                                        {
                                            itemAnswer.Append('"' + answer.OptionText + '"');
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrWhiteSpace(answer.AnswerText))
                                            {
                                                itemAnswer.Append('"' + answer.AnswerText + '"');
                                            }
                                        }
                                    }
                                    if (k < itemAnswers.Length - 1)
                                        itemAnswer.Append(",");

                                }
                            }
                        }
                        
                        if (i < responseItems.ResultData.Length - 1)
                        {
                            itemAnswer.Append(",");
                        }
                                        
                }
                answers.Add(itemAnswer.ToString());
            }

            
            

            String fileName = String.Concat(surveyName,".csv");
            StreamWriter sw = new StreamWriter(fileName);
            foreach (String answer in answers)
            {
                sw.WriteLine(answer);
            }
            sw.Close();
            surveyClient.Close();
            authClient.Close();
            responseDataClient.Close();
            Environment.Exit(1);
        }
    }
}