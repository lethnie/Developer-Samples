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
            
            if (responses == null)
            {
                Environment.Exit(1);
            }

            for (int j = 0; j < responses.Length; j++)
            {
                var responseGuid = responses[j].Guid;
                var responseItems = responseDataClient.GetAnswersForResponseByGuid(authTicket.ResultData, surveyInfo.ResultData.Id, "en-US",
                                                                                     responseGuid);
                itemAnswer = new StringBuilder();
                //itemAnswer.Append('"' + responses[j].Id.ToString() + '"' + ',' + '"' + responses[j].Guid.ToString() + '"' + ',');
                itemAnswer.Append('"');
                for (int i = 0; i < responseItems.ResultData.Length; i++)
                {                   
                    var item = responseItems.ResultData[i];                   
                    for (int k = 0; k < item.Answers.Length; k++)
                    {
                        var answer = item.Answers[k];
                        if (answer.OptionId.HasValue && !string.IsNullOrWhiteSpace(answer.AnswerText))
                        {
                            itemAnswer.Append("Other: " + answer.AnswerText + '"');
                            if (k < item.Answers.Length - 1)
                                itemAnswer.Append(","+'"');
                        }

                        if (answer.OptionId.HasValue)
                        {
                            itemAnswer.Append(answer.OptionText + '"');
                            if (k < item.Answers.Length - 1)
                                itemAnswer.Append("," + '"');
                            continue;
                        }

                        if (!string.IsNullOrWhiteSpace(answer.AnswerText))
                        {
                            itemAnswer.Append(answer.AnswerText + '"');
                            if (k < item.Answers.Length - 1)
                                itemAnswer.Append("," + '"');
                        }
                    }
                    if (i < responseItems.ResultData.Length - 1)
                    {
                        itemAnswer.Append("," + '"');
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
            Environment.Exit(1);
        }
    }
}