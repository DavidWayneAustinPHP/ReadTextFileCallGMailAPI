using System;
using System.Text.RegularExpressions;
using MimeKit;
using MimeKit.Utils;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using System.IO;
using System.Threading;
using Google.Apis.Services;
using Google.Apis.Util.Store;

// Install the following packages in Package Manager Console
// Install-Package MimeKit
// Install-Package Google.Apis.Oauth2.v1 -Version 1.9.0.860
// Install-Package Google.Apis.Gmail.v1 -Version 1.38.2.1543
// Also configure the Google Api manager to allow use of an email address for SMTP trafic emails to be sent from this program
// https://console.cloud.google.com/apis/dashboard

namespace SampleProject
{
    class Program
    {
        static void Main(string[] args)
        {
            // Declare starting variables and ask for user input
            string locationFile = "D:\\locationOfFile\\Sample.txt";
            Console.WriteLine("This program reads data from a text file and outputs the result");
            Console.WriteLine("If the result is 0 or the operation in invalid and email is sent");
            Console.WriteLine("Press enter for the program to run");
            Console.ReadLine();

            // Call the method readfile to read the data in the text file
            readfile(locationFile);

            Console.ReadLine();
        }

        // This method reads in the data from the text file and performs the calculations and writes out the output with the date time stamp
        public static void readfile(string locationFile)
        {
            var fileData = System.IO.File.ReadAllLines(locationFile);
            
            foreach (string line in fileData)
            {
                // Read in the data from the text file
                string[] splitFileData = Regex.Split(line, ",");
                int resultCalculation;

                if (splitFileData[3] == "Multiplication")
                {
                    int value1 = Int32.Parse(splitFileData[0]);
                    int value2 = Int32.Parse(splitFileData[1]);
                    int value3 = Int32.Parse(splitFileData[2]);
                    resultCalculation = value1 * value2 * value3;

                    Console.WriteLine("{0} * {1} * {2} = {3} TimeStamp: {4}", value1, value2, value3, resultCalculation, DateTime.Now.ToString("{HH:MM:ss, dddd, MMMM d, yyyy}"));

                    // If there is an result of less than 0 then call the menthod send email
                    if (resultCalculation < 0)
                    {
                        // Send email using the method sendEmailLessThanZero
                        sendEmailLessThanZero(splitFileData[3]);
                    }
                }

                if (splitFileData[3] == "Addition")
                {
                    int value1 = Int32.Parse(splitFileData[0]);
                    int value2 = Int32.Parse(splitFileData[1]);
                    int value3 = Int32.Parse(splitFileData[2]);
                    resultCalculation = value1 + value2 + value3;

                    Console.WriteLine("{0} + {1} + {2} = {3} TimeStamp: {4}", value1, value2, value3, resultCalculation, DateTime.Now.ToString("{HH:MM:ss, dddd, MMMM d, yyyy}"));

                    // If there is an result of less than 0 then call the menthod send email
                    if (resultCalculation < 0)
                    {
                        // Send email using the method sendEmailLessThanZero
                        sendEmailLessThanZero(splitFileData[3]);
                    }
                }

                if (splitFileData[3] == "Subtraction")
                {
                    int value1 = Int32.Parse(splitFileData[0]);
                    int value2 = Int32.Parse(splitFileData[1]);
                    int value3 = Int32.Parse(splitFileData[2]);
                    resultCalculation = value1 - value2 - value3;

                    Console.WriteLine("{0} - {1} - {2} = {3} TimeStamp: {4}", value1, value2, value3, resultCalculation, DateTime.Now.ToString("{HH:MM:ss, dddd, MMMM d, yyyy}"));

                    // If there is an result of less than 0 then call the menthod send email
                    if (resultCalculation < 0)
                    {
                        // Send email using the method sendEmailLessThanZero
                        sendEmailLessThanZero(splitFileData[3]);
                    }
                }

                if (splitFileData[3] == "Division")
                {
                    int value1 = Int32.Parse(splitFileData[0]);
                    int value2 = Int32.Parse(splitFileData[1]);
                    int value3 = Int32.Parse(splitFileData[2]);

                    // Using a try catch statement to catch if there is a divide by Zero error
                    try
                    {
                        resultCalculation = value1 / value2 / value3;

                        Console.WriteLine("{0} / {1} / {2} = {3} TimeStamp: {4}", value1, value2, value3, resultCalculation, DateTime.Now.ToString("{HH:MM:ss, dddd, MMMM d, yyyy}"));

                        // If there is an result of less than 0 then call the menthod send email
                        if (resultCalculation < 0)
                        {
                            // Send email using the method sendEmailLessThanZero
                            sendEmailLessThanZero(splitFileData[3]);
                        }
                    }
                    catch (Exception e)
                    {
                        string failedCalculation = "Divide by Zero Error";
                        
                        Console.WriteLine("{0} / {1} / {2} = {3} TimeStamp: {4}", value1, value2, value3, failedCalculation, DateTime.Now.ToString("{HH:MM:ss, dddd, MMMM d, yyyy}"));

                        // Send email using the method sendEmailLessThanZero
                        sendEmailLessThanZero("Divide by Zero Error");
                    }
                }
            }
            
        }

        // This method using the Google API to send emails
        public static void sendEmailLessThanZero(string emailErrorResult)
        {
            Console.WriteLine("Trying to send an Email because error with {0} has occurred", emailErrorResult);

            // Declaring the email addresses used in this method, have not entered the correct emails addresses as for security reason
            // In a live version of this method the email addresses will be held in a more secure model
            string youremailName = "EnterName";
            string youremailAddress = "emailName@gmail.com";
            string sendToName = "EnterNameSendingto";
            string sendToEmailAddress = "emailNameSendTo@gmail.com";
            UserCredential credential;

            // A JSON secret file needs to be be generated in the Google API Manager to allow for authentication
            string locationGoogleClientSecretFile = "Enter location of the JSON Google API client secret file";
            using (var stream = new FileStream(locationGoogleClientSecretFile, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, new[] {
                    GmailService.Scope.MailGoogleCom,
                    GmailService.Scope.GmailCompose,
                    GmailService.Scope.GmailModify, GmailService.Scope.GmailSend },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                var service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "GmailService API .NET Quickstart",
                });

                // The below code builds the message body of the mail, builder.HtmlBody can be used to build a HTML email if needed
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(youremailName, youremailAddress));
                message.To.Add(new MailboxAddress(sendToName, sendToEmailAddress));

                message.Subject = "An error has occurred with the calculations";
                var builder = new BodyBuilder();
                DateTime dateTime = DateTime.Now;
                string strDate = dateTime.ToString("HH:MM:ss, dddd, MMMM d, yyyy");

                builder.TextBody = @"An error has occurred with " + emailErrorResult + " at the followinf date time " + strDate;
                var contentId = MimeUtils.GenerateMessageId();

                // The email needs to be encoded before sending via the Google API
                message.Body = builder.ToMessageBody();
                var rawMessage = "";
                using (var streem = new MemoryStream())
                {
                    message.WriteTo(streem);
                    rawMessage = Convert.ToBase64String(streem.GetBuffer(), 0,
                        (int)streem.Length)
                        .Replace('+', '-')
                        .Replace('/', '_')
                        .Replace("=", "");
                }

                var gmailMessage = new Google.Apis.Gmail.v1.Data.Message
                {
                    Raw = rawMessage
                };

                // Send the encoded email and wrap in a Try Catch encase the email send is not successfully sent and report an erro has occured
                try
                {
                    service.Users.Messages.Send(gmailMessage, sendToEmailAddress).Execute();
                }
                catch (Exception e)
                {
                    Console.WriteLine("There was a problem sending the email via the Google Api");
                }

            }
        }
    }
}
