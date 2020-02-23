﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CaptureUploader
{
    class GoogleCredential
    {


        private readonly BackgroundWorker worker = new BackgroundWorker();
        private static AutoResetEvent resetEvent = new AutoResetEvent(false);
        private string[] GoogleScope = null;
        public void GetGoogleCredential(string[] _Scopes)
        {
            this.GoogleScope = _Scopes;
            worker.DoWork += Do_GettingCredentialWork;
            worker.RunWorkerCompleted += Done_GettingCredentialWork;
            worker.RunWorkerAsync();
            resetEvent.WaitOne();
        }

        private void Do_GettingCredentialWork(object sender, DoWorkEventArgs e)
        {
            UserCredential credential;

            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            string clientID = Path.Combine(projectDirectory, "Resources\\client_id.json");

            using (var stream =
                new FileStream(clientID, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/credentials.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    GoogleScope,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            resetEvent.Set();
        }

        private void Done_GettingCredentialWork(object sender, RunWorkerCompletedEventArgs e)
        {
            worker.DoWork -= Do_GettingCredentialWork;
            worker.RunWorkerCompleted -= Done_GettingCredentialWork;
        }
    }
}