using Google.Apis.Auth.OAuth2;
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
        static EventWaitHandle _waitHandle = new AutoResetEvent(false);
        private string[] GoogleScope = null;
        public UserCredential GetGoogleCredential(string[] _Scopes)
        {
            this.GoogleScope = _Scopes;

            var task = Task.Run(() =>
            {
                this.Do_GettingCredentialWork(null, null);
            });

            bool isTimeOut = task.Wait(TimeSpan.FromMilliseconds(5000));
            if (isTimeOut)
            {
                Console.WriteLine("GetGoogleCredential done.");
            }
            else
            {
                //throw new TimeoutException("The function has taken longer than the maximum time allowed.");
                Console.WriteLine("The function has taken longer than the maximum time allowed.");
            }

            //_waitHandle.WaitOne();
            //worker.DoWork += Do_GettingCredentialWork;
            //worker.RunWorkerCompleted += Done_GettingCredentialWork;
            //worker.RunWorkerAsync();

            return credential;
        }
        

        private UserCredential credential = null;
        private void Do_GettingCredentialWork(object sender, DoWorkEventArgs e)
        {
            Console.WriteLine("Do_GettingCredentialWork called");

            //string workingDirectory = Environment.CurrentDirectory;
            string workingDirectory = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Console.WriteLine($"workingDirectory: {workingDirectory}");
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            string clientID = Path.Combine(projectDirectory, "Resources\\credentials.json");

            using (var stream =
                new FileStream(clientID, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/token.json");

                Console.WriteLine($"credPath: {credPath}");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    GoogleScope,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
        }

        private void Done_GettingCredentialWork(object sender, RunWorkerCompletedEventArgs e)
        {
            worker.DoWork -= Do_GettingCredentialWork;
            worker.RunWorkerCompleted -= Done_GettingCredentialWork;
            _waitHandle.Set();
        }
    }
}
