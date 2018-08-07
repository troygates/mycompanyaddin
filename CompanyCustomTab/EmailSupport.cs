using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using System;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management;
using System.Net.Mail;
using System.Reflection;

namespace CompanyCustomTab
{
    /// <summary>
    /// Creates an email with user, computer, and Revit info to send to a company support email address
    /// NOTE: To change the email address, change in Properties-> Settings-> EmailSupport
    /// </summary>
    class EmailSupport
    {
        Document doc;
        Application app;
        private string UserName;
        private string UserFullName;
        private string Email;
        private string Computer;
        private string RevitVersion;
        private string RevitServerAccelerator;
        private string LocalModel;
        private string CentralModel;
        private string modelSize;
        private string totalMemory;
        private string freeMemory;

        public EmailSupport(Document document, Application application)
        {
            doc = document;
            app = application;

            // Collect info about user
            UserName = Environment.UserName;

            // Following only works when computer is on a domain
            try
            {
                UserFullName = UserPrincipal.Current.DisplayName;
                Email = UserPrincipal.Current.EmailAddress;
            }
            catch 
            {
                UserFullName = UserName;
                Email = "";
            }     
            
            Computer = Environment.MachineName;

            // Collect info about Revit
            RevitVersion = app.VersionName + " | " + app.VersionBuild;
            RevitServerAccelerator = app.CurrentRevitServerAccelerator;
            LocalModel = doc.PathName;

            // Check to see if model is a worksharing model
            if (doc.GetWorksharingCentralModelPath() != null)
            {
                CentralModel = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            }
            else
            {
                CentralModel = "Central model couldn't be found";
            }

            ModelSize();
            GetMemory();
        }

        /// <summary>
        /// Compute the file size of the Revit model
        /// </summary>
        private void ModelSize()
        {
            try
            {
                string[] sizes = { "B", "KB", "MB", "GB", "TB" };
                double len = new FileInfo(LocalModel).Length;
                int order = 0;
                while (len >= 1024 && order < sizes.Length - 1)
                {
                    order++;
                    len = len / 1024;
                }
                modelSize = String.Format("{0:0.#} {1}", len, sizes[order]);

            }
            catch
            {
                modelSize = "Error in computing model size";
            }
        }

        /// <summary>
        /// Calculate Total and Free computer Memory
        /// </summary>
        private void GetMemory()
        {
            ManagementObjectSearcher mgmtObjects = new ManagementObjectSearcher("Select * from CIM_OperatingSystem");
            foreach (var item in mgmtObjects.Get())
            {
                double totalMemoryD = Convert.ToDouble(item.Properties["TotalVisibleMemorySize"].Value.ToString());
                double freeMemoryD = Convert.ToDouble(item.Properties["FreePhysicalMemory"].Value.ToString());
                totalMemoryD = totalMemoryD / 1024 / 1024;
                freeMemoryD = freeMemoryD / 1024 / 1024;

                totalMemory = String.Format("{0:0} {1}", totalMemoryD, "GB");
                freeMemory = String.Format("{0:0.#} {1}", freeMemoryD, "GB");
            }
        }

        /// <summary>
        /// Create the email file
        /// </summary>
        /// <returns></returns>
        public string CreateEmail()
        {
            string EmailFrom = Email;
            if (string.IsNullOrEmpty(EmailFrom))
                EmailFrom = Properties.Settings.Default.EmailSupport;

            var mailMessage = new MailMessage(EmailFrom, Properties.Settings.Default.EmailSupport);
            mailMessage.Subject = "Revit Support";
            mailMessage.IsBodyHtml = true;

            // Create the email body of text
            string message = "<p>&lt;<i>PLEASE DESCRIBE ISSUE HERE</i>&gt;</p><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><p>";
            message = message + "<b><i>DIAGNOSTIC INFO</i></b><br /><b>===============</b>";
            message = message + "<br /><b>USER:</b> " + UserFullName;
            message = message + "<br /><b>EMAIL:</b> " + Email;
            message = message + "<br /><b>COMPUTER:</b> " + Computer;
            message = message + "<br /><b>REVIT VERSION:</b> " + RevitVersion;
            message = message + "<br /><b>TOTAL MEMORY:</b> " + totalMemory;
            message = message + "<br /><b>FREE MEMORY:</b> " + freeMemory;
            message = message + "<br /><br /><b>LOCAL MODEL:</b> " + LocalModel;
            message = message + "<br /><b>CENTRAL MODEL:</b> " + CentralModel;
            message = message + "<br /><b>MODEL FILE SIZE:</b> " + modelSize;
            message = message + "<br /><b>REVIT SERVER ACCELERATOR:</b> " + RevitServerAccelerator;
            message = message + "<br /><br /><br />";

            mailMessage.Body = message;

            // Attach the journal file
            string journalFile = app.RecordingJournalFilename;
            File.Copy(journalFile, Path.GetTempPath() + Path.GetFileName(journalFile), true);
            mailMessage.Attachments.Add(new Attachment(Path.GetTempPath() + Path.GetFileName(journalFile)));

            // Attach the central log file
            if (doc.GetWorksharingCentralModelPath() != null)
            {
                try
                {
                    string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                    string centralName = Path.GetFileName(centralPath);
                    string centralLog = centralName.Replace(".rvt", @"_backup\" +
                        Path.GetFileNameWithoutExtension(centralPath) + ".slog");
                    File.Copy(centralLog, Path.GetTempPath() + Path.GetFileName(centralLog), true);
                    mailMessage.Attachments.Add(new Attachment(Path.GetTempPath() + Path.GetFileName(centralLog)));
                }
                catch
                {

                }
            }

            //Save the email file
            string emailFile = Path.GetTempPath() + Path.GetRandomFileName() + ".eml";
            mailMessage.Save(emailFile);

            return emailFile;
        }
    }

    /// <summary>
    /// Class for creating an email file and tagging it not sent
    /// </summary>
    public static class EmailUtility
    {
        public static void Save(this MailMessage message, string filename, bool addUnsentHeader = true)
        {
            using (var filestream = File.Open(filename, FileMode.Create))
            {
                if (addUnsentHeader)
                {
                    var binaryWriter = new BinaryWriter(filestream);
                    binaryWriter.Write(System.Text.Encoding.UTF8.GetBytes("X-Unsent: 1" + Environment.NewLine));
                }

                var assembly = typeof(SmtpClient).Assembly;
                var mailWriterType = assembly.GetType("System.Net.Mail.MailWriter");
                var mailWriterConstructor = mailWriterType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(Stream) }, null);
                var mailWriter = mailWriterConstructor.Invoke(new object[] { filestream });
                var sendMethod = typeof(MailMessage).GetMethod("Send", BindingFlags.Instance | BindingFlags.NonPublic);
                sendMethod.Invoke(message, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { mailWriter, true, true }, null);
                var closeMethod = mailWriter.GetType().GetMethod("Close", BindingFlags.Instance | BindingFlags.NonPublic);
                closeMethod.Invoke(mailWriter, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { }, null);
            }
        }
    }
}
