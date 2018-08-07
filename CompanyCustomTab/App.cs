using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace CompanyCustomTab
{
    /// <summary>
    /// Startup Class for Revit Addin
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        static AddInId appId = new AddInId(new Guid("4493C2D7-6FB2-4D8B-A889-7CC21EB9F3B7"));

        /// <summary>
        /// Revit Startup Commands
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication application)
        {
            // Add Ribbon Tab, Panels, and Buttons
            AddRibbon(application);
            return Result.Succeeded;
        }

        /// <summary>
        /// Revit Shutdown Commands
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Add Ribbon Tab, Panels, and Buttons
        /// </summary>
        /// <param name="app"></param>
        private void AddRibbon(UIControlledApplication app)
        {
            string dll = Assembly.GetExecutingAssembly().Location;

            // CHANGESETTINGS: Change the Ribbon Tab to your company name in Properties-> Settings -> RevitTab
            string ribbonTab = Properties.Settings.Default.RevitTab;
            app.CreateRibbonTab(ribbonTab);

            // Add some Ribbone Panels
            RibbonPanel panelSupport = app.CreateRibbonPanel(ribbonTab, "Support");
            RibbonPanel panelWebsites = app.CreateRibbonPanel(ribbonTab, "Websites");
            RibbonPanel panelTools = app.CreateRibbonPanel(ribbonTab, "Tools");

            // Button for EmailSupport
            // CHANGESETTINGS: Change the email address for support in Properties-> Settings -> EmailSupport
            PushButton btnEmailSupport = (PushButton)panelSupport.AddItem(new PushButtonData("email_support", "Email\nSupport", dll, "CompanyCustomTab.Email"));
            btnEmailSupport.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.Email32.png");
            btnEmailSupport.Image = GetEmbeddedImage("CompanyCustomTab.Images.Email32.png");
            btnEmailSupport.ToolTip = "Create a new email to email to support";

            // Split Button for Journal file and folder
            PushButtonData btnJournalFile = new PushButtonData("JournalFile", "Open\nJournal File", dll, "CompanyCustomTab.JournalFile");
            PushButtonData btnJournalFolder = new PushButtonData("JournalFolder", "Open\nJournal Folder", dll, "CompanyCustomTab.JournalFolder");
            btnJournalFile.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.JournalFile32.png");
            btnJournalFolder.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.JournalFolder32.png");
            btnJournalFile.Image = GetEmbeddedImage("CompanyCustomTab.Images.JournalFile16.png");
            btnJournalFolder.Image = GetEmbeddedImage("CompanyCustomTab.Images.JournalFolder16.png");
            btnJournalFile.ToolTip = "Open the current model's journal file in notepad";
            btnJournalFolder.ToolTip = "Open the folder containing the journal files for this version of Revit";
            SplitButtonData sbdJournal = new SplitButtonData("splitJournal", "Journal File & Folder");
            SplitButton sbJournal = panelSupport.AddItem(sbdJournal) as SplitButton;
            sbJournal.AddPushButton(btnJournalFile);
            sbJournal.AddPushButton(btnJournalFolder);

            // Button for Company Intranet
            // CHANGESETTINGS: Change the address to your company's intranet address in Properties-> Settings -> IntranetAddress
            // CHANGESETTINGS: Change the name of this button in Properties-> Settings -> IntranetName
            PushButton btnIntranet = (PushButton)panelWebsites.AddItem(new PushButtonData("intranet", Properties.Settings.Default.IntranetName, dll, "CompanyCustomTab.CompanyIntranet"));
            btnIntranet.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.Intranet32.png");
            btnIntranet.Image = GetEmbeddedImage("CompanyCustomTab.Images.Intranet16.png");
            btnIntranet.ToolTip = "Launch the company's intranet";

            // Button for website #1
            // CHANGESETTINGS: Change the web address for this website in Properties-> Settings -> Website1Address
            // CHANGESETTINGS: Change the name of this button in Properties-> Settings -> Website1Name
            PushButton btnWebsite1 = (PushButton)panelWebsites.AddItem(new PushButtonData("website1", Properties.Settings.Default.Website1Name, dll, "CompanyCustomTab.Website1"));
            btnWebsite1.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.Website32.png");
            btnWebsite1.Image = GetEmbeddedImage("CompanyCustomTab.Images.Website16.png");
            btnWebsite1.ToolTip = "Launch " + Properties.Settings.Default.Website1Name + " website";

            // Button for website #2
            // CHANGESETTINGS: Change the web address for this website in Properties-> Settings -> Website2Address
            // CHANGESETTINGS: Change the name of this button in Properties-> Settings -> Website2Name
            PushButton btnWebsite2 = (PushButton)panelWebsites.AddItem(new PushButtonData("website2", Properties.Settings.Default.Website2Name, dll, "CompanyCustomTab.Website2"));
            btnWebsite2.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.Website32.png");
            btnWebsite2.Image = GetEmbeddedImage("CompanyCustomTab.Images.Website16.png");
            btnWebsite2.ToolTip = "Launch " + Properties.Settings.Default.Website2Name + " website";

            // Button for website #3
            // CHANGESETTINGS: Change the web address for this website in Properties-> Settings -> Website3Address
            // CHANGESETTINGS: Change the name of this button in Properties-> Settings -> Website3Name
            PushButton btnWebsite3 = (PushButton)panelWebsites.AddItem(new PushButtonData("website3", Properties.Settings.Default.Website3Name, dll, "CompanyCustomTab.Website3"));
            btnWebsite3.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.Website32.png");
            btnWebsite3.Image = GetEmbeddedImage("CompanyCustomTab.Images.Website16.png");
            btnWebsite3.ToolTip = "Launch " + Properties.Settings.Default.Website3Name + " website";

            // Button for Revit Server Admin
            // CHANGESETTINGS: Set the Revit Server Name in Properties-> Settings -> RevitServerAdmin
            // CHANGESETTINGS: If no Revit Server in company, change setting to FALSE for Properties-> Settings -> ShowRevitServerAdmin
            if (Properties.Settings.Default.ShowRevitServerAdmin)
            {
                PushButton btnRevitServer = (PushButton)panelTools.AddItem(new PushButtonData("revit server", "Revit Server\nAdmin", dll, "CompanyCustomTab.RevitServer"));
                btnRevitServer.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.RevitServer32.png");
                btnRevitServer.Image = GetEmbeddedImage("CompanyCustomTab.Images.RevitServer16.png");
                btnRevitServer.ToolTip = "Open Revit Server Admin interface in a web browser";
            }

            // project folder
            PushButton btnProjectFolder = (PushButton)panelTools.AddItem(new PushButtonData("project folder", "Project\nFolder", dll, "CompanyCustomTab.ProjectFolder"));
            btnProjectFolder.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.ProjectFolder32.png");
            btnProjectFolder.Image = GetEmbeddedImage("CompanyCustomTab.Images.ProjectFolder16.png");
            btnProjectFolder.ToolTip = "Open Windows Explorer to the central file's location";

            // Created & Last Edited By
            PushButton btnCreatedEditedBy = (PushButton)panelTools.AddItem(new PushButtonData("created edited by", "Created &\nLast Edited By", dll, "CompanyCustomTab.CreatedLastEdited"));
            btnCreatedEditedBy.LargeImage = GetEmbeddedImage("CompanyCustomTab.Images.CreatedEditedBy32.png");
            btnCreatedEditedBy.Image = GetEmbeddedImage("CompanyCustomTab.Images.CreatedEditedBy16.png");
            btnCreatedEditedBy.ToolTip = "Use this tool to select an Element to find out who Created and Last Edited it.";
        }

        /// <summary>
        /// Retrieves embedded images for button icons
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        static BitmapSource GetEmbeddedImage(string name)
        {
            try
            {
                Assembly a = Assembly.GetExecutingAssembly();
                Stream s = a.GetManifestResourceStream(name);
                return BitmapFrame.Create(s);
            }
            catch
            {
                return null;
            }
        }
    }
}
