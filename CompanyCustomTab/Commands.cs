using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using System.IO;

namespace CompanyCustomTab
{
    /// <summary>
    /// Email Support
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Email : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            Application app = commandData.Application.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                EmailSupport emailSupport = new EmailSupport(doc, app);
                string email = emailSupport.CreateEmail();

                Process.Start(email);
                return Result.Succeeded;
            }
            catch 
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Open Journal File for Current Model
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class JournalFile : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            Application app = commandData.Application.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                Process.Start(app.RecordingJournalFilename);
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }

        }
    }

    /// <summary>
    /// Open Journal Folder
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class JournalFolder : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            Application app = commandData.Application.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                Process.Start(Path.GetDirectoryName(app.RecordingJournalFilename));
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }

        }
    }

    /// <summary>
    /// Company Intranet Launcher
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CompanyIntranet : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                Process.Start("http://" + Properties.Settings.Default.IntranetAddress);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Website 1 Launcher
    /// CHANGESETTINGS: Change the website address in Properties-> Settings -> Website1Address
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Website1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                Process.Start("http://" + Properties.Settings.Default.Website1Address);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Website 2 Launcher
    /// CHANGESETTINGS: Change the website address in Properties-> Settings -> Website2Address
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Website2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                Process.Start("http://" + Properties.Settings.Default.Website2Address);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Website 3 Launcher
    /// CHANGESETTINGS: Change the website address in Properties-> Settings -> Website2Address
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Website3 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                Process.Start("http://" + Properties.Settings.Default.Website3Address);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Project Folder
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ProjectFolder : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                string projectPath = null;

                if (doc.GetWorksharingCentralModelPath() != null)
                {
                    projectPath = Path.GetDirectoryName(ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath()));
                }
                else
                {
                    projectPath = Path.GetDirectoryName(doc.PathName);
                }

                Process.Start(projectPath);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }

        }
    }

    /// <summary>
    /// Revit Server
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RevitServer : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                // CHANGESETTINGS: Set the Revit Server Name in Properties-> Settings -> RevitServerAdmin
                string rs = Properties.Settings.Default.RevitServerAdmin;
#if RELEASE2017
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://" + rs + "/revitserveradmin2017/");
#elif RELEASE2018
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://" + rs + "/revitserveradmin2018/");
#elif RELEASE2019
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://" + rs + "/revitserveradmin2019/");
#endif
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Show a dialog for who created and last edited an object
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreatedLastEdited : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            if (doc.IsWorkshared)
            {
                try
                {
                    Reference reference = uidoc.Selection.PickObject(ObjectType.Element, "Select an element to get info about who created and last edited it:");
                    Element element = doc.GetElement(reference);
                    ElementId elementId = element.Id;

                    WorksharingTooltipInfo worksharingTooltipInfo = WorksharingUtils.GetWorksharingTooltipInfo(doc, elementId);

                    TaskDialog taskDialog = new TaskDialog("Element Created and Last Edited By");
                    taskDialog.MainInstruction = "Element: " + element.Name + "\n"
                        + "ElementId: " + elementId.ToString() + "\n\n"
                        + "Created by: " + worksharingTooltipInfo.Creator + "\n"
                        + "Last Edited by: " + worksharingTooltipInfo.LastChangedBy;
                    taskDialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
                    taskDialog.CommonButtons = TaskDialogCommonButtons.Close;
                    taskDialog.DefaultButton = TaskDialogResult.Close;

                    TaskDialogResult taskDialogResult = taskDialog.Show();
                    
                    return Result.Succeeded;
                }
                catch
                {
                    return Result.Failed;
                }
            }
            else
            {
                TaskDialog.Show("Workshare Info", "This tool only works with Workshare Enabled Models.");
                return Result.Succeeded;
            }

        }
    }
}
