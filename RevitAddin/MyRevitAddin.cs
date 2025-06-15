using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;

namespace MyRevitAddin
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            const string tabName = "Custom Tab";
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {
                // Tab might already exist
            }

            RibbonPanel panel = app.CreateRibbonPanel(tabName, "Main Panel");

            PulldownButtonData pbd = new PulldownButtonData("pulldown", "Choose UI");
            PulldownButton pulldown = panel.AddItem(pbd) as PulldownButton;

            for (int i = 1; i <= 5; i++)
            {
                PushButtonData pData = new PushButtonData($"cmd{i}", $"UI {i}", typeof(App).Assembly.Location, typeof(DummyCommand).FullName);
                pulldown.AddPushButton(pData);
            }

            for (int i = 6; i <= 10; i++)
            {
                PushButtonData pData = new PushButtonData($"cmd{i}", $"UI {i}", typeof(App).Assembly.Location, typeof(DummyCommand).FullName);
                panel.AddItem(pData);
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class DummyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("MyRevitAddin", "Button executed.");
            return Result.Succeeded;
        }
    }
}
