using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using PipeAutoCreate.RevitCommand;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace PipeAutoCreate
{
    [Transaction(TransactionMode.Manual)]
    public class RevitApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            application.CreateRibbonTab("工具");

            var panel = application.CreateRibbonPanel("工具", "工具");

            var pushButtonData = new PushButtonData("12", "市政创建", typeof(CreateGeberator).Assembly.Location, "PipeAutoCreate.RevitCommand.CreateGeberator");
            
            var button = panel.AddItem(pushButtonData) as PushButton;

            button.LargeImage = new BitmapImage(new Uri("pack://application:,,,/PipeAutoCreate;component/Pic/市政.png"));
            return Result.Succeeded;
        }
    }
}
