using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace ExcelDataExport
{
    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            // Register your Syncfusion license key here (get one from your Syncfusion account).
            // Without it the exported file shows a trial/evaluation message.
            // Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");

            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
        }
    }
}
