using System.Web;
using System.Web.Mvc;

namespace Convert_Excel_to_PDF
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
