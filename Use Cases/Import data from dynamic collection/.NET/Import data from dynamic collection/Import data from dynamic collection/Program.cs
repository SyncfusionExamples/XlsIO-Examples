using Syncfusion.XlsIO;
using System.Dynamic;

namespace ImportDynamicCollection
{
    /// <summary>
    /// Custom dynamic object class
    /// </summary>
    public class CustomDynamicObject : DynamicObject
    {
        /// <summary>
        /// The dictionary property used store the data
        /// </summary>
        internal Dictionary<string, object> properties = new Dictionary<string, object>();
        /// <summary>
        /// Provides the implementation for operations that get member values.
        /// </summary>
        /// <param name="binder">Get Member Binder object</param>
        /// <param name="result">The result of the get operation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = default(object);

            if (properties.ContainsKey(binder.Name))
            {
                result = properties[binder.Name];
                return true;
            }
            return false;
        }
        /// <summary>
        /// Provides the implementation for operations that set member values.
        /// </summary>
        /// <param name="binder">Set memeber binder object</param>
        /// <param name="value">The value to set to the member</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            properties[binder.Name] = value;
            return true;
        }
        /// <summary>
        /// Return all dynamic member names
        /// </summary>
        /// <returns>the property name list</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return properties.Keys;
        }
    }
    class Program
    {
        /// <summary>
        /// Generates a dynamic collection of data representing members reports
        /// </summary>
        /// <returns>A list of ExpandoObject representing members</returns>
        public static List<ExpandoObject> GetMembersReport()
        {
            List<ExpandoObject> report = new List<ExpandoObject>();

            ExpandoObject obj = new ExpandoObject();
            string propertyName = "Id";
            object propertyValue = 01;
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Name";
            propertyValue = "Karen Fillippe";
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Age";
            propertyValue = 23;
            AddDynamicProperty(obj, propertyName, propertyValue);
            report.Add(obj);

            obj = new ExpandoObject();
            propertyName = "Id";
            propertyValue = 02;
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Name";
            propertyValue = "Andy Bernard";
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Age";
            propertyValue = 20;
            AddDynamicProperty(obj, propertyName, propertyValue);
            report.Add(obj);

            obj = new ExpandoObject();
            propertyName = "Id";
            propertyValue = 03;
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Name";
            propertyValue = "Jim Halpert";
            AddDynamicProperty(obj, propertyName, propertyValue);
            propertyName = "Age";
            propertyValue = 21;
            AddDynamicProperty(obj, propertyName, propertyValue);
            report.Add(obj);

            return report;
        }
        /// <summary>
        /// Adds a dynamic property to an ExpandoObject
        /// </summary>
        /// <param name="dynamicObj">The ExpandoObject to which the property will be added</param>
        /// <param name="propertyName">The name of the property</param>
        /// <param name="propertyValue">The value of the property</param>
        public static void AddDynamicProperty(ExpandoObject dynamicObj, string propertyName, object propertyValue)
        {
            var expandoDict = dynamicObj as IDictionary<string, object>;
            expandoDict[propertyName] = propertyValue;
        }
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the excel application object.
                IApplication application = excelEngine.Excel;

                //The workbook is created
                IWorkbook workbook = application.Workbooks.Create(1);

                //Access the first worksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Import the dynamic collection
                worksheet.ImportData(GetMembersReport(), 1, 1, true);

                //Auto-fit the columns 
                worksheet.UsedRange.AutofitColumns();

                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}
