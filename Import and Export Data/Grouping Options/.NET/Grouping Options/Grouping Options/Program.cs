using System.IO;
using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.ComponentModel;

namespace Grouping_Options
{
    class Program
    {
        static void Main(string[] args)
        {
            ImportData();
        }
        //Main method to import data from nested collection to Excel worksheet. 
        private static void ImportData()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            IList<Brand> vehicles = GetVehicleDetails();

            ExcelImportDataOptions importDataOptions = new ExcelImportDataOptions();

            //Imports from 4th row.
            importDataOptions.FirstRow = 4;

            //Imports column headers.
            importDataOptions.IncludeHeader = true;

            //Set layout options.
            importDataOptions.NestedDataLayoutOptions = ExcelNestedDataLayoutOptions.Default;

            //Set grouping option. Available GroupingOptions are Collapse and Expand
            importDataOptions.NestedDataGroupOptions = ExcelNestedDataGroupOptions.Collapse;

            //Set collapse level.
            //GroupingOption must set to ‘Collapse’ before applying ‘CollapseLevel’.
            importDataOptions.CollapseLevel = 2;

            //Import data from the nested collection.
            worksheet.ImportData(vehicles, importDataOptions);

            //Apply style to headers 
            worksheet["A1:C2"].Merge();
            worksheet["A1"].Text = "Automobile Brands in the US";

            worksheet.UsedRange.AutofitColumns();

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream("ImportData.xlsx", FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
        }
        //Helper method to load data from XML file and add them in collections. 
        private static IList<Brand> GetVehicleDetails()
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(BrandObjects));

            //Read data from XML file. 
            FileStream stream = new FileStream(Path.GetFullPath(@"Data/ExportData.xml"), FileMode.Open, FileAccess.Read);
            TextReader textReader = new StreamReader(stream);
            BrandObjects brands = (BrandObjects)deserializer.Deserialize(textReader);

            //Initialize parent collection to add data from XML file. 
            List<Brand> list = new List<Brand>();

            string brandName = brands.BrandObject[0].BrandName;
            string vehicleType = brands.BrandObject[0].VahicleType;
            string modelName = brands.BrandObject[0].ModelName;

            //Parent class 
            Brand brand = new Brand(brandName);
            brand.VehicleTypes = new List<VehicleType>();

            VehicleType vehicle = new VehicleType(vehicleType);
            vehicle.Models = new List<Model>();

            Model model = new Model(modelName);
            brand.VehicleTypes.Add(vehicle);

            list.Add(brand);

            foreach (BrandObject brandObj in brands.BrandObject)
            {
                if (brandName == brandObj.BrandName)
                {
                    if (vehicleType == brandObj.VahicleType)
                    {
                        vehicle.Models.Add(new Model(brandObj.ModelName));
                        continue;
                    }
                    else
                    {
                        vehicle = new VehicleType(brandObj.VahicleType);
                        vehicle.Models = new List<Model>();
                        vehicle.Models.Add(new Model(brandObj.ModelName));
                        brand.VehicleTypes.Add(vehicle);
                        vehicleType = brandObj.VahicleType;
                    }
                    continue;
                }
                else
                {
                    brand = new Brand(brandObj.BrandName);
                    vehicle = new VehicleType(brandObj.VahicleType);
                    vehicle.Models = new List<Model>();
                    vehicle.Models.Add(new Model(brandObj.ModelName));
                    brand.VehicleTypes = new List<VehicleType>();
                    brand.VehicleTypes.Add(vehicle);
                    vehicleType = brandObj.VahicleType;
                    list.Add(brand);
                    brandName = brandObj.BrandName;
                }
            }

            textReader.Close();
            return list;
        }
    }
    //Parent Class 
    public class Brand
    {
        private string m_brandName;

        [DisplayNameAttribute("Brand")]
        public string BrandName
        {
            get { return m_brandName; }
            set { m_brandName = value; }
        }

        //Vehicle Types Collection 
        private IList<VehicleType> m_vehicleTypes;

        public IList<VehicleType> VehicleTypes
        {
            get { return m_vehicleTypes; }
            set { m_vehicleTypes = value; }
        }

        public Brand(string brandName)
        {
            m_brandName = brandName;
        }
    }

    //Child Class 
    public class VehicleType
    {
        private string m_vehicleName;

        [DisplayNameAttribute("Vehicle Type")]
        public string VehicleName
        {
            get { return m_vehicleName; }
            set { m_vehicleName = value; }
        }

        //Models collection 
        private IList<Model> m_models;
        public IList<Model> Models
        {
            get { return m_models; }
            set { m_models = value; }
        }

        public VehicleType(string vehicle)
        {
            m_vehicleName = vehicle;
        }
    }

    //Sub-child Class 
    public class Model
    {
        private string m_modelName;

        [DisplayNameAttribute("Model")]
        public string ModelName
        {
            get { return m_modelName; }
            set { m_modelName = value; }
        }

        public Model(string name)
        {
            m_modelName = name;
        }
    }

    //Helper Classes 
    [XmlRootAttribute("BrandObjects")]
    public class BrandObjects
    {
        [XmlElement("BrandObject")]
        public BrandObject[] BrandObject { get; set; }
    }

    public class BrandObject
    {
        public string BrandName { get; set; }
        public string VahicleType { get; set; }
        public string ModelName { get; set; }
    }
}

