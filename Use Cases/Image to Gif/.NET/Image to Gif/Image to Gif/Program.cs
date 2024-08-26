using ImageMagick;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Image_to_Gif
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                
                //Load an input template
                using(FileStream file = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = application.Workbooks.Open(file);

                    using (MagickImageCollection collection = new MagickImageCollection())
                    {
                        for (int i = 0; i < workbook.Worksheets.Count; i++)
                        {
                            IWorksheet worksheet = workbook.Worksheets[i];

                            //Initialize XlsIORenderer
                            application.XlsIORenderer = new XlsIORenderer();

                            //Create a memory stream to save the image
                            using (MemoryStream stream = new MemoryStream())
                            {
                                //Convert worksheet to image and save it to stream
                                worksheet.ConvertToImage(1,1,27,11, stream);

                                //Set stream postition
                                stream.Position = 0;

                                //Load the image into MagickImage
                                MagickImage image = new MagickImage(stream);

                                //Add the image to the collection
                                collection.Add(image);
                            }
                        }

                        //Set the delay between frames (100 = 1 second)
                        foreach (MagickImage image in collection)
                        {
                            image.AnimationDelay = 100;
                        }

                        //Create an animated GIF
                        collection.Write(@"Output.gif");
                    }
                }
            }
        }
    }
}





