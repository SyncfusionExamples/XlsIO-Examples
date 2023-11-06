// See https://aka.ms/new-console-template for more information
//Reads the template Excel document.
using System.Net;

FileStream imageStream = new FileStream("../../../AdventureCycles-Logo.png", FileMode.Open, FileAccess.Read);
imageStream.Position = 0;

//Saves the Excel document in memory stream.
MemoryStream inputStream = new MemoryStream();
imageStream.CopyTo(inputStream);
inputStream.Position = 0;

try
{
    Console.WriteLine("Please enter your Azure Functions URL :");
    string functionURL = Console.ReadLine();

    //Create HttpWebRequest with hosted azure functions URL.                
    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(functionURL);

    //Set request method as POST
    req.Method = "POST";

    //Get the request stream to save the Excel document stream
    Stream stream = req.GetRequestStream();

    //Write the Excel document stream into request stream
    stream.Write(inputStream.ToArray(), 0, inputStream.ToArray().Length);

    //Gets the responce from the Azure Functions.
    HttpWebResponse res = (HttpWebResponse)req.GetResponse();

    //Saves the Excel stream.
    FileStream excelStream = File.Create("Sample.xlsx");
    res.GetResponseStream().CopyTo(excelStream);

    //Dispose the streams
    inputStream.Dispose();
    excelStream.Dispose();
}
catch (Exception ex)
{
    throw;
}
