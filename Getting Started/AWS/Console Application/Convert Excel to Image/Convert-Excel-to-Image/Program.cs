using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Amazon;
using Amazon.Lambda;
using Amazon.Lambda.Model;
using Newtonsoft.Json;

namespace Convert_Excel_to_Image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter your AWS Access Key ID :");
            string awsAccessKeyID = Console.ReadLine();
            Console.WriteLine("Please enter your AWS Secret Access Key :");
            string awsSecretAccessKey = Console.ReadLine();
            Console.WriteLine("Please enter your Function Name :");
            string functionName = Console.ReadLine();
            //Create a new AmazonLambdaClient
            AmazonLambdaClient client = new AmazonLambdaClient(awsAccessKeyID, awsSecretAccessKey, RegionEndpoint.USEast1);

            //Create new InvokeRequest with published function name.
            InvokeRequest invoke = new InvokeRequest
            {
                FunctionName = functionName,
                InvocationType = InvocationType.RequestResponse,
                Payload = "\"Test\""
            };
            //Get the InvokeResponse from client InvokeRequest.
            InvokeResponse response = client.Invoke(invoke);

            //Read the response stream
            var stream = new StreamReader(response.Payload);
            JsonReader reader = new JsonTextReader(stream);
            var serilizer = new JsonSerializer();
            var responseText = serilizer.Deserialize(reader);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(responseText.ToString());
            FileStream fileStream = new FileStream("Image.Jpeg", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
            System.Diagnostics.Process.Start("Image.Jpeg");
        }
    }
}
