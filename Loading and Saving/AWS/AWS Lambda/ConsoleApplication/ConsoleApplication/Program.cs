﻿using Amazon.Lambda.Model;
using Amazon.Lambda;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Amazon;

namespace ConsoleApplication
{
    internal class Program
    {
        static async Task Main(string[] args)
        {      
            //Create a new AmazonLambdaClient
            AmazonLambdaClient client = new AmazonLambdaClient("awsaccessKeyID", "awsSecreteAccessKey", RegionEndpoint.USEast1);

            //Create new InvokeRequest with published function name.
            InvokeRequest invoke = new InvokeRequest
            {
                FunctionName = "MyNewFunction",
                InvocationType = InvocationType.RequestResponse,
                Payload = "\"Test\""
            };

            // Get the InvokeResponse from client InvokeRequest.
            InvokeResponse response = await client.InvokeAsync(invoke);   

            //Read the response stream
            var stream = new StreamReader(response.Payload);
            JsonReader reader = new JsonTextReader(stream);
            var serilizer = new JsonSerializer();
            var responseText = serilizer.Deserialize(reader);     

            //Convert Base64String into Excel document
            byte[] bytes = Convert.FromBase64String(responseText.ToString());         
            FileStream fileStream = new FileStream("Sample.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close(); 
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "Sample.xlsx",
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
    }
}