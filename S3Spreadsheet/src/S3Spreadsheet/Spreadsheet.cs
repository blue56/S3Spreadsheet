using System.Text.Json.Nodes;
using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using LargeXlsx;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using Newtonsoft.Json.Linq;

namespace S3Spreadsheet;

public class Spreadsheet
{
    private IAmazonS3 _s3Client;

    private RegionEndpoint _region = null;

    public Spreadsheet(string Region)
    {
        // Set the AWS region where your S3 bucket is located
        _region = RegionEndpoint.GetBySystemName(Region);

        // Create an S3 client
        _s3Client = new AmazonS3Client(_region);
    }

    /// <summary>
    /// Create spreadsheet based on json
    /// </summary>
    public void Create(string Bucketname, string Source, string Result)
    {
        // Fetch json document 

        // Parse json file
        // Read JSON file content
        string content = GetFileContentFromS3(Bucketname, Source).Result;

        JObject node = JObject.Parse(content);

        var stream = new MemoryStream();

        var sheetsNode = (JObject)node["Sheets"];

        foreach (var sheet in sheetsNode)
        {
            string sheetName = sheet.Key;

            var sheetObject = sheetsNode[sheet.Key];

            // Rows is a Json array
            // Each element in array is a row, represented by a element
            // with properties.
            var rowsArray = sheetObject["Rows"];

            var rowList = new List<Dictionary<string, object>>();

            foreach (JObject n in rowsArray.Children())
            {
                //  { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                var rowvalues = new Dictionary<string, object>();

                foreach (var kv in n)
                {
                    string column = kv.Key;
                    var columnValue = kv.Value;

                    rowvalues.Add(column, columnValue);
                }

                rowList.Add(rowvalues);
            }

            var config = new OpenXmlConfiguration
            {
                DynamicColumns = new DynamicExcelColumn[] {
                    new DynamicExcelColumn("id"){Ignore=true},
                    new DynamicExcelColumn("name"){Index=1,Width=10},
                    new DynamicExcelColumn("createdate"){Index=0,Format="yyyy-MM-dd",Width=15},
                    new DynamicExcelColumn("point"){Index=2,Name="Account Point"},
                }
            };

            MiniExcel.SaveAs(stream, rowList, true, sheetName);
        }

        SaveFile(_s3Client, Bucketname, Result, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    public void CreateV2(string Bucketname, string Source, string Result)
    {
        // Fetch json document 

        // Parse json file
        // Read JSON file content
        string content = GetFileContentFromS3(Bucketname, Source).Result;

        JObject node = JObject.Parse(content);

        var stream = new MemoryStream();

        var xlsxWriter = new XlsxWriter(stream);
        /*        using var xlsxWriter = new XlsxWriter(stream);
                xlsxWriter
                    .BeginWorksheet("Sheet 1")
                    .BeginRow().Write("Name").Write("Location").Write("Height (m)")
                    .BeginRow().Write("Kingda Ka").Write("Six Flags Great Adventure").Write(139)
                    .BeginRow().Write("Top Thrill Dragster").Write("Cedar Point").Write(130)
                    .BeginRow().Write("Superman: Escape from Krypton").Write("Six Flags Magic Mountain").Write(126);
        */

        var sheetsNode = (JObject)node["Sheets"];

        foreach (var sheet in sheetsNode)
        {
            string sheetName = sheet.Key;

            xlsxWriter = xlsxWriter.BeginWorksheet(sheetName);

            var sheetObject = sheetsNode[sheet.Key];

            // Rows is a Json array
            // Each element in array is a row, represented by a element
            // with properties.
            var rowsArray = sheetObject["Rows"];

            foreach (JArray cellArray in rowsArray.Children())
            {
                // New row
                xlsxWriter = xlsxWriter.BeginRow();

                foreach (var c in cellArray.Children())
                {
                    if (c["Value"] != null)
                    {
                        var v = c["Value"];

                        if (v.Type == JTokenType.Integer)
                        {
                            int columnValue = v.ToObject<int>();
                            xlsxWriter = xlsxWriter.Write(columnValue);
                        }
                        else if (v.Type == JTokenType.Float)
                        {
                            float columnValue = v.ToObject<float>();
                            xlsxWriter = xlsxWriter.Write(columnValue);
                        }
                        else
                        {
                            string columnValue = v.ToString();
                            xlsxWriter = xlsxWriter.Write(columnValue);
                        }
                    }
                    else if (c["Formula"] != null)
                    {
                        string formula = c["Formula"].ToString();

                        xlsxWriter = xlsxWriter.WriteFormula(formula);
                    }
                }
            }
        }

        xlsxWriter.Dispose();

        SaveFile(_s3Client, Bucketname, Result, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }


    public void SaveFile(IAmazonS3 _s3Client, string Bucketname,
        string S3Path, Stream Stream, string ContentType)
    {
        var putRequest = new PutObjectRequest
        {
            BucketName = Bucketname,
            Key = S3Path,
            ContentType = ContentType,
            InputStream = Stream
        };

        _s3Client.PutObjectAsync(putRequest).Wait();
    }

    private async Task<string> GetFileContentFromS3(string bucketName, string key)
    {
        try
        {
            GetObjectRequest request = new GetObjectRequest
            {
                BucketName = bucketName,
                Key = key
            };

            using (GetObjectResponse response = await _s3Client.GetObjectAsync(request))
            using (Stream responseStream = response.ResponseStream)
            using (StreamReader reader = new StreamReader(responseStream))
            {
                return await reader.ReadToEndAsync();
            }
        }
        catch (AmazonS3Exception e)
        {
            // Handle S3 exception
            return $"Error getting template: {e.Message}";
        }
    }
}