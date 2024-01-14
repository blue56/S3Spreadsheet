using Amazon.Lambda.Core;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace S3Spreadsheet;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(Request Request, ILambdaContext context)
    {
        Spreadsheet spreadsheet = new Spreadsheet(Request.Region);
        
        spreadsheet.CreateV2(Request.Bucketname, Request.Source, Request.Result);

        return "OK";
    }
}
