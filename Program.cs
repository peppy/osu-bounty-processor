// See https://aka.ms/new-console-template for more information

using System.Collections.ObjectModel;
using System.Diagnostics;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

UserCredential credential;

const string credentials_file = "credentials.json";

using (var stream = new FileStream(credentials_file, FileMode.Open, FileAccess.Read))
{
    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        GoogleClientSecrets.FromStream(stream).Secrets,
        new[] { SheetsService.Scope.Spreadsheets },
        "user",
        CancellationToken.None,
        new FileDataStore("token.json", true)).Result;
}

using (
    var service = new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "bounty-processor",
    }))
{
    const string spreadsheet_id = "1x8SaXZ9DFL9PARkmDBmffPC33lzc7tR3Dy-cL5Emo5c";
    const string range = "Responses!A:N";

    ValueRange response = service.Spreadsheets.Values.Get(spreadsheet_id, range).Execute();

    for (var i = 0; i < response.Values.Count; i++)
    {
        var row = response.Values[i];
        int rowNumber = i + 1;

        if (row.Count >= 12 && !string.IsNullOrEmpty(row[11].ToString()))
            continue; // already dealt with.

        Console.ForegroundColor = ConsoleColor.Gray;

        Console.WriteLine();
        Console.WriteLine("---");
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Bounty #{rowNumber}");
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine($"Date: {row[0]}");
        Console.WriteLine($"Email: {row[1]}");
        Console.WriteLine($"Github: {row[2]}");
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Project: {row[4]}");
        Console.WriteLine($"Time spent: {row[6]}");
        Console.WriteLine($"Compensation: {row[7]}");
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine($"{row[5]}");

        ask:

        Console.Write("Response (enter to open browser): ");
        var responseText = Console.ReadLine();

        if (string.IsNullOrEmpty(responseText))
        {
            Process.Start("open", row[5].ToString() ?? throw new InvalidOperationException());
            goto ask;
        }


        var update = new ValueRange
        {
            Values = new Collection<IList<object>>
            {
                new List<object> { responseText }
            }
        };

        var updateRequest = service.Spreadsheets.Values.Update(update, spreadsheet_id, $"Responses!L{rowNumber}");
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        updateRequest.Execute();
    }
}
