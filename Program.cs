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

    Console.WriteLine();
    Console.WriteLine($"Processing bounties up to #{response.Values.Count + 1}");

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

        Console.Write("Accept? (enter to open browser) ");
        string? acceptText = Console.ReadLine();
        string? amount = string.Empty;

        if (string.IsNullOrEmpty(acceptText))
        {
            Process.Start("open", row[5].ToString() ?? throw new InvalidOperationException());
            goto ask;
        }

        if (acceptText == "yes")
        {
            while (string.IsNullOrEmpty(amount))
            {
                Console.Write("Amount: ");
                amount = Console.ReadLine();
            }
        }

        var update = new ValueRange
        {
            Values = new Collection<IList<object>>
            {
                new List<object> { acceptText, amount }
            }
        };

        var updateRequest = service.Spreadsheets.Values.Update(update, spreadsheet_id, $"Responses!L{rowNumber}:N{rowNumber}");
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.Execute();
    }
}
