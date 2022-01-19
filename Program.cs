// See https://aka.ms/new-console-template for more information

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

    Console.WriteLine(response);
}
