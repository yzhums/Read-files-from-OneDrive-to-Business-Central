pageextension 50100 CustomerListExt extends "Customer List"
{
    actions
    {
        addafter("Sent Emails")
        {
            action(ImportFileFromOneDrive)
            {
                Caption = 'Import File From OneDrive';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = GetActionMessages;

                trigger OnAction()
                var
                    OneDriveHandler: Codeunit OneDriveHandler;
                begin
                    OneDriveHandler.Run();
                end;
            }
        }
    }
}

codeunit 50120 OneDriveHandler
{
    trigger OnRun()
    begin
        ImportFilesFromOneDrive();
    end;

    procedure ImportFilesFromOneDrive()
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Headers: HttpHeaders;
        JsonResponse: JsonObject;
        JsonArray: JsonArray;
        JsonToken: JsonToken;
        JsonTokenLoop: JsonToken;
        JsonValue: JsonValue;
        JsonObjectLoop: JsonObject;
        AuthToken: SecretText;
        OneDriveFolderUrl: Text;
        ResponseText: Text;
        FileName: Text;
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        // Define the OneDrive folder URL

        // delegated permissions
        //OneDriveFolderUrl := 'https://graph.microsoft.com/v1.0/me/drive/root/children';

        // application permissions (replace with the actual user principal name)
        OneDriveFolderUrl := 'https://graph.microsoft.com/v1.0/users/Admin@2qcj3x.onmicrosoft.com/drive/root/children/OneDriveAPITest/children';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(OneDriveFolderUrl);
        HttpRequestMessage.Method := 'GET';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));

        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            // Log the status code for debugging
            //Message('HTTP Status Code: %1', HttpResponseMessage.HttpStatusCode());

            if HttpResponseMessage.IsSuccessStatusCode() then begin
                HttpResponseMessage.Content.ReadAs(ResponseText);
                JsonResponse.ReadFrom(ResponseText);

                if JsonResponse.Get('value', JsonToken) then begin
                    JsonArray := JsonToken.AsArray();

                    foreach JsonTokenLoop in JsonArray do begin
                        JsonObjectLoop := JsonTokenLoop.AsObject();
                        if JsonObjectLoop.Get('name', JsonTokenLoop) then begin
                            JsonValue := JsonTokenLoop.AsValue();
                            if FileName = '' then begin
                                FileName := JsonValue.AsText();
                            end else begin
                                FileName := FileName + '\' + JsonValue.AsText();
                            end;
                        end;
                    end;
                    Message(FileName);
                end;

            end else begin
                //Report errors!
                HttpResponseMessage.Content.ReadAs(ResponseText);
                Error('Failed to fetch files from OneDrive: %1 %2', HttpResponseMessage.HttpStatusCode(), ResponseText);
            end;
        end else
            Error('Failed to send HTTP request to OneDrive');
    end;

    procedure GetOAuthToken() AuthToken: SecretText
    var
        ClientID: Text;
        ClientSecret: Text;
        TenantID: Text;
        AccessTokenURL: Text;
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        ClientID := 'b4fe1687-f1ab-4bfa-b494-0e2236ed50bd';
        ClientSecret := 'huL8Q~edsQZ4pwyxka3f7.WUkoKNcPuqlOXv0bww';
        TenantID := '7e47da45-7f7d-448a-bd3d-1f4aa2ec8f62';
        AccessTokenURL := 'https://login.microsoftonline.com/' + TenantID + '/oauth2/v2.0/token';
        Scopes.Add('https://graph.microsoft.com/.default');
        if not OAuth2.AcquireTokenWithClientCredentials(ClientID, ClientSecret, AccessTokenURL, '', Scopes, AuthToken) then
            Error('Failed to get access token from response\%1', GetLastErrorText());
    end;
}
