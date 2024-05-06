codeunit 51010 "SharePoint File Management"
{
    trigger OnRun()
    begin
    end;

    var
        SharePointClient: Codeunit "SharePoint Client";
        FileMgt: Codeunit "File Management";
        Authorization: Interface "SharePoint Authorization";
        FileDoesNotExistErr: Label 'The file %1 does not exist.', Comment = '%1 File Path';
        DirectoryDoesNotExistErr: Label 'The directory %1 does not exist.', Comment = '%1 Directory Path';
        FileAlreadyExistErr: Label 'The file name %1 already exists.', Comment = '%1 File Path';
        InvalidFileNameErr: Label 'Invalid file name %1.', Comment = '%1 File Name';
        NoFileExtensionErr: Label 'No file %1 extension specified.', Comment = '%1 File Name';


    //Gets "SharePoint Connection Setup" table record to initialize connection with SharePoint
    local procedure InitializeConnection()
    var
        SharePointSetup: Record "SharePoint Connector Setup";
    begin
        SharePointSetup.Get();
        InitializeConnection(SharePointSetup."Sharepoint URL", SharePointSetup."Client ID", SharePointSetup."Client Secret");
    end;

    //Connection is initialized out of procedure parameters. Can be accesed outside of codeunit.
    procedure InitializeConnection(SharePointUrl: Text[250]; ClientID: Text[250]; ClientSecret: Text[250]): Boolean
    var
        AadTenantId: Text;
        Diag: Interface "HTTP Diagnostics";
        SharePointList: Record "SharePoint List" temporary;
    begin
        AadTenantId := GetAadTenantNameFromBaseUrl(SharePointUrl); //Used to get an Azure Active Directory ID from a URL
        Authorization := GetSharePointAuthorization(AadTenantId, ClientID, ClientSecret);
        SharePointClient.Initialize(SharePointUrl, Authorization); //Initializes the client
        if not SharePointClient.GetLists(SharePointList) then //We need to perform at least one action to get diagnostics data
            ErrorCheck('lists get'); //Optional: used to get diagnostics, useful for debugging errors

        exit(true);
    end;

    local procedure GetSharePointAuthorization(AadTenantId: Text; ClientID: Text[250]; ClientSecret: Text[250]): Interface "SharePoint Authorization"
    var
        SharePointAuth: Codeunit "SharePoint Auth.";
        Scopes: List of [Text];
    begin
        Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default'); //Using a default scope provided as an example
        //We return an authorization code that will be used to initialize the Sharepoint Client
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, ClientID, ClientSecret, Scopes));
    end;

    //Gets tenant name from base url. Result will be - host(without .sharepoint.com).onmicrosoft.com
    local procedure GetAadTenantNameFromBaseUrl(BaseUrl: Text): Text
    var
        Uri: Codeunit Uri;
        MySiteHostSuffixTxt: Label '-my.sharepoint.com', Locked = true;
        SharePointHostSuffixTxt: Label '.sharepoint.com', Locked = true;
        OnMicrosoftTxt: Label '.onmicrosoft.com', Locked = true;
        UrlInvalidErr: Label 'The Base Url %1 does not seem to be a valid SharePoint Online Url.', Comment = '%1=BaseUrl';
        Host: Text;
    begin
        // SharePoint Online format:  https://tenantname.sharepoint.com/SiteName/LibraryName/
        // SharePoint My Site format: https://tenantname-my.sharepoint.com/personal/user_name/
        Uri.Init(BaseUrl);
        Host := Uri.GetHost();
        if not Host.EndsWith(SharePointHostSuffixTxt) then
            Error(UrlInvalidErr, BaseUrl);
        if Host.EndsWith(MySiteHostSuffixTxt) then
            exit(CopyStr(Host, 1, StrPos(Host, MySiteHostSuffixTxt) - 1) + OnMicrosoftTxt);
        exit(CopyStr(Host, 1, StrPos(Host, SharePointHostSuffixTxt) - 1) + OnMicrosoftTxt);
    end;

    //Replaces File Management - UploadFile
    procedure UploadFile(DialogTitle: Text[50]; DirectoryPath: Text) FilePath: Text
    var
        InStream: InStream;
        SharePointFile: Record "SharePoint File" temporary;
        SharePointFolder: Record "SharePoint Folder" temporary;
    begin
        UploadIntoStream(DialogTitle, '', '', FilePath, InStream);

        if SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile) then
            FilePath := SharePointFile."Server Relative Url";
    end;

    //Replaces File - Upload, File Management - UploadFileWithFilters
    procedure UploadFileWithFilters(DialogTitle: Text[50]; DirectoryPath: Text; Filter: Text) FilePath: Text
    var
        InStream: InStream;
        SharePointFile: Record "SharePoint File" temporary;
    begin
        UploadIntoStream(DialogTitle, '', Filter, FilePath, InStream);

        if SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile) then
            FilePath := SharePointFile."Server Relative Url";
    end;


    //Replaces File - Create
    procedure CreateFile(FilePath: Text; InStream: InStream): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        DirectoryPath: text;
    begin
        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(FilePath));
        exit(SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile));
    end;

    //Used locally in CreateFile, UploadFile, UploadFileWithFilters, ReplaceFileContent, CopyFile, BLOBExportToFile
    local procedure SaveFile(DirectoryPath: text; FileName: Text; InStream: InStream; var SharePointFile: Record "SharePoint File" temporary) Saved: Boolean
    var
        FilePath: Text;
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        DirectoryPath := GetRelativeUrl(DirectoryPath);

        if FileName = '' then
            Error('No file name provided');

        if FileMgt.GetExtension(FileName) = '' then
            Error(NoFileExtensionErr, FileName);

        if not DirectoryExists(DirectoryPath) then
            Error(DirectoryDoesNotExistErr, DirectoryPath);

        FilePath := DirectoryPath + '/' + FileName;

        if FileExists(FilePath) then
            Error(FileAlreadyExistErr, FilePath);

        Saved := SharePointClient.AddFileToFolder(DirectoryPath, FileName, InStream, SharePointFile);
        if not Saved then
            ErrorCheck('file save');
    end;

    //Replaces File - Download, File Management - DownloadHandler
    procedure DownloadFile(FromFilePath: Text; DialogTitle: Text; ToFileName: Text) Downloaded: Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        FromFilePath := GetRelativeUrl(FromFilePath);

        if not FileExists(FromFilePath) then
            Error(FileDoesNotExistErr, FromFilePath);

        if ToFileName = '' then
            ToFileName := FileMgt.GetFileName(FromFilePath);

        if (FileMgt.GetExtension(ToFileName) = '') or not (FileMgt.GetExtension(ToFileName) = FileExt) then
            ToFileName := FileMgt.GetFileNameWithoutExtension(ToFileName) + FileExt;

        Downloaded := SharePointClient.DownloadFileContentByServerRelativeUrl(FromFilePath, ToFileName);
        if not Downloaded then
            ErrorCheck('file download');
    end;

    //Replaces File - Rename
    procedure RenameFile(OldFilePath: Text; NewFileName: Text) Renamed: Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        OldFileName: Text;
        NewFilePath: Text;
    begin
        OldFilePath := GetRelativeUrl(OldFilePath);
        OldFileName := FileMgt.GetFileNameWithoutExtension(FileMgt.GetFileName(OldFilePath));
        NewFileName := FileMgt.GetFileNameWithoutExtension(NewFileName);

        NewFilePath := OldFilePath.Replace(OldFileName, NewFileName);

        CopyFile(OldFilePath, NewFilePath, false, SharePointFile);
        if SharePointFile.FindFirst() then begin
            DeleteFile(OldFilePath);
            Renamed := true;
        end;
    end;

    //Can be used instead File - Close.
    //Could be harmful. To replace file content procedure deletes file and saves again with changed content.
    //If any error occured while this procedure and file was deleted, it is saved to Sharepoint Documents Temp folder with original content.
    procedure ReplaceFileContent(FilePath: Text; InStream: InStream): Boolean
    var
        DirectoryPath: Text;
        TempDirectoryPath: Text;
        TempSharePointFile: Record "SharePoint File" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        SourceInStream: InStream;
        SourceTempBlob: Codeunit "Temp Blob";
    begin
        FilePath := GetRelativeUrl(FilePath);
        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(FilePath));
        TempDirectoryPath := GetRootFolder() + '/' + 'Temp';

        if not DirectoryExists(TempDirectoryPath) then
            CreateDirectory(TempDirectoryPath);

        if not FileExists(FilePath) then
            Error(FileDoesNotExistErr, FilePath);

        if GetFileContentBlob(FilePath, SourceTempBlob) then begin
            SourceTempBlob.CreateInStream(SourceInStream);
            if SaveFile(TempDirectoryPath, FileMgt.GetFileName(FilePath), SourceInStream, TempSharePointFile) then begin
                DeleteFile(FilePath);
                if SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile) then begin
                    DeleteFile(TempDirectoryPath + '/' + FileMgt.GetFileName(FilePath));
                    exit(true);
                end;
            end;
        end;
        exit(false);
    end;

    //Can be used instead File - Close.
    //Could be harmful. To replace file content procedure deletes file and saves again with changed content.
    //If any error occured while this procedure and file was deleted, it is saved to Sharepoint Documents Temp folder with original content.
    procedure ReplaceFileContent(FilePath: Text; TempBlob: Codeunit "Temp Blob"): Boolean
    var
        DirectoryPath: Text;
        TempDirectoryPath: Text;
        SharePointFile: Record "SharePoint File" temporary;
        TempSharePointFile: Record "SharePoint File" temporary;
        InStream: InStream;
        SourceInStream: InStream;
        SourceTempBlob: Codeunit "Temp Blob";
    begin
        TempBlob.CreateInStream(InStream);

        FilePath := GetRelativeUrl(FilePath);
        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(FilePath));
        TempDirectoryPath := GetRootFolder() + '/' + 'Temp';

        if not DirectoryExists(TempDirectoryPath) then
            CreateDirectory(TempDirectoryPath);

        if not FileExists(FilePath) then
            Error(FileDoesNotExistErr, FilePath);

        if GetFileContentBlob(FilePath, SourceTempBlob) then begin
            SourceTempBlob.CreateInStream(SourceInStream);
            if SaveFile(TempDirectoryPath, FileMgt.GetFileName(FilePath), SourceInStream, TempSharePointFile) then begin
                DeleteFile(FilePath);
                if SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile) then begin
                    DeleteFile(TempDirectoryPath + '/' + FileMgt.GetFileName(FilePath));
                    exit(true);
                end;
            end;
        end;
        exit(false);
    end;

    //Replaces File - Copy
    //Copies an existing file to a new file. Overwriting a file of the same name is allowed.
    procedure CopyFile(SourceFilePath: Text; TargetFilePath: Text; Overwrite: Boolean): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
    begin
        CopyFile(SourceFilePath, TargetFilePath, Overwrite, SharePointFile);
        if SharePointFile.FindFirst() then
            exit(true);
        exit(false);
    end;

    //Replaces File - Copy, File Management - CopyServerFile
    //Copies an existing file to a new file. Overwriting a file of the same name is not allowed.
    procedure CopyFile(SourceFilePath: Text; TargetFilePath: Text): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
    begin
        CopyFile(SourceFilePath, TargetFilePath, false, SharePointFile);
        if SharePointFile.FindFirst() then
            exit(true);
        exit(false);
    end;

    //Locally used in procedures public CopyFile, RenameFile
    local procedure CopyFile(SourceFilePath: Text; TargetFilePath: Text; Overwrite: Boolean; var SharePointFile: Record "SharePoint File" temporary)
    var
        TempBlob: Codeunit "Temp Blob";
        InStream: InStream;
        DirectoryPath: Text;
    begin
        SourceFilePath := GetRelativeUrl(SourceFilePath);
        TargetFilePath := GetRelativeUrl(TargetFilePath);
        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(TargetFilePath));

        if not DirectoryExists(DirectoryPath) then
            Error(DirectoryDoesNotExistErr, DirectoryPath);

        if not FileExists(SourceFilePath) then
            Error(FileDoesNotExistErr, SourceFilePath);

        if FileMgt.GetExtension(FileMgt.GetFileName(TargetFilePath)) = '' then
            Error(NoFileExtensionErr, FileMgt.GetFileName(TargetFilePath));

        if GetFileContentBlob(SourceFilePath, TempBlob) then begin
            TempBlob.CreateInStream(InStream);
            if Overwrite then
                if FileExists(TargetFilePath) then
                    DeleteFile(TargetFilePath);

            SaveFile(DirectoryPath, FileMgt.GetFileName(TargetFilePath), InStream, SharePointFile);
        end;
    end;

    //Replaces File - Exists, File Management - ServerFileExists
    procedure FileExists(FilePath: Text): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        DirectoryPath: Text;
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        FilePath := GetRelativeUrl(FilePath);

        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(FilePath));
        if SharePointClient.GetFolderFilesByServerRelativeUrl(DirectoryPath, SharePointFile) then begin
            SharePointFile.SetRange(Name, FileMgt.GetFileName(FilePath));
            if SharePointFile.FindFirst() then
                exit(true);
        end;
        exit(false);
    end;

    //Replaces File - Erase, File Management - DeleteServerFile
    procedure DeleteFile(FilePath: text): Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        FilePath := GetRelativeUrl(FilePath);

        if not FileExists(FilePath) then
            exit(false);

        if SharePointClient.DeleteFileByServerRelativeUrl(FilePath) then
            exit(true)
        else
            ErrorCheck('file delete');
    end;

    //Replaces File Management - ServerDirectoryExists
    procedure DirectoryExists(DirectoryPath: Text): Boolean
    var
        SharePointList: Record "SharePoint List" temporary;
        SharepointFolder: Record "SharePoint Folder" temporary;
        RelativeUrl: Text;
        Filter: Text;
        FolderName: Text;
        NewSharepointFolder: Record "SharePoint Folder" temporary;
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        RelativeUrl := GetRootFolder();

        if DirectoryPath.StartsWith('/sites') then
            DirectoryPath := DelStr(DirectoryPath, 1, StrLen(RelativeUrl));

        if DirectoryPath <> '' then
            if StrPos(DirectoryPath, '/') = 1 then
                DirectoryPath := DelStr(DirectoryPath, 1, 1);

        repeat
            if DirectoryPath <> '' then begin
                if StrPos(DirectoryPath, '/') = 0 then begin
                    FolderName := DirectoryPath;
                    if SharePointClient.GetSubFoldersByServerRelativeUrl(RelativeUrl, NewSharepointFolder) then begin
                        if NewSharepointFolder.Exists then begin
                            Filter := '@' + RelativeUrl + '/' + FolderName;
                            NewSharepointFolder.SetFilter("Server Relative Url", Filter);
                            if not NewSharepointFolder.FindFirst() then
                                exit(false);
                        end;
                    end;
                    RelativeUrl := NewSharepointFolder."Server Relative Url";
                    DirectoryPath := '';
                end else begin
                    FolderName := CopyStr(DirectoryPath, 1, StrPos(DirectoryPath, '/') - 1);
                    if SharePointClient.GetSubFoldersByServerRelativeUrl(RelativeUrl, NewSharepointFolder) then begin
                        if NewSharepointFolder.Exists then begin
                            Filter := '@' + RelativeUrl + '/' + FolderName;
                            NewSharepointFolder.SetFilter("Server Relative Url", Filter);
                            if not NewSharepointFolder.FindFirst() then
                                exit(false);
                        end;
                    end;
                    RelativeUrl := NewSharepointFolder."Server Relative Url";
                    DirectoryPath := CopyStr(DirectoryPath, StrPos(DirectoryPath, '/') + 1, StrLen(DirectoryPath) - StrLen(CopyStr(DirectoryPath, 1, StrPos(DirectoryPath, '/'))));
                end;
            end;
        until DirectoryPath = '';
        exit(true);
    end;

    //Replaces File Management - ServerCreateDirectory
    procedure CreateDirectory(DirectoryPath: Text) Created: Boolean
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
    begin
        DirectoryPath := GetRelativeUrl(DirectoryPath);

        if not DirectoryExists(DirectoryPath) then
            Created := CreateFolder(DirectoryPath, SharePointFolder)
    end;

    //Used locally in CreateDirectory procedure
    local procedure CreateFolder(ServerRelativeUrl: Text; var SharePointFolder: Record "SharePoint Folder" temporary) Created: Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        Created := SharePointClient.CreateFolder(ServerRelativeUrl, SharePointFolder);
        if not Created then
            ErrorCheck('');
    end;

    //Can be used instead File Management - RemoveServerDirectory
    // Deletes directory and all files stored in it
    procedure RemoveDirectory(DirectoryPath: Text): Boolean
    begin
        DirectoryPath := GetRelativeUrl(DirectoryPath);

        if DirectoryExists(DirectoryPath) then
            exit(DeleteFolder(DirectoryPath));
    end;

    //Replaces File Management - RemoveDirectory procedure
    local procedure DeleteFolder(ServerRelativeUrl: text) Deleted: Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        Deleted := SharePointClient.DeleteFolderByServerRelativeUrl(ServerRelativeUrl);
        if not Deleted then
            ErrorCheck('folder delete');
    end;

    //Replaces File Management - GetServerDirectoryFilesList
    procedure GetDirectoryFilesList(var SharePointFile: Record "SharePoint File"; DirectoryPath: Text)
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        DirectoryPath := GetRelativeUrl(DirectoryPath);

        SharePointFile.Reset();
        SharePointFile.DeleteAll();

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(DirectoryPath, SharePointFile) then
            Error(DirectoryDoesNotExistErr, DirectoryPath);
    end;

    //Replaces File Management - GetServerDirectoryFilesListInclSubDirs
    procedure GetDirectoryFilesListInclSubDirs(var SharePointFile: Record "SharePoint File" temporary; DirectoryPath: Text)
    begin
        DirectoryPath := GetRelativeUrl(DirectoryPath);

        SharePointFile.Reset();
        SharePointFile.DeleteAll();

        if DirectoryExists(DirectoryPath) then
            GetDirectoryFilesListInclSubDirsInner(SharePointFile, DirectoryPath);
    end;

    //Replaces File Management - GetServerDirectoryFilesListInclSubDirsInner
    local procedure GetDirectoryFilesListInclSubDirsInner(var SharePointFile: Record "SharePoint File"; ServerRelativeUrl: Text)
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile2: Record "SharePoint File" temporary;
    begin
        GetDirectoryFilesList(SharePointFile2, ServerRelativeUrl);
        if SharePointFile2.FindSet() then
            repeat
                SharePointFile.Init();
                SharePointFile.Copy(SharePointFile2);
                SharePointFile.Insert();
            until SharePointFile2.Next() = 0;

        GetDirectoryFoldersList(SharePointFolder, ServerRelativeUrl);
        if SharePointFolder.FindSet() then
            repeat
                GetDirectoryFilesListInclSubDirsInner(SharePointFile, SharePointFolder."Server Relative Url");
            until SharePointFolder.Next() = 0;
    end;

    //Used locally in GetDirectoryFilesListInclSubDirsInner procedure
    procedure GetDirectoryFoldersList(var SharePointFolder: Record "SharePoint Folder" temporary; DirectoryPath: Text)
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        DirectoryPath := GetRelativeUrl(DirectoryPath);

        if not SharePointClient.GetSubFoldersByServerRelativeUrl(DirectoryPath, SharePointFolder) then
            Error(DirectoryDoesNotExistErr, DirectoryPath);
    end;

    //Gets single "SharePoint File" record
    procedure GetFile(DirectoryPath: Text; FileName: Text; var SharePointFile: Record "SharePoint File" temporary)
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        DirectoryPath := GetRelativeUrl(DirectoryPath);

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(DirectoryPath, SharePointFile) then
            Error(DirectoryDoesNotExistErr, DirectoryPath);

        SharePointFile.SetRange(Name, FileName);
    end;

    //Replaces File Management - BLOBImportFromFile
    procedure BLOBImportFromFile(var TempBlob: Codeunit "Temp Blob"; FilePath: Text)
    begin
        FilePath := GetRelativeUrl(FilePath);

        if not FileExists(FilePath) then
            Error(FileDoesNotExistErr, FilePath);

        GetFileContentBlob(FilePath, TempBlob);
    end;


    //Can be used instead File - Open
    procedure InStreamImportFromFile(var InStream: InStream; FilePath: Text): Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        FilePath := GetRelativeUrl(FilePath);

        if not FileExists(FilePath) then
            Error(FileDoesNotExistErr, FilePath);

        if SharePointClient.DownloadFileContentByServerRelativeUrl(FilePath, InStream) then
            exit(true)
        else
            ErrorCheck('download file content instream');
        exit(false);
    end;

    //Used locally in GetFileContents, ReplaceFileContent, CopyFile
    local procedure GetFileContentBlob(ServerRelativeUrl: Text; var TempBlob: Codeunit "Temp Blob"): Boolean
    begin
        if not ConnectionIsInitialized() then
            InitializeConnection();

        if SharePointClient.DownloadFileContentByServerRelativeUrl(ServerRelativeUrl, TempBlob) then
            exit(true)
        else
            ErrorCheck('download file content instream');
        exit(false);
    end;

    //Replaces File Management - BLOBExportToFile
    procedure BLOBExportToFile(var TempBlob: Codeunit "Temp Blob"; FilePath: Text): Boolean
    var
        InStream: InStream;
        DirectoryPath: Text;
        SharePointFile: Record "SharePoint File" temporary;
    begin
        FilePath := GetRelativeUrl(FilePath);

        TempBlob.CreateInStream(InStream);

        if FileExists(FilePath) then
            Error(FileAlreadyExistErr, FilePath);

        DirectoryPath := FixPathForSharePoint(FileMgt.GetDirectoryName(FilePath));
        exit(SaveFile(DirectoryPath, FileMgt.GetFileName(FilePath), InStream, SharePointFile));
    end;

    //Replaces File Management - IsServerDirectoryEmpty
    procedure IsDirectoryEmpty(Path: Text): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
    begin
        Path := GetRelativeUrl(Path);

        if DirectoryExists(Path) then begin
            GetDirectoryFilesList(SharePointFile, Path);
            exit(SharePointFile.IsEmpty);
        end;
        exit(false);
    end;

    //Replaces File Management - GetFileContents
    procedure GetFileContents(FilePath: Text) Result: Text
    begin
        GetFileContents(FilePath, TextEncoding::UTF8);
    end;

    procedure GetFileContents(FilePath: Text; Encoding: TextEncoding) Result: Text
    var
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        FilePath := GetRelativeUrl(FilePath);

        if not FileExists(FilePath) then
            exit;

        GetFileContentBlob(FilePath, TempBlob);
        TempBlob.CreateInStream(InStr, Encoding);
        InStr.Read(Result);
    end;

    //Gets SharePoint site Documents server relative url
    procedure GetRootFolder() FolderPath: Text
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
    begin
        GetRootFolder(SharePointFolder);
        if SharePointFolder.FindFirst() then
            FolderPath := SharePointFolder."Server Relative Url"
        else
            FolderPath := '';
    end;

    //Gets SharePoint site Documents as "SharePoint Folder"
    local procedure GetRootFolder(var SharePointFolder: Record "SharePoint Folder" temporary)
    var
        SharePointList: Record "SharePoint List" temporary;
    begin
        if not ConnectionIsInitialized() then //Checking for connection
            InitializeConnection(); // calling InitializeConnection when it is not initialized
        if SharePointClient.GetLists(SharePointList) then begin //GetLists writes data to SharePointList
            SharePointList.SetRange(Title, 'Documents'); //We filter out the Documents list to get a list of files.
            if SharePointList.FindFirst() then
                //We then proceed the folder of the Documents list so that we may get a list of files within the "Documents" directory
                if not SharePointClient.GetDocumentLibraryRootFolder(SharePointList.OdataId, SharePointFolder) then
                    ErrorCheck('lists get');
        end;
    end;

    //Used locally to fix path after "File Management" GetDirectoryName procedure result.
    //Could be used outside codeunit before calling other procedures.
    procedure FixPathForSharePoint(Path: Text): Text
    var
        SPPath: Text;
    begin
        SPPath := Path.Replace('\', '/');
        SPPath := SPPath.Replace('//', '/');
        if StrPos(SPPath, ':') > 0 then
            SPPath := DelStr(SPPath, 1, StrPos(SPPath, ':'));
        if SPPath.EndsWith('/') then
            SPPath := DelStr(SPPath, StrLen(SPPath), 1);
        exit(SPPath);
    end;

    //Forms path if it is not start with SharePoint site Documents server relative url.
    procedure GetRelativeUrl(Path: Text): Text
    var
        RootFolder: Text;
    begin
        RootFolder := GetRootFolder();
        if not (RootFolder = '') then
            if not Path.StartsWith(RootFolder) then begin
                if StrPos(Path, '/') = 1 then
                    Path := RootFolder + Path
                else
                    Path := RootFolder + '/' + Path;
            end;
        exit(Path);
    end;

    //Used locally in every procedure before "SharePoint Client" is used
    local procedure ConnectionIsInitialized(): Boolean
    var
        Diag: Interface "HTTP Diagnostics";
    begin
        Diag := SharePointClient.GetDiagnostics();
        exit(Diag.IsSuccessStatusCode())
    end;

    //Used locally in procedures where "SharePoint Client" is used
    local procedure ErrorCheck(Process: Text)
    var
        Diag: Interface "HTTP Diagnostics";
        SaveFailedErr: Label 'SharePoint Client get failed in %5.\ErrorMessage: %1\HttpRetryAfter: %2\HttpStatusCode: %3\ResponseReasonPhrase: %4', Comment = '%1=GetErrorMessage; %2=GetHttpRetryAfter; %3=GetHttpStatusCode; %4=GetResponseReasonPhrase';
    begin
        Diag := SharePointClient.GetDiagnostics();
        if not Diag.IsSuccessStatusCode() then
            Error(SaveFailedErr, Diag.GetErrorMessage(), Diag.GetHttpRetryAfter(), Diag.GetHttpStatusCode(), Diag.GetResponseReasonPhrase(), Process);
    end;
}

