page 51112 "SharePoint Connector Setup"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "SharePoint Connector Setup";
    ModifyAllowed = true;
    InsertAllowed = false;
    DeleteAllowed = false;

    layout
    {
        area(Content)
        {
            group(Setup)
            {
                field("Client ID"; Rec."Client ID")
                {
                    ApplicationArea = All;
                }
                field("Client Secret"; Rec."Client Secret")
                {
                    ApplicationArea = All;
                    ExtendedDatatype = Masked;
                }
                field("Sharepoint URL"; Rec."Sharepoint URL")
                {
                    ApplicationArea = All;
                }
            }
            group(File)
            {
                field(FileName; FileName)
                {
                    ApplicationArea = All;
                    Caption = 'File Name';
                }
                field(FileContents; FileContents)
                {
                    ApplicationArea = All;
                    Caption = 'File Contents';
                    MultiLine = true;
                }
                field(DirectoryPath; DirectoryPath)
                {
                    ApplicationArea = All;
                    Caption = 'Directory Path';
                }
            }
            part(SCList; "Sharepoint Connector List")
            {
            }
        }
    }

    actions
    {
        area(Processing)
        {
            //CONNECT TO A SITE
            action(Connect)
            {
                ApplicationArea = All;
                Caption = 'Connect to a site';

                trigger OnAction()
                begin
                    if SharePointMgt.InitializeConnection(Rec."Sharepoint URL", Rec."Client ID", Rec."Client Secret") then
                        Message('Connection established');
                end;
            }
            action(GetList)
            {
                ApplicationArea = All;
                Caption = 'Get Files List';

                trigger OnAction()
                var
                    AadTenantId: Text;
                    SCListRec: Record "Sharepoint Connector List";
                    SharePointFile: Record "SharePoint File" temporary;
                begin
                    RootFolderPath := SharePointMgt.GetRootFolder();
                    if not (RootFolderPath = '') then begin
                        SCListRec.DeleteAll(); //List is not a temporary record, delete to clear any data.
                        SharePointMgt.GetDirectoryFilesList(SharePointFile, RootFolderPath);
                        if SharePointFile.FindSet() then
                            repeat
                                //Let's create a list of all items within the Documents directory
                                //This list is used to display the files.
                                SCListRec.Init();
                                SCListRec.Id := SharePointFile."Unique Id";
                                SCListRec.Title := SharePointFile.Name;
                                SCListRec.OdataId := SharePointFile.OdataId;
                                SCListRec."Server Relative Url" := SharePointFile."Server Relative Url";
                                if SCListRec.Insert() then; //We initialize and insert a record for each Sharepoint List
                            until SharePointFile.Next() = 0;
                    end;

                    CurrPage.SCList.Page.Update(); //Update the current page to refresh the list
                end;
            }
            action(GetAllList)
            {
                ApplicationArea = All;
                Caption = 'Get Sub Folders Files List';

                trigger OnAction()
                var
                    AadTenantId: Text;
                    SCListRec: Record "Sharepoint Connector List";
                    SharePointFile: Record "SharePoint File" temporary;
                begin
                    RootFolderPath := SharePointMgt.GetRootFolder();
                    if not (RootFolderPath = '') then begin
                        SCListRec.DeleteAll(); //List is not a temporary record, delete to clear any data.
                        SharePointMgt.GetDirectoryFilesListInclSubDirs(SharePointFile, RootFolderPath);
                        if SharePointFile.FindSet() then
                            repeat
                                //Let's create a list of all items within the Documents directory
                                //This list is used to display the files.
                                SCListRec.Init();
                                SCListRec.Id := SharePointFile."Unique Id";
                                SCListRec.Title := SharePointFile.Name;
                                SCListRec.OdataId := SharePointFile.OdataId;
                                SCListRec."Server Relative Url" := SharePointFile."Server Relative Url";
                                if SCListRec.Insert() then; //We initialize and insert a record for each Sharepoint List
                            until SharePointFile.Next() = 0;
                    end;

                    CurrPage.SCList.Page.Update(); //Update the current page to refresh the list
                end;
            }
            action(ShowDirectory)
            {
                ApplicationArea = All;
                Caption = 'Show Directory';
                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                    SharePointFolder: Record "SharePoint Folder" temporary;
                    SharePointFile: Record "SharePoint File" temporary;
                begin
                    SCListRec.DeleteAll();

                    SharePointMgt.GetDirectoryFoldersList(SharePointFolder, DirectoryPath);
                    SharePointMgt.GetDirectoryFilesList(SharePointFile, DirectoryPath);

                    if SharePointFolder.FindSet() then begin
                        repeat
                            //Let's create a list of all items within the Documents directory
                            //This list is used to display the folders.
                            SCListRec.Init();
                            SCListRec.Id := SharePointFolder."Unique Id";
                            SCListRec.Title := SharePointFolder.Name;
                            SCListRec.OdataId := SharePointFolder.OdataId;
                            SCListRec."Server Relative Url" := SharePointFolder."Server Relative Url";
                            if SCListRec.Insert() then; //We initialize and insert a record for each folder
                        until SharePointFolder.Next() = 0;
                    end;

                    if SharePointFile.FindSet() then begin
                        repeat
                            //Let's create a list of all items within the Documents directory
                            //This list is used to display the files.
                            SCListRec.Init();
                            SCListRec.Id := SharePointFile."Unique Id";
                            SCListRec.Title := SharePointFile.Name;
                            SCListRec.OdataId := SharePointFile.OdataId;
                            SCListRec."Server Relative Url" := SharePointFile."Server Relative Url";
                            if SCListRec.Insert() then; //We initialize and insert a record for each file
                        until SharePointFile.Next() = 0;
                    end;

                    CurrPage.SCList.Page.Update(); //Update the current page to refresh the list
                end;
            }
            action(UploadFile)
            {
                ApplicationArea = All;
                Caption = 'Upload File';

                trigger OnAction()
                var
                    FilePath: Text;
                begin
                    SharePointMgt.UploadFile('Dialog ', DirectoryPath);
                    if not (FilePath = '') then begin
                        Message('File - %1 sucessfully uploaded to %2!', FilePath, DirectoryPath);
                    end;
                end;
            }
            action(UploadFileWithFilters)
            {
                ApplicationArea = All;
                Caption = 'Upload File With Filters';

                trigger OnAction()
                var
                    FilePath: Text;
                begin
                    SharePointMgt.UploadFileWithFilters('My Dialog ', DirectoryPath, 'Text file (*.txt)|*.txt');
                    if not (FilePath = '') then begin
                        Message('Text file - %1 sucessfully uploaded to %2!', FilePath, DirectoryPath);
                    end;
                end;
            }
            action(GetFileContent)
            {
                ApplicationArea = All;
                Caption = 'Get Selected File Content';

                trigger OnAction()
                var
                    InStream: InStream;
                    TempBlob: Codeunit "Temp Blob";
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    SharePointMgt.InStreamImportFromFile(InStream, SCListRec."Server Relative Url");

                    InStream.ReadText(FileContents);

                end;
            }
            action(RenameFile)
            {
                ApplicationArea = All;
                Caption = 'Rename Selected File';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    if SharePointMgt.RenameFile(SCListRec."Server Relative Url", FileName) then
                        Message('File - %1 sucessfully renamed to %2!', SCListRec.Title, FileName);
                end;
            }
            action(ReplaceFile)
            {
                ApplicationArea = All;
                Caption = 'Replace Selected File Content';

                trigger OnAction()
                var
                    InStream: InStream;
                    OutStream: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    OutStream := TempBlob.CreateOutStream();
                    OutStream.Write(FileContents);
                    InStream := TempBlob.CreateInStream();

                    if SharePointMgt.ReplaceFileContent(SCListRec."Server Relative Url", InStream) then
                        Message('File - %1 content sucessfully replaced!', SCListRec.Title);
                end;
            }
            action(CreateFile)
            {
                ApplicationArea = All;
                Caption = 'Create File';

                trigger OnAction()
                var
                    SharePointFile: Record "SharePoint File" temporary;
                    IS: InStream;
                    OS: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    FilePath: Text;
                begin
                    OS := TempBlob.CreateOutStream();
                    OS.Write(FileContents);
                    IS := TempBlob.CreateInStream();

                    FilePath := SharePointMgt.CreateFile(DirectoryPath, FileName, IS);
                    if not (FilePath = '') then
                        Message('File - %1 sucessfully created!', FilePath);
                end;
            }
            action(CreateFolder)
            {
                ApplicationArea = All;
                Caption = 'Create Folder';

                trigger OnAction()
                var
                    IS: InStream;
                    OS: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                begin
                    OS := TempBlob.CreateOutStream();
                    OS.Write(FileContents);
                    IS := TempBlob.CreateInStream();

                    if SharePointMgt.CreateDirectory(DirectoryPath) then
                        Message('Folder - %1 sucessfully created!', DirectoryPath);
                end;
            }

            action(DownloadFile)
            {
                ApplicationArea = All;
                Caption = 'Download Selected File';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    SharePointMgt.DownloadFile(SCListRec."Server Relative Url", 'Download', SCListRec.Title);
                end;
            }
            action(DeleteFile)
            {
                ApplicationArea = All;
                Caption = 'Delete Selected File';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    if SharePointMgt.DeleteFile(SCListRec."Server Relative Url") then
                        Message('File - %1 deleted sucessfully!', SCListRec.Title);
                end;
            }
            action(DeleteFolder)
            {
                ApplicationArea = All;
                Caption = 'Delete Selected Folder';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    if SharePointMgt.RemoveDirectory(SCListRec."Server Relative Url") then
                        Message('Folder - %1 deleted sucessfully!', SCListRec.Title);
                end;
            }
            action(CopyFile)
            {
                ApplicationArea = All;
                Caption = 'Copy Selected File';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                    NewSharePointFile: Record "SharePoint File" temporary;
                    FilePath: Text;
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);

                    if DirectoryPath = '' then
                        DirectoryPath := SharePointMgt.FixPathForSharePoint(FileMgt.GetDirectoryName(SCListRec."Server Relative Url"));
                    if FileName = '' then
                        FileName := SCListRec.Title;
                    FilePath := DirectoryPath + '/' + FileName;
                    if FilePath = SCListRec."Server Relative Url" then
                        FileName := SCListRec.Title.Replace('.', ' (Copy).');

                    FilePath := DirectoryPath + '/' + FileName;
                    FilePath := SharePointMgt.FixPathForSharePoint(FilePath);
                    if SharePointMgt.CopyFile(SCListRec."Server Relative Url", FilePath, false) then
                        Message('File - %1 sucessfully copied to - %2!', SCListRec.Title, FilePath);
                end;
            }
        }
    }

    trigger OnOpenPage()
    begin
        if not Rec.Get() then begin
            Rec.Init();
            Rec.Insert();
        end;
    end;

    var
        FileName: Text;
        FileContents: Text;
        DirectoryPath: Text;
        SharePointMgt: Codeunit "SharePoint File Management";
        RootFolderPath: Text;
        FileMgt: Codeunit "File Management";

}