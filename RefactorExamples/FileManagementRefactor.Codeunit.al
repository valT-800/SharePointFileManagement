codeunit 51118 "File Mgt. Refactor Exmamples"
{

    var
        SharePointMgt: Codeunit "SharePoint File Management";
        FileMgt: Codeunit "File Management";

    procedure UploadFile(ClientFilePath: Text)
    begin
        // SPLN1.00 - Start
        //FileMgt.UploadFile('My dialog',ClientFilePath); // File uploaded from passed ClientFilePath
        SharePointMgt.UploadFile('My dialog', ClientFilePath); //File uploaded by client while procedure runs
        //SPLN1.00 - End
    end;

    procedure DownloadFile(FilePath: Text)
    var
        ToFileName: Text;
    begin
        ToFileName := FileMgt.GetFileName(FilePath);
        //SPLN1.00 - Start
        // FileMgt.DownloadToFile(ToFileName, FilePath);
        FilePath := SharePointMgt.FixPathForSharePoint(FilePath);
        SharePointMgt.DownloadFile(FilePath, '', ToFileName);
        //SPLN1.00 - End
    end;

    procedure CopyAndDeleteFile(SourceFilePath: Text; TargetFilePath: Text)
    begin
        // SPLN1.00 - Start
        // FileMgt.CopyServerFile(SourceFilePath, TargetFilePath, true);
        // FileMgt.DeleteServerFile(SourceFilePath);
        if SharePointMgt.CopyFile(SourceFilePath, TargetFilePath, true) then
            SharePointMgt.DeleteFile(SourceFilePath);
        // SPLN1.00 - End
    end;

    procedure GetDirectoryFilesList(ServerFolderPath: Text)
    var
        // TempNameValueBuffer: Record "Name/Value Buffer" temporary;
        // File: File;
        InStream: InStream;
        TempSharePointFile: Record "SharePoint File" temporary;
    begin
        //SPLN1.00 - Start
        // FileMgt.GetServerDirectoryFilesList(TempNameValueBuffer, ServerFolderPath);
        // if TempNameValueBuffer.FindSet then
        //     repeat
        //         File.Open(TempNameValueBuffer.Name);
        //         File.CreateInStream(InStream);
        //         XMLPORT.Import(XMLPORT::"CAL Test Coverage Map", InStream);
        //         File.Close;
        //     until TempNameValueBuffer.Next = 0;
        SharePointMgt.GetDirectoryFilesList(TempSharePointFile, ServerFolderPath);
        if TempSharePointFile.FindSet then
            repeat
                if SharePointMgt.InStreamImportFromFile(InStream, TempSharePointFile."Server Relative Url") then
                    XMLPORT.Import(XMLPORT::"CAL Test Coverage Map", InStream);
            until TempSharePointFile.Next = 0;
        //SPLN1.00 - End
    end;

}