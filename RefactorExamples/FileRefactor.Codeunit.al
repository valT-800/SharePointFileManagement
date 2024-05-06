codeunit 51117 "File Refactor Examples"
{
    SingleInstance = true;

    var
        SharePointMgt: Codeunit "SharePoint File Management";

    procedure CreateAndWrite(FilePath: Text)
    var
        // File: File;
        TempBlob: Codeunit "Temp Blob";
        OutStream: OutStream;
        InStream: InStream;
        DirectoryPath: Text;
    begin
        //SPLN1.00 - Start
        // File.WriteMode(true);
        // File.TextMode(true);
        // File.Create(FilePath);
        // File.Write('Line 1');
        // File.Write('');
        // File.Write('Line 3');
        // File.Close();
        TempBlob.CreateOutStream(OutStream);
        OutStream.WriteText('Line 1');
        OutStream.WriteText('');
        OutStream.WriteText('Line 3');
        TempBlob.CreateInStream(InStream);
        SharePointMgt.CreateFile(FilePath, InStream);
        //SPLN1.00 - End
    end;

    procedure OpenAndWrite(FilePath: Text)
    var
        // File: File;
        TempBlob: Codeunit "Temp Blob";
        OutStream: OutStream;
        DirectoryPath: Text;
    begin
        //SPLN1.00 - Start
        // File.WriteMode(true);
        // File.TextMode(true);
        // File.Open(FilePath);
        // File.Write('Line 1');
        // File.Write('');
        // File.Write('Line 3');
        // File.Close();
        SharePointMgt.BLOBImportFromFile(TempBlob, FilePath);
        TempBlob.CreateOutStream(OutStream);
        OutStream.WriteText('Line 1');
        OutStream.WriteText('');
        OutStream.WriteText('Line 3');
        SharePointMgt.ReplaceFileContent(FilePath, TempBlob);
        //SPLN1.00 - End
    end;

    procedure OpenAndRead(FilePath: Text) Text: Text
    var
        // File: File;
        TempBlob: Codeunit "Temp Blob";
        InStream: InStream;
        DirectoryPath: Text;
    begin
        //SPLN1.00 - Start
        // File.TextMode(true);
        // File.Open(FilePath);
        // File.Read(Text);
        // File.Close();
        SharePointMgt.InStreamImportFromFile(InStream, FilePath);
        InStream.ReadText(Text);
        //SPLN1.00 - End
    end;

    procedure CopyFileAndDelete(SourceFilePath: Text; TargetFilePath: Text)
    begin
        //SPLN1.00 - Start
        // if Copy(SourceFilePath, TargetFilePath) then
        // Erase(SourceFilePath);
        if SharePointMgt.CopyFile(SourceFilePath, TargetFilePath) then
            SharePointMgt.DeleteFile(SourceFilePath);
        //SPLN1.00 - End
    end;
}