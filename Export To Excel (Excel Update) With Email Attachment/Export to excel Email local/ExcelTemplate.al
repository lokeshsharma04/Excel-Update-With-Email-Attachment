table 50103 "Excel Template"
{
    fields
    {
        field(1; Name; Text[250]) { DataClassification = CustomerContent; }
        field(2; Filename; Text[80]) { DataClassification = CustomerContent; }
        field(3; ReportID; Integer)
        { }
        field(4; "Blob Key"; BigInteger) { DataClassification = CustomerContent; }
    }

    var
        PersistentBlob: Codeunit "Persistent Blob";
        DialogCaptionTxt: Label 'Select a file';
        FileFilterTxt: Label 'Excel Files (*.xlsx)|*.xlsx';
        ExtFilterTxt: Label 'xlsx';
        CouldNotStoreExcelFileErr: Label 'Could not store Excel file';

    trigger OnDelete()
    begin
        DeletePersistentBlob();
    end;

    procedure HasContent(): Boolean
    begin
        if "Blob Key" <> 0 then
            exit(PersistentBlob.Exists("Blob Key"));
    end;

    procedure ImportTemplateFile()
    var
        TempBlob: Codeunit "Temp Blob";
        FileMgt: Codeunit "File Management";
        OutStr: OutStream;
    begin
        Filename := FileMgt.BLOBImportWithFilter(TempBlob, DialogCaptionTxt, '', FileFilterTxt, ExtFilterTxt);
        if Filename <> '' then
            StoreBlob(TempBlob, '', 0);
        if Name = '' then
            Name := Filename;
        if not Modify() then
            Insert();
    end;

    procedure GetTemplateFileAsTempBlob(var TempBlob: Codeunit "Temp Blob"): Boolean
    var
        OutStr: OutStream;
    begin
        if "Blob Key" = 0 then
            exit;
        TempBlob.CreateOutStream(OutStr);
        PersistentBlob.CopyToOutStream("Blob Key", OutStr);
        exit(true);
    end;

    procedure StoreBlob(var TempBlob: Codeunit "Temp Blob"; FileName: text; ReportID: integer)
    var
        InStr: InStream;
    begin
        DeletePersistentBlob();
        Rec."Blob Key" := PersistentBlob.Create();
        TempBlob.CreateInStream(InStr);
        if not PersistentBlob.CopyFromInStream("Blob Key", InStr) then
            Error(CouldNotStoreExcelFileErr);
        Rec.Name := FileName;
        Rec.ReportID := ReportID;
        if not Rec.Modify() then
            Rec.Insert();
    end;

    local procedure DeletePersistentBlob()
    begin
        if "Blob Key" <> 0 then
            PersistentBlob.Delete("Blob Key");
    end;
}