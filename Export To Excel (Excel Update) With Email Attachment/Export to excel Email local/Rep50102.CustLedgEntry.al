namespace ALBC.ALBC;

using Microsoft.Sales.Receivables;
using System.IO;
using Microsoft.Sales.Customer;
using System.Email;
using System.Utilities;
report 50102 "Cust Ledg Entry"
{
    ApplicationArea = All;
    Caption = 'Cust Ledg Entry';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;
    dataset
    {
        dataitem(CustLedgerEntry; Customer)
        {
        }
    }
    requestpage
    {
        layout
        {
            area(Content)
            {
                group(GroupName)
                {
                }
            }
        }
        actions
        {
            area(Processing)
            {
            }
        }
    }
    trigger OnPostReport()
    var
        CustLedgEnt: Record "Cust. Ledger Entry";
    begin
        // CustLedgEnt.SetRange("Entry No.", 1, 1000);
        CustLedgEnt.SetRange("Customer No.", '10000');
        ExportCustLedgerEntries(CustLedgEnt);

    end;


    procedure ExportCustLedgerEntries(var CustLedgerEntryRec: Record "Cust. Ledger Entry")
    var
        CustLedgerEntriesLbl: Label 'Customer Ledger Entries';
        ExcelFileName: Label 'CustomerLedgerEntries_%1_%2';
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Entry No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Posting Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Document Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Customer No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Customer Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Currency Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Original Amount"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Amount), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Amount (LCY)"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Remaining Amount"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Remaining Amt. (LCY)"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Open), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        if CustLedgerEntryRec.FindSet() then
            repeat
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Entry No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Posting Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Document Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Document No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Customer No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Customer Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Currency Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Original Amount", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Amount, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Amount (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Remaining Amount", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Remaining Amt. (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Open, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
            until CustLedgerEntryRec.Next() = 0;
        // CreateExcelSheet('CustLedg', true);
        TempExcelBuffer.CreateNewBook(CustLedgerEntriesLbl);
        TempExcelBuffer.WriteSheet(CustLedgerEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();

        // CreateAndSendEmail(TempExcelBuffer, 'CustomerLedg.xlsx');
        StoreExcelToBlob(ExcelFileName, 50102);//New
    end;
    //New
    procedure StoreExcelToBlob(ExcelFileName: text; ReportID: Integer)
    var
        TempBlob: Codeunit "Temp Blob";
        OutStr: OutStream;
        ExcelTemp: Record "Excel Template";
    begin
        TempBlob.CreateOutStream(OutStr);
        TempExcelBuffer.SaveToStream(OutStr, true);
        ExcelTemp.StoreBlob(TempBlob, ExcelFileName, ReportID);
    end;

    // procedure CreateAndSendEmail(var TempExcelBuf: Record "Excel Buffer" temporary;
    //        BookName: Text)
    // var
    //     //SMTPMail: Codeunit "SMTP Mail";
    //     Recipients: List of [Text];
    //     Email: Codeunit Email;
    //     EmailMessage: Codeunit "Email Message";
    // begin
    //     Recipients.Add('lucksa463@gmail.com');
    //     // SMTPMail.CreateMessage(
    //     //     'Business Central Mail',
    //     //     'bc@cronus.company',
    //     //     Recipients,
    //     //     'Test Export to Excel and email',
    //     //     'This is an text to export data to an Excel file and email it.',
    //     //     false);


    //     // AddAttachment(SMTPMail, TempExcelBuf, GetFriendlyFilename(BookName));
    //     // SMTPMail.Send();
    //     // Message(EmailSentTxt);

    //     //TempExcelBuf.
    //     EmailMessage.create('lucksa463@gmail.com', 'This is the subject', 'This is the body');
    //     AddAttachment(EmailMessage, TempExcelBuf, GetFriendlyFilename(BookName));
    //     //EmailMessage.AddAttachment('FileName.pdf', 'PDF', InStr);
    //     Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
    //     Message('EmailSentTxt');
    // end;

    // procedure AddAttachment(
    //     //var SMTPMail: Codeunit "SMTP Mail";
    //     var EmailMessage: Codeunit "Email Message";
    //     var TempExcelBuf: Record "Excel Buffer" temporary;
    //     BookName: Text)
    // var
    //     TempBlob: Codeunit "Temp Blob";
    //     InStr: InStream;
    // begin
    //     ExportExcelFileToBlob(TempExcelBuf, TempBlob);
    //     TempBlob.CreateInStream(InStr);
    //     //SMTPMail.AddAttachmentStream(InStr, BookName);
    //     // EmailMessage.AddAttachment('FileName.xlsx', 'xlsx', InStr);//Transfer to Codeunit
    // end;

    // local procedure ExportExcelFileToBlob(
    //     var TempExcelBuf: Record "Excel Buffer" temporary;
    //     var TempBlob: Codeunit "Temp Blob")
    // var
    //     OutStr: OutStream;
    // begin
    //     TempBlob.CreateOutStream(OutStr);
    //     TempExcelBuf.SaveToStream(OutStr, true);
    // end;

    // local procedure GetFriendlyFilename(BookName: Text): Text
    // var
    //     FileManagement: Codeunit "File Management";
    // begin
    //     if BookName = '' then
    //         exit(Book1Txt + ExcelFileExtensionTok);

    //     exit(FileManagement.StripNotsupportChrInFileName(BookName) + ExcelFileExtensionTok);
    // end;

    var
        ExcelFileExtensionTok: Label '.xlsx', Locked = true;
        EmailSentTxt: Label 'File has been sent by email';
        Book1Txt: Label 'Book1';

        TempExcelBuffer: Record "Excel Buffer" temporary;
}
