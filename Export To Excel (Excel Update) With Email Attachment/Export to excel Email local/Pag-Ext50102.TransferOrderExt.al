namespace ALBC.ALBC;

using Microsoft.Inventory.Transfer;
using System.Email;
using System.IO;
using Microsoft.Sales.Reports;
using System.Utilities;

pageextension 50102 "Transfer Order Ext" extends "Transfer Order"
{
    actions
    {
        addafter(PreviewPosting)
        {
            action(ImportServerFile)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    CCMailList: List of [text];
                    bCCMailList: List of [text];
                    Recipient: text;
                    Tempblob: Codeunit "Temp Blob";
                    Ins: InStream;
                    EmailMess: Codeunit "Email Message";
                    Email: Codeunit Email;
                    FileMgmt: codeunit "File Management";
                    FilePath: Text;
                begin
                    CCMailList.Add('lucksa463@gmail.com');
                    bCCMailList.Add('luckysharmasharma336@gmail.com');
                    // Recipient.Add('lucksa463@gmail.com');
                    Recipient := 'lucksa463@gmail.com';
                    EmailMess.Create(Recipient, 'Test Email Import File', 'Body');
                    FilePath := 'D:\CU 51 NAV 2013 R2 W1\MicrosoftDynamicsNAV2013R2_ImportExportData.pdf';
                    FileMgmt.BLOBImportFromServerFile(Tempblob, FilePath);
                    Tempblob.CreateInStream(Ins);
                    EmailMess.AddAttachment('Document.pdf', 'PDF', Ins);
                    Email.Send(EmailMess, ENUM::"Email Scenario"::Default);
                end;
            }
            action(EmailNew)
            {
                trigger OnAction()
                var
                    TempBlob: codeunit "Temp Blob";
                    Email: Codeunit Email;
                    EmailMessage: Codeunit "Email Message";
                    ReportExample: Report "Customer - List";
                    InStr: Instream;
                    OutStr: OutStream;
                    ReportParameters: Text;
                begin

                    begin
                        // EmailMessage.Create('', 'This is the subject', 'This is the body');
                        // Email.Send(EmailMessage, Enum::"Email Scenario"::Default);

                        begin
                            // ReportParameters := ReportExample.RunRequestPage();
                            TempBlob.CreateOutStream(OutStr);
                            ReportExample.SaveAs('', ReportFormat::Excel, OutStr);
                            TempBlob.CreateInStream(InStr);

                            emailmessage.create('lucksa463@gmail.com', 'This is the subject', 'This is the body');
                            EmailMessage.AddAttachment('FileName.pdf', 'PDF', InStr);
                            Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
                        end;
                    end;
                end;

            }
            action(ExportToExcel)
            {
                Caption = 'Export to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;
                trigger OnAction()
                var
                    Email: Codeunit Email;
                    EmailMessage: Codeunit "Email Message";
                    ExcelBuffer: Record "Excel Buffer" temporary;
                begin
                    ExcelTemp.DeleteAll();
                    Repport.Run();
                    CreateAndSendEmail(1, '');//NEW
                end;
            }
        }
    }
    //NEW
    procedure CreateAndSendEmail(ReportID: Integer; BookName: Text)
    var
        Email: Codeunit Email;
        EmailMessage: Codeunit "Email Message";
    begin
        EmailMessage.create('lucksa463@gmail.com', 'This is the subject', 'This is the body');
        AddAttachment(EmailMessage, 'CustLedgEntry', 50102);
        Email.Send(EmailMessage, Enum::"Email Scenario"::Default);
    end;



    //NEW
    procedure AddAttachment(var EmailMessage: Codeunit "Email Message"; FileName: Text; ReportID: Integer)
    var
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        ExcelTemp.Reset();
        ExcelTemp.SetRange(ReportID, ReportID);
        ExcelTemp.FindLast();
        if not ExcelTemp.GetTemplateFileAsTempBlob(TempBlob) then
            exit;

        TempBlob.CreateInStream(InStr);
        EmailMessage.AddAttachment(FileName, 'xlsx', InStr);//Transfer to Codeunit
    end;

    var
        Repport: report "Cust Ledg Entry";
        ExcelTemp: Record "Excel Template";

}