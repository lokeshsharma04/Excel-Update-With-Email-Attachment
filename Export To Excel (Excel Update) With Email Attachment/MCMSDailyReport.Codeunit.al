
Codeunit 90405 "MCMS Daily Report 2"
{
    trigger OnRun()
    var
        Report90054: report 90054;
        Report50109: report 50109;
        Exceltemp: Record "Excel Template";
    begin
        Exceltemp.DeleteAll();
        SalesSetup.Get;
        Filename := SalesSetup."MCMS Report Path" + 'AgedRecv.xls';
        // Compile pending
        Clear(MCAgingReport);
        //  MCAgingReport.sav
        // saveasexcel();//NEW
        // MCAgingReport.CreateExcelbook;//NEW
        // Commit();
        Clear(Receipients);
        // Receipients.Add(Email001);
        Receipients.Add('luckysharmasharma336@gmail.com');
        Clear(BccMailList);
        Clear(bcc);
        // bcc := 'MCMS-AllStaff@nissinfoods.com.hk';
        // bcc := 'MCMS-AllStaff@nissinfoods.com.hk';
        BccMailList.Add(bcc);
        EmailMessage.Create(Receipients, '',
        Format(Today) + ' MCMS Daily Reports', false, CcMailList, bCcMailList);

        // EmailMessage.AddRecipient('masamune.komori@mcmshk.com' );
        EmailMessage.AppendToBody(Format(Today) + ' MCMS Daily Reports');
        // EmailMessage.AddAttachment('AgedRecvDetail.xlsx', 'xlsx', Instrm);//NEW
        Clear(MCAgingReport);
        Clear(Instrm);
        Clear(Outstrm);
        //IF FILE.COPY(SalesSetup."MCMS Report Path" + 'Template\DailySales.xlsx', SalesSetup."MCMS Report Path" + 'DailySales.xlsx') THEN;

        //ClientFileHelper.Copy(SalesSetup."MCMS Report Path" + 'Template\DailySales.xlsx', SalesSetup."MCMS Report Path" + 'DailySales.xlsx');
        //FileManagement.MoveAndRenameClientFile('\\10.1.8.5\MCMS-Report\Template\DailySales.xlsx', 'DailySales.xlsx','\\10.1.8.5\MCMS-Report\') ;

        DateRange := (Format(CalcDate('<CM-1M+1D>', Today)) + '..' + Format(CalcDate('<CM>', Today)));

        Commit();
        report.Run(90054, false, false);
        AddAttachment(EmailMessage, 'DailySales.xlsx', 90054);
        Commit();
        report.Run(50109, false, false);
        AddAttachment(EmailMessage, 'DailySales(Mobile).xlsx', 50109);

        EmailSend.Send(EmailMessage, Enum::"Email Scenario"::Default);
        // Compile pending
    end;

    var
        MCAgingReport: Report 50202;
        SalesSetup: Record "Sales & Receivables Setup";
        //MCDailySales: Report 50111; Compile pending
        Filename: Text;
        DateRange: Text;
        //SMTPMail: Codeunit "SMTP Mail"; Compile pending
        Email001: label 'itd@nissinfoods.com.hk';
        Email002: label 'Daily Report';
        bcc: Text;
        //ClientFileHelper: dotnet File; Compile pending
        EmailSend: Codeunit Email;
        EmailMessage: Codeunit "Email Message";
        EmailAccount: Record "Email Account";
        BodyMessage: Text;
        AddBodyMessage: Text;
        Receipients: List of [Text];
        CcMailList: List of [Text];
        BccMailList: List of [Text];
        Instrm: InStream;
        Outstrm: OutStream;
        TempBlob_lRec: Codeunit "Temp Blob";

    procedure AddAttachment(var EmailMessage: Codeunit "Email Message"; FileName: Text; ReportID: Integer)
    var
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
        ExcelTemp: Record "Excel Template";
    begin
        ExcelTemp.Reset();
        ExcelTemp.SetRange(ReportID, ReportID);
        ExcelTemp.FindLast();
        if not ExcelTemp.GetTemplateFileAsTempBlob(TempBlob) then
            exit;

        TempBlob.CreateInStream(InStr);
        EmailMessage.AddAttachment(FileName, 'xlsx', InStr);//Transfer to Codeunit
    end;

    procedure saveasexcel()
    var
        Out: OutStream;
        RecRef: RecordRef;
        FileManagement_lCdu: Codeunit "File Management";
        Cust_Rec: Record Customer;
        ReportParam: text;
    begin
        Clear(Instrm);
        Clear(MCAgingReport);
        TempBlob_lRec.CreateOutStream(Out, TEXTENCODING::UTF8);
        Cust_Rec.Get('C000111');
        RecRef.GetTable(Cust_Rec);
        RecRef.Get(Cust_Rec.RecordId);
        // Report.SaveAs(50202, '', ReportFormat::Excel, Out, RecRef);//TEC.VJ
        MCAgingReport.SaveAs(ReportParam, ReportFormat::Excel, Out, RecRef);
        FileManagement_lCdu.BLOBExport(TempBlob_lRec, Filename, true);
        TempBlob_lRec.CreateInStream(Instrm);
    end;
}

