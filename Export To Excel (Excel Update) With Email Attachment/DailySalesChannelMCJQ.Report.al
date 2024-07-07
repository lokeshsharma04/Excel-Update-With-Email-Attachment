report 50109 "Daily Sales - Channel MC JQ"
{
    ProcessingOnly = true;
    ApplicationArea = All;

    dataset
    {
        dataitem("Sales Line"; "Sales Line")
        {
            //"Document Type", "Shipment Date", "Order Type"
            DataItemTableView = SORTING()
                                WHERE(Type = CONST(Item),
                                      Amount = FILTER(<> 0));

            trigger OnAfterGetRecord()
            var
                L_BrandCode: Code[20];
            begin
                IF (MonthEnd = FALSE) THEN BEGIN
                    RecCust.GET("Sales Line"."Sell-to Customer No.");

                    L_BrandCode := '';

                    IF NOT TempTableForReports.GET(RecCust."User Defined Field 5", RecCust."User Defined Field 1", '') THEN BEGIN
                        TempTableForReports.INIT;
                        TempTableForReports."Value 1" := RecCust."User Defined Field 5";
                        TempTableForReports."Value 2" := RecCust."User Defined Field 1";
                        IF TempTableForReports.INSERT THEN;
                    END;
                END;


                L_BrandCode := GetBrandDimValue("Sales Line"."Dimension Set ID");

                IF L_BrandCode <> '' THEN BEGIN
                    TempDimVal.INIT;
                    TempDimVal."Dimension Code" := 'BRAND';
                    TempDimVal.Code := L_BrandCode;
                    IF TempDimVal.INSERT THEN;
                END;
            end;

            trigger OnPostDataItem()
            begin

                //MESSAGE('%1',TempTableForReports.COUNT);
            end;

            trigger OnPreDataItem()
            begin
                SETFILTER("Shipment Date", PostingDateFilter);
            end;
        }
        dataitem(dtaitem37; "Sales Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Type = CONST(Item),
                                      Amount = FILTER(<> 0));

            trigger OnAfterGetRecord()
            var
                L_BrandCode: Code[20];
            begin
                //RecCust.GET("Sales Invoice Line"."Sell-to Customer No.");
                L_BrandCode := '';

                IF NOT TempTableForReports.GET(RecCust."User Defined Field 5", RecCust."User Defined Field 1", '') THEN BEGIN
                    TempTableForReports.INIT;
                    TempTableForReports."Value 1" := RecCust."User Defined Field 5";
                    TempTableForReports."Value 2" := RecCust."User Defined Field 1";
                    IF TempTableForReports.INSERT THEN;
                END;


                //L_BrandCode := GetBrandDimValue("Sales Invoice Line"."Dimension Set ID");
                //MESSAGE('%1',L_BrandCode);
                IF L_BrandCode <> '' THEN BEGIN
                    TempDimVal.INIT;
                    TempDimVal."Dimension Code" := 'BRAND';
                    TempDimVal.Code := L_BrandCode;
                    IF TempDimVal.INSERT THEN;
                END;
            end;

            trigger OnPostDataItem()
            begin
                //  MESSAGE('%1',TempTableForReports.COUNT);
                // ERROR('%1',TempDimVal.COUNT);

                //MESSAGE('%1',TempTableForReports.COUNT);
            end;

            trigger OnPreDataItem()
            begin
                SETFILTER("Posting Date", PostingDateFilter);
            end;
        }
        dataitem(DataItem2; "Sales Cr.Memo Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Type = CONST(Item),
                                      Amount = FILTER(<> 0));

            trigger OnAfterGetRecord()
            var
                L_BrandCode: Code[20];
            begin
                RecCust.GET("Sell-to Customer No.");

                L_BrandCode := '';

                IF NOT TempTableForReports.GET(RecCust."User Defined Field 5", RecCust."User Defined Field 1", '') THEN BEGIN
                    TempTableForReports.INIT;
                    TempTableForReports."Value 1" := RecCust."User Defined Field 5";
                    TempTableForReports."Value 2" := RecCust."User Defined Field 1";
                    IF TempTableForReports.INSERT THEN;
                END;

                //L_BrandCode := GetBrandDimValue("Sales Cr.Memo Line"."Dimension Set ID");
                IF L_BrandCode <> '' THEN BEGIN
                    TempDimVal.INIT;
                    TempDimVal."Dimension Code" := 'BRAND';
                    TempDimVal.Code := L_BrandCode;
                    IF TempDimVal.INSERT THEN;
                END;
            end;

            trigger OnPostDataItem()
            begin
                //MESSAGE('%1',TempTableForReports.COUNT);
                TempTableForReports.RESET;
                IF TempTableForReports.FINDFIRST THEN
                    REPEAT
                        TempTableForReports2 := TempTableForReports;
                        IF TempTableForReports2.INSERT THEN;
                    UNTIL TempTableForReports.NEXT = 0;
            end;

            trigger OnPreDataItem()
            begin
                //SETFILTER("Posting Date", PostingDateFilter);
            end;
        }
        dataitem(DataItem3; Integer)
        {
            DataItemTableView = SORTING(Number);
            MaxIteration = 1;

            trigger OnAfterGetRecord()
            var
                L_CalculatedVal: Decimal;
                Cnt: Integer;
            begin
                TempTableForReports.RESET;
                IF TempTableForReports.FINDFIRST THEN
                    REPEAT
                        L_CalculatedVal := 0;
                        TempTableForReports2.GET(TempTableForReports."Value 1", TempTableForReports."Value 2", '');

                        IF LastVal1 <> TempTableForReports."Value 1" THEN BEGIN
                            ExcelBuffer.NewRow;
                            ExcelBuffer.AddColumn(TempTableForReports."Value 1", FALSE, '', TRUE, FALSE, FALSE, '', 1);
                        END;

                        //calculated case value >>
                        ColTotal := 0;
                        Cnt := 0;
                        ExcelBuffer.NewRow;
                        ExcelBuffer.AddColumn(TempTableForReports."Value 2", FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('CASE', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        IF TempDimVal.FINDFIRST THEN
                            REPEAT
                                Cnt := Cnt + 1;
                                L_CalculatedVal := GetCaseValue(TempDimVal.Code, TempTableForReports."Value 1", TempTableForReports."Value 2", 1);
                                ColTotal += L_CalculatedVal;
                                GrpCase[Cnt] += L_CalculatedVal;
                                GrandCase[Cnt] += L_CalculatedVal;
                                IF L_CalculatedVal <> 0 THEN
                                    ExcelBuffer.AddColumn(L_CalculatedVal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0)
                                ELSE
                                    ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1)

UNTIL TempDimVal.NEXT = 0;

                        ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn(ColTotal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);


                        //calculated case value <<



                        // calculated KG/LTR value
                        ColTotal := 0;
                        Cnt := 0;
                        ExcelBuffer.NewRow;
                        ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('KG/LITRE', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        IF TempDimVal.FINDFIRST THEN
                            REPEAT
                                Cnt := Cnt + 1;
                                L_CalculatedVal := 0;
                                L_CalculatedVal := GetCaseValue(TempDimVal.Code, TempTableForReports."Value 1", TempTableForReports."Value 2", 2);
                                ColTotal += L_CalculatedVal;
                                GrpKG[Cnt] += L_CalculatedVal;
                                GrandKG[Cnt] += L_CalculatedVal;

                                IF L_CalculatedVal <> 0 THEN
                                    ExcelBuffer.AddColumn(L_CalculatedVal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0)
                                ELSE
                                    ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1)
UNTIL TempDimVal.NEXT = 0;

                        ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn(ColTotal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);


                        // calculaetd $ value
                        ColTotal := 0;
                        Cnt := 0;
                        ExcelBuffer.NewRow;
                        ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('$', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        IF TempDimVal.FINDFIRST THEN
                            REPEAT
                                Cnt := Cnt + 1;
                                L_CalculatedVal := 0;
                                L_CalculatedVal := GetCaseValue(TempDimVal.Code, TempTableForReports."Value 1", TempTableForReports."Value 2", 3);
                                ColTotal += L_CalculatedVal;
                                GrpDollar[Cnt] += L_CalculatedVal;
                                GrandDollar[Cnt] += L_CalculatedVal;

                                IF L_CalculatedVal <> 0 THEN
                                    ExcelBuffer.AddColumn(L_CalculatedVal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0)
                                ELSE
                                    ExcelBuffer.AddColumn('-', FALSE, '', FALSE, FALSE, FALSE, '', 1)
UNTIL TempDimVal.NEXT = 0;
                        ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
                        ExcelBuffer.AddColumn(ColTotal, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);



                        ExcelBuffer.NewRow;

                        IF TempTableForReports2.NEXT <> 0 THEN BEGIN
                            IF TempTableForReports2."Value 1" <> TempTableForReports."Value 1" THEN BEGIN
                                ShowGroupTotal(TempTableForReports."Value 1");

                            END;
                        END ELSE
                            ShowGroupTotal(TempTableForReports."Value 1");


                        LastVal1 := TempTableForReports."Value 1";

                    UNTIL TempTableForReports.NEXT = 0;
            end;

            trigger OnPostDataItem()
            var
                L_Cnt: Integer;
                Tot_Tot: Decimal;
            begin
            end;

            trigger OnPreDataItem()
            begin
                ExcelBuffer.NewRow;
                ExcelBuffer.AddColumn('Channel', FALSE, '', TRUE, FALSE, TRUE, '', 1);
                ExcelBuffer.AddColumn('Factor', FALSE, '', TRUE, FALSE, TRUE, '', 1);
                ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                TempDimVal.RESET;
                IF TempDimVal.FINDFIRST THEN
                    REPEAT
                        ExcelBuffer.AddColumn(TempDimVal.Code, FALSE, '', TRUE, FALSE, TRUE, '', 1);
                    UNTIL TempDimVal.NEXT = 0;

                ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuffer.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 1);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field("Delivery Date"; PostingDateFilter)
                    {
                        Caption = 'PostingDateFilter';
                        ApplicationArea = All;

                        trigger OnValidate()
                        begin
                            RecDate.SETFILTER("Period Start", PostingDateFilter);
                            PostingDateFilter := RecDate.GETFILTER("Period Start");
                        end;
                    }
                    field("User Defined 5"; UserDefined5Filter)
                    {
                        Caption = 'User Defined 5';
                        ApplicationArea = All;
                    }
                }
                field("Month End (Y/N)"; MonthEnd)
                {
                    Visible = false;
                    Caption = 'MonthEnd';
                    ApplicationArea = All;
                }
            }
        }

        actions
        {
        }

        trigger OnInit()
        begin
            UserDefined5Filter := '<>WRITEOFF';
        end;
    }

    labels
    {
    }

    trigger OnPostReport()
    begin
        ShowGrandTotal;
        SalesSetup.GET;
        //PostingDateFilter :=  (FORMAT(CALCDATE('<CM-1M+1D>',TODAY))+'..'+FORMAT(CALCDATE('<CM>',TODAY)));
        ServerFileName := SalesSetup."MCMS Report Path" + 'DailySales(Mobile).xlsx';


        ExcelBuffer.CreateNewBook('Daily Sales By channel');
        //ExcelBuffer.UpdateBook(ServerFileName, 'Daily Sales By channel');
        ExcelBuffer.WriteSheet('Daily Sales By channel', COMPANYNAME, USERID);
        ExcelBuffer.CloseBook();
        ExcelBuffer.SetFriendlyFilename(StrSubstNo('Daily Sales By channel', CurrentDateTime, UserId));
        if GuiAllowed then
            ExcelBuffer.OpenExcel();
        //ExcelBuffer.OpenExcel;
        //ExcelBuffer.GiveUserControl;
        StoreExcelToBlob('DailySales.xlsx', 50109);//TEC.VJ
    end;

    trigger OnPreReport()
    begin
        PostingDateFilter := FORMAT(CALCDATE('<-CM>', TODAY)) + '..' + FORMAT(TODAY);
        //PostingDateFilter := FORMAT(TODAY);

        ExcelBuffer.DELETEALL;
        TempTableForReports.DELETEALL;

        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('DAILY SALES BY CHANNEL', FALSE, '', TRUE, FALSE, FALSE, '', 1);

        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn(PostingDateFilter, FALSE, '', TRUE, FALSE, FALSE, '', 1);

        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('Print on ' + FORMAT(TODAY) + '  ' + FORMAT(TIME), FALSE, '', TRUE, FALSE, FALSE, '', 1);

        ExcelBuffer.NewRow;
        ExcelBuffer.NewRow;
    end;

    var
        TempTableForReports: Record "Temp Table for Reports" temporary;
        RecCust: Record Customer;
        PostingDateFilter: Text;
        RecDate: Record Date;
        ExcelBuffer: Record "Excel Buffer" temporary;
        LastVal1: Code[20];
        TempTableForReports2: Record "Temp Table for Reports" temporary;
        DimensionVal: Record "Dimension Value";
        TempDimVal: Record "Dimension Value" temporary;
        ColTotal: Decimal;
        GrpCase: array[100] of Decimal;
        GrpKG: array[100] of Decimal;
        GrpDollar: array[100] of Decimal;
        GrandCase: array[100] of Decimal;
        GrandKG: array[100] of Decimal;
        GrandDollar: array[100] of Decimal;
        BrandDimensionFilter: Text;
        UserDefined5Filter: Text;
        MonthEnd: Boolean;
        ServerFileName: Text;
        SalesSetup: Record "Sales & Receivables Setup";
    //TEC.VJ 18062024>>
    procedure StoreExcelToBlob(ExcelFileName: text; ReportID: Integer)
    var
        TempBlob: Codeunit "Temp Blob";
        OutStr: OutStream;
        ExcelTemp: Record "Excel Template";
    begin
        TempBlob.CreateOutStream(OutStr);
        ExcelBuffer.SaveToStream(OutStr, true);
        ExcelTemp.StoreBlob2(TempBlob, ExcelFileName, ReportID);
    end;
    //TEC.VJ 18062024<<

    procedure GetBrandDimValue(DimSetId: Integer): Code[20]
    var
        L_DimSetEntry: Record "Dimension Set Entry";
    begin
        L_DimSetEntry.RESET;
        L_DimSetEntry.SETRANGE(L_DimSetEntry."Dimension Set ID", DimSetId);
        L_DimSetEntry.SETRANGE(L_DimSetEntry."Dimension Code", 'BRAND');
        IF L_DimSetEntry.FINDFIRST THEN
            EXIT(L_DimSetEntry."Dimension Value Code");
    end;

    //
    procedure GetCaseValue(BrandCode: Code[20]; UserDefined5: Code[20]; UserDefined1: Code[20]; Option: Integer): Decimal
    var
        SIL: Record "Sales Invoice Line";
        L_Cust: Record Customer;
        ToT_CaseQty: Decimal;
        SCL: Record "Sales Cr.Memo Line";
        L_Item: Record Item;
        SIH: Record "Sales Invoice Header";
        SCH: Record "Sales Cr.Memo Header";
        SL: Record "Sales Line";
        SH: Record "Sales Header";
    begin
        ToT_CaseQty := 0;
        // 20191203 JO
        IF (MonthEnd = FALSE) THEN BEGIN
            //JO 20190926
            SL.COPYFILTERS("Sales Line");
            //SL.SETFILTER("Shipment Date",PostingDateFilter);
            IF SL.FINDFIRST THEN
                REPEAT
                    L_Cust.GET(SL."Sell-to Customer No.");
                    IF (L_Cust."User Defined Field 5" = UserDefined5) AND (L_Cust."User Defined Field 1" = UserDefined1) THEN BEGIN
                        IF BrandCode = GetBrandDimValue(SL."Dimension Set ID") THEN BEGIN
                            IF Option = 1 THEN
                                ToT_CaseQty += SL."Case Qty."
                            ELSE
                                IF Option = 2 THEN BEGIN
                                    L_Item.GET(SL."No.");
                                    ToT_CaseQty += SL."Case Qty." * L_Item."Net Weight";
                                END ELSE
                                    IF Option = 3 THEN BEGIN
                                        SH.GET(SL."Document Type", SL."Document No.");
                                        //
                                        IF BrandCode IN ['BADO', 'EVIA', 'VOLV', 'FERRAR'] THEN BEGIN
                                            IF SH."Currency Factor" <> 0 THEN
                                                ToT_CaseQty += (SL.Amount - SL."Discounted Amt. 1" - SL."Discounted Amt. 4" - SL."Discounted Amt. 5") / SH."Currency Factor"
                                            ELSE
                                                ToT_CaseQty += (SL.Amount - SL."Discounted Amt. 1" - SL."Discounted Amt. 4" - SL."Discounted Amt. 5");
                                        END ELSE
                                            IF SH."Currency Factor" <> 0 THEN
                                                ToT_CaseQty += (SL.Amount - SL."Discounted Amt. 1" - SL."Discounted Amt. 2" - SL."Discounted Amt. 3" - SL."Discounted Amt. 4" - SL."Discounted Amt. 5") / SH."Currency Factor"
                                            ELSE
                                                ToT_CaseQty += (SL.Amount - SL."Discounted Amt. 1" - SL."Discounted Amt. 2" - SL."Discounted Amt. 3" - SL."Discounted Amt. 4" - SL."Discounted Amt. 5");
                                        //
                                    END;
                        END;
                    END;
                UNTIL SL.NEXT = 0;
            //JO 20190926
        END;
        // 20191203 JO

        //SIL.COPYFILTERS("Sales Invoice Line");
        //SIL.SETFILTER("Posting Date",PostingDateFilter);
        IF SIL.FINDFIRST THEN
            REPEAT
                L_Cust.GET(SIL."Sell-to Customer No.");
                IF (L_Cust."User Defined Field 5" = UserDefined5) AND (L_Cust."User Defined Field 1" = UserDefined1) THEN BEGIN
                    IF BrandCode = GetBrandDimValue(SIL."Dimension Set ID") THEN BEGIN
                        IF Option = 1 THEN
                            ToT_CaseQty += SIL."Case Qty."
                        ELSE
                            IF Option = 2 THEN BEGIN
                                L_Item.GET(SIL."No.");
                                ToT_CaseQty += SIL."Case Qty." * L_Item."Net Weight";
                            END ELSE
                                IF Option = 3 THEN BEGIN
                                    SIH.GET(SIL."Document No.");
                                    IF BrandCode IN ['BADO', 'EVIA', 'VOLV', 'FERRAR'] THEN BEGIN
                                        IF SIH."Currency Factor" <> 0 THEN
                                            ToT_CaseQty += (SIL.Amount - SIL."Discounted Amt. 1" - SIL."Discounted Amt. 4" - SIL."Discounted Amt. 5") / SIH."Currency Factor"
                                        ELSE
                                            ToT_CaseQty += (SIL.Amount - SIL."Discounted Amt. 1" - SIL."Discounted Amt. 4" - SIL."Discounted Amt. 5");
                                    END ELSE
                                        IF SIH."Currency Factor" <> 0 THEN
                                            ToT_CaseQty += (SIL.Amount - SIL."Discounted Amt. 1" - SIL."Discounted Amt. 2" - SIL."Discounted Amt. 3" - SIL."Discounted Amt. 4" - SIL."Discounted Amt. 5") / SIH."Currency Factor"
                                        ELSE
                                            ToT_CaseQty += (SIL.Amount - SIL."Discounted Amt. 1" - SIL."Discounted Amt. 2" - SIL."Discounted Amt. 3" - SIL."Discounted Amt. 4" - SIL."Discounted Amt. 5");

                                END;

                    END;
                END;
            UNTIL SIL.NEXT = 0;

        //SCL.COPYFILTERS("Sales Cr.Memo Line");
        //SCL.SETFILTER("Posting Date",PostingDateFilter);
        IF SCL.FINDFIRST THEN
            REPEAT
                L_Cust.GET(SCL."Sell-to Customer No.");
                IF (L_Cust."User Defined Field 5" = UserDefined5) AND (L_Cust."User Defined Field 1" = UserDefined1) THEN BEGIN
                    IF BrandCode = GetBrandDimValue(SCL."Dimension Set ID") THEN BEGIN
                        IF Option = 1 THEN
                            ToT_CaseQty -= SCL."Case Qty."
                        ELSE
                            IF Option = 2 THEN BEGIN
                                L_Item.GET(SCL."No.");
                                ToT_CaseQty -= SCL."Case Qty." * L_Item."Net Weight";
                            END ELSE
                                IF Option = 3 THEN BEGIN
                                    SCH.GET(SCL."Document No.");
                                    IF BrandCode IN ['BADO', 'EVIA', 'VOLV', 'FERRAR'] THEN BEGIN
                                        IF SCH."Currency Factor" <> 0 THEN
                                            ToT_CaseQty -= (SCL.Amount - SCL."Discounted Amt. 1" - SCL."Discounted Amt. 4" - SCL."Discounted Amt. 5") / SCH."Currency Factor"
                                        ELSE
                                            ToT_CaseQty -= (SCL.Amount - SCL."Discounted Amt. 1" - SCL."Discounted Amt. 4" - SCL."Discounted Amt. 5");
                                    END ELSE
                                        IF SCH."Currency Factor" <> 0 THEN
                                            ToT_CaseQty -= (SCL.Amount - SCL."Discounted Amt. 1" - SCL."Discounted Amt. 2" - SCL."Discounted Amt. 3" - SCL."Discounted Amt. 4" - SCL."Discounted Amt. 5") / SCH."Currency Factor"
                                        ELSE
                                            ToT_CaseQty -= (SCL.Amount - SCL."Discounted Amt. 1" - SCL."Discounted Amt. 2" - SCL."Discounted Amt. 3" - SCL."Discounted Amt. 4" - SCL."Discounted Amt. 5");

                                END;
                    END;
                END;
            UNTIL SCL.NEXT = 0;

        EXIT(ToT_CaseQty);
    end;

    //
    procedure ShowGroupTotal(Val: Code[20])
    var
        L_Cnt: Integer;
        Tot_Tot: Decimal;
    begin

        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn(Val, FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('CASE', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;
                IF GrpCase[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrpCase[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                Tot_Tot += GrpCase[L_Cnt];
                GrpCase[L_Cnt] := 0;
            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);


        // KG
        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('KG/LITRE', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;

                IF GrpKG[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrpKG[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#.##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                Tot_Tot += GrpKG[L_Cnt];
                GrpKG[L_Cnt] := 0;


            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);


        //Dollar

        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('$', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //ExcelBuffer.NewRow;
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;

                IF GrpDollar[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrpDollar[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                Tot_Tot += GrpDollar[L_Cnt];
                GrpDollar[L_Cnt] := 0;
            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);
    end;

    //
    procedure ShowGrandTotal()
    var
        L_Cnt: Integer;
        Tot_Tot: Decimal;
    begin
        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('GRAND', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('CASE', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;
                IF GrandCase[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrandCase[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                Tot_Tot += GrandCase[L_Cnt];
            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);


        // KG
        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('KG/LITRE', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;

                IF GrandKG[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrandKG[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                Tot_Tot += GrandKG[L_Cnt];
            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);


        //Dollar

        L_Cnt := 0;
        Tot_Tot := 0;
        ExcelBuffer.NewRow;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('$', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //ExcelBuffer.NewRow;
        IF TempDimVal.FINDFIRST THEN
            REPEAT
                L_Cnt += 1;

                IF GrandDollar[L_Cnt] <> 0 THEN
                    ExcelBuffer.AddColumn(GrandDollar[L_Cnt], FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0)
                ELSE
                    ExcelBuffer.AddColumn('-', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                Tot_Tot += GrandDollar[L_Cnt];
            UNTIL TempDimVal.NEXT = 0;
        ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuffer.AddColumn(Tot_Tot, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);
    end;

    //
    procedure InitDate()
    begin
        PostingDateFilter := FORMAT(TODAY);
    end;
}

