report 60101 "Import RFQ From Excel"
{

    ProcessingOnly = true;

    dataset
    {
        dataitem(DataItem1; "Purchase Header")
        {
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
                    field(FileName; FileName)
                    {
                        Caption = 'WorkBook FileName';
                        AssistEdit = true;
                        trigger OnAssistEdit();
                        begin
                            TempExcelBuffer.Init();
                            //UploadFile_GFnc;

                            If UploadIntoStream('My Upload Dialog', 'C:\TEMP', 'All Files (*.*)|*.*', FileName, exInstream) then begin
                                //TempExcelBuffer.OpenBook(FileName, 'RFQ');
                                TempExcelBuffer.OpenBookStream(exInstream, TempExcelBuffer.SelectSheetsNameStream(exInstream));
                                TempExcelBuffer.ReadSheet();
                            end;
                        end;
                    }
                    field(SheetName; SheetName)
                    {
                        Caption = 'WorkBook SheetName';
                        AssistEdit = true;
                        Visible = false;

                        trigger OnAssistEdit();
                        begin
                            //SheetName := TempExcelBuffer.SelectSheetsNameStream(exInstream);
                            //SheetName := TempExcelBuffer.SelectSheetsName(FileName);
                        end;
                    }
                    field(QuotationNo; QuotationNo)
                    {
                        Caption = 'Quotation No.';
                    }
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPreReport();
    var
        ExcelBuffer_LRec: Record "Excel Buffer";
        Text001_LTxt: Label 'Insert Requisition no. for all lines.';
        Text002_LTxt: Label 'You cant import excel for requisition no. %1 in requisition No. %2';
    begin
        //IF NOT EXISTS(FileName) OR (SheetName = '') THEN
        //    ERROR(Text002_GTxt);

        TempExcelBuffer.RESET;
        TempExcelBuffer.LOCKTABLE;
        TempExcelBuffer.SETCURRENTKEY("Row No.", "Column No.");
        HeaderExists := TRUE;

        IF NOT TempExcelBuffer.GET(1, 4) OR (TempExcelBuffer."Cell Value as Text" <> Text003_GTxt) THEN
            HeaderExists := FALSE;

        IF NOT HeaderExists THEN
            ERROR(Text005_GTxt);
        ImportPurchOrderData_GFnc;
        COMMIT;
        MESSAGE(Text006_GTxt, PurchHeader."No.");
    end;

    var
        Txt001: Label 'Request for Quotation';
        Txt002: Label 'Vendor No.';
        Txt003: Label 'Vendor Name';
        Txt004: Label 'Requested Terms';
        Txt005: Label 'Department';
        Txt006: Label '"Phone "';
        Txt007: Label 'Quote Due By:';
        Txt008: Label 'All vendor quantities are subject to the provisions of the attached General Contract Terms and Conditions (Form R-1). Vendor quotations will not be accepted unless all requested information has been provided and the quotation is signed';
        Txt009: Label 'Shipping Address';
        Txt010: Label 'Import Cancelled as precess terminated by user.';
        PmtTerms_Caption: Label 'Payment Terms';
        ShipVia_Caption: Label 'Ship VIA';
        FOB_Caption: Label 'F.O.B.';
        FreightTerms_Caption: Label 'Freight Terms';
        Delivery_Caption: Label 'Delivery';
        QuotedBy_Caption: Label 'Quoted By:';
        Phone_Caption: Label 'Phone:';
        Date_Caption: Label 'Date:';
        FileName: Text[1024];
        SheetName: Text[1024];
        ImportFilePath: Text[250];
        ServerFileName: Text[250];
        CommonDialogMgt: Codeunit "File Management";
        Text001_GTxt: Label 'Import Excel File';
        Text002_GTxt: Label 'Please specify a valid file name and worksheet name for RFQ import.';
        TempExcelBuffer: Record "Excel Buffer";
        HeaderExists: Boolean;
        Text003_GTxt: Label 'Request for Quotation';
        Text004_GTxt: Label 'TO';
        Text005_GTxt: Label 'The selected input file with RFQ Import data has uncommon layout. Please select a relevant file for import.';
        Window: Dialog;
        Text006_GTxt: Label 'Purchase Quote Created For No. : %1';
        Maxi: Integer;
        Maxj: Integer;
        Text009_GTxt: Label 'Creating Purchase Entries.....';
        i: Integer;
        PurchHeader: Record "Purchase Header";
        PurchLines: Record "Purchase Line";
        QuotationNo: Option "Quotation 1","Quotation 2","Quotation 3","Quotation 4","Quotation 5";
        ImportInReqNo: Code[20];
        RecNoSeries: Record "No. Series";
        Text50001: Label 'Please Enter the No series No.';
        exInstream: InStream;

    procedure UploadFile_GFnc();
    begin
        ImportFilePath := CommonDialogMgt.OpenFileDialog(Text001_GTxt, '', '');
        IF ImportFilePath <> '' THEN BEGIN
            FileName := CommonDialogMgt.GetFileName(ImportFilePath);
            ServerFileName := CommonDialogMgt.UploadFileSilent(ImportFilePath);
            FileName := ServerFileName;
        END;
    end;

    procedure ImportPurchOrderData_GFnc();
    var
        PurchSetup: Record "Purchases & Payables Setup";
        NoSeriesMgt: Codeunit NoSeriesManagement;
        TxtErr: Label 'Please Specify Vendor No. in Excel Sheet';
        LineNo: Integer;
        ReqNo: Code[20];
        ReqLineNo: Integer;
        Qty: Decimal;
        DirectUnitCost: Decimal;
        Amount: Decimal;
        TxtErr1: Label 'Please Specify Job No. in Excel Sheet';
        ReqLines: Record "Requisition Lines";
    begin
        PurchSetup.GET;
        i := 17;
        IF TempExcelBuffer.GET(i, 5) THEN BEGIN
            LineNo := 0;
            PurchHeader.INIT;
            PurchHeader.VALIDATE("Document Type", PurchHeader."Document Type"::Quote);
            PurchHeader."No." := NoSeriesMgt.GetNextNo(PurchSetup."Quote Nos.", WORKDATE, TRUE);
            IF TempExcelBuffer.GET(7, 4) THEN
                PurchHeader.VALIDATE("Buy-from Vendor No.", TempExcelBuffer."Cell Value as Text")
            ELSE
                ERROR(TxtErr);

            PurchHeader.VALIDATE("Order Date", WORKDATE);
            PurchHeader.VALIDATE("Posting Date", WORKDATE);

            PurchHeader.INSERT;

            IF TempExcelBuffer.GET(9, 4) THEN BEGIN
                PurchHeader.VALIDATE("Shortcut Dimension 1 Code", TempExcelBuffer."Cell Value as Text");
                PurchHeader.MODIFY;
            END;
            REPEAT
                PurchLines.INIT;
                PurchLines.VALIDATE("Document Type", PurchHeader."Document Type");
                PurchLines.VALIDATE("Document No.", PurchHeader."No.");
                LineNo += 10000;
                PurchLines.VALIDATE("Line No.", LineNo);
                PurchLines.VALIDATE("Buy-from Vendor No.", PurchHeader."Buy-from Vendor No.");
                PurchLines.VALIDATE(Type, PurchLines.Type::Item);
                IF TempExcelBuffer.GET(i, 3) THEN
                    PurchLines.VALIDATE("No.", TempExcelBuffer."Cell Value as Text");
                IF TempExcelBuffer.GET(i, 5) THEN BEGIN
                    EVALUATE(Qty, TempExcelBuffer."Cell Value as Text");
                    PurchLines.VALIDATE(Quantity, Qty);
                END;
                IF TempExcelBuffer.GET(i, 6) THEN
                    PurchLines.VALIDATE("Unit of Measure Code", TempExcelBuffer."Cell Value as Text");
                IF TempExcelBuffer.GET(i, 10) THEN
                    PurchLines.VALIDATE("Item Category Code", TempExcelBuffer."Cell Value as Text");
                IF TempExcelBuffer.GET(i, 8) THEN
                    PurchLines.VALIDATE("Variant Code", TempExcelBuffer."Cell Value as Text");
                IF TempExcelBuffer.GET(i, 7) THEN BEGIN
                    EVALUATE(DirectUnitCost, TempExcelBuffer."Cell Value as Text");
                    PurchLines.VALIDATE("Direct Unit Cost", DirectUnitCost);
                END;
                //IF TempExcelBuffer.GET(i, 14) THEN BEGIN
                //    EVALUATE(Amount, TempExcelBuffer."Cell Value as Text");
                //    PurchLines.VALIDATE(Amount, Amount);
                //END;
                IF TempExcelBuffer.GET(i, 1) THEN
                    ReqNo := TempExcelBuffer."Cell Value as Text";
                IF TempExcelBuffer.GET(i, 2) THEN
                    EVALUATE(ReqLineNo, TempExcelBuffer."Cell Value as Text");
                PurchLines.VALIDATE("Int Source No.", ReqNo);
                PurchLines.VALIDATE("Int Source Line No.", ReqLineNo);
                PurchLines.INSERT;

                ReqLines.RESET;
                ReqLines.SETRANGE("Requisition No.", ReqNo);
                ReqLines.SETRANGE("Line No.", ReqLineNo);
                IF ReqLines.FINDFIRST THEN BEGIN
                    ModifyReqLines(ReqLines, QuotationNo);
                END;
                i += 1;
            UNTIL NOT TempExcelBuffer.GET(i, 5);
        END;
    end;

    procedure ModifyReqLines(var RequisitionLines: Record "Requisition Lines"; QuotationNo: Option "1","2","3","4","5");
    var
        Text001_LTxt: Label 'Do you want to change the %1 in Requisition No. %2 Line No. %3';
    begin
        IF QuotationNo = 0 THEN BEGIN
            IF RequisitionLines."Quotaion No. 1" <> '' THEN BEGIN
                IF CONFIRM(Text001_LTxt, FALSE, RequisitionLines.FIELDCAPTION("Quotaion No. 1"), RequisitionLines."Requisition No.", RequisitionLines."Line No.") THEN BEGIN
                    RequisitionLines."Quotaion No. 1" := PurchLines."Document No.";
                    RequisitionLines.MODIFY;
                END ELSE
                    ERROR(Txt010);
            END ELSE BEGIN
                RequisitionLines."Quotaion No. 1" := PurchLines."Document No.";
                RequisitionLines.MODIFY;
            END;
        END;

        IF QuotationNo = 1 THEN BEGIN
            IF RequisitionLines."Quotaion No. 2" <> '' THEN BEGIN
                IF CONFIRM(Text001_LTxt, FALSE, RequisitionLines.FIELDCAPTION("Quotaion No. 2"), RequisitionLines."Requisition No.", RequisitionLines."Line No.") THEN BEGIN
                    RequisitionLines."Quotaion No. 2" := PurchLines."Document No.";
                    RequisitionLines.MODIFY;
                END ELSE
                    ERROR(Txt010);
            END ELSE BEGIN
                RequisitionLines."Quotaion No. 2" := PurchLines."Document No.";
                RequisitionLines.MODIFY;
            END;
        END;

        IF QuotationNo = 2 THEN BEGIN
            IF RequisitionLines."Quotaion No. 3" <> '' THEN BEGIN
                IF CONFIRM(Text001_LTxt, FALSE, RequisitionLines.FIELDCAPTION("Quotaion No. 3"), RequisitionLines."Requisition No.", RequisitionLines."Line No.") THEN BEGIN
                    RequisitionLines."Quotaion No. 3" := PurchLines."Document No.";
                    RequisitionLines.MODIFY;
                END ELSE
                    ERROR(Txt010);
            END ELSE BEGIN
                RequisitionLines."Quotaion No. 3" := PurchLines."Document No.";
                RequisitionLines.MODIFY;
            END;
        END;

        IF QuotationNo = 3 THEN BEGIN
            IF RequisitionLines."Quotaion No. 4" <> '' THEN BEGIN
                IF CONFIRM(Text001_LTxt, FALSE, RequisitionLines.FIELDCAPTION("Quotaion No. 4"), RequisitionLines."Requisition No.", RequisitionLines."Line No.") THEN BEGIN
                    RequisitionLines."Quotaion No. 4" := PurchLines."Document No.";
                    RequisitionLines.MODIFY;
                END ELSE
                    ERROR(Txt010);
            END ELSE BEGIN
                RequisitionLines."Quotaion No. 4" := PurchLines."Document No.";
                RequisitionLines.MODIFY;
            END;
        END;

        IF QuotationNo = 4 THEN BEGIN
            IF RequisitionLines."Quotaion No. 5" <> '' THEN BEGIN
                IF CONFIRM(Text001_LTxt, FALSE, RequisitionLines.FIELDCAPTION("Quotaion No. 5"), RequisitionLines."Requisition No.", RequisitionLines."Line No.") THEN BEGIN
                    RequisitionLines."Quotaion No. 5" := PurchLines."Document No.";
                    RequisitionLines.MODIFY;
                END ELSE
                    ERROR(Txt010);
            END ELSE BEGIN
                RequisitionLines."Quotaion No. 5" := PurchLines."Document No.";
                RequisitionLines.MODIFY;
            END;
        END;
    end;

    procedure SetRequisitionNo(RequsitionNo: Code[20]);
    begin
        ImportInReqNo := RequsitionNo;
    end;
}

