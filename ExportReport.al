report 50102 CustomerreportExcel
{
    ProcessingOnly = true;
    UseRequestPage = false;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All, Basic, Suite;

    dataset
    {

        dataitem(CustomerNew; CustomerNew)
        {
            DataItemTableView = sorting (CustCode) where (Custcode = filter (<> ''));
            RequestFilterFields = CustCode;
            trigger OnPreDataItem();
            var
            //Excel : Integer;
            begin
                ExcelBuffer_gRec.DeleteAll();
                ExcelBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn('Customer Balance Sheet', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn('Customer Code', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn('Customer Name', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn('Customer Balance', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                //ExcelBuffer_gRec.NewRow;
            end;

            trigger OnAfterGetRecord();
            begin
                CustomerNew.CalcFields(CustBalance);
                ExceLBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn(CustCode, false, '', False, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn(CustName, false, '', False, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn(CustBalance, false, '', false, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);

            end;

            trigger OnPostDataItem();
            begin


            end;
        }


    }



    var
        ExcelBuffer_gRec: Record "Excel Buffer" temporary;
        //GeneralLedgersetupExt_gRec: Record "General Ledger Setup Extension";

    trigger OnInitReport();
    begin

    end;

    trigger OnPreReport();
    begin
        ExcelBuffer_gRec.DeleteAll;
    end;

    trigger OnPostReport();
    begin
        //GeneralLedgersetupExt_gRec.get;
        ExcelBuffer_gRec.CreateNewBook('Customer');
        ExcelBuffer_gRec.WriteSheet('Customer', CompanyName, UserId);
        ExcelBuffer_gRec.CloseBook;
        ExcelBuffer_gRec.OpenExcel;
    end;
}

