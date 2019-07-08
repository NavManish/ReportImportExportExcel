table 50101 CustomerNew
{
    DataClassification = CustomerContent;

    fields
    {
        field(1; CustCode; Code[10])
        {
            DataClassification = CustomerContent;
            //TableRelation = Customer."No.";
            TableRelation = Customer."No.";
            trigger OnValidate();
            begin
                Customer_gRec.Reset();
                Customer_gRec.SetRange("No.", CustCode);
                if Customer_gRec.FindFirst() then begin
                    CustName := Customer_gRec.Name;
                end;
            end;

        }
        field(2; CustName; Text[100])
        {
            DataClassification = CustomerContent;
        }
        field(3; CustBalance; Decimal)
        {
            FieldClass = FlowField;
            CalcFormula = Sum ("Detailed Cust. Ledg. Entry".Amount WHERE ("Customer No." = field (CustCode)));
        }
    }

    keys
    {
        key(PK; CustCode)
        {
            Clustered = true;
        }
    }

    var
        Customer_gRec: Record Customer;

    trigger OnInsert();
    begin
    end;

    trigger OnModify();
    begin
    end;

    trigger OnDelete();
    begin
    end;

    trigger OnRename();
    begin
    end;

}

page 50103 CustomerCardNew
{
    PageType = Card;
    SourceTable = CustomerNew;

    layout
    {
        area(content)
        {
            group(General)
            {
                field(CustomerCode; CustCode)
                {
                  ApplicationArea = All;
                }
                field(CustomerName; CustName)
                {
                  ApplicationArea = All;
                }
                field(Balance; CustBalance)
                {
                  ApplicationArea = All;
                }
            }
        }
    }

    actions
    {

        area(Processing)
        {

            action("Customer report excel")
            {
                Image = Customer;

                trigger OnAction();
                begin
                    CustNew_gRec.reset();
                    CustNew_gRec.SetRange(CustNew_gRec.CustCode, CustCode);
                    if CustNew_gRec.FindFirst() then begin
                        CustomerReportExcel_gRep.SetTableView(CustNew_gRec);
                        CustomerReportExcel_gRep.Run();
                    END;
                end;
            }
            action("Customer Report PDF")
            {
                Image = SendAsPDF;
                trigger OnAction();
                begin

                end;
            }
        }
    }


    var
        myInt: Integer;
        msg: TextConst ENU = 'hello';
        CustomerReportExcel_gRep: Report CustomerreportExcel;
        CustNew_gRec: Record CustomerNew;


}

page 50104 CustomerNew
{

    //ApplicationArea= Purchasing;
    UsageCategory = Lists;
    PageType = List;
    SourceTable = CustomerNew;
    CardPageId = 50103;

    layout
    {
        area(content)
        {
            repeater(Group)
            {
                field(CustomerNo; CustCode)
                {

                }
                field(CustomerName; CustName)
                {

                }
                field(Balance; CustBalance)
                {

                }
            }
        }
        area(factboxes)
        {
        }
    }

    actions
    {
        area(processing)
        {
            action(ActionName)
            {
                trigger OnAction();
                begin
                end;
            }
        }
    }
}
profile AL
{
    Description = 'AL Objects';
    RoleCenter = "Order Processor Role Center";
    Customizations =;
}