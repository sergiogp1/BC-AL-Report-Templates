report 50105 "ABC_Sales - Shipment"
{
    Caption = 'Sales - Shipment';
    DefaultRenderingLayout = WordLayout;
    PreviewMode = PrintLayout;
    WordMergeDataItem = CopyLoop;

    dataset
    {
        dataitem(CopyLoop; Integer)
        {
            DataItemTableView = sorting(Number);
            
            dataitem(Header; "Sales Shipment Header")
            {
                DataItemTableView = sorting("No.");
                RequestFilterFields = "No.", "Sell-to Customer No.";
                RequestFilterHeading = 'Sales - Shipment';
                
                column(No_; "No.") { IncludeCaption = true; }
                column(Posting_Date; Format("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Document_Date; Format("Document Date", 0, '<Day,2>-<Month Text,3>-<Year4>')) {}
                column(Posting_Date_Lbl; Header.FieldCaption("Posting Date")) {}
                column(Due_Date; Format("Due Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Due_Date_Lbl; Header.FieldCaption("Due Date")) {}
                column(Order_No_;"Order No.") { IncludeCaption = true; }
                column(Your_Reference; "Your Reference") { IncludeCaption = true; }
                column(External_Document_No_;"External Document No.") { IncludeCaption = true; }
                column(Sell_to_Customer_No_;"Sell-to Customer No.") { IncludeCaption = true; }
                column(Bill_to_Customer_No_;"Bill-to Customer No.") { IncludeCaption = true; }
                column(Bill_to_Name;"Bill-to Name") { IncludeCaption = true; }
                column(Bill_to_Address;"Bill-to Address") { IncludeCaption = true; }
                column(Bill_to_Address_2;"Bill-to Address 2") { IncludeCaption = true; }
                column(Bill_to_City;"Bill-to City") { IncludeCaption = true; }
                column(Bill_to_County;"Bill-to County") { IncludeCaption = true; }
                column(Bill_to_Post_Code;"Bill-to Post Code") { IncludeCaption = true; }
                column(Bill_to_Country_Region_Code;"Bill-to Country/Region Code") { IncludeCaption = true; }
                column(Bill_to_Country_Region_Name; BillToCountryName) {}
                column(Bill_to_Contact;"Bill-to Contact") { IncludeCaption = true; } 
                column(Ship_to_Name;"Ship-to Name") { IncludeCaption = true; }
                column(Ship_to_Address;"Ship-to Address") { IncludeCaption = true; }
                column(Ship_to_Address_2;"Ship-to Address 2") { IncludeCaption = true; }
                column(Ship_to_City;"Ship-to City") { IncludeCaption = true; }
                column(Ship_to_County;"Ship-to County") { IncludeCaption = true; }
                column(Ship_to_Post_Code;"Ship-to Post Code") { IncludeCaption = true; }
                column(Ship_to_Country_Region_Code;"Ship-to Country/Region Code") { IncludeCaption = true; }
                column(Ship_to_Country_Region_Name; ShipToCountryName) {}
                column(Ship_to_Contact;"Ship-to Contact") { IncludeCaption = true; }
                column(VAT_Registration_No_;"VAT Registration No.") { IncludeCaption = true; }
                column(CompanyName; CompanyInfo.Name) {}
                column(CompanyAddress; CompanyInfo.Address) {}
                column(CompanyAddress2; CompanyInfo."Address 2") {}
                column(CompanyCity; CompanyInfo.City) {}
                column(CompanyPostCode; CompanyInfo."Post Code") {}
                column(CompanyCounty; CompanyInfo.County) {}
                column(CompanyCountryCode; CompanyInfo."Country/Region Code") {}
                column(CompanyCountryName; CompanyInfoCountryName) {}
                column(CompanyHomePage; CompanyInfo."Home Page") {}
                column(CompanyHomePage_Lbl; CompanyInfo.FieldCaption("Home Page")) {}
                column(CompanyEMail; CompanyInfo."E-Mail") {}
                column(CompanyEMail_Lbl; CompanyInfo.FieldCaption("E-Mail")) {}
                column(CompanyPicture; CompanyInfo.Picture) {}
                column(CompanyPhoneNo; CompanyInfo."Phone No.") {}
                column(CompanyPhoneNo_lbl; CompanyInfo.FieldCaption("Phone No.")) {}
                column(CompanyBankName; CompanyInfo."Bank Name") {}
                column(CompanyBankNameLbl; CompanyInfo.FieldCaption("Bank Name")) {}
                column(CompanyBankAccountNo; CompanyInfo."Bank Account No.") {}
                column(CompanyBankAccountNo_Lbl; CompanyInfo.FieldCaption("Bank Account No.")) {}
                column(CompanyIBAN; CompanyIBAN) {}
                column(CompanyIBAN_Lbl; CompanyInfo.FieldCaption(IBAN)) {}
                column(CompanySWIFT; CompanyInfo."SWIFT Code") {}
                column(CompanySWIFT_Lbl; CompanyInfo.FieldCaption("SWIFT Code")) {}
                column(CompanyVATRegistrationNo; CompanyInfo."VAT Registration No.") {}
                column(CompanyVATRegistrationNo_Lbl; CompanyInfo.FieldCaption("VAT Registration No.")) {}
                column(CompanyFaxNo; CompanyInfo."Fax No.") {}
                column(PaymentTermsDescription; PaymentTermsDescription) {}
                column(PaymentTerms_Lbl; PaymentTerms.TableCaption()) {}
                column(PaymentMethodDescription; PaymentMethodDescription) {}
                column(PaymentMethod_Lbl; PaymentMethod.TableCaption()) {}
                column(SalesPersonName; SalespersonPurchaserName) {}
                column(SalesPersonEmail; SalespersonPurchaserEmail) {}
                column(SalesPerson_Lbl; SalespersonPurchaser.TableCaption()) {}
                column(CurrencySymbol; CurrencySymbol) {}
                column(Shipment_Method_Code;"Shipment Method Code") {}
                column(Sell_to_Phone_No_;"Sell-to Phone No.") {}
                column(Shipping_Agent_Code;"Shipping Agent Code") { IncludeCaption = true; }
                dataitem(Line; "Sales Shipment Line")
                {
                    DataItemLinkReference = Header;
                    DataItemLink = "Document No." = field("No.");
                    DataItemTableView = sorting("Document No.", "Line No.");
                    
                    column(Description_Line; Description) { IncludeCaption = true; }
                    column(LineDiscountPercent_Line; LineDiscountTxt) {}
                    column(LineDiscountPercent_Line_Lbl; Line.FieldCaption("Line Discount %")) {}
                    column(LineAmount_Line; LineAmount_Line) {}
                    column(LineAmount_Line_Lbl; FieldCaption("VAT Base Amount")) {}
                    column(No_Line; No_Line) {}
                    column(No_Line_Lbl; FieldCaption("No.")) {}
                    column(Quantity_Line; Quantity_Line) {}
                    column(Quantity_Line_Lbl; FieldCaption(Quantity)) {}
                    column(Unit_Price_Line; UnitPrice_Line) {}
                    column(Unit_Price_Line_Lbl; FieldCaption("Unit Price")) {}
                    column(Unit_Price_With_Disc_Line; UnitPriceWithDisc_Line) { }

                    trigger OnAfterGetRecord()
                    begin
                        No_Line := '';
                        if (Type <> Type::"G/L Account") and (Type <> Type::" ") then
                            No_Line := "No.";
                            
                        LineDiscountTxt := '';
                        if (Type <> Type::" ") and ("Line Discount %" <> 0) then
                            LineDiscountTxt := Format("Line Discount %") + '%';
                                                
                        Quantity_Line := '';
                        if (Quantity <> 0) then
                            Quantity_Line := Format(Quantity);
                    end;
                }
                
                trigger OnPreDataItem()
                begin
                    SetRange("No.", RunningHeader."No.");
                end;

                trigger OnAfterGetRecord()
                begin
                    TotalSubTotal := 0;
                    TotalInvDiscAmount := 0;
                    TotalAmount := 0;
                    TotalVATAmount := 0;
                    TotalAmountInclVAT := 0;
                    TotalPaymentDiscount := 0;

                    BillToCountryName := '';
                    if CountryRegion.Get(Header."Bill-to Country/Region Code") then
                        BillToCountryName := CountryRegion.Name;

                    ShipToCountryName := '';
                    if CountryRegion.Get(Header."Ship-to Country/Region Code") then
                        ShipToCountryName := CountryRegion.Name;
                        
                    PaymentTermsDescription := '';
                    if PaymentTerms.Get(Header."Payment Terms Code") then
                        PaymentTermsDescription := PaymentTerms.Description;
                        
                    PaymentMethodDescription := '';
                    CompanyIBAN := '';
                    if PaymentMethod.Get(Header."Payment Method Code") then
                        PaymentMethodDescription := PaymentMethod.Description;
                        
                    SalespersonPurchaserName := '';
                    if SalespersonPurchaser.Get(Header."Salesperson Code") then begin
                        SalespersonPurchaserName := SalespersonPurchaser.Name;
                        SalespersonPurchaserEmail := SalespersonPurchaser."E-Mail";
                    end;  

                    if Header."Currency Code" = '' then
                        Header."Currency Code" := GeneralLedgerSetup."LCY Code";                      
                end;
            }
            
            trigger OnPreDataItem()
            begin
                SetFilter(Number, '%1..%2', 1, Header.Count() * (NoOfCopies + 1));
                
                OldHeader.Copy(Header);
                RunningHeader.Copy(OldHeader);
            end;
            
            trigger OnAfterGetRecord()
            begin
                if RunningHeader.Next() = 0 then begin                  
                    RunningHeader.Copy(OldHeader);
                    RunningHeader.FindFirst();
                end;  
            end;
        }
    }
    
    requestpage
    {
        layout
        {
            area(Content)
            {
                group(General)
                {
                    field(NoOfCopies2;NoOfCopies)
                    {
                        ApplicationArea = All;
                        Caption = 'No. of copies';
                    }
                }
            }
        }
    }
    
    rendering
    {
        layout(WordLayout)
        {
            Type = Word;
            Caption = 'Sales - Shipment';
            LayoutFile = './src/report/Layout/SalesShipment.docx';
        }
    }
    
    labels
    {
        PageLbl = 'Page';
        Total_VatAmountLine_Lbl = 'Total';
        InvoiceLbl = 'Delivery Note';
    }

    trigger OnPreReport()
    begin
        if Header.GetFilters() = '' then
            Error(NoFilterSetErr);
            
        CompanyInfo.Get();
        CompanyInfo.CalcFields(Picture);
        GeneralLedgerSetup.Get();
        
        CompanyInfoCountryName := '';
        if CountryRegion.Get(CompanyInfo."Country/Region Code") then
            CompanyInfoCountryName := CountryRegion.Name;
    end;

    var
        NoFilterSetErr: Label 'You must specify one or more filters to prevent accidental printing of all documents.';
        PaymentTerms: Record "Payment Terms";
        PaymentMethod: Record "Payment Method";
        SalespersonPurchaser: Record "Salesperson/Purchaser";
        CompanyInfo: Record "Company Information";
        CountryRegion: Record "Country/Region";
        GeneralLedgerSetup: Record "General Ledger Setup";
        OldHeader, RunningHeader: Record "Sales Shipment Header";
        Quantity_Line, LineAmount_Line, UnitPrice_Line, No_Line, PaymentTermsDescription, PaymentMethodDescription, CompanyIBAN: Text;
        TotalSubTotal, TotalAmount, TotalAmountInclVAT, TotalVATAmount, TotalInvDiscAmount, TotalPaymentDiscount, LineAmount: Decimal;
        CurrencySymbol, CurrencySymbol_Line: Code[10];
        SalespersonPurchaserName, LineDiscountTxt, BillToCountryName, ShipToCountryName, CompanyInfoCountryName, SalespersonPurchaserEmail, UnitPriceWithDisc_Line: Text;
        TotalSubTotalTxt, TotalAmountTxt, TotalVATAmountTxt, TotalAmountInclVATTxt: Text;
        NoOfCopies: Integer;
        SalesLine: Record "Sales Line";
}