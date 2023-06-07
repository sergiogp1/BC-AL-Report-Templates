report 50102 "ABC_Sales - Invoice"
{
    Caption = 'Sales - Invoice';
    DefaultRenderingLayout = WordLayout;
    PreviewMode = PrintLayout;
    WordMergeDataItem = CopyLoop;

    dataset
    {
        dataitem(CopyLoop; Integer)
        {
            DataItemTableView = sorting(Number);
            
            dataitem(Header; "Sales Invoice Header")
            {
                DataItemTableView = sorting("No.");
                RequestFilterFields = "No.", "Sell-to Customer No.";
                RequestFilterHeading = 'Posted Sales Invoice';
                
                column(No_; "No.") { IncludeCaption = true; }
                column(Posting_Date; Format("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Posting_Date_Lbl; Header.FieldCaption("Posting Date")) {}
                column(Due_Date; Format("Due Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Due_Date_Lbl; Header.FieldCaption("Due Date")) {}
                column(Order_No_;"Order No.") { IncludeCaption = true; }
                column(Your_Reference; "Your Reference") { IncludeCaption = true; }
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
                column(PaymentTermsDescription; PaymentTermsDescription) {}
                column(PaymentTerms_Lbl; PaymentTerms.TableCaption()) {}
                column(PaymentMethodDescription; PaymentMethodDescription) {}
                column(PaymentMethod_Lbl; PaymentMethod.TableCaption()) {}
                column(SalesPersonName; SalespersonPurchaserName) {}
                column(SalesPerson_Lbl; SalespersonPurchaser.TableCaption()) {}
                column(Remaining_Amount; "Remaining Amount") { IncludeCaption = true; }
                column(CurrencySymbol; CurrencySymbol) {}
                dataitem(Line; "Sales Invoice Line")
                {
                    DataItemLinkReference = Header;
                    DataItemLink = "Document No." = field("No.");
                    DataItemTableView = sorting("Document No.", "Line No.");
                    
                    column(Description_Line; Description) { IncludeCaption = true; }
                    column(LineDiscountPercent_Line; LineDiscountTxt) {}
                    column(LineDiscountPercent_Line_Lbl; Line.FieldCaption("Line Discount %")) {}
                    column(LineAmount_Line; LineAmount_Line) {}
                    column(LineAmount_Line_Lbl; FieldCaption("Line Amount")) {}
                    column(No_Line; No_Line) {}
                    column(No_Line_Lbl; FieldCaption("No.")) {}
                    column(Quantity_Line; Quantity_Line) {}
                    column(Quantity_Line_Lbl; FieldCaption(Quantity)) {}
                    column(Unit_Price_Line; UnitPrice_Line) {}
                    column(Unit_Price_Line_Lbl; FieldCaption("Unit Price")) {}
                    column(CurrencySymbol_Line; CurrencySymbol_Line) {}

                    trigger OnAfterGetRecord()
                    begin
                        CurrencySymbol_Line := '';
                                                
                        No_Line := '';
                        if (Type <> Type::"G/L Account") and (Type <> Type::" ") then
                            No_Line := "No.";
                            
                        LineDiscountTxt := '';
                        if (Type <> Type::" ") and ("Line Discount %" <> 0) then
                            LineDiscountTxt := Format("Line Discount %") + '%';
                                                
                        Quantity_Line := '';
                        if Quantity <> 0 then
                            Quantity_Line := Format(Quantity);
                            
                        UnitPrice_Line := '';
                        if "Unit Price" <> 0 then begin
                            UnitPrice_Line := Format("Unit Price", 0, AutoFormat.ResolveAutoFormat(Enum::"Auto Format"::UnitAmountFormat, Header."Currency Code"));
                            CurrencySymbol_Line := CurrencySymbol;
                        end;
                        
                        LineAmount_Line := '';
                        if "Line Amount" <> 0 then    
                            LineAmount_Line := Format("Line Amount", 0, AutoFormat.ResolveAutoFormat(Enum::"Auto Format"::AmountFormat, Header."Currency Code"));

                        VATAmountLine.Init();
                        VATAmountLine."VAT Identifier" := Format("VAT %");
                        VATAmountLine."VAT Calculation Type" := "VAT Calculation Type";
                        VATAmountLine."Tax Group Code" := "Tax Group Code";
                        VATAmountLine."VAT %" := "VAT %";
                        VATAmountLine."EC %" := "EC %";
                        VATAmountLine."VAT Base" := Amount;
                        VATAmountLine."Amount Including VAT" := "Amount Including VAT";
                        VATAmountLine."Line Amount" := Abs("Line Amount");
                        if "Allow Invoice Disc." then
                            VATAmountLine."Inv. Disc. Base Amount" := "Line Amount";
                        VATAmountLine."Invoice Discount Amount" := "Inv. Discount Amount";
                        VATAmountLine."VAT Clause Code" := "VAT Clause Code";
                        VATAmountLine.InsertLine();

                        TotalInvDiscAmount -= "Inv. Discount Amount";
                        TotalSubTotal += "Line Amount";
                        TotalAmount += Amount;
                        TotalVATAmount += "Amount Including VAT" - Amount;
                        TotalAmountInclVAT += "Amount Including VAT";
                    end;

                    trigger OnPreDataItem()
                    begin
                        VATAmountLine.DeleteAll();
                    end;
                }
                dataitem(VATAmountLine; "VAT Amount Line")
                {
                    DataItemTableView = sorting("VAT Identifier", "VAT Calculation Type", "Tax Group Code", "Use Tax", Positive);
                    UseTemporary = true;
                    
                    column(InvoiceDiscountAmount_VATAmountLine; "Invoice Discount Amount")
                    {
                        AutoFormatExpression = Header."Currency Code";
                        AutoFormatType = 1;
                        IncludeCaption = true;
                    }
                    column(VATAmount_VatAmountLine; "VAT Amount")
                    {
                        AutoFormatExpression = Header."Currency Code";
                        AutoFormatType = 1;
                        IncludeCaption = true;
                    }
                    column(VATBase_VatAmountLine; "VAT Base")
                    {
                        AutoFormatExpression = Header."Currency Code";
                        AutoFormatType = 1;
                        IncludeCaption = true;
                    }
                    column(VATPct_VatAmountLine; "VAT %")
                    {
                        DecimalPlaces = 0 : 5;
                        AutoFormatExpression = Header."Currency Code";
                        AutoFormatType = 1;
                        IncludeCaption = true;
                    }
                    column(Total_VatAmountLine; "VAT Base" + "VAT Amount")
                    {
                        DecimalPlaces = 0 : 5;
                        AutoFormatExpression = Header."Currency Code";
                        AutoFormatType = 1;
                    }
                    column(CurrencySymbol_VATAmountLine; CurrencySymbol) {}
                }
                dataitem(Totals; "Integer")
                {
                    DataItemTableView = sorting(Number) where(Number = const(1));
                    
                    column(TotalSubTotal; TotalSubTotal) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalInvoiceDiscountAmount; TotalInvDiscAmount) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalAmount; TotalAmount) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalVATAmount; TotalVATAmount) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalAmountIncludingVAT; TotalAmountInclVAT) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(CurrencySymbol_Totals; CurrencySymbol) {}
                }
                
                trigger OnPreDataItem()
                begin
                    SetAutoCalcFields("Remaining Amount");
                    SetRange("No.", RunningHeader."No.");
                end;

                trigger OnAfterGetRecord()
                begin
                    TotalSubTotal := 0;
                    TotalInvDiscAmount := 0;
                    TotalAmount := 0;
                    TotalVATAmount := 0;
                    TotalAmountInclVAT := 0;
                    
                    CurrencySymbol := RetrieveCurrencySymbol("Currency Code");

                    BillToCountryName := '';
                    if CountryRegion.Get(Header."Bill-to Country/Region Code") then
                        BillToCountryName := CountryRegion.Name;
                        
                    PaymentTermsDescription := '';
                    if PaymentTerms.Get(Header."Payment Terms Code") then
                        PaymentTermsDescription := PaymentTerms.Description;
                        
                    PaymentMethodDescription := '';
                    CompanyIBAN := '';
                    if PaymentMethod.Get(Header."Payment Method Code") then begin
                        PaymentMethodDescription := PaymentMethod.Description;
                        
                        if PaymentMethod."Collection Agent" = PaymentMethod."Collection Agent"::Bank then
                            CompanyIBAN := DepositInLbl + CompanyInfo.IBAN;
                    end;
                        
                    SalespersonPurchaserName := '';
                    if SalespersonPurchaser.Get(Header."Salesperson Code") then
                        SalespersonPurchaserName := SalespersonPurchaser.Name;
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
            Caption = 'Sales - Invoice';
            LayoutFile = './SalesInvoice.docx';
        }
        layout(WordLayoutWithBars)
        {
            Type = Word;
            Caption = 'Sales - Invoice';
            LayoutFile = './SalesInvoiceWithBars.docx';
        }
    }
    
    labels
    {
        PageLbl = 'Page';
        Total_VatAmountLine_Lbl = 'Total';
        InvoiceLbl = 'Sales invoice';
        OtherTaxesLbl = 'Other taxes';
        VATBreakdownLbl = 'VAT breakdown';
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
        DepositInLbl: Label 'Deposit in ';
        PaymentTerms: Record "Payment Terms";
        PaymentMethod: Record "Payment Method";
        SalespersonPurchaser: Record "Salesperson/Purchaser";
        CompanyInfo: Record "Company Information";
        CountryRegion: Record "Country/Region";
        GeneralLedgerSetup: Record "General Ledger Setup";
        OldHeader, RunningHeader: Record "Sales Invoice Header";
        Quantity_Line, LineAmount_Line, UnitPrice_Line, No_Line, PaymentTermsDescription, PaymentMethodDescription, CompanyIBAN: Text;
        TotalSubTotal, TotalAmount, TotalAmountInclVAT, TotalVATAmount, TotalInvDiscAmount: Decimal;
        CurrencySymbol, CurrencySymbol_Line: Code[10];
        SalespersonPurchaserName, LineDiscountTxt, BillToCountryName, CompanyInfoCountryName: Text;
        NoOfCopies: Integer;
        AutoFormat: Codeunit "Auto Format";

    local procedure RetrieveCurrencySymbol(CurrencyCode: Code[10]): Code[10]
    var
        Currency: Record Currency;
    begin
        if (CurrencyCode = '') or (CurrencyCode = GeneralLedgerSetup."LCY Code") then
            exit(GeneralLedgerSetup.GetCurrencySymbol());

        if Currency.Get(CurrencyCode) then
            exit(Currency.GetCurrencySymbol());

        exit(CurrencyCode);
    end;
}