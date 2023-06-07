report 50103 "ABC_Purchase - Order"
{
    DefaultRenderingLayout = WordLayout;
    PreviewMode = PrintLayout;
    WordMergeDataItem = CopyLoop;
    Caption = 'Purchase - Order';

    dataset
    {
        dataitem(CopyLoop; Integer)
        {
            DataItemTableView = sorting(Number);
            
            dataitem(Header; "Purchase Header")
            {
                DataItemTableView = sorting("Document Type", "No.") where("Document Type" = const(Order));
                RequestFilterFields = "No.", "Buy-from Vendor No.", "No. Printed";
                RequestFilterHeading = 'Purchase Order';
                
                column(No_;"No.") { IncludeCaption = true; }
                column(Your_Reference;"Your Reference") { IncludeCaption = true; }
                column(Posting_Date; Format("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Posting_Date_Lbl; Header.FieldCaption("Posting Date")) {}
                column(Document_Date; Format("Document Date", 0, '<Day,2>/<Month,2>/<Year4>')) {}
                column(Document_Date_Lbl; Header.FieldCaption("Document Date")) {}
                column(Buy_from_Address;"Buy-from Address") { IncludeCaption = true; }
                column(Buy_from_Address_2;"Buy-from Address 2") { IncludeCaption = true; }
                column(Buy_from_City;"Buy-from City") { IncludeCaption = true; }
                column(Buy_from_County;"Buy-from County") { IncludeCaption = true; }
                column(Buy_from_Post_Code;"Buy-from Post Code") { IncludeCaption = true; }
                column(Buy_from_Vendor_Name;"Buy-from Vendor Name") { IncludeCaption = true; }
                column(Buy_from_Vendor_No_;"Buy-from Vendor No.") { IncludeCaption = true; }
                column(Buy_from_Country_Region_Code;"Buy-from Country/Region Code") { IncludeCaption = true; }
                column(Buy_from_Country_Region_Name;BillToCountryName) {}
                column(VAT_Registration_No_;"VAT Registration No.") { IncludeCaption = true; }
                column(Vendor_Order_No_;"Vendor Order No.") { IncludeCaption = true; }
                column(Order_Date;"Order Date") { IncludeCaption = true; }
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
                column(CompanySWIFT; CompanyInfo."SWIFT Code") {}
                column(CompanySWIFT_Lbl; CompanyInfo.FieldCaption("SWIFT Code")) {}
                column(CompanyVATRegistrationNo; CompanyInfo."VAT Registration No.") {}
                column(CompanyVATRegistrationNo_Lbl; CompanyInfo.FieldCaption("VAT Registration No.")) {}
                column(PaymentTermsDescription; PaymentTermsDescription) {}
                column(PaymentTerms_Lbl; PaymentTerms.TableCaption()) {}
                column(PaymentMethodDescription; PaymentMethodDescription) {}
                column(PaymentMethod_Lbl; PaymentMethod.TableCaption()) {}
                column(CurrencySymbol; CurrencySymbol) {}
                
                dataitem("Purchase Line"; "Purchase Line")
                {
                    DataItemLink = "Document Type" = field("Document Type"), "Document No." = field("No.");
                    DataItemTableView = sorting("Document Type", "Document No.", "Line No.");
                    
                    column(No_Line; No_Line) {}
                    column(No_Line_Lbl; FieldCaption("No.")) {}
                    column(Description_Line; Description) { IncludeCaption = true; }
                    column(Quantity_Line; Quantity_Line) {}
                    column(Quantity_Line_Lbl; FieldCaption("Quantity")) {}
                    column(DirUnitCost_Line; FormattedDirectUnitCost) {}
                    column(DirUnitCost_Line_Lbl; FieldCaption("Direct Unit Cost")) {}
                    column(LineDisc_Line; LineDiscountTxt) {}
                    column(LineDisc_Line_Lbl; FieldCaption("Line Discount %")) {}
                    column(LineAmt_Line; LineAmount_Line) {}
                    column(LineAmt_Line_Lbl; FieldCaption("Line Amount")) {}
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
                            
                        FormattedDirectUnitCost := '';
                        if "Direct Unit Cost" <> 0 then begin
                            FormattedDirectUnitCost := Format("Direct Unit Cost", 0, AutoFormat.ResolveAutoFormat(Enum::"Auto Format"::UnitAmountFormat, Header."Currency Code"));
                            CurrencySymbol_Line := CurrencySymbol;
                        end;                            
                            
                        Quantity_Line := '';
                        if Quantity <> 0 then
                            Quantity_Line := Format(Quantity);
                        
                        LineAmount_Line := '';
                        if "Line Amount" <> 0 then    
                            LineAmount_Line := Format("Line Amount", 0, AutoFormat.ResolveAutoFormat(Enum::"Auto Format"::AmountFormat, Header."Currency Code"));
                        
                        TotalInvDiscAmount -= "Inv. Discount Amount";
                        TotalSubTotal += "Line Amount";
                        TotalAmount += Amount;
                        TotalVATAmount += "Amount Including VAT" - Amount;
                        TotalAmountInclVAT += "Amount Including VAT";
                    end;
                }
                dataitem(Totals; "Integer")
                {
                    DataItemTableView = sorting(Number) where(Number = const(1));
                    
                    column(TotalSubTotal; TotalSubTotal) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalAmount; TotalAmount) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalVATAmount; TotalVATAmount) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(TotalAmountInclVAT; TotalAmountInclVAT) { AutoFormatExpression = Header."Currency Code"; AutoFormatType = 1; }
                    column(CurrencySymbol_Totals; CurrencySymbol) {}
                }

                trigger OnAfterGetRecord()
                begin
                    TotalSubTotal := 0;
                    TotalInvDiscAmount := 0;
                    TotalAmount := 0;
                    TotalVATAmount := 0;
                    TotalAmountInclVAT := 0;
                    
                    CurrencySymbol := RetrieveCurrencySymbol("Currency Code");
                    
                    BillToCountryName := '';
                    if CountryRegion.Get(Header."Buy-from Country/Region Code") then
                        BillToCountryName := CountryRegion.Name;
                        
                    PaymentTermsDescription := '';
                    if PaymentTerms.Get(Header."Payment Terms Code") then
                        PaymentTermsDescription := PaymentTerms.Description;
                        
                    PaymentMethodDescription := '';
                    if PaymentMethod.Get(Header."Payment Method Code") then
                        PaymentMethodDescription := PaymentMethod.Description;
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
            Caption = 'Purchase - Order';
            LayoutFile = './PurchaseOrder.docx';
        }
        layout(WordLayoutWithBars)
        {
            Type = Word;
            Caption = 'Purchase - Order';
            LayoutFile = './PurchaseOrderWithBars.docx';
        }
    }
    
    labels
    {
        PageLbl = 'Page';
        BaseLbl = 'Amount';
        VatAmount_Lbl = 'VAT';
        InvoiceLbl = 'Purchase order';
        OtherTaxesLbl = 'Other taxes';
        TotalLbl = 'Total';
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
        CompanyInfo: Record "Company Information";
        NoOfCopies: Integer;
        TotalAmountInclVAT, TotalAmount, TotalSubTotal, TotalInvoiceDiscountAmount, TotalInvDiscAmount, TotalVATAmount: Decimal;
        LineDiscountTxt, BillToCountryName, Quantity_Line, No_Line, FormattedDirectUnitCost, PaymentTermsDescription: Text;
        PaymentMethodDescription, SalespersonPurchaserName, LineAmount_Line, CompanyInfoCountryName: Text;
        CurrencySymbol, CurrencySymbol_Line: Code[10];
        OldHeader, RunningHeader: Record "Purchase Header";
        GeneralLedgerSetup: Record "General Ledger Setup";
        CountryRegion: Record "Country/Region";
        PaymentMethod: Record "Payment Method";
        PaymentTerms: Record "Payment Terms";
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