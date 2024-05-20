report 55000 "Import CardNet Azul Amex"
{
    ApplicationArea = Suite, Basic;
    Caption = 'Import CardNet Azul Amex';
    ProcessingOnly = true;
    UsageCategory = Tasks;
    Permissions =
        tabledata "Dimension Value" = R,
        tabledata Integer = R;

    dataset
    {
        dataitem("Integer"; "Integer")
        {
            DataItemTableView = SORTING(Number) WHERE(Number = CONST(1));

            trigger OnAfterGetRecord()
            var
                Text001: Label 'You must select payment processor type';
                Text002: Label 'You must insert a supplier code';
                Text003: Label 'You must insert an account number';
            begin
                if TipoProcesadorPago = TipoProcesadorPago::" " then
                    Error(Text001);

                if TipoProcesadorPago <> TipoProcesadorPago::Amex then begin

                    if CodigoProveedorFactura = '' then
                        Error(Text002);

                    if NumeroCuentaFactura = '' then
                        Error(Text002);

                    /*if dim1 = '' then
                        Error(Text002);*/


                end;


                if CodigoProveedorFactura = '' then
                    Error(Text002);

                if ProveedorCompensacion = '' then
                    Error(Text003);

                if BancoCompesacion = '' then
                    Error(Text003);

                if CuentaContableLiquidacion = '' then
                    Error(Text003);

                if ClienteLiquidacion = '' then
                    Error(Text003);

                if BancoLiquidacion = '' then
                    Error(Text003);

                ImportDocumentos.SetParameters(CodigoProveedorFactura, NumeroCuentaFactura, ProveedorCompensacion, BancoCompesacion, CuentaContableLiquidacion, ClienteLiquidacion, BancoLiquidacion, dim1, dim2, dim3);

                if TipoProcesadorPago = TipoProcesadorPago::Azul then begin
                    ImportDocumentos.ImportarAzul(TipoProcesadorPago::Azul);
                end;

                if TipoProcesadorPago = TipoProcesadorPago::Cardnet then begin
                    ImportDocumentos.ImportarCardnet(TipoProcesadorPago::Cardnet);
                end;

                if TipoProcesadorPago = TipoProcesadorPago::Amex then begin
                    ImportDocumentos.ImportarAmex(TipoProcesadorPago::Amex);
                end;
                if TipoProcesadorPago = TipoProcesadorPago::VisaNet then begin
                    ImportDocumentos.ImportVisaNet(TipoProcesadorPago::VisaNet);
                end;



            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                group(General)
                {
                    field(TipoProcesadorPago; TipoProcesadorPago)
                    {
                        Caption = 'Payment Processor Type';
                        ApplicationArea = basic, suite;
                    }
                }
                group("Factura Compra")
                {
                    field(CodigoProveedorFactura; CodigoProveedorFactura)
                    {
                        Caption = 'Code Supplier';
                        TableRelation = Vendor;
                        ApplicationArea = basic, suite;
                    }
                    field(NumeroCuentaFactura; NumeroCuentaFactura)
                    {
                        Caption = 'No. Account';
                        TableRelation = "G/L Account";
                        ApplicationArea = basic, suite;
                    }

                }
                group(Compensacion)
                {
                    Caption = 'Compensación';
                    field(ProveedorCompensacion; ProveedorCompensacion)
                    {
                        Caption = 'Proveedor compesación';
                        TableRelation = Vendor;
                        ApplicationArea = basic, suite;
                    }
                    field(BancoCompesacion; BancoCompesacion)
                    {
                        Caption = 'Caja compensación';
                        TableRelation = "Bank Account";
                        ApplicationArea = basic, suite;
                    }
                }
                group("Liquidacion Tarjeta")
                {
                    field(CuentaContableLiquidacion; CuentaContableLiquidacion)
                    {
                        Caption = 'Cuenta contable';
                        TableRelation = "G/L Account";
                        ApplicationArea = basic, suite;
                    }
                    field(ClienteLiquidacion; ClienteLiquidacion)
                    {
                        Caption = 'Caja Liquidacion';
                        TableRelation = "Bank Account";
                        ApplicationArea = basic, suite;
                    }
                    field(BancoLiquidacion; BancoLiquidacion)
                    {
                        Caption = 'Banco Liquidación';
                        TableRelation = "Bank Account";
                        ApplicationArea = basic, suite;
                    }
                }
                group(Dimensiones)
                {
                    Caption = 'Dimensiones';
                    field(dim1; dim1)
                    {
                        Caption = 'LINNEGOCIO';
                        TableRelation = "Dimension Value".Code WHERE("Dimension Code" = FILTER('LINNEGOCIO'));
                        ApplicationArea = basic, suite;
                    }
                    field(dim2; dim2)
                    {
                        Caption = 'DEPARTAMENTO';
                        TableRelation = "Dimension Value".Code WHERE("Dimension Code" = FILTER('DEPARTAMENTO'));
                        ApplicationArea = basic, suite;
                    }
                    field(dim3; dim3)
                    {
                        Caption = 'SUCURSAL';
                        TableRelation = "Dimension Value".Code WHERE("Dimension Code" = FILTER('SUCURSAL'));
                        ApplicationArea = basic, suite;
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

    var

        ImportDocumentos: Codeunit "Import Excel";
        NumeroCuentaFactura: Code[20];
        CodigoProveedorFactura: Code[20];
        TipoProcesadorPago: enum DSNTipoProcesadorDePago;
        CuentaContableLiquidacion: Code[20];
        ClienteLiquidacion: Code[20];
        BancoLiquidacion: Code[20];
        BancoCompesacion: Code[20];
        ProveedorCompensacion: Code[20];
        dim1: Code[20];
        dim2: Code[20];
        dim3: Code[20];
}

