report 55001 "Import Cardnet"
{
    ApplicationArea = Suite, Basic;
    Caption = 'Importar CardNet';
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


                if CodigoProveedorFactura = '' then
                    Error(Text002);

                if NumeroCuentaFactura = '' then
                    Error(Text002);

                /*if dim1 = '' then
                    Error(Text002);*/


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
                ImportDocumentos.Import();




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

        ImportDocumentos: Codeunit "ImportCardNet";
        NumeroCuentaFactura: Code[20];
        CodigoProveedorFactura: Code[20];
        CuentaContableLiquidacion: Code[20];
        ClienteLiquidacion: Code[20];
        BancoLiquidacion: Code[20];
        BancoCompesacion: Code[20];
        ProveedorCompensacion: Code[20];
        dim1: Code[20];
        dim2: Code[20];
        dim3: Code[20];
}

