page 55000 DSNConfigCargosTarjetas
{
    ApplicationArea = All;
    Caption = 'Config. Cargos Tarjetas';
    AdditionalSearchTerms = 'Config. Cargos, Cof Car, Configuracion Cargos Tarjetas';
    PageType = List;
    SourceTable = "Config Cargos Tarjetas";
    Permissions =
        tabledata "Config Cargos Tarjetas" = RIMD;
    //UsageCategory = Lists;

    layout
    {
        area(content)
        {
            repeater(General)
            {
                field("Procesador de pagos"; Rec."Procesador de pagos")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Procesador de pagos field.';
                }
                field("Columna Comisión"; Rec."Columna Comisión")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna Comisión field.';
                }
                field("Columna Fecha"; Rec."Columna Fecha")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna Fecha field.';
                }
                field("Columna ITBIS"; Rec."Columna ITBIS")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna ITBIS field.';
                }
                field("Columna Monto Liquidación"; Rec."Columna Monto Liquidación")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna Monto Liquidación field.';
                }
                field("Columna Monto Lote"; Rec."Columna Monto Lote")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna Monto Lote field.';
                }
                field("Columna No Lote"; Rec."Columna No Lote")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna No Lote field.';
                }
                field("Columna Importe Descuento"; Rec."Columna Importe Descuento")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Columna No Lote field.';
                }

            }
        }
    }
}
