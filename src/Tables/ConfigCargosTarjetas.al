table 55001 "Config Cargos Tarjetas"
{
    Caption = 'Config Cargos Tarjetas';
    DataClassification = CustomerContent;

    fields
    {
        field(1; "Procesador de pagos"; Enum DSNTipoProcesadorDePago)
        {
            Caption = 'Procesador de pagos';
        }
        field(2; "Columna Fecha"; Text[4])
        {
            Caption = 'Columna Fecha';
        }
        field(3; "Columna No Lote"; Text[4])
        {
            Caption = 'Columna No Lote';
        }
        field(4; "Columna Comisión"; Text[4])
        {
            Caption = 'Columna Comisión';
        }
        field(5; "Columna ITBIS"; Text[4])
        {
            Caption = 'Columna ITBIS';
        }
        field(6; "Columna Monto Lote"; Text[4])
        {
            Caption = 'Columna Monto Lote';
        }
        field(7; "Columna Monto Liquidación"; Text[4])
        {
            Caption = 'Columna Monto Liquidación';
        }
        field(8; "Columna Importe Descuento"; text[4])
        {
            Caption = 'Columna Importe Descuento';
        }

    }
    keys
    {
        /*key(PK; "Procesador de pagos")
        {
            Clustered = true;
        }*/
    }
}
