table 55000 EstadosAfiliados
{
    Caption = 'EstadosAfiliados';
    DataClassification = ToBeClassified;

    fields
    {
        field(1; ID; Integer)
        {
            Caption = 'ID';
            DataClassification = CustomerContent;
        }
        field(2; "Fecha de Entrada"; Date)
        {
            Caption = 'Fecha de Entrada';
            DataClassification = CustomerContent;
        }
        field(3; "Terminal ID"; Integer)
        {
            Caption = 'Terminal ID';
            DataClassification = CustomerContent;
        }
        field(4; "No. de Lote"; Integer)
        {
            Caption = 'No. de Lote';
            DataClassification = CustomerContent;
        }
        field(5; "Monto Lote"; Decimal)
        {
            Caption = 'Monto Lote';
            DataClassification = CustomerContent;
        }
        field(6; Comision; Decimal)
        {
            Caption = 'Comision';
            DataClassification = CustomerContent;
        }
        field(7; "ITBIS Retenido"; Decimal)
        {
            Caption = 'ITBIS Retenido';
            DataClassification = CustomerContent;
        }
        field(8; "Monto a liquidar"; Decimal)
        {
            Caption = 'Monto a liquidar';
            DataClassification = CustomerContent;
        }
        field(9; "DIM SUC"; Text[30])
        {
            Caption = 'DIM SUC';
            DataClassification = CustomerContent;
        }
        field(10; Cuenta; Code[20])
        {
            Caption = 'Cuenta';
            DataClassification = CustomerContent;
        }
        field(11; NCF; Code[20])
        {
            Caption = 'NCF';
            DataClassification = CustomerContent;
        }
        field(12; Tipo; enum DSNTipoProcesadorDePago)
        {
            Caption = 'Tipo';
            DataClassification = CustomerContent;
        }
        field(13; "Deposito Bruto"; Decimal)
        {
            Caption = 'Deposito Bruto';
            DataClassification = CustomerContent;
        }
        field(14; "Retencion ITBIS"; Decimal)
        {
            Caption = 'Retencion ITBIS';
            DataClassification = CustomerContent;
        }
        field(15; Seccion; enum DSNTipoSeccionEstadoAfiliado)
        {
            Caption = 'Seccion';
            DataClassification = CustomerContent;

        }
        field(16; "Total pagado Visanet"; Decimal)
        {
            DataClassification = CustomerContent;
        }
    }
    keys
    {
        key(PK; ID)
        {
            Clustered = true;
        }
    }
}
