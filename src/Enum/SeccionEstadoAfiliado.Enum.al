enum 55000 DSNTipoSeccionEstadoAfiliado
{
    Extensible = true;

    value(0; "Distribucion de pago")
    {
        Caption = 'Distribucion de pago';
    }
    value(1; "Detalle de Transacciones")
    {
        Caption = ' Detalle de Transacciones';
    }
    value(2; "Comprobante fiscal por cargos")
    {
        Caption = ' Comprobante fiscal por cargos';
    }
    value(3; "Resumen por Lote")
    {
        Caption = ' Resumen por Lote';
    }
    value(4; VisaNetPagoComercio)
    {
        Caption = 'VisaNet Pago Comercio';
    }
    value(5; VisaNetCompraNormal)
    {
        Caption = 'VisaNet Compra Normal';
    }
    value(6; DetalleLote)
    { }
}
