permissionset 55000 PermisosTarjeta
{
    Assignable = true;
    Permissions = tabledata "Config Cargos Tarjetas" = RIMD,
        tabledata EstadosAfiliados = RIMD,
        table "Config Cargos Tarjetas" = X,
        table EstadosAfiliados = X,
        report "Import Cardnet Azul Amex" = X,
        codeunit "Import Excel" = X,
        page DSNConfigCargosTarjetas = X,
        codeunit ColumnasExcel = X,
        tabledata "DatosCardNet" = RIMD,
        table "DatosCardNet" = X,
        report "Import Cardnet" = X,
        codeunit "ImportCardNet" = X;
}