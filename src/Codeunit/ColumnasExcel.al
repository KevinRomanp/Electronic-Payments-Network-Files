codeunit 55001 ColumnasExcel
{
    procedure GetComision(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        cm: Decimal;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cm, Buffer."Cell Value as Text");
            exit(cm);
        end;

    end;

    procedure GetDate(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Date
    var
        d: Date;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(D, Buffer."Cell Value as Text");
            exit(D);
        end;
    end;

    procedure GetNCF(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        cv: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cv, Buffer."Cell Value as Text");
            exit(cv);
        end;
    end;

    procedure GetColumnNumber(ColumnName: Text): Integer
    var
        columnIndex: Integer;
        factor: Integer;
        pos: Integer;
    begin
        factor := 1;
        for pos := strlen(ColumnName) downto 1 do
            if ColumnName[pos] >= 65 then begin
                columnIndex += factor * ((ColumnName[pos] - 65) + 1);
                factor *= 26;
            end;

        exit(columnIndex);
    end;
}
