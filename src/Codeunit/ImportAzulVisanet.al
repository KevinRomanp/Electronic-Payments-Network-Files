codeunit 55000 "Import Excel"
{
    Permissions =
        tabledata "Config. Empresas" = R,
        tabledata "Default Dimension" = R,
        tabledata "Dimension Value" = R,
        tabledata EstadosAfiliados = RID,
        tabledata "G/L Account" = R,
        tabledata "Gen. Journal Line" = RIM,
        tabledata "General Ledger Setup" = R,
        tabledata "Purchase Header" = RIM,
        tabledata "Purchase Line" = RIM;

    var
        GenJournalLine: Record "Gen. Journal Line";
        ExcelBuff: Record "Excel Buffer" temporary;
        ExcelBuffer: Record "Excel Buffer" temporary;
        ExcelBuffer1: Record "Excel Buffer" temporary;
        ExcelBuffer2: Record "Excel Buffer" temporary;
        ExcelBuffer3: Record "Excel Buffer" temporary;
        DimensionSetEntry: Record "Dimension Set Entry";
        recDimVal: Record "Dimension Value";
        TempDimEntry: Record "Dimension Set Entry" temporary;
        recDfltDim: Record "Default Dimension";
        ConfContab: record "General Ledger Setup";
        PurchaseLine: Record "Purchase Line";
        EstadosAfiliados2: Record EstadosAfiliados;
        GenJnlLine: Record "Gen. Journal Line";
        EstadosAfiliados: Record EstadosAfiliados;
        EstadosAfiliados3: Record EstadosAfiliados;
        PurchaseHeader: Record "Purchase Header";
        Contador: Integer;
        Cantidad: Integer;


        Ventana: Dialog;
        LineNo: Integer;
        JournalBatchName: Text;
        JournalTemplateName: Text;

        cduDim: Codeunit DimensionManagement;
        DimSetID: Integer;
        YYYY: text[4];
        MM: text[2];
        DD: text[2];
        iYYYY: Integer;
        iMM: Integer;
        iDD: Integer;
        FechaFac: Date;
        NCFVisaNet: Code[25];
        TotalColumns: Integer;
        TotalRows: Integer;

        LastNo: Integer;
        Filename: Text[250];
        NLote: Integer;
        Cuentas: Text;
        MLote: Decimal;
        Comision: Decimal;
        TotalPagadoVisanet: Decimal;
        ITBIS: Decimal;
        MLiquidacion: Decimal;
        "Tipo doc": Text;
        I: Integer;
        Iinit: Integer;
        TempString: Text;
        TempString2: Date;
        Day: Integer;
        Month: Integer;
        Year: Integer;
        FechaEnt: Date;
        Terminal: Integer;
        Suc: Text;
        Text002: Label 'Processing UOM @1@@@@@@@@@@@@@\Processing ICR  @2@@@@@@@@@@@@@\Processing SP  @3@@@@@@@@@@@@@\#4##########';
        SheetName: Text;
        GLAccount: Record "G/L Account";
        NombreCuenta: Text;
        FinalRow: Integer;
        NCF: Code[20];

        NumeroFactura: Code[20];
        Seccion: enum DSNTipoSeccionEstadoAfiliado;
        ProveedorFacturaCompra: Code[20];
        NumeroCuentaFacturaCompra: Code[20];
        ProveedorCompensacion: Code[20];
        BancoCompesacion: Code[20];
        CuentaContableLiquidacion: Code[20];
        CajaLiquidacion: Code[20];
        BancoLiquidacion: Code[20];
        LinNegocio: Code[20];
        Departamento: Code[20];
        Sucursal: code[20];
        TDepositoBruto: Decimal;
        TRetencionITBIS: Decimal;
        TdescuentoComision: Decimal;
        SigFechaFac: Date;
        FechaRegistro: Date;
        InS: InStream;



    procedure InsertDataExpress(RowNo: Integer)
    begin
        //EstadosAfiliados.DELETEALL;

        EstadosAfiliados3.RESET;
        Clear(LastNo);
        Clear(NLote);
        Clear(Cuentas);
        //CLEAR(FechaEnt);
        //CLEAR(Terminal);
        Clear(MLote);
        Clear(Comision);
        Clear(ITBIS);
        Clear(MLiquidacion);
        Clear("Tipo doc");
        Clear(Contador);
        Clear(Cantidad);
        if EstadosAfiliados3.FINDLAST then
            LastNo := EstadosAfiliados3.Id;

        ExcelBuffer.SetFilter("Row No.", '=%1', RowNo);
        Cantidad := ExcelBuff.Count;
        if ExcelBuffer.FindSet then begin
            repeat
                Contador := Contador + 1;
                Ventana.Update(2, 'Insertando en Intermedia' + Format(Round(Contador / Cantidad * 10000, 1)));
                if ExcelBuffer1.Get(10, 10) then begin
                    if CopyStr(ExcelBuffer1."Cell Value as Text", 1, 4) = 'VALE' then
                        Cuentas := 'BDI RD$'
                    else
                        Cuentas := 'BPD';
                end;
                if (Cuentas = '') then begin
                    if ExcelBuffer1.Get(9, 10) then begin
                        if CopyStr(ExcelBuffer1."Cell Value as Text", 1, 4) = 'VALE' then
                            Cuentas := 'BDI RD$'
                        else
                            Cuentas := 'BPD';
                    end;

                    if ExcelBuffer1.Get(9, 1) then begin
                        if CopyStr(ExcelBuffer1."Cell Value as Text", 1, 4) = 'VALE' then
                            Cuentas := 'BDI RD$'
                        else
                            Cuentas := 'BPD';
                    end;
                end;

                if ExcelBuffer1.Get(9, 23) then begin
                    if StrLen(ExcelBuffer1."Cell Value as Text") > 0 then
                        Cuentas := ExcelBuffer1."Cell Value as Text";
                end;

                if ExcelBuffer."Column No." = ReturnColumNo('Fecha de Entrada') then begin
                    if (CopyStr(ExcelBuffer."Cell Value as Text", 1, 15) = 'Total liquidado') then
                        exit;
                end;
                if ExcelBuffer."Column No." = ReturnColumNo('Fecha de Entrada') then begin
                    Year := 0;
                    //EVALUATE(FechaEnt,ExcelBuff."Cell Value as Text");
                    if ExcelBuffer."Cell Value as Text" in ['Liquidación: RD$', 'Fecha de Entrada'] then begin
                        break;
                    end;
                    if ReturnColumNo('Fecha de Entrada') = 1 then begin
                        TempString := ConvertStr(ExcelBuffer."Cell Value as Text", '/', ',');
                        Evaluate(Day, SelectStr(1, TempString));
                        Evaluate(Month, SelectStr(2, TempString));
                        Evaluate(Year, SelectStr(3, TempString));
                        Year := Year + 2000;
                        FechaEnt := DMY2Date(Month, Day, Year);
                    end else begin
                        TempString := ConvertStr(ExcelBuffer."Cell Value as Text", '/', ',');
                        Evaluate(Day, SelectStr(1, TempString));
                        Evaluate(Month, SelectStr(2, TempString));
                        Evaluate(Year, SelectStr(3, TempString));
                        FechaEnt := DMY2Date(Day, Month, Year);
                    end;
                end;
                Evaluate(FechaEnt, TempString);
                if (ExcelBuffer."Column No." = ReturnColumNo('Terminal ID')) and (StrLen(ExcelBuffer."Cell Value as Text") > 0) and (StrLen(ExcelBuffer."Cell Value as Text") < 10) then begin

                    Evaluate(Terminal, ExcelBuffer."Cell Value as Text");
                end;

                case ExcelBuffer."Column No." of
                    ReturnColumNo('No. de Lote'):
                        begin
                            if (ExcelBuffer."Column No." = ReturnColumNo('No. de Lote')) and (ExcelBuffer."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(NLote, ExcelBuffer."Cell Value as Text");
                        end;
                    ReturnColumNo('Monto Lote'):
                        begin
                            if (ExcelBuffer."Column No." = ReturnColumNo('Monto Lote')) and (ExcelBuffer."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(MLote, ExcelBuffer."Cell Value as Text");

                        end;
                    ReturnColumNo('Comisión'):
                        begin
                            if (ExcelBuffer."Column No." = ReturnColumNo('Comisión')) and (ExcelBuffer."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(Comision, ExcelBuffer."Cell Value as Text");

                        end;
                    ReturnColumNo('ITBIS Retenido'):
                        begin
                            if (ExcelBuffer."Column No." = ReturnColumNo('ITBIS Retenido')) and (ExcelBuffer."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(ITBIS, ExcelBuffer."Cell Value as Text");

                        end;
                    ReturnColumNo('Monto a Liquidar'):
                        begin
                            if (ExcelBuffer."Column No." = ReturnColumNo('Monto a Liquidar')) and (ExcelBuffer."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(MLiquidacion, ExcelBuffer."Cell Value as Text");

                        end;
                end;

                if ExcelBuffer1.Get(9, 24) then
                    Suc := ExcelBuffer1."Cell Value as Text";

                if ExcelBuffer."Column No." = ReturnColumNo('Monto a Liquidar') then begin
                    EstadosAfiliados.INIT;
                    LastNo := LastNo + 1;
                    EstadosAfiliados.Id := LastNo;
                    EstadosAfiliados."Fecha de Entrada" := FechaEnt;
                    EstadosAfiliados."Terminal ID" := Terminal;
                    EstadosAfiliados."No. de Lote" := NLote;
                    EstadosAfiliados."Monto Lote" := MLote;
                    EstadosAfiliados.Comision := Comision;
                    EstadosAfiliados."ITBIS Retenido" := ITBIS;
                    EstadosAfiliados."Monto a Liquidar" := MLiquidacion;
                    EstadosAfiliados.Cuenta := Cuentas;
                    EstadosAfiliados."DIM SUC" := Suc;
                    if NLote = 0 then begin
                        break;
                    end;
                    EstadosAfiliados.INSERT;
                end;
            until ExcelBuffer.Next = 0;
        end;

        //EVALUATE(ExcelBuffer.
    end;

    procedure ImportarExpress()
    begin
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();



            Ventana.Open(Text002);
            ExcelBuffer1.LockTable;
            ExcelBuffer1.OpenBookStream(InS, SheetName);
            ExcelBuffer1.ReadSheet;
            GetLastRowandColumn;

            //ExcelBuffer1.COPY(ExcelBuffer);
            EstadosAfiliados.DELETEALL;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Fecha de Entrada');
            if ExcelBuffer2.FindFirst then begin
                Iinit := (ExcelBuffer2."Row No." + 1);
            end else
                Iinit := 34;

            for I := Iinit to (TotalRows - 2) do begin
                InsertDataExpress(I);
                //I := (I + 1);
            end;
            // InsertGL('DEPOSITO TARJETA AMEX');
            Ventana.Close;
            ExcelBuffer.DeleteAll;
            EstadosAfiliados.DELETEALL;
            Message('Import is completed');
        end;
    end;


    procedure ImportarAzul(TipoProcesadorPago: enum DSNTipoProcesadorDePago)
    begin
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();
            TotalRows := 0;
            I := 0;
            Clear(FinalRow);
            //Ventana.OPEN(Text002);
            ExcelBuffer.DeleteAll;
            Commit;
            ExcelBuffer.LockTable;
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet;
            ExcelBuff.DeleteAll;
            ExcelBuffer1.DeleteAll;
            ExcelBuffer2.DeleteAll;
            ExcelBuffer3.DeleteAll;
            ExcelBuffer.Find('-');
            repeat
                ExcelBuff.TransferFields(ExcelBuffer);
                ExcelBuffer1.TransferFields(ExcelBuffer);
                ExcelBuffer2.TransferFields(ExcelBuffer);
                ExcelBuffer3.TransferFields(ExcelBuffer);
                ExcelBuff.Insert();
                ExcelBuffer1.Insert();
                ExcelBuffer2.Insert();
                ExcelBuffer3.Insert();
            until ExcelBuffer.Next() = 0;
            GetLastRowandColumn;

            EstadosAfiliados.DELETEALL;
            Commit();

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Comprobante fiscal por cargos');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 3);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Resumen Mensual de Facturación y Certificación de Retención de ITBIS');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 4);

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Comprobante fiscal por cargos", I);
            end;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Resumen por Lote');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 2);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Total');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 2);

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Resumen por Lote", I);
            end;


            CrearFactura(TipoProcesadorPago);
            InsertarDiario(TipoProcesadorPago);
            Message('Import is completed');
        end;
    end;


    procedure InsertDataAzul(RowNo: Integer)
    begin


        EstadosAfiliados2.RESET;
        Clear(LastNo);
        Clear(NLote);
        //CLEAR(FechaEnt);
        //CLEAR(Terminal);
        Clear(MLote);
        Clear(Comision);
        Clear(ITBIS);
        Clear(MLiquidacion);
        Clear("Tipo doc");
        if EstadosAfiliados2.FINDLAST then;
        LastNo := EstadosAfiliados2.Id;
        Cantidad := ExcelBuffer.Count;
        ExcelBuff.SetFilter("Row No.", '=%1', RowNo);
        if ExcelBuff.FindSet then begin
            repeat

                if ExcelBuff."Column No." = 9 then begin
                    if ExcelBuff."Cell Value as Text" = 'Total' then begin
                        I := TotalRows;
                        exit;
                    end;
                    if CopyStr(ExcelBuff."Cell Value as Text", 1, 7) = 'Resumen' then begin
                        break;
                    end;
                    if CopyStr(ExcelBuff."Cell Value as Text", 1, 5) = 'Fecha' then begin
                        break;
                    end;

                    Cuentas := 'BPD';
                    if ExcelBuffer1.Get(14, 71) then begin
                        if CopyStr(ExcelBuffer1."Cell Value as Text", 1, 4) = 'VALE' then
                            Cuentas := 'BDI RD$'
                        else
                            Cuentas := 'BPD';
                    end;

                    if ExcelBuffer1.Get(18, 69) then begin
                        if StrLen(ExcelBuffer1."Cell Value as Text") > 0 then
                            Cuentas := ExcelBuffer1."Cell Value as Text";
                    end;

                    Contador := Contador + 1;


                    if ReturnColumNo('Fecha de cierre') = 9 then begin

                        TempString := ConvertStr(ExcelBuff."Cell Value as Text", '/', ',');
                        Evaluate(Day, SelectStr(1, TempString));
                        Evaluate(Month, SelectStr(2, TempString));
                        Evaluate(Year, SelectStr(3, TempString));
                        Year := Year + 2000;
                        FechaEnt := DMY2Date(Day, Month, Year);
                    end else begin
                        TempString := ConvertStr(ExcelBuff."Cell Value as Text", '/', ',');
                        Evaluate(Day, SelectStr(1, TempString));
                        Evaluate(Month, SelectStr(2, TempString));
                        Evaluate(Year, SelectStr(3, TempString));
                        FechaEnt := DMY2Date(Day, Month, Year);
                    end;
                end;

                if (ExcelBuff."Column No." = 20) and (ExcelBuff."Cell Value as Text" <> '') then
                    Evaluate(Terminal, ExcelBuff."Cell Value as Text");

                case ExcelBuff."Column No." of
                    33:
                        begin
                            if (ExcelBuff."Column No." = 33) and (ExcelBuff."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(NLote, ExcelBuff."Cell Value as Text");
                        end;

                    67:
                        begin
                            if (ExcelBuff."Column No." = 67) and (ExcelBuff."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(Comision, ExcelBuff."Cell Value as Text");

                        end;
                    74:
                        begin
                            if (ExcelBuff."Column No." = 74) and (ExcelBuff."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(ITBIS, ExcelBuff."Cell Value as Text");

                        end;
                    86:
                        begin  //76
                            if (ExcelBuff."Column No." = 86) and (ExcelBuff."Cell Value as Text" = '') then begin
                                break;
                            end;
                            Evaluate(MLiquidacion, ExcelBuff."Cell Value as Text");
                            Evaluate(MLote, ExcelBuff."Cell Value as Text");

                        end;
                end;
                if ExcelBuffer1.Get(18, 76) then
                    Suc := ExcelBuffer1."Cell Value as Text";

                if ExcelBuff."Column No." = 86 then begin
                    EstadosAfiliados.INIT;
                    LastNo := LastNo + 1;
                    EstadosAfiliados.Id := LastNo;
                    EstadosAfiliados."Fecha de Entrada" := FechaEnt;
                    EstadosAfiliados."Terminal ID" := Terminal;
                    EstadosAfiliados."No. de Lote" := NLote;
                    EstadosAfiliados."Monto Lote" := MLote;
                    EstadosAfiliados.Comision := Comision;
                    EstadosAfiliados."ITBIS Retenido" := ITBIS;
                    EstadosAfiliados."Monto a Liquidar" := MLiquidacion;
                    EstadosAfiliados.Cuenta := Cuentas;
                    EstadosAfiliados."DIM SUC" := Suc;
                    if NLote = 0 then begin
                        break;
                    end;
                    EstadosAfiliados.INSERT;
                end;
            until ExcelBuff.Next = 0;
        end;



    end;

    procedure ImportarCardnet(TipoProcesadorPago: enum DSNTipoProcesadorDePago)
    begin
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();
            TotalRows := 0;
            I := 0;
            //Ventana.OPEN(Text002);
            ExcelBuffer.DeleteAll;
            Commit;
            ExcelBuffer.LockTable;
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet;

            ExcelBuff.DeleteAll;
            ExcelBuffer1.DeleteAll;
            ExcelBuffer2.DeleteAll;
            ExcelBuffer3.DeleteAll;

            ExcelBuffer.Find('-');
            repeat
                ExcelBuff.TransferFields(ExcelBuffer);
                ExcelBuffer1.TransferFields(ExcelBuffer);
                ExcelBuffer2.TransferFields(ExcelBuffer);
                ExcelBuffer3.TransferFields(ExcelBuffer);
                ExcelBuff.Insert();
                ExcelBuffer1.Insert();
                ExcelBuffer2.Insert();
                ExcelBuffer3.Insert();
            until ExcelBuffer.Next() = 0;

            GetLastRowandColumn;

            EstadosAfiliados.DELETEALL;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DISTRIBUCION DE PAGO');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 1);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DETALLE DE TRANSACCIONES');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 2);

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Distribucion de pago", I);
            end;



            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DETALLE DE TRANSACCIONES');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 2);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DETALLE DESCUENTOS');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 1);

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Detalle de Transacciones", I);
            end;

            //Ventana.CLOSE;
            // InsertGL('DEPOSITO TARJETA AZUL ');



            CrearFactura(TipoProcesadorPago);
            InsertarDiario(TipoProcesadorPago);

            Message('Import is completed');
        end;
    end;

    local procedure InsertData(TipoProcesadorPago: enum DSNTipoProcesadorDePago; Seccion: enum DSNTipoSeccionEstadoAfiliado; Rowno: Integer)
    var
        DepositoBruto: Decimal;
        RetencionITBIS: Decimal;
        MontoTotal: Decimal;
        descuentoComision: Decimal;

        CompraVisanet: Boolean;

    begin
        EstadosAfiliados2.RESET;
        Clear(YYYY);
        Clear(MM);
        Clear(DD);
        Clear(iYYYY);
        Clear(iMM);
        Clear(iDD);
        Clear(LastNo);
        Clear(NLote);
        Clear(FechaEnt);
        Clear(NCF);
        Clear(Terminal);
        Clear(MLote);
        Clear(Comision);
        Clear(ITBIS);
        Clear(MLiquidacion);
        Clear("Tipo doc");
        Clear(TotalPagadoVisanet);
        CompraVisanet := false;
        if EstadosAfiliados2.FINDLAST then;
        LastNo := EstadosAfiliados2.Id;
        Cantidad := ExcelBuffer.Count;
        ExcelBuff.Reset();
        ExcelBuff.SetFilter("Row No.", '=%1', Rowno);

        // CARDNET
        if TipoProcesadorPago = TipoProcesadorPago::Cardnet then begin
            if ExcelBuff.FindSet() then begin
                repeat

                    if Seccion = Seccion::"Distribucion de pago" then begin
                        case ExcelBuff."Column No." of
                            1:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'FECHA') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    //                    TempString := CONVERTSTR(ExcelBuff."Cell Value as Text",'-',',');
                                    //                    EVALUATE(Day,SELECTSTR(1,TempString));
                                    //                    EVALUATE(Month,SELECTSTR(2,TempString));
                                    //                    EVALUATE(Year,SELECTSTR(3,TempString));
                                    //                    Year  := Year +2000;
                                    //                    FechaEnt  :=DMY2DATE(Day, Month, Year);
                                    Evaluate(FechaEnt, ExcelBuff."Cell Value as Text");
                                end;
                            2:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'NCF') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    NCF := ExcelBuff."Cell Value as Text";
                                end;
                            5:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'COMISION DESCONTADA') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(Comision, ExcelBuff."Cell Value as Text");
                                end;
                            15:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'MONTO') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(MontoTotal, ExcelBuff."Cell Value as Text");
                                end;
                        end;

                        if ExcelBuff."Column No." = 15 then begin
                            EstadosAfiliados.INIT;
                            LastNo := LastNo + 1;
                            EstadosAfiliados.Id := LastNo;
                            EstadosAfiliados."Fecha de Entrada" := FechaEnt;
                            //        EstadosAfiliados."Terminal ID"  :=  Terminal;
                            //        EstadosAfiliados."No. de Lote"  := NLote;
                            //        EstadosAfiliados."Monto Lote":=MLote;
                            EstadosAfiliados.Comision := Comision;
                            EstadosAfiliados.NCF := NCF;
                            EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::Cardnet;
                            EstadosAfiliados.Seccion := Seccion::"Distribucion de pago";
                            EstadosAfiliados."Monto a Liquidar" := MontoTotal;
                            //        EstadosAfiliados."ITBIS Retenido" := ITBIS;
                            //        EstadosAfiliados."Monto a Liquidar" := MLiquidacion;
                            //        EstadosAfiliados.Cuenta := Cuentas;
                            //        EstadosAfiliados."DIM SUC"  := Suc;
                            //        IF NLote  = 0 THEN BEGIN
                            //          BREAK;
                            //        END;
                            if MontoTotal > 0 then
                                EstadosAfiliados.INSERT;

                            exit;
                        end;
                    end; // end Seccion

                    if Seccion = Seccion::"Detalle de Transacciones" then begin

                        if ExcelBuff."Cell Value as Text" = '*TOTAL POR DIA' then begin
                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No." - 1), '1');
                            Evaluate(FechaFac, ExcelBuffer1."Cell Value as Text");

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '6');
                            Evaluate(DepositoBruto, ExcelBuffer1."Cell Value as Text");



                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '8');
                            Evaluate(descuentoComision, ExcelBuffer1."Cell Value as Text");

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '11');
                            Evaluate(RetencionITBIS, ExcelBuffer1."Cell Value as Text");



                            EstadosAfiliados.INIT;
                            LastNo := LastNo + 1;
                            EstadosAfiliados.Id := LastNo;
                            EstadosAfiliados."Deposito Bruto" := DepositoBruto;
                            EstadosAfiliados."Fecha de Entrada" := FechaFac;
                            EstadosAfiliados."Retencion ITBIS" := RetencionITBIS;
                            EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::Cardnet;
                            EstadosAfiliados.Seccion := Seccion::"Detalle de Transacciones";
                            EstadosAfiliados.Comision := descuentoComision;

                            EstadosAfiliados.INSERT(true);
                            exit;
                        end;

                    end; // end seccion

                until ExcelBuff.Next = 0;
            end;
        end;

        // AZUL

        if TipoProcesadorPago = TipoProcesadorPago::Azul then begin
            if ExcelBuff.FindSet then begin
                repeat
                    if Seccion = Seccion::"Comprobante fiscal por cargos" then begin

                        case ExcelBuff."Column No." of
                            10:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'Fecha') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;


                                    //                          TempString := CONVERTSTR(ExcelBuff."Cell Value as Text",'/',',');
                                    //                          EVALUATE(Day,SELECTSTR(1,TempString));
                                    //                          EVALUATE(Month,SELECTSTR(2,TempString));
                                    //                          EVALUATE(Year,SELECTSTR(3,TempString));
                                    //                          Year  := Year +2000;
                                    //                          FechaEnt  :=DMY2DATE(Day, Month, Year);
                                    Evaluate(FechaEnt, ExcelBuff."Cell Value as Text");
                                end;
                            15:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'NCF') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;

                                    NCF := ExcelBuff."Cell Value as Text";

                                end;
                            78:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'Monto (RD$)') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(Comision, ExcelBuff."Cell Value as Text");


                                end;
                        end;





                        if ExcelBuff."Column No." = 78 then begin
                            EstadosAfiliados.INIT;
                            LastNo := LastNo + 1;
                            EstadosAfiliados.Id := LastNo;
                            EstadosAfiliados."Fecha de Entrada" := FechaEnt;
                            EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::Azul;
                            EstadosAfiliados.Seccion := Seccion::"Comprobante fiscal por cargos";

                            EstadosAfiliados.Comision := Comision;
                            EstadosAfiliados.NCF := NCF;

                            if Comision > 0 then
                                EstadosAfiliados.INSERT;

                            exit;
                        end;


                    end; // fin de la seccion


                    if Seccion = Seccion::"Resumen por Lote" then begin

                        if ExcelBuff."Cell Value as Text" = 'Total Por Dia' then begin

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No." - 1), '9');
                            Evaluate(FechaFac, ExcelBuffer1."Cell Value as Text");

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '50');
                            Evaluate(DepositoBruto, ExcelBuffer1."Cell Value as Text");

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '67');
                            Evaluate(descuentoComision, ExcelBuffer1."Cell Value as Text");

                            Clear(ExcelBuffer1);
                            ExcelBuffer1.Reset;
                            ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '74');
                            Evaluate(RetencionITBIS, ExcelBuffer1."Cell Value as Text");



                            EstadosAfiliados.INIT;
                            LastNo := LastNo + 1;
                            EstadosAfiliados.Id := LastNo;
                            EstadosAfiliados."Fecha de Entrada" := FechaFac;
                            EstadosAfiliados."Deposito Bruto" := DepositoBruto;
                            EstadosAfiliados."Retencion ITBIS" := RetencionITBIS;
                            EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::Azul;
                            EstadosAfiliados.Seccion := Seccion::"Resumen por Lote";
                            EstadosAfiliados.Comision := descuentoComision;

                            //if RetencionITBIS > 0 then
                            EstadosAfiliados.INSERT(true);

                            exit;


                        end;



                    end; // fin de la seccion



                until ExcelBuff.Next = 0;
            end;
        end; //fin azul

        if TipoProcesadorPago = TipoProcesadorPago::Amex then begin

            if ExcelBuff.FindSet() then begin
                repeat
                    // IF ExcelBuff.FINDSET THEN BEGIN
                    //    REPEAT
                    //IF ExcelBuff."Cell Value as Text" = 'Fecha de Entrada' THEN BEGIN
                    if ExcelBuff."Row No." > 2 then begin
                        case
                            ExcelBuff."Column No." of
                            1:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'Fecha de Entrada') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(FechaFac, ExcelBuff."Cell Value as Text");
                                end;
                            5:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'Monto Lote') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(DepositoBruto, ExcelBuff."Cell Value as Text");
                                end;
                            6:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'Comisión') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(descuentoComision, ExcelBuff."Cell Value as Text");
                                end;
                            7:
                                begin
                                    if (ExcelBuff."Cell Value as Text" = 'ITBIS') or (ExcelBuff."Cell Value as Text" = '') then
                                        exit;
                                    Evaluate(RetencionITBIS, ExcelBuff."Cell Value as Text");
                                end;
                        end;


                        Clear(ExcelBuffer1);
                        ExcelBuffer1.Reset;
                        ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '1');
                        Evaluate(FechaFac, ExcelBuffer1."Cell Value as Text");

                        Clear(ExcelBuffer1);
                        ExcelBuffer1.Reset;
                        ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '5');
                        Evaluate(DepositoBruto, ExcelBuffer1."Cell Value as Text");
                        //TDepositoBruto += DepositoBruto;

                        Clear(ExcelBuffer1);
                        ExcelBuffer1.Reset;
                        ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '6');
                        Evaluate(descuentoComision, ExcelBuffer1."Cell Value as Text");
                        //TdescuentoComision += descuentoComision;

                        Clear(ExcelBuffer1);
                        ExcelBuffer1.Reset;
                        ExcelBuffer1.Get(Format(ExcelBuff."Row No."), '7');
                        Evaluate(RetencionITBIS, ExcelBuffer1."Cell Value as Text");
                        //TRetencionITBIS += RetencionITBIS;



                        if FechaFac <> 0D then
                            FechaRegistro := FechaFac;


                        if (FechaFac <> 0D) or (ExcelBuffer1."Cell Value as Text" = 'Total liquidado: American Express') then begin

                            EstadosAfiliados.INIT;
                            LastNo := LastNo + 1;
                            EstadosAfiliados.Id := LastNo;
                            EstadosAfiliados."Fecha de Entrada" := FechaFac;
                            EstadosAfiliados."Deposito Bruto" := DepositoBruto;
                            EstadosAfiliados."Retencion ITBIS" := RetencionITBIS;
                            EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::Amex;
                            //EstadosAfiliados.Seccion:=Seccion::"Resumen por Lote";
                            EstadosAfiliados.Comision := descuentoComision;



                            //if TRetencionITBIS > 0 then
                            EstadosAfiliados.INSERT(true);

                            /*Clear(FechaRegistro);
                            Clear(TDepositoBruto);
                            Clear(TdescuentoComision);
                            Clear(TRetencionITBIS);*/
                            exit;
                        end;

                    end;


                UNTIL ExcelBuff.NEXT = 0;
            end;

        end;

        if TipoProcesadorPago = TipoProcesadorPago::VisaNet then begin
            if ExcelBuff.FindSet() then begin
                repeat
                    case ExcelBuff."Column No." of
                        5:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Tipo Trans.') or (ExcelBuff."Cell Value as Text" = '') then
                                    exit;
                                if (ExcelBuff."Cell Value as Text" = 'S') then begin
                                    CompraVisanet := true;
                                    exit;
                                end else
                                    CompraVisanet := false;
                                //Evaluate(Comision, ExcelBuff."Cell Value as Text");
                            end;
                        9:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Monto Bruto') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                Evaluate(DepositoBruto, ExcelBuff."Cell Value as Text");
                            end;
                        10:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Descuento') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                Evaluate(Comision, ExcelBuff."Cell Value as Text");
                            end;
                        11:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Impuestos') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                Evaluate(RetencionITBIS, ExcelBuff."Cell Value as Text");
                            end;
                        12:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Total Pagado') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                Evaluate(TotalPagadoVisanet, ExcelBuff."Cell Value as Text");
                            end;
                        13:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'NCF') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                NCF := ExcelBuff."Cell Value as Text";

                            end;
                        14:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Fecha de Pago') or (ExcelBuff."Cell Value as Text" = '') or (CompraVisanet = true) then
                                    exit;
                                YYYY := CopyStr(ExcelBuff."Cell Value as Text", 1, 4);
                                MM := CopyStr(ExcelBuff."Cell Value as Text", 6, 2);
                                DD := CopyStr(ExcelBuff."Cell Value as Text", 9, 2);
                                Evaluate(iYYYY, YYYY);
                                Evaluate(iDD, DD);
                                Evaluate(iMM, MM);
                                FechaFac := DMY2Date(iDD, iMM, iYYYY);
                                //Evaluate(FechaFac, ExcelBuff."Cell Value as Text");
                            end;
                    end;

                    if (ExcelBuff."Column No." = 14) then begin
                        EstadosAfiliados.INIT;
                        LastNo := LastNo + 1;
                        EstadosAfiliados.Id := LastNo;
                        EstadosAfiliados."Fecha de Entrada" := FechaFac;
                        EstadosAfiliados.Tipo := EstadosAfiliados.Tipo::VisaNet;
                        EstadosAfiliados."Deposito Bruto" := DepositoBruto;
                        EstadosAfiliados."Total pagado Visanet" := TotalPagadoVisanet;
                        EstadosAfiliados.Seccion := Seccion::"Distribucion de pago"; //TODO : arreglar
                        EstadosAfiliados.Comision := Comision;
                        EstadosAfiliados.NCF := NCF;

                        if Comision > 0 then
                            EstadosAfiliados.INSERT;
                        exit;
                    end;
                until ExcelBuff.Next() = 0;
            end;
        end;
    end;

    local procedure InsertarDiario(TipoProcesadorPago: Enum DSNTipoProcesadorDePago)
    var
        ConfigEmpresa: Record "Config. Empresas";
    begin
        Clear(Contador);
        Clear(Cantidad);
        //View_EstadosAfiliados.RESET;
        GenJnlLine.Reset;
        ConfigEmpresa.Reset;
        ConfigEmpresa.Get;
        EstadosAfiliados.RESET;
        ConfContab.Get();
        // fermo
        //compensacion
        if EstadosAfiliados.FINDSET then begin
            repeat
                EstadosAfiliados.Reset();
                if (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Distribucion de pago") or
                (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Comprobante fiscal por cargos") or
                (TipoProcesadorPago = TipoProcesadorPago::Amex) or
                (TipoProcesadorPago = TipoProcesadorPago::VisaNet) then begin

                    Clear(NumeroFactura);
                    Clear(PurchaseHeader);
                    if (TipoProcesadorPago <> TipoProcesadorPago::Amex) then begin

                        if TipoProcesadorPago = TipoProcesadorPago::VisaNet then begin
                            PurchaseHeader.Reset;
                            PurchaseHeader.SetFilter("Buy-from Vendor No.", ProveedorCompensacion);
                            PurchaseHeader.SetFilter("Vendor Invoice No.", NCFVisaNet);
                            PurchaseHeader.FindFirst;
                            NumeroFactura := PurchaseHeader."No.";
                        end
                        else begin
                            PurchaseHeader.Reset;
                            PurchaseHeader.SetFilter("Buy-from Vendor No.", ProveedorCompensacion);
                            PurchaseHeader.SetFilter("Vendor Invoice No.", EstadosAfiliados.NCF);
                            PurchaseHeader.FindFirst;
                            NumeroFactura := PurchaseHeader."No.";
                        end;
                    end;

                    //primera linea
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch Cobro");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template Cobro");

                    if GenJnlLine.FindLast then;

                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch Cobro";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template Cobro";

                    if GenJournalLine.Insert(true) then begin
                        GenJournalLine."Posting Date" := EstadosAfiliados."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::Vendor); // customer
                        if TipoProcesadorPago = TipoProcesadorPago::Amex then
                            GenJournalLine."Document No." := 'AMEX-' + Format(EstadosAfiliados."Fecha de Entrada")
                        else
                            GenJournalLine."Document No." := NumeroFactura;


                        GenJournalLine.Validate("Account No.", ProveedorCompensacion); //cliente cambiar
                        GenJournalLine.Validate(Amount, Round(EstadosAfiliados.Comision));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;

                    // Termina primera linea

                    // segunda linea proveedor
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch Cobro");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template Cobro");
                    if GenJnlLine.FindLast then;
                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch Cobro";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template Cobro";

                    if GenJournalLine.Insert(true) then begin
                        GenJournalLine."Posting Date" := EstadosAfiliados."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");
                        if TipoProcesadorPago = TipoProcesadorPago::Amex then
                            GenJournalLine."Document No." := 'AMEX-' + Format(EstadosAfiliados."Fecha de Entrada")
                        else
                            GenJournalLine."Document No." := NumeroFactura;

                        GenJournalLine.Validate("Account No.", BancoCompesacion);
                        GenJournalLine.Validate(Amount, Round(EstadosAfiliados.Comision * -1));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;
                    //    Termino de segunda linea

                end; // end distribucion de pago

                if (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Detalle de Transacciones") or (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Comprobante fiscal por cargos") or (TipoProcesadorPago = TipoProcesadorPago::Amex) then begin

                    if (TipoProcesadorPago <> TipoProcesadorPago::Amex) or (TipoProcesadorPago <> TipoProcesadorPago::VisaNet) then begin
                        EstadosAfiliados2.RESET;
                        EstadosAfiliados2.SETRANGE("Fecha de Entrada", EstadosAfiliados."Fecha de Entrada");

                        if (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Detalle de Transacciones") then
                            EstadosAfiliados2.SETRANGE(Seccion, EstadosAfiliados2.Seccion::"Distribucion de pago");
                        if (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Resumen por Lote") then
                            EstadosAfiliados2.SETRANGE(Seccion, EstadosAfiliados2.Seccion::"Comprobante fiscal por cargos");



                        if EstadosAfiliados2.FINDFIRST then;

                        Clear(PurchaseHeader);


                    end;
                end
            until EstadosAfiliados.NEXT = 0;
        end;


        //primera linea liquidacion
        EstadosAfiliados.Reset();
        if EstadosAfiliados.FINDSET then
            repeat
                EstadosAfiliados.Reset();
                if (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Resumen por Lote") or
                (EstadosAfiliados.Seccion = EstadosAfiliados.Seccion::"Detalle de Transacciones") or
                (TipoProcesadorPago = TipoProcesadorPago::Amex) or (TipoProcesadorPago = TipoProcesadorPago::VisaNet) then begin
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch liquidacion");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template liquidacion");


                    if EstadosAfiliados."Retencion ITBIS" <> 0 then
                        if GenJnlLine.FindLast then;

                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch liquidacion";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template liquidacion";
                    if GenJournalLine.Insert(true) then begin



                        //azul
                        GenJournalLine."Posting Date" := EstadosAfiliados."Fecha de Entrada";



                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"G/L Account");
                        if TipoProcesadorPago = TipoProcesadorPago::Amex then
                            GenJournalLine."Document No." := 'AMEX-' + Format(EstadosAfiliados."Fecha de Entrada")
                        else
                            GenJournalLine."Document No." := NumeroFactura;

                        GenJournalLine.Validate("Account No.", CuentaContableLiquidacion);
                        GenJournalLine.Validate(Amount, Round(EstadosAfiliados."Retencion ITBIS"));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;

                    // Termina primera linea

                    // segunda linea cliente
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch liquidacion");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template liquidacion");

                    if GenJnlLine.FindLast then;

                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch liquidacion";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template liquidacion";
                    if GenJournalLine.Insert(true) then begin



                        //azul
                        GenJournalLine."Posting Date" := EstadosAfiliados."Fecha de Entrada";


                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");
                        if TipoProcesadorPago = TipoProcesadorPago::Amex then
                            GenJournalLine."Document No." := 'AMEX-' + Format(EstadosAfiliados."Fecha de Entrada")
                        else
                            GenJournalLine."Document No." := NumeroFactura;
                        GenJournalLine.Validate("Account No.", CajaLiquidacion);
                        if TipoProcesadorPago = TipoProcesadorPago::VisaNet then
                            GenJournalLine.Validate(Amount, Round((EstadosAfiliados."Total pagado Visanet") * -1))
                        else
                            GenJournalLine.Validate(Amount, Round((EstadosAfiliados."Deposito Bruto" - EstadosAfiliados.Comision) * -1));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);

                    end;
                    //    Termino de segunda linea

                    //    Tercera linea
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch liquidacion");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template liquidacion");

                    if GenJnlLine.FindLast then;

                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch liquidacion";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template liquidacion";

                    if GenJournalLine.Insert(true) then begin

                        //azul
                        GenJournalLine."Posting Date" := EstadosAfiliados."Fecha de Entrada";


                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");
                        if TipoProcesadorPago = TipoProcesadorPago::Amex then
                            GenJournalLine."Document No." := 'AMEX-' + Format(EstadosAfiliados."Fecha de Entrada")
                        else
                            GenJournalLine."Document No." := NumeroFactura;
                        GenJournalLine.Validate("Account No.", BancoLiquidacion);
                        if TipoProcesadorPago = TipoProcesadorPago::VisaNet then
                            GenJournalLine.Validate(Amount, Round((EstadosAfiliados."Total pagado Visanet")))
                        else
                            GenJournalLine.Validate(Amount, Round((EstadosAfiliados."Deposito Bruto" - EstadosAfiliados."Retencion ITBIS" - EstadosAfiliados.Comision)));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;
                end

            until estadosafiliados.Next() = 0;
    end;

    procedure ImportarAmex(TipoProcesadorPago: enum DSNTipoProcesadorDePago)
    begin
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();
            TotalRows := 0;
            I := 0;
            //Ventana.OPEN(Text002);
            ExcelBuffer.DeleteAll;
            Commit;
            ExcelBuffer.LockTable;
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet;
            ExcelBuff.DeleteAll;
            ExcelBuffer1.DeleteAll;
            ExcelBuffer2.DeleteAll;
            ExcelBuffer3.DeleteAll;

            ExcelBuffer.Find('-');
            repeat
                ExcelBuff.TransferFields(ExcelBuffer);
                ExcelBuffer1.TransferFields(ExcelBuffer);
                ExcelBuffer2.TransferFields(ExcelBuffer);
                ExcelBuffer3.TransferFields(ExcelBuffer);
                ExcelBuff.Insert();
                ExcelBuffer1.Insert();
                ExcelBuffer2.Insert();
                ExcelBuffer3.Insert();
            until ExcelBuffer.Next() = 0;
            GetLastRowandColumn;

            EstadosAfiliados.DELETEALL;
            Commit;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Fecha de Entrada');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 1);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Total liquidado: American Express');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 2);

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Distribucion de pago", I);
            end;

            InsertarDiario(TipoProcesadorPago);

            Message('Import is completed');
        end;
    end;

    procedure ImportVisaNet(TipoProcesadorPago: enum DSNTipoProcesadorDePago)
    begin
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();
            TotalRows := 0;
            I := 0;
            //Ventana.OPEN(Text002);
            ExcelBuffer.DeleteAll;
            Commit;
            ExcelBuffer.LockTable;
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet;
            ExcelBuff.DeleteAll;
            ExcelBuffer1.DeleteAll;
            ExcelBuffer2.DeleteAll;
            ExcelBuffer3.DeleteAll;

            ExcelBuffer.Find('-');
            repeat
                ExcelBuff.TransferFields(ExcelBuffer);
                ExcelBuffer1.TransferFields(ExcelBuffer);
                ExcelBuffer2.TransferFields(ExcelBuffer);
                ExcelBuffer3.TransferFields(ExcelBuffer);
                ExcelBuff.Insert();
                ExcelBuffer1.Insert();
                ExcelBuffer2.Insert();
                ExcelBuffer3.Insert();
            until ExcelBuffer.Next() = 0;
            GetLastRowandColumn;

            EstadosAfiliados.DELETEALL;
            Commit;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Sucursal');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 1);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Column No.", 2);
            ExcelBuffer2.FindLast();
            FinalRow := (ExcelBuffer2."Row No.");

            for I := Iinit to FinalRow do begin
                InsertData(TipoProcesadorPago, Seccion::"Distribucion de pago", I);
            end;
            CrearFacturaVisanet();
            InsertarDiario(TipoProcesadorPago);

            Message('Import is completed');
        end;
    end;

    local procedure ReturnColumNo(ColumnName: Text) ColumnN: Integer
    begin

        ExcelBuffer3.Reset;

        ExcelBuffer3.Reset;
        ExcelBuffer3.SetRange("Cell Value as Text", ColumnName);
        if ExcelBuffer3.FindFirst then;
        ColumnN := ExcelBuffer3."Column No.";
    end;

    local procedure CrearFactura(TipoProcesadorPago: enum DSNTipoProcesadorDePago)
    var
        PurchaseHeader: Record "Purchase Header";
        PurchaseLine: Record "Purchase Line";
        EstadosAfiliados: Record EstadosAfiliados;
    begin

        EstadosAfiliados.RESET;
        if TipoProcesadorPago = TipoProcesadorPago::Cardnet then
            EstadosAfiliados.SETRANGE(Seccion, Seccion::"Distribucion de pago");

        if TipoProcesadorPago = TipoProcesadorPago::Azul then
            EstadosAfiliados.SETRANGE(Seccion, Seccion::"Comprobante fiscal por cargos");

        if EstadosAfiliados.FINDSET then begin
            repeat
                // cabecera
                Clear(PurchaseHeader);
                Clear(PurchaseLine);
                PurchaseHeader.Reset;
                PurchaseHeader.Init;
                PurchaseHeader.Validate("Document Type", PurchaseHeader."Document Type"::Invoice);
                PurchaseHeader.Insert(true);
                PurchaseHeader.Validate("Buy-from Vendor No.", ProveedorFacturaCompra);
                PurchaseHeader.Validate("Posting Date", EstadosAfiliados."Fecha de Entrada");
                PurchaseHeader.Validate("Vendor Invoice No.", EstadosAfiliados.NCF);
                PurchaseHeader.Validate("DSNCod. Clasificacion Gasto", '07');
                PurchaseHeader.Validate("DSNNo. Comprobante Fiscal", EstadosAfiliados.NCF);
                PurchaseHeader.Modify(true);

                // lineas
                PurchaseLine.Reset;
                PurchaseLine.Init;
                PurchaseLine.Validate("Document Type", PurchaseLine."Document Type"::Invoice);
                PurchaseLine.Validate("Document No.", PurchaseHeader."No.");
                PurchaseLine.Validate("Line No.", 10000);
                PurchaseLine.Insert(true);
                PurchaseLine.Validate(Type, PurchaseLine.Type::"G/L Account");
                PurchaseLine.Validate("No.", NumeroCuentaFacturaCompra);
                PurchaseLine.Validate(Quantity, 1);
                PurchaseLine.Validate("Direct Unit Cost", EstadosAfiliados.Comision);
                InsertarDimTemp('DEPARTAMENTO', Departamento);
                InsertarDimTemp('SUCURSAL', Sucursal);
                InsertarDimTemp('LINNEGOCIO', LinNegocio);
                //InsertarDimTempDef(39);
                PurchaseLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                PurchaseLine.validate("Dimension Set ID");
                PurchaseLine.Modify(true);

            until EstadosAfiliados.NEXT = 0
        end;
    end;

    procedure CrearFacturaVisanet()
    var

        x: Codeunit ColumnasExcel;
    begin
        Clear(YYYY);
        Clear(MM);
        Clear(DD);
        Clear(iYYYY);
        Clear(iMM);
        Clear(iDD);
        //NCF
        ExcelBuffer2.Reset();
        ExcelBuffer2.SetRange(xlColID, 'M');
        ExcelBuffer2.SetFilter("Cell Value as Text", '<>%1', '');
        ExcelBuffer2.SetFilter("Cell Value as Text", '<>%1', 'NCF');
        ExcelBuffer2.FindFirst();
        NCFVisaNet := ExcelBuffer2."Cell Value as Text";

        //Fecha
        ExcelBuffer2.Reset();
        ExcelBuffer2.SetRange(xlColID, 'E');
        ExcelBuffer2.SetRange("Cell Value as Text", 'P');
        ExcelBuffer2.FindFirst();
        ExcelBuff.Reset();
        ExcelBuff.SetRange("Row No.", ExcelBuffer2."Row No.");
        ExcelBuff.SetRange(xlColID, 'N');
        ExcelBuff.FindFirst();
        YYYY := CopyStr(ExcelBuff."Cell Value as Text", 1, 4);
        MM := CopyStr(ExcelBuff."Cell Value as Text", 6, 2);
        DD := CopyStr(ExcelBuff."Cell Value as Text", 9, 2);
        Evaluate(iYYYY, YYYY);
        Evaluate(iDD, DD);
        Evaluate(iMM, MM);
        FechaFac := DMY2Date(iDD, iMM, iYYYY);

        //cabecera
        ExcelBuffer2.Reset();
        ExcelBuffer2.SetRange(xlColID, 'E');
        ExcelBuffer2.SetRange("Cell Value as Text", 'P');
        ExcelBuffer2.FindFirst();
        Clear(PurchaseHeader);
        Clear(PurchaseLine);
        PurchaseHeader.Reset;
        PurchaseHeader.Init;
        PurchaseHeader.Validate("Document Type", PurchaseHeader."Document Type"::Invoice);
        PurchaseHeader.Insert(true);
        PurchaseHeader.Validate("Buy-from Vendor No.", ProveedorFacturaCompra);
        PurchaseHeader.Validate("Posting Date", FechaFac);
        PurchaseHeader.Validate("Vendor Invoice No.", x.GetNCF(ExcelBuffer2, 'M', ExcelBuffer2."Row No."));
        PurchaseHeader.Validate("DSNCod. Clasificacion Gasto", '07');
        PurchaseHeader.Validate("DSNNo. Comprobante Fiscal", x.GetNCF(ExcelBuffer2, 'M', ExcelBuffer2."Row No."));
        PurchaseHeader.Modify(true);

        //lineas
        ExcelBuffer2.Reset();
        ExcelBuffer2.SetRange(xlColID, 'E');
        ExcelBuffer2.SetRange("Cell Value as Text", 'P');
        ExcelBuffer2.FindSet();
        repeat
            PurchaseLine.Reset;
            PurchaseLine.Init;
            PurchaseLine.Validate("Document Type", PurchaseLine."Document Type"::Invoice);
            PurchaseLine.Validate("Document No.", PurchaseHeader."No.");
            PurchaseLine."Line No." += 1;
            PurchaseLine.Insert(true);
            PurchaseLine.Validate(Type, PurchaseLine.Type::"G/L Account");
            PurchaseLine.Validate("No.", NumeroCuentaFacturaCompra);
            PurchaseLine.Validate(Quantity, 1);
            PurchaseLine.Validate("Direct Unit Cost", x.GetComision(ExcelBuffer2, 'J', ExcelBuffer2."Row No."));
            InsertarDimTemp('DEPARTAMENTO', Departamento);
            InsertarDimTemp('SUCURSAL', Sucursal);
            InsertarDimTemp('LINNEGOCIO', LinNegocio);
            //InsertarDimTempDef(39);
            PurchaseLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);

            if ConfContab."Global Dimension 1 Code" = recdfltdim."Dimension Code" then
                PurchaseLine."Shortcut Dimension 1 Code" := recdfltdim."Dimension Value Code"
            else
                if ConfContab."Global Dimension 2 Code" = recdfltdim."Dimension Code" then
                    PurchaseLine."Shortcut Dimension 2 Code" := recdfltdim."Dimension Value Code";
            PurchaseLine.Modify(true);
        until ExcelBuffer2.Next() = 0;
    end;

    procedure InsertGL2(NombreCuentas: Text)
    begin
        Clear(Contador);
        Clear(Cantidad);
        EstadosAfiliados2.RESET;
        GenJnlLine.Reset;
        Cantidad := EstadosAfiliados2.COUNT;
        if EstadosAfiliados2.FINDSET then begin
            repeat
                //Contador  := Contador + 1;
                //Ventana.UPDATE(2,'Insertando en GJL' + FORMAT(ROUND(Contador / Cantidad * 10000,1)));

                Clear(LineNo);
                GenJnlLine.SetRange("Journal Batch Name", JournalBatchName);
                GenJnlLine.SetRange("Journal Template Name", JournalTemplateName);
                if GenJnlLine.FindLast then;
                LineNo := GenJnlLine."Line No." + 10000;

                GenJournalLine.Init;
                GenJournalLine.Validate("Line No.", LineNo);
                GenJournalLine."Journal Batch Name" := JournalBatchName;
                GenJournalLine."Journal Template Name" := JournalTemplateName;


                if GenJournalLine.Insert then begin
                    GenJournalLine."Posting Date" := EstadosAfiliados2."Fecha de Entrada";
                    GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::Payment);
                    GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");
                    GenJournalLine."Document No." := 'Tarjeta' + Format(EstadosAfiliados2."Fecha de Entrada", 0, '<Day,2>');
                    GenJournalLine.Validate("Account No.", EstadosAfiliados2.Cuenta);
                    NombreCuenta := NombreCuentas;
                    GenJournalLine.Description := NombreCuenta;
                    GenJournalLine.Validate(Amount, EstadosAfiliados2."Monto a Liquidar");
                    GenJournalLine.Validate("Shortcut Dimension 1 Code", EstadosAfiliados2."DIM SUC");
                    GenJournalLine.Modify;
                end;

                Clear(NombreCuenta);
                if EstadosAfiliados2."Monto Lote" <> 0 then begin
                    GenJournalLine.Init;
                    LineNo := LineNo + 10000;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := JournalBatchName;
                    GenJournalLine."Journal Template Name" := JournalTemplateName;

                    if GenJournalLine.Insert then begin
                        GenJournalLine."Posting Date" := EstadosAfiliados2."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::Payment);
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"G/L Account");
                        GenJournalLine."Document No." := 'Tarjeta' + Format(EstadosAfiliados2."Fecha de Entrada", 0, '<Day,2>');
                        GenJournalLine.Validate("Account No.", '1047');
                        if GLAccount.Get('1047') then
                            NombreCuenta := GLAccount.Name;
                        GenJournalLine.Description := NombreCuenta;
                        GenJournalLine.Validate(Amount, (EstadosAfiliados2."Monto Lote" * -1));
                        GenJournalLine.Validate("Shortcut Dimension 1 Code", EstadosAfiliados2."DIM SUC");
                        GenJournalLine.Modify;
                    end;
                end;
                Clear(NombreCuenta);
                if EstadosAfiliados2."ITBIS Retenido" <> 0 then begin
                    GenJournalLine.Init;
                    LineNo := LineNo + 10000;

                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := JournalBatchName;
                    GenJournalLine."Journal Template Name" := JournalTemplateName;

                    if GenJournalLine.Insert then begin
                        GenJournalLine."Posting Date" := EstadosAfiliados2."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::Payment);
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"G/L Account");
                        GenJournalLine."Document No." := 'Tarjeta' + Format(EstadosAfiliados2."Fecha de Entrada", 0, '<Day,2>');
                        GenJournalLine.Validate("Account No.", '2032');
                        if GLAccount.Get('2032') then
                            NombreCuenta := GLAccount.Name;
                        GenJournalLine.Description := NombreCuenta;
                        GenJournalLine.Validate(Amount, EstadosAfiliados2."ITBIS Retenido");
                        GenJournalLine.Validate("Shortcut Dimension 1 Code", EstadosAfiliados2."DIM SUC");
                        GenJournalLine.Modify;
                    end;
                end;
                Clear(NombreCuenta);
                if EstadosAfiliados2.Comision <> 0 then begin
                    GenJournalLine.Init;
                    LineNo := LineNo + 10000;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := JournalBatchName;
                    GenJournalLine."Journal Template Name" := JournalTemplateName;

                    if GenJournalLine.Insert then begin
                        GenJournalLine."Posting Date" := EstadosAfiliados2."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::Payment);
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"G/L Account");
                        GenJournalLine."Document No." := 'Tarjeta' + Format(EstadosAfiliados2."Fecha de Entrada", 0, '<Day,2>');
                        GenJournalLine.Validate("Account No.", '6102');
                        if GLAccount.Get('6102') then
                            NombreCuenta := GLAccount.Name;
                        GenJournalLine.Description := NombreCuenta;
                        GenJournalLine.Validate(Amount, EstadosAfiliados2.Comision);
                        GenJournalLine.Validate("Shortcut Dimension 1 Code", EstadosAfiliados2."DIM SUC");
                        GenJournalLine.Modify;
                    end;
                end;
            until EstadosAfiliados2.NEXT = 0;
        end;

        //Rec.COPY(GenJournalLine);
        //MESSAGE('Import is completed');
    end;


    procedure GetLastRowandColumn()

    begin
        ExcelBuff.SetRange("Row No.", 1);
        TotalColumns := ExcelBuff.Count;

        ExcelBuff.Reset;
        if ExcelBuff.FindLast then
            TotalRows := ExcelBuff."Row No.";
    end;

    procedure SetParameters(ProveedorFacturaCompraP: Code[20]; NumeroCuentaFacturaCompraP: Code[20]; ProveedorCompensacionP: Code[20]; BancoCompesacionP: Code[20]; CuentaContableLiquidacionP: Code[20]; CajaLiquidacionP: Code[20]; BancoLiquidacionP: Code[20]; LinNegocioP: Code[20]; DepartamentoP: Code[20]; SucursalP: Code[20])
    begin
        ProveedorFacturaCompra := ProveedorFacturaCompraP;
        NumeroCuentaFacturaCompra := NumeroCuentaFacturaCompraP;
        ProveedorCompensacion := ProveedorCompensacionP;
        BancoCompesacion := BancoCompesacionP;
        CuentaContableLiquidacion := CuentaContableLiquidacionP;
        CajaLiquidacion := CajaLiquidacionP;
        BancoLiquidacion := BancoLiquidacionP;
        LinNegocio := LinNegocioP;
        Departamento := DepartamentoP;
        Sucursal := SucursalP;
    end;

    procedure InsertarDimTemp(DimCode: Code[20]; DimValue: Code[20])
    begin
        recDimVal.Get(DimCode, DimValue);
        if not TempDimEntry.Get(recdimval."Dimension value Id", DimCode) then begin
            Clear(TempDimEntry);
            TempDimEntry.Validate("Dimension Code", DimCode);
            TempDimEntry.Validate("Dimension Value Code", DimValue);
            TempDimEntry.Validate("Dimension Value ID", recDimVal."Dimension Value ID");
            if TempDimEntry.Insert(true) then;
        end;
    end;

    procedure InsertarDimTempDef(intPrmTabla: Integer)
    begin
        ConfContab.Get();
        recDfltDim.Reset();
        recDfltDim.SetRange("Table ID", intPrmTabla);
        //recDfltDim.SetRange("No.", PurchaseLine."Document No.");
        if recDfltDim.FindSet() then
            repeat
                InsertarDimTemp(recDfltDim."Dimension Code", recDfltDim."Dimension Value Code");

                if ConfContab."Global Dimension 1 Code" = recdfltdim."Dimension Code" then
                    PurchaseLine."Shortcut Dimension 1 Code" := recdfltdim."Dimension Value Code"
                else
                    if ConfContab."Global Dimension 2 Code" = recdfltdim."Dimension Code" then
                        PurchaseLine."Shortcut Dimension 2 Code" := recdfltdim."Dimension Value Code";
            until recDfltDim.Next() = 0;
    end;
}

