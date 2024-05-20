codeunit 55002 "ImportCardNet"
{
    Permissions =
        tabledata "Config. Empresas" = R,
        tabledata "DatosCardNet" = RID,
        tabledata "Default Dimension" = R,
        tabledata "Dimension Value" = R,
        tabledata "Gen. Journal Line" = RIM,
        tabledata "General Ledger Setup" = R,
        tabledata "Purchase Header" = RIM,
        tabledata "Purchase Line" = RIM;

    var
        recDimVal: Record "Dimension Value";
        DatosCardnetUP: record "DatosCardNet";
        DatosCardnetUP2: record "DatosCardNet";
        GenJnlLine: Record "Gen. Journal Line";
        recDfltDim: Record "Default Dimension";
        PurchaseLine: Record "Purchase Line";

        GenJournalLine: Record "Gen. Journal Line";
        ExcelBuff: Record "Excel Buffer" temporary;
        ExcelBuffer: Record "Excel Buffer" temporary;
        ExcelBuffer2: Record "Excel Buffer" temporary;
        ExcelBuffer3: Record "Excel Buffer" temporary;
        ConfContab: record "General Ledger Setup";
        TempDimEntry: Record "Dimension Set Entry" temporary;
        cduDim: Codeunit DimensionManagement;
        Seccion: enum DSNTipoSeccionEstadoAfiliado;
        LineNo: Integer;


        FechaFac: Date;
        LastNo: Integer;
        Filename: Text[250];
        Comision: Decimal;

        ITBIS: Decimal;

        I: Integer;
        Iinit: Integer;
        SheetName: Text;

        FinalRow: Integer;
        NCF: Code[20];
        PurchaseHeader: Record "Purchase Header";
        NumeroFactura: Code[20];
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
        InS: InStream;
        DepositoBruto: Decimal;

    procedure Import()
    begin

        ConfContab.Get();
        ExcelBuffer.DeleteAll();
        if UploadIntoStream('Escoja un archivo', '', '', Filename, ins) then begin
            if SheetName = '' then
                SheetName := ExcelBuffer.SelectSheetsNameStream(InS);
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet();
            I := 0;
            //Ventana.OPEN(Text002);
            ExcelBuffer.DeleteAll;
            Commit;
            ExcelBuffer.LockTable;
            ExcelBuffer.OpenBookStream(InS, SheetName);
            ExcelBuffer.ReadSheet;
            ExcelBuff.DeleteAll();
            ExcelBuffer2.DeleteAll;
            ExcelBuffer3.DeleteAll();


            ExcelBuffer.Find('-');
            repeat
                ExcelBuff.TransferFields(ExcelBuffer);
                ExcelBuffer2.TransferFields(ExcelBuffer);
                ExcelBuffer3.TransferFields(ExcelBuffer);
                ExcelBuff.Insert();
                ExcelBuffer2.Insert();
                ExcelBuffer3.Insert();
            until ExcelBuffer.Next() = 0;

            DatosCardnetUP.DELETEALL;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DISTRIBUCIÓN DE PAGO');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 2);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DETALLES DE LOTE');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 2);

            for I := Iinit to FinalRow do begin
                InsertData(Seccion::"Distribucion de pago", I);
            end;

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'DETALLES DE LOTE');
            if ExcelBuffer2.FindFirst then
                Iinit := (ExcelBuffer2."Row No." + 2);

            ExcelBuffer2.Reset;
            ExcelBuffer2.SetRange("Cell Value as Text", 'Total');
            if ExcelBuffer2.FindFirst then
                FinalRow := (ExcelBuffer2."Row No." - 1);


            for I := Iinit to FinalRow do begin
                InsertData(Seccion::DetalleLote, I);
            end;

            CrearFactura();
            InsertarDiario();

            Message('Import is completed');
        end;
    end;

    procedure insertData(seccion: enum DSNTipoSeccionEstadoAfiliado; Rowno: Integer)
    begin
        clear(DepositoBruto);
        Clear(FechaFac);
        Clear(LastNo);
        if DatosCardnetUP2.FINDLAST then;
        LastNo := DatosCardnetUP2.Id;

        ExcelBuff.Reset();
        ExcelBuff.SetFilter("Row No.", '=%1', Rowno);
        if ExcelBuff.FindSet() then
            repeat
                if seccion = seccion::"Distribucion de pago" then begin
                    case ExcelBuff."Column No." of
                        1:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Fecha') or (ExcelBuff."Cell Value as Text" = '') then
                                    exit;
                                Evaluate(FechaFac, ExcelBuff."Cell Value as Text");
                            end;
                        2:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'NCF') or (ExcelBuff."Cell Value as Text" = '') then
                                    exit;
                                Evaluate(NCF, ExcelBuff."Cell Value as Text");
                            end;
                        3:
                            begin
                                if (ExcelBuff."Cell Value as Text" = 'Comisión') or (ExcelBuff."Cell Value as Text" = '') then
                                    exit;
                                Evaluate(Comision, ExcelBuff."Cell Value as Text");
                            end;
                    end;
                    if ExcelBuff."Column No." = 3 then begin
                        DatosCardnetUP.INIT;
                        LastNo := LastNo + 1;
                        DatosCardnetUP.Id := LastNo;
                        DatosCardnetUP."Fecha de Entrada" := FechaFac + 1;
                        DatosCardnetUP.Seccion := seccion::"Distribucion de pago";
                        DatosCardnetUP.Comision := Comision;
                        DatosCardnetUP.NCF := NCF;

                        if NCF <> '' then
                            DatosCardnetUP.Insert();
                        exit;
                    end
                end;

                if seccion = seccion::DetalleLote then
                    if ExcelBuff."Cell Value as Text" = 'Total por Día' then begin
                        Clear(ExcelBuffer3);
                        ExcelBuffer3.Reset;
                        ExcelBuffer3.Get(Format(ExcelBuff."Row No." - 1), '1');
                        Evaluate(FechaFac, ExcelBuffer3."Cell Value as Text");

                        Clear(ExcelBuffer3);
                        ExcelBuffer3.Reset;
                        ExcelBuffer3.Get(Format(ExcelBuff."Row No."), '3');
                        Evaluate(DepositoBruto, ExcelBuffer3."Cell Value as Text");

                        Clear(ExcelBuffer3);
                        ExcelBuffer3.Reset;
                        ExcelBuffer3.Get(Format(ExcelBuff."Row No."), '4');
                        Evaluate(Comision, ExcelBuffer3."Cell Value as Text");

                        Clear(ExcelBuffer3);
                        ExcelBuffer3.Reset;
                        ExcelBuffer3.Get(Format(ExcelBuff."Row No."), '5');
                        Evaluate(ITBIS, ExcelBuffer3."Cell Value as Text");
                        if (FechaFac <> 0D) then begin

                            DatosCardnetUP.INIT;
                            LastNo := LastNo + 1;
                            DatosCardnetUP.Id := LastNo;
                            DatosCardnetUP."Fecha de Entrada" := FechaFac + 1; //No se por que se estan restando 1 dia
                            DatosCardnetUP."Deposito Bruto" := DepositoBruto;
                            DatosCardnetUP.Comision := Comision;
                            DatosCardnetUP.Seccion := seccion::DetalleLote;
                            DatosCardnetUP."Retencion ITBIS" := ITBIS;
                            DatosCardnetUP.INSERT(true);
                            exit;
                        end;
                    end;
            until ExcelBuff.Next() = 0;
    end;

    local procedure InsertarDiario()
    var
        ConfigEmpresa: Record "Config. Empresas";
    begin
        //View_EstadosAfiliados.RESET;
        GenJnlLine.Reset;
        ConfigEmpresa.Reset;
        ConfigEmpresa.Get;
        DatosCardnetUP.RESET;
        //compensacion
        if DatosCardnetUP.FINDSET then
            repeat
                DatosCardnetUP.Reset();
                if (DatosCardnetUP.Seccion = DatosCardnetUP.Seccion::"Distribucion de pago") then begin
                    Clear(NumeroFactura);
                    Clear(PurchaseHeader);

                    PurchaseHeader.Reset;
                    PurchaseHeader.SetFilter("Buy-from Vendor No.", ProveedorCompensacion);
                    PurchaseHeader.SetFilter("Vendor Invoice No.", DatosCardnetUP.NCF);
                    PurchaseHeader.FindFirst;
                    NumeroFactura := PurchaseHeader."No.";

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
                        GenJournalLine."Posting Date" := DatosCardnetUP."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::Vendor); // customer
                        GenJournalLine."Document No." := NumeroFactura;

                        GenJournalLine.Validate("Account No.", ProveedorCompensacion);
                        GenJournalLine.Validate(Amount, Round(DatosCardnetUP.Comision));
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
                        GenJournalLine."Posting Date" := DatosCardnetUP."Fecha de Entrada";
                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");
                        GenJournalLine."Document No." := NumeroFactura;

                        GenJournalLine.Validate("Account No.", BancoCompesacion);
                        GenJournalLine.Validate(Amount, Round(DatosCardnetUP.Comision * -1));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;
                    //    Termino de segunda linea

                    // end distribucion de pago

                    Clear(PurchaseHeader);
                end;

            until DatosCardnetUP.NEXT = 0;
        //primera linea liquidacion
        DatosCardnetUP.Reset();
        if DatosCardnetUP.FINDSET then
            repeat
                DatosCardnetUP.Reset();
                if DatosCardnetUP.Seccion = DatosCardnetUP.Seccion::DetalleLote then begin
                    Clear(GenJnlLine);
                    Clear(LineNo);
                    GenJnlLine.Reset;
                    GenJnlLine.SetRange("Journal Batch Name", ConfigEmpresa."Journal Batch liquidacion");
                    GenJnlLine.SetRange("Journal Template Name", ConfigEmpresa."Journal Template liquidacion");


                    if DatosCardnetUP."Retencion ITBIS" <> 0 then
                        if GenJnlLine.FindLast then;

                    LineNo := GenJnlLine."Line No." + 10000;
                    GenJournalLine.Init;
                    GenJournalLine.Validate("Line No.", LineNo);
                    GenJournalLine."Journal Batch Name" := ConfigEmpresa."Journal Batch liquidacion";
                    GenJournalLine."Journal Template Name" := ConfigEmpresa."Journal Template liquidacion";
                    if GenJournalLine.Insert(true) then begin

                        GenJournalLine."Posting Date" := DatosCardnetUP."Fecha de Entrada";



                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"G/L Account");

                        GenJournalLine."Document No." := NumeroFactura;

                        GenJournalLine.Validate("Account No.", CuentaContableLiquidacion);
                        GenJournalLine.Validate(Amount, Round(DatosCardnetUP."Retencion ITBIS"));
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

                        GenJournalLine."Posting Date" := DatosCardnetUP."Fecha de Entrada";


                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");

                        GenJournalLine."Document No." := NumeroFactura;
                        GenJournalLine.Validate("Account No.", CajaLiquidacion);

                        GenJournalLine.Validate(Amount, Round((DatosCardnetUP."Deposito Bruto" - DatosCardnetUP.Comision) * -1));
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
                        GenJournalLine."Posting Date" := DatosCardnetUP."Fecha de Entrada";


                        GenJournalLine.Validate("Document Type", GenJournalLine."Document Type"::" ");
                        GenJournalLine.Validate("Account Type", GenJournalLine."Account Type"::"Bank Account");

                        GenJournalLine."Document No." := NumeroFactura;
                        GenJournalLine.Validate("Account No.", BancoLiquidacion);

                        GenJournalLine.Validate(Amount, Round((DatosCardnetUP."Deposito Bruto" - DatosCardnetUP."Retencion ITBIS" - DatosCardnetUP.Comision)));
                        InsertarDimTemp('DEPARTAMENTO', Departamento);
                        InsertarDimTemp('SUCURSAL', Sucursal);
                        InsertarDimTemp('LINNEGOCIO', LinNegocio);
                        GenJournalLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                        GenJournalLine.Validate("Dimension Set ID");
                        GenJournalLine.Modify(true);
                    end;
                end;
            until DatosCardnetUP.Next() = 0;
    end;

    procedure CrearFactura()
    begin
        DatosCardnetUP.Reset();
        DatosCardnetUP.SetRange(Seccion, DatosCardnetUP.Seccion::"Distribucion de pago");
        if DatosCardnetUP.FINDSET then begin
            repeat
                // cabecera
                Clear(PurchaseHeader);
                Clear(PurchaseLine);
                PurchaseHeader.Reset;
                PurchaseHeader.Init;
                PurchaseHeader.Validate("Document Type", PurchaseHeader."Document Type"::Invoice);
                PurchaseHeader.Insert(true);
                PurchaseHeader.Validate("Buy-from Vendor No.", ProveedorFacturaCompra);
                PurchaseHeader.Validate("Posting Date", DatosCardnetUP."Fecha de Entrada");
                PurchaseHeader.Validate("Vendor Invoice No.", DatosCardnetUP.NCF);
                PurchaseHeader.Validate("DSNCod. Clasificacion Gasto", '07');
                PurchaseHeader.Validate("DSNNo. Comprobante Fiscal", DatosCardnetUP.NCF);
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
                PurchaseLine.Validate("Direct Unit Cost", DatosCardnetUP.Comision);
                InsertarDimTemp('DEPARTAMENTO', Departamento);
                InsertarDimTemp('SUCURSAL', Sucursal);
                InsertarDimTemp('LINNEGOCIO', LinNegocio);
                // InsertarDimTempDef(39);
                PurchaseLine."Dimension Set ID" := cduDim.GetDimensionSetID(TempDimEntry);
                PurchaseLine.Validate("Dimension Set ID");
                PurchaseLine.Modify(true);
            until DatosCardnetUP.NEXT = 0
        end;
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
