pageextension 55000 AccountantRoleCenter extends "Accountant Role Center"
{
    actions
    {
        addlast(reporting)
        {
            action(ImportCardNetAzulAmex)
            {
                Caption = 'Import CardNet Azul Amex';
                ApplicationArea = all;
                Image = ExportToExcel;
                RunObject = report "Import Cardnet Azul Amex";
            }
            action(ImportCardNet)
            {
                Caption = 'Import CardNet';
                ApplicationArea = all;
                Image = ExportToExcel;
                RunObject = report "Import Cardnet";
            }
        }
    }
}
