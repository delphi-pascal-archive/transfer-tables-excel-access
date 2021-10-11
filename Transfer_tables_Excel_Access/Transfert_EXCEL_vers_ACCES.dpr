program Transfert_EXCEL_vers_ACCES;

uses
  Forms,
  Transfert_EXCEL_vers_ACCES_u in 'Transfert_EXCEL_vers_ACCES_u.pas' {F_Transfert},
  bd_u in 'bd_u.pas' {BD_};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TF_Transfert, F_Transfert);
  Application.CreateForm(TBD_, BD_);
  Application.Run;
end.
