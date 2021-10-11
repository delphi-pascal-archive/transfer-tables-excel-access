unit Transfert_EXCEL_vers_ACCES_u;
{
 Juillet 2010

 EXEMPLE DE RANSFERT DE TABLES ENTRE une base EXCEL et  ACCESS :
  avec SQL et les composants ADO...

 copie d'une table  EXCEL  vers une nouvelle table   ACCESS
 copie d'une table  ACCESS vers une nouvelle table   EXCEL
 copie d'une table  ACCESS vers une nouvelle table   ACCESS
 copie d'une table  EXCEL  vers une nouvelle table   EXCEL

 insertion d'une table  EXCEL  vers  table existante  ACCESS
 insertion d'une table  ACCESS vers  table existante  EXCEL
 insertion d'une table  EXCEL  vers  table existante  EXCEL
 insertion d'une table  ACCESS vers  table existante  ACCESS

 effacement de tables dans EXCEL et ACCESS

 la couleur bleu --> source
            vert --> destination
}
interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,ShellAPI,
  Dialogs, StdCtrls, Buttons, Menus, ExtCtrls, DBCtrls, Grids, DBGrids,
  ComCtrls;

type
  TF_Transfert = class(TForm)
    Source_EXCEL: TEdit;
    Destination_ACCESS: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    OpenDialog1: TOpenDialog;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu1: TMainMenu;
    Fini1: TMenuItem;
    Rpertoire1: TMenuItem;
    BD1: TMenuItem;
    Voir1: TMenuItem;
    ComboBox1: TComboBox;
    Label4: TLabel;
    ComboBox2: TComboBox;
    GroupBox2: TGroupBox;
    SQL_cd: TEdit;
    Label5: TLabel;
    SpeedButton3: TSpeedButton;
    StatusBar1: TStatusBar;
    Nouvelle_table: TEdit;
    SpeedButton4: TSpeedButton;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    ComboBox3: TComboBox;
    SpeedButton7: TSpeedButton;
    edit6: TEdit;
    Edit13: TEdit;
    Label10: TLabel;
    Label9: TLabel;
    GroupBox5: TGroupBox;
    SpeedButton8: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton9: TSpeedButton;
    GroupBox6: TGroupBox;
    Label3: TLabel;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    CheckBox2: TCheckBox;
    Timer1: TTimer;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    RichEdit1: TRichEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Fini1Click(Sender: TObject);
    procedure Rpertoire1Click(Sender: TObject);
    procedure Voir1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SQL_cdDblClick(Sender: TObject);
    procedure ComboBox3Change(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton12Click(Sender: TObject);
    procedure Nouvelle_tableKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SpeedButton13Click(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure SpeedButton10MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure GroupBox6MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton9MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure GroupBox5MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton6MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton7MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton8MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton11MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure CheckBox2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure SpeedButton15Click(Sender: TObject);
    procedure SpeedButton15MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure SpeedButton16Click(Sender: TObject);
    procedure SpeedButton16MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  F_Transfert: TF_Transfert;
  Tab_source, Tab_destination : string;
implementation

uses bd_u;

{$R *.dfm}

procedure TF_Transfert.SpeedButton1Click(Sender: TObject);
begin
  IF OpenDialog1.Execute THEN
   Source_EXCEL.Text := OpenDialog1.FileName;
end;

procedure TF_Transfert.SpeedButton2Click(Sender: TObject);
begin
  IF OpenDialog1.Execute THEN
   Destination_ACCESS.Text := OpenDialog1.FileName;
end;

procedure TF_Transfert.FormCreate(Sender: TObject);
begin
   Source_EXCEL.Text       := ExtractFilePath(Application.ExeName)+'Dossier_EXCEL.xls';

   Destination_ACCESS.Text := ExtractFilePath(Application.ExeName)+'Dossier_ACCESS.mdb';;

   //
   SQL_cd.Text :=  'SELECT * INTO '+ Nouvelle_table.Text + ' IN "'+   Destination_ACCESS.Text +'" FROM [' + combobox1.Text +']';

   left := 5;
end;

procedure TF_Transfert.Fini1Click(Sender: TObject);
begin
    Close
end;

procedure TF_Transfert.Rpertoire1Click(Sender: TObject);
begin

  shellExecute(handle,'open',pchar(ExtractFilePath(Source_EXCEL.Text)),NIL,nil,sw_show) ;

end;

procedure TF_Transfert.Voir1Click(Sender: TObject);
begin
        bd_.show
end;

procedure TF_Transfert.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
      bd_.ADOConnection1.Close;
      bd_.ADOConnection2.Close;
end;

procedure TF_Transfert.SpeedButton3Click(Sender: TObject);
var sCopy : string;
begin
// copie une table EXCEL  dans la base ACCESS
{ voir site
http://www.access-programmers.co.uk/forums/showthread.php?t=159409
http://www.w3schools.com/Sql/sql_select_into.asp
}
// SOURCE ET DESTINATION
   Tab_source     := Source_EXCEL.Text ;  //  Destination_ACCESS.Text
   Tab_destination:= Nouvelle_table.Text;

 if not BD_.ADOConnection1.Connected then
  begin
    showmessage('base EXCEL A CONNECTER');
    bd_.show;
    EXIT;
  end;
//
// commande SQL
  sCopy := SQL_cd.Text  ;
// existance de la table dans ACCESS
  SpeedButton3.Tag := combobox2.Items.IndexOf( Tab_destination ) ;
  if SpeedButton3.Tag   > 0 then
  BEGIN
   showmessage(' la table : ' +Tab_destination+ #10'  existe dans la base ACCESS');
   exit;
  END;
 //   source dans ACCESS
  caption :=  sCopy;
  try
   BD_.ADOQuery1.SQL.Text:=sCopy;
   BD_.ADOQuery1.ExecSQL;
   StatusBar1.SimpleText:= 'Transfert EXCEL --> ACCESS ... FAIT';
  except
   StatusBar1.SimpleText:= sCopy;
  end;

// initialise la listed des tables ACCESS
  SpeedButton4Click(Self);
end;

procedure TF_Transfert.SpeedButton4Click(Sender: TObject);
begin
    bd_.ADOConnection2.Close;
    bd_.ADOConnection2.Connected := true;
end;

procedure TF_Transfert.SpeedButton5Click(Sender: TObject);
begin

  SpeedButton5.Tag := combobox2.Items.IndexOf( combobox2.Text ) ;
  if SpeedButton5.Tag   <= 0 then
  BEGIN
   showmessage('La table : ' +combobox2.Text+ #10'Est inexistante dans la base ACCESS');
   exit;
  END;

  STATUSBAR1.SimpleText :=  BD_.AdoQuery2.SQL.Text ;

  try
   BD_.ADOQuery2.SQL.Text:='DROP TABLE ['+combobox2.Text+']';;
   BD_.ADOQuery2.ExecSQL;
   StatusBar1.SimpleText:= 'Table ACCESS effacée ....';
   combobox2.ItemIndex:= 0;
  except
   StatusBar1.SimpleText:= 'Erreur à l''éffacement';
  end;

end;

procedure TF_Transfert.SpeedButton6Click(Sender: TObject);
begin
//création de la commande SQL
//SELECT * INTO copie_access IN "D:\Projets 2010\Programmes\Mouvement du personnel\BD_Equipements_DMPH.mdb" FROM [Feuil1$]

 SQL_cd.Enabled := true;

 SQL_cd.Text :=  'SELECT * INTO '+ Nouvelle_table.Text + ' IN "'+   Destination_ACCESS.Text +'" FROM [' + combobox1.Text +']';
 SpeedButton3.Visible := true;
end;

procedure TF_Transfert.SQL_cdDblClick(Sender: TObject);
begin
   ComboBox3.Visible := not ComboBox3.Visible
end;

procedure TF_Transfert.ComboBox3Change(Sender: TObject);
begin
    SQL_cd.Text := ComboBox3.Text ;

    ComboBox3.Visible := false;
end;

procedure TF_Transfert.SpeedButton7Click(Sender: TObject);
var sCopy : string;
begin
// copie une table ACCESS dans la base ACCESS

   Tab_source     := edit13.Text ;// table ACCESS existante
   Tab_destination:= edit6.Text; // table EXCEL nouvelle

  sCopy := 'SELECT * INTO ["Excel 8.0;Database=' + Source_EXCEL.Text + '"].['+Tab_destination+'] FROM '+  Tab_source;

  // LA TABLE EXCEL EXISTE ?
  SpeedButton7.Tag := combobox1.Items.IndexOf(Tab_destination ) ;
  if  SpeedButton7.Tag   > 0 then
  BEGIN
   showmessage(' la table EXCEL : ' +Tab_destination+ #10'  existe ');
   exit;
  END;

// COPIE  ACCESS VERS EXCEL   POSSIBLE
  caption :=  sCopy;
  TRY
  BD_.ADOQuery2.SQL.Text:=sCopy;  // liée avec ACCESS
  BD_.ADOQuery2.ExecSQL;
    StatusBar1.SimpleText:= 'Transfert ACCESSS vers EXCEL ... FAIT';
  EXCEPT
   StatusBar1.SimpleText:= sCopy;
  END;
  
end;

procedure TF_Transfert.SpeedButton8Click(Sender: TObject);
var sCopy : string;
begin
// copie une table EXCEL dans la base EXCEL
// source      TABLE ACCESS  Authors
// destination TABLE EXCEL   Authors
{
pour le transfert d'une feuille mettre entre crochet [] le nom de la feuille
}
   Tab_source     := combobox1.text ;// table excel existante
   Tab_destination:= edit6.Text; // table EXCEL nouvelle


  sCopy := 'SELECT * INTO ['+Tab_destination+']  FROM ['+ Tab_source +']';

  // la nouvelle est-elle présente  ?
  SpeedButton8.Tag := combobox1.Items.IndexOf(Tab_destination );
  STATUSBAR1.SimpleText := '!';

  if BD_.ADOConnection1.Connected then
  begin
  
  IF SpeedButton8.Tag <0 then
  begin
  try
  STATUSBAR1.SimpleText := sCopy ;
  BD_.ADOQuery1.SQL.Text:= sCopy;
  BD_.ADOQuery1.ExecSQL;
  except
    showmessage('Erreur lors du transfert ')
  end;
  end else showmessage('La table destination : '+ Tab_destination +#10' existe ')

  end else  bd_.Show;
end;

procedure TF_Transfert.SpeedButton10Click(Sender: TObject);
var sAppend : string;
begin
// insertion d'une table EXCEL vers une table ACCESS
// La table ACCESS doit exister et avoir des champs identiques

   Tab_source     := combobox1.text ; // table excel
   Tab_destination:= Nouvelle_table.Text; // table ACCESS déjas existante

 sAppend:='INSERT INTO ['+Tab_destination+'] IN "' + DESTINATION_ACCESS.Text + '" SELECT * FROM  ['+ Tab_source+']';
 STATUSBAR1.SimpleText := sAppend ;

    if CheckBox2.Checked then sAppend  :=  RichEdit1.Lines[0]  
     else  RichEdit1.Text := sAppend;

     RichEdit1.Lines.Add(   sAppend   ) ;
{
 il est possible dans SELECT * FROM
 de remplacer * par une liste de champ présent dans les tables
 ceci permet de ne traiter que ces champs
}
// teste de la présence ou non de la table dans la base EXCEL
   SpeedButton10.Tag := combobox1.Items.IndexOf( Tab_source ) ;
   if SpeedButton10.Tag   < 0 then
   BEGIN
   showmessage(' la table : ' +Tab_source+ #10'  N''existe pas dans la base EXCEL');
   exit;
   END;

// teste de la présence ou non de la table dans la base ACCESS
   SpeedButton10.Tag := combobox2.Items.IndexOf( Tab_destination ) ;
   if SpeedButton10.Tag   < 0 then
   BEGIN
   showmessage(' la table : ' +Tab_destination+ #10'  N''existe pas dans la base ACCESS');
   exit;
   END;

   try
    bd_.ADOQuery1.SQL.Text:=sAppend;
    bd_.AdoQuery1.ExecSQL;
   except
    StatusBar1.SimpleText:= 'Erreur d''insertion ...';
   end;
   
end;

procedure TF_Transfert.SpeedButton11Click(Sender: TObject);
var  sAppend :string;
begin
// insertion d'une table ACCESS dans une table EXCEL existante

   Tab_source     := edit13.text ;// table ACCESS existante
   Tab_destination:= edit6.Text; // table EXCEL nouvelle

  sAppend:='INSERT INTO ['+Tab_destination+'] IN "' + Source_EXCEL.Text + '" "Excel 8.0;" SELECT * FROM '+ Tab_source;

// insertion  ACCESS VERS EXCEL   

  TRY
  BD_.ADOQuery2.SQL.Text:= sAppend;  //
  BD_.ADOQuery2.ExecSQL;
  StatusBar1.SimpleText:= 'Insertion ACCESSS vers EXCEL ... FAIT';
  EXCEPT
   StatusBar1.SimpleText:= sAppend;
  END;
end;

procedure TF_Transfert.SpeedButton9Click(Sender: TObject);
var sCopy : string;
begin
// copie une table  ACCESS dans la base  ACCESS

   Tab_source     := combobox2.text ;
   Tab_destination:= Nouvelle_table.Text;

// teste la présence de la table  destination
  SpeedButton11.Tag := combobox2.Items.IndexOf( Tab_destination ) ;
  if  SpeedButton11.Tag   > 0 then
  BEGIN
   showmessage(' la table : ' +Tab_destination+ #10'  existe dans la base ACESS');
   exit;
  END;
               //           destination              source
  sCopy := 'SELECT * INTO ['+ Tab_destination +']  FROM '+ Tab_source;

  // copie 
  try
  bd_.ADOQuery2.SQL.Text:=sCopy;
  bd_.AdoQuery2.ExecSQL;

  StatusBar1.SimpleText:= sCopy
  except
     StatusBar1.SimpleText:= '!';
  end;
end;

procedure TF_Transfert.SpeedButton12Click(Sender: TObject);
begin
       bd_.show;
       GroupBox5.Visible:= true;GroupBox6.Visible:= true;
       bd_.SpeedButton1Click(self);
end;

procedure TF_Transfert.Nouvelle_tableKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
SpeedButton3.Visible := false;
SpeedButton6.Visible := true;
end;

procedure TF_Transfert.SpeedButton13Click(Sender: TObject);
begin
    bd_.ADOConnection1.Close;
    bd_.ADOConnection1.Connected := true;
end;

procedure TF_Transfert.SpeedButton14Click(Sender: TObject);
begin

// efface une table EXCEL

  SpeedButton14.Tag := combobox1.Items.IndexOf( combobox1.Text ) ;
  if SpeedButton14.Tag   <= 0 then
  BEGIN
   showmessage('La table : ' +combobox1.Text+ #10'Est inexistante dans la base EXCEL');
   exit;
  END;

  STATUSBAR1.SimpleText :=  BD_.AdoQuery2.SQL.Text ;

  try
   BD_.ADOQuery1.SQL.Text:='DROP TABLE ['+combobox1.Text+']';;
   BD_.ADOQuery1.ExecSQL;
   StatusBar1.SimpleText:= 'Tabel EXCEL effacée ..';
   combobox2.ItemIndex:= 0;
  except
   StatusBar1.SimpleText:= 'Erreur à l''éffacement';
  end;

    SpeedButton13Click(Self);

end;

procedure TF_Transfert.SpeedButton10MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin          
     combobox1.Color := clAqua   ; // table ACCESS
 
     Nouvelle_table.Color:= clGreen;
end;

procedure TF_Transfert.GroupBox6MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
     combobox1.Color := clwhite;
     combobox2.Color := clwhite;
     combobox2.Font.Color:= clBlack; // table excel
     combobox1.Font.Color:= clBlack; // table excel
     Nouvelle_table.Font.Color:= clBlack;

  //   Nouvelle_table.Focused ;
    SQL_cd.Color := clWhite;
    SQL_cd.Font.Color := clblack;
    Destination_ACCESS.Font.Color := clblack;
    edit13.Color := clWHite;
    NOUVELLe_table.Color  := clWhite;
end;

procedure TF_Transfert.SpeedButton9MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin

     combobox2.Color := clAqua   ;
  //   combobox2.Font.Color:=clwhite; ; // table excel
     Nouvelle_table.Color:= clGreen;
end;

procedure TF_Transfert.GroupBox5MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
     edit13.Font.Color := clblack ;//
     edit6.Font.Color  := clblack;
     edit13.Color := clWhite ;//
     edit6.Color  := clWhite;
     edit6.Font.Color  := clBlack;

     combobox1.Font.Color := clblack; ;// table excel existante
     combobox1.Color := clWhite;
     combobox2.Color := clWhite;


     combobox1.Font.Color:= clblack; // table excel
     combobox2.Font.Color:= clblack; // table excel
     Nouvelle_table.Font.Color:= clblack;
     Nouvelle_table.Focused ;

      Source_EXCEL.Color := clWhite;
end;

procedure TF_Transfert.SpeedButton3MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
   SQL_cd.Color := clAqua;
   Nouvelle_table.Color := clGreen;
   
end;

procedure TF_Transfert.SpeedButton6MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
   SQL_cd.Font.Color := clBlue;

   combobox1.Color := clAqua;
   Destination_ACCESS.Font.Color := clBlue;

   Nouvelle_table.Font.Color := clGreen;
end;

procedure TF_Transfert.SpeedButton7MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    edit13.Color := clAqua ;// table ACCESS existante
    edit6.Color := clGreen;
end;

procedure TF_Transfert.SpeedButton8MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    combobox1.Color := clAqua; ;// // table excel existante
    edit6.Color := clGreen;     // table excel nouvelle
end;

procedure TF_Transfert.SpeedButton11MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    edit13.Color := clAqua;// table ACCESS existante
    edit6.Color := clGreen;// table EXCEL nouvelle
    Source_EXCEL.Color := clAqua;
end;

procedure TF_Transfert.CheckBox2Click(Sender: TObject);
begin
   if CheckBox2.Checked then CheckBox2.Caption :='sql memo'
   else  CheckBox2.Caption :='sql nouveau'


end;

procedure TF_Transfert.Timer1Timer(Sender: TObject);
begin
    Timer1.Enabled := false;
     SpeedButton12Click(Self);
     bd_.Left := screen.Width - bd_.Width;

end;

procedure TF_Transfert.SpeedButton15Click(Sender: TObject);
var sAppend : string;
begin
// insertion d'une table ACCESS vers une table ACCESS
// La table ACCESS doit exister et avoir des champs identiques

   Tab_source     := combobox1.text ; // table excel
   Tab_destination:= Nouvelle_table.Text; // table ACCESS déjas existante

 sAppend:='INSERT INTO [Nouvelle_table] IN ''' + DESTINATION_ACCESS.Text +   ''' SELECT * FROM  [T_ACCESS]' ;

    STATUSBAR1.SimpleText := sAppend ;
    if CheckBox2.Checked then
    begin
       sAppend  :=  RichEdit1.Lines[0] ;

    end else  RichEdit1.Text := sAppend;
     RichEdit1.Lines.Add(   sAppend   ) ;

   SpeedButton10.Tag := combobox1.Items.IndexOf( Tab_source ) ;
   if SpeedButton10.Tag   < 0 then
   BEGIN
   showmessage(' la table : ' +Tab_source+ #10'  N''existe pas dans la base EXCEL');
   exit;
   END;


   SpeedButton10.Tag := combobox2.Items.IndexOf( Tab_destination ) ;
   if SpeedButton10.Tag   < 0 then
   BEGIN
   showmessage(' la table : ' +Tab_destination+ #10'  N''existe pas dans la base ACCESS');
   exit;
   END;

   try
    begin
    bd_.ADOQuery2.SQL.Text:=sAppend;
    bd_.AdoQuery2.ExecSQL;
    end;
   except
    StatusBar1.SimpleText:= 'Erreur d''insertion ...';
   end;

end;

procedure TF_Transfert.SpeedButton15MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    combobox2.Color := clAqua ; // table excel
    Nouvelle_table.Color := clGreen ;
    edit13.Color :=clAqua   ;

   DESTINATION_ACCESS.Color :=clAqua;
end;

procedure TF_Transfert.SpeedButton16Click(Sender: TObject);
var  sAppend :string;
begin
// insertion d'une table EXCEL  dans une table  EXCEL  existante

   Tab_source     := combobox1.text ;// table ACCESS existante
   Tab_destination:= edit6.Text; // table EXCEL nouvelle


// insertion une table access vers une feuille excel
                          //destination          nom ficheier excel
  sAppend:='INSERT INTO ['+Tab_destination+'] IN "' + Source_EXCEL.Text + '" "Excel 8.0;" SELECT * FROM ['+ Tab_source+']';

  RichEdit1.Lines.Add(SaPPEND) ;

// insertion  ACCESS VERS EXCEL   

  TRY
  BD_.ADOQuery1.SQL.Text:= sAppend;  //
  BD_.ADOQuery1.ExecSQL;
  StatusBar1.SimpleText:= 'Insertion ACCESSS vers EXCEL ... FAIT';
  EXCEPT
   StatusBar1.SimpleText:= sAppend;
  END;
end;

procedure TF_Transfert.SpeedButton16MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
combobox1.Color := clAqua;
edit6.Color := clGreen;

Source_EXCEL.Color := clAqua;
end;

end.
