unit bd_u;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Buttons, StdCtrls, ExtCtrls, DBCtrls, Grids, DBGrids,
  ComCtrls;

type
  TBD_ = class(TForm)
    ADOConnection1: TADOConnection;
    ADOConnection2: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DataSource1: TDataSource;
    DataSource2: TDataSource;
    SpeedButton1: TSpeedButton;
    GroupBox1: TGroupBox;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    GroupBox2: TGroupBox;
    DBGrid2: TDBGrid;
    DBNavigator2: TDBNavigator;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    Label3: TLabel;
    Label4: TLabel;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    StatusBar1: TStatusBar;
    SpeedButton6: TSpeedButton;
    Label5: TLabel;
    Label6: TLabel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ADOConnection1AfterConnect(Sender: TObject);
    procedure ADOConnection2AfterConnect(Sender: TObject);
    procedure ADOConnection2AfterDisconnect(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure ADOQuery1AfterOpen(DataSet: TDataSet);
    procedure ADOQuery2AfterOpen(DataSet: TDataSet);
    procedure RedimGrille_(var DB_Gille : TDBGrid);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure ADOConnection1AfterDisconnect(Sender: TObject);
    procedure ADOQuery1AfterClose(DataSet: TDataSet);
    procedure ADOQuery2AfterClose(DataSet: TDataSet);
     
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  BD_: TBD_;
  Base_EXCEL , Base_ACCESS : string;

implementation

uses Transfert_EXCEL_vers_ACCES_u;

{$R *.dfm}

procedure TBD_.SpeedButton1Click(Sender: TObject);
var strConn :  widestring;
begin
//  connection EXCEL
  AdoConnection1.LoginPrompt:=False;
  AdoQuery1.Connection:=AdoConnection1;
  DataSource1.DataSet :=AdoQuery1;
  DBGrid1.DataSource  :=DataSource1;
  DBNavigator1.DataSource:=DataSource1;
//
  strConn:='Provider=Microsoft.Jet.OLEDB.4.0;' +
           'Data Source=' + Base_EXCEL + ';' +
           'Extended Properties=Excel 8.0;';

  AdoConnection1.Connected:=False;
  AdoConnection1.ConnectionString:=strConn;
  try
    AdoConnection1.Open;
  except
    ShowMessage('Unable to connect to Excel, make sure the workbook ' + Base_EXCEL + ' exist!');
    raise;
  end;

  //connect to an Access database to send the data to Excel
  AdoConnection2.Close ;
  AdoConnection2.LoginPrompt:=False;
  AdoQuery2.Connection:=AdoConnection2;
  DataSource2.DataSet :=AdoQuery2;
  DBGrid2.DataSource  :=DataSource2;
  DBNavigator2.DataSource:=DataSource2;
  //
  AdoConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+ Base_access+ ';Persist Security Info=False';

  // connexion ACCESS
  AdoConnection2.Open ;

//

end;

procedure TBD_.FormCreate(Sender: TObject);
begin
  Base_EXCEL  := F_Transfert.Source_EXCEL.Text   ;
  Base_access := F_Transfert.Destination_ACCESS.Text   ;
end;

procedure TBD_.ADOConnection1AfterConnect(Sender: TObject);
begin
GroupBox1.Visible := true;
AdoConnection1.GetTableNames(combobox1.Items , true);
AdoConnection1.GetTableNames(F_Transfert.ComboBox1.Items, true);
SpeedButton1.Visible := false;
SpeedButton6.Visible := true;
  F_Transfert.GroupBox3.Enabled := true;
  F_Transfert.GroupBox4.Enabled := true;
end;

procedure TBD_.ADOConnection2AfterConnect(Sender: TObject);
begin
GroupBox2.Visible := true;
AdoConnection2.GetTableNames(combobox3.Items , true);
AdoConnection2.GetTableNames(F_Transfert.ComboBox2.Items, true);
F_Transfert.ComboBox2.Visible := true;
F_Transfert.SpeedButton4.Visible := true;
F_Transfert.SpeedButton5.Visible := true;
  F_Transfert.GroupBox3.Enabled := true;
end;

procedure TBD_.ADOConnection2AfterDisconnect(Sender: TObject);
begin
GroupBox2.Visible := FALSE;
F_Transfert.SpeedButton4.Visible := false;
F_Transfert.SpeedButton5.Visible := false;

  F_Transfert.GroupBox3.Enabled := false;
end;

procedure TBD_.SpeedButton2Click(Sender: TObject);
begin
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add('SELECT * FROM ['+Combobox1.Text+']');
  ADOQuery1.Open;
  statusbar1.SimpleText :=  ADOQuery1.SQL.Text ;
end;


procedure TBD_.RedimGrille_(var DB_Gille : TDBGrid);
var n  , m , l : integer;
begin  // redimensionne la grille
     l := 25;   m := DB_Gille.Width -l;

   if  DB_Gille.Columns.Count < 2 then exit;

   m := m div ( DB_Gille.Columns.Count )  ;

   for n := 0 to DB_Gille.Columns.Count-1 do
    if (DB_Gille.Columns[n].FieldName ='N°') or
       (DB_Gille.Columns[n].FieldName ='Numéro') THEN DB_Gille.Columns[n].Width := 5
    ELSE IF (DB_Gille.Columns[n].FieldName ='Ordre') then DB_Gille.Columns[n].Width := 35
     else DB_Gille.Columns[n].Width := m ;
    
end ;
procedure TBD_.ADOQuery1AfterOpen(DataSet: TDataSet);
begin
  ADOQuery1.GetFieldNames(combobox2.Items);
  RedimGrille_( DBGrid1);
  LABEL5.Caption := inttostr( adoquery1.RecordCount);
end;

procedure TBD_.ADOQuery2AfterOpen(DataSet: TDataSet);
begin
  ADOQuery2.GetFieldNames(combobox4.Items);
  RedimGrille_( DBGrid2);
  LABEL6.Caption := inttostr( adoquery2.RecordCount);
end;

procedure TBD_.SpeedButton4Click(Sender: TObject);
begin
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('SELECT * FROM ['+Combobox3.Text+']');
  ADOQuery2.Open;
  statusbar1.SimpleText :=  ADOQuery2.SQL.Text ;
end;

procedure TBD_.SpeedButton5Click(Sender: TObject);
begin
     ADOQuery2.Close
end;

procedure TBD_.SpeedButton3Click(Sender: TObject);
begin
     ADOQuery1.CLOSE
end;

procedure TBD_.SpeedButton6Click(Sender: TObject);
begin
    ADOConnection1.Close;
    ADOConnection2.Close;
end;

procedure TBD_.ADOConnection1AfterDisconnect(Sender: TObject);
begin
SpeedButton1.Visible := true;
SpeedButton6.Visible := false;

  F_Transfert.GroupBox4.Enabled := false;
end;

procedure TBD_.ADOQuery1AfterClose(DataSet: TDataSet);
begin
 LABEL5.Caption := 'Nb...';
end;

procedure TBD_.ADOQuery2AfterClose(DataSet: TDataSet);
begin
LABEL6.Caption := 'nb..';
end;

end.
