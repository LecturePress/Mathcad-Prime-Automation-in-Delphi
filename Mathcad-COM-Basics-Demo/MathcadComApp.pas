unit MathcadComApp;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
//  Winapi.ActiveX,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, System.Win.ComObj,
  Ptc_MathcadPrime_Automation_TLB, MathcadComEventSink,
  Vcl.StdCtrls, Vcl.OleServer, Vcl.ExtCtrls;

type
  TMainForm = class(TForm)
    BtnStartMathcad: TButton;
    BtnCloseMathcad: TButton;
    BtnShowMathcad: TButton;
    BtnHideMathcad: TButton;
    BtnGetActiveWorksheet: TButton;
    BtnOpenWorksheet: TButton;
    BtnOpenNewWorksheet: TButton;
    BtnSaveWorksheet: TButton;
    BtnSaveAsWorskheet: TButton;
    BtnSaveAsMctx: TButton;
    BtnCloseWorksheet: TButton;
    RGEvents: TRadioGroup;
    procedure BtnStartMathcadClick(Sender: TObject);
    procedure BtnCloseMathcadClick(Sender: TObject);
    procedure BtnShowMathcadClick(Sender: TObject);
    procedure BtnHideMathcadClick(Sender: TObject);
    procedure BtnGetActiveWorksheetClick(Sender: TObject);
    procedure BtnOpenWorksheetClick(Sender: TObject);
    procedure BtnOpenNewWorksheetClick(Sender: TObject);
    procedure BtnSaveWorksheetClick(Sender: TObject);
    procedure BtnSaveAsWorskheetClick(Sender: TObject);
    procedure BtnSaveAsMctxClick(Sender: TObject);
    procedure BtnCloseWorksheetClick(Sender: TObject);
    procedure RGEventsClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    MPApplication : IMathcadPrimeApplication3;
    MPWorksheet : IMathcadPrimeWorksheet3;
    MPEvents : IMathcadPrimeEvents2;

    FConnectionToken: integer;
  public
    { Déclarations publiques }
  end;

var
  MainForm: TMainForm;



implementation

{$R *.dfm}

procedure TMainForm.BtnCloseMathcadClick(Sender: TObject);
begin
if (MPApplication <> nil) then
   begin
      try
      MPApplication.Quit(SaveOption_spSaveChanges);
      except
      ShowMessage('Can not close Mathcad Prime, make sure it is actually running');
      end;
   end;

end;

procedure TMainForm.BtnCloseWorksheetClick(Sender: TObject);
begin
if (MPApplication <> nil) and (MPWorksheet <> nil) then
   begin
      try
      MPWorksheet.Close(SaveOption_spDiscardChanges);
      except
      ShowMessage('Unable to close worksheet. Please check that Mathcad Prime is running');
      end;

   if (MPWorksheet <> nil) then MPWorksheet := nil;

   end;
end;

procedure TMainForm.BtnGetActiveWorksheetClick(Sender: TObject);
begin
if (MPApplication <> nil) then
  MPWorksheet := MPApplication.ActiveWorksheet as IMathcadPrimeWorksheet3;
end;

procedure TMainForm.BtnHideMathcadClick(Sender: TObject);
begin
if (MPApplication <> nil) then
   begin
      try
      MPApplication.Visible := False;
      except
      ShowMessage('Can not make Mathcad Prime invisible. Make sure Mathcad Prime is running.');
      end;
   end;
end;

procedure TMainForm.BtnOpenNewWorksheetClick(Sender: TObject);
begin
if (MPApplication <> nil) then
MPWorksheet := MPApplication.Open('') as IMathcadPrimeWorksheet3;
end;

procedure TMainForm.BtnOpenWorksheetClick(Sender: TObject);
var
OpenDialog : TOpenDialog;
begin
if (MPApplication = nil) then Exit;

  try
  OpenDialog := TOpenDialog.Create(Self);
  OpenDialog.Filter := 'MCDX files (*.mcdx)|.mcdx|All files (*.*)|*.*';
  OpenDialog.FilterIndex := 2;

  if OpenDialog.Execute then
    begin
      try
      MPWorksheet := MPApplication.Open(OpenDialog.FileName) as IMathcadPrimeWorksheet3;
      except
      ShowMessage('Couldn''t open Mathcad Prime file. Please check that Mathcad Prime is running and you have proper permissions');
      end;
    end;
  finally
  OpenDialog.Free;
  end;
end;  

procedure TMainForm.BtnSaveAsMctxClick(Sender: TObject);
var
SaveDialog : TSaveDialog;

begin
if (MPApplication = nil) then Exit;

if (MPWorksheet = nil) then
ShowMessage('There is no selected worksheet: use ''Get Active Worksheet''');

  try
  SaveDialog := TSaveDialog.Create(Self);
  SaveDialog.DefaultExt := 'mctx';
  SaveDialog.Filter := 'MCTX files (*.mctx)|.mctx';

  if SaveDialog.Execute then
    begin
      try
      MPWorksheet.SaveAs(SaveDialog.FileName);
      except
      ShowMessage('Unable to save worksheet. Please check permissions');
      end;
    end;
  finally
  SaveDialog.Free;
  end;
end;

procedure TMainForm.BtnSaveAsWorskheetClick(Sender: TObject);
var
SaveDialog : TSaveDialog;

begin
if (MPApplication = nil) then Exit;

if (MPWorksheet = nil) then 
   begin
   ShowMessage('There is no selected worksheet: use ''Get Active Worksheet''');
   Exit;
   end;

  try
  SaveDialog := TSaveDialog.Create(Self);
  SaveDialog.DefaultExt := 'mcdx';
  SaveDialog.Filter := 'MCDX files (*.mcdx)|.mcdx|All files (*.*)|*.*';

  if SaveDialog.Execute then
    begin
      try
      MPWorksheet.SaveAs(SaveDialog.FileName);
      except
      ShowMessage('Unable to save worksheet. Please check permissions');
      end;
    end;
  finally
  SaveDialog.Free;
  end;
end;

procedure TMainForm.BtnSaveWorksheetClick(Sender: TObject);
begin
if (MPApplication <> nil) and (MPWorksheet <> nil) then
   begin
      try
      MPWorksheet.Save;
      except
      ShowMessage('Can not save worksheet. If it is ''Untitled'' Use SaveAs');
      end;
   end;
end;

procedure TMainForm.BtnShowMathcadClick(Sender: TObject);
begin
if (MPApplication <> nil) then
   begin
      try
      MPApplication.Visible := True;

      // Activate Prime to bring it forward
      MPApplication.Activate;
      except
      ShowMessage('Can not Activate worksheet. Make sure Mathcad Prime is running and worksheet is loaded');
      end;
   end;
end;

procedure TMainForm.BtnStartMathcadClick(Sender: TObject);
var
isMathcadPrimeEvents2Initialized : integer;

begin
   try
   MPApplication := CoApplication.Create;
   MPApplication.Visible := True;

   InterfaceConnect(MPApplication, IMathcadPrimeApplication3, MPEvents, FConnectionToken);

   if (MPApplication <> nil) then
      begin
      isMathcadPrimeEvents2Initialized := MPApplication.InitializeEvents2(MPEvents, True);
      if (MPEvents = nil) or (isMathcadPrimeEvents2Initialized <> 0) then
         ShowMessage('Events initialization failed');
      end

   else
      begin
      raise Exception.Create('Problem to connect Mathcad Prime');
      end;

   except
   on E: Exception do
      ShowMessage('Problem to connect Mathcad Prime: ' + E.Message);
   end;

end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   try
   InterfaceDisconnect(MPApplication, IMathcadPrimeApplication3, FConnectionToken);
   MPApplication := nil;
   MPEvents := nil;
   except
   Close;
   end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
MPEvents := TMathcadPrimeEvents.Create;

end;

procedure TMainForm.RGEventsClick(Sender: TObject);
begin
case RGEvents.ItemIndex of
0: begin         //enable
      try
      if (MPEvents <> nil) and (MPEvents is TMathcadPrimeEvents) then
      TMathcadPrimeEvents(MPEvents).ShowEvents := True;
      except
      ShowMessage('Can not initial events. Make sure Mathcad Prime is running and worksheet is loaded');
      end;

   end;

1: begin        //disable
   if (MPEvents <> nil) and (MPEvents is TMathcadPrimeEvents) then
   TMathcadPrimeEvents(MPEvents).ShowEvents := False;
   end;
  
end;
end;

end.
