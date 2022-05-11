program MathcadComBasicsDemo;

uses
  Vcl.Forms,
  MathcadComApp in 'MathcadComApp.pas' {MainForm},
  MathcadComEventSink in 'MathcadComEventSink.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
