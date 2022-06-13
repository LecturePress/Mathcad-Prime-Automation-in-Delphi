unit MathcadComEventSink;

interface

uses
 Winapi.Windows, System.Classes, System.SysUtils, System.Variants,
 System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX, 
 Ptc_MathcadPrime_Automation_TLB, Vcl.Dialogs;



type
  TMathcadPrimeEvents = class(TInterfacedObject, IMathcadPrimeEvents2)
    ShowEvents : Boolean;

    function OnWorksheetSaved(const documentFullNameArg: WideString): HResult; stdcall;
    function OnWorksheetClosed(const documentFullNameArg: WideString;
                               const documentNameArg: WideString): HResult; stdcall;
    function OnWorksheetModified(const documentFullNameArg: WideString;
                                 const documentNameArg: WideString; isModifiedArg: WordBool): HResult; stdcall;
    function OnWorksheetRenamed(const previousFullNameArg: WideString;
                                const currentFullNameArg: WideString;
                                const previousDocNameArg: WideString;
                                const currentDocNameArg: WideString): HResult; stdcall;
    function OnWorksheetInputsOutputsSelected(const documentFullNameArg: WideString;
                                              const documentNameArg: WideString;
                                              const inputsArg: IMathcadPrimeInputs;
                                              const outputsArg: IMathcadPrimeOutputs): HResult; stdcall;
    function OnExit: HResult; stdcall;
    function OnWorksheetStatesGenerated(const documentFullNameArg: WideString;
                                        const documentNameArg: WideString;
                                        operationsArg: WorksheetOperations;
                                        const itemsStatesArg: IMathcadPrimeInputsOutputsStates;
                                        const conflictsArg: IMathcadPrimeInputsOutputsConflicts): HResult; stdcall;
    function OnWorksheetStatesGenerating(const documentFullNameArg: WideString;
                                         const documentNameArg: WideString;
                                         operationsArg: WorksheetOperations;
                                         const itemsStatesArg: IMathcadPrimeInputsOutputsStates;
                                         const conflictsArg: IMathcadPrimeInputsOutputsConflicts): HResult; stdcall;
    function OnWorksheetRequestToUpdateInputs(const documentFullNameArg: WideString;
                                              const documentNameArg: WideString;
                                              const setterArg: IMathcadPrimeValuesSetter): HResult; stdcall;

    constructor Create;

  end;




implementation



{ TMathcadPrimeEvents }

constructor TMathcadPrimeEvents.Create;
begin
ShowEvents := False;
end;

function TMathcadPrimeEvents.OnExit: HResult;
begin
if (ShowEvents = True) then
ShowMessage('Mathcad Prime is closed by User');
end;

function TMathcadPrimeEvents.OnWorksheetClosed(const documentFullNameArg,
  documentNameArg: WideString): HResult;
begin
if (ShowEvents = True) then
ShowMessage('Worksheet was closed:' + sLineBreak + sLineBreak + #09 + documentFullNameArg + sLineBreak + #09 + documentNameArg);
end;

function TMathcadPrimeEvents.OnWorksheetInputsOutputsSelected(
  const documentFullNameArg, documentNameArg: WideString;
  const inputsArg: IMathcadPrimeInputs;
  const outputsArg: IMathcadPrimeOutputs): HResult;
begin

end;

function TMathcadPrimeEvents.OnWorksheetModified(const documentFullNameArg,
  documentNameArg: WideString; isModifiedArg: WordBool): HResult;
begin
if (ShowEvents = True) then
ShowMessage('Worksheet was modified:' + sLineBreak + sLineBreak + #09 + documentFullNameArg + sLineBreak + #09 + documentNameArg);
end;

function TMathcadPrimeEvents.OnWorksheetRenamed(const previousFullNameArg,
  currentFullNameArg, previousDocNameArg,
  currentDocNameArg: WideString): HResult;
begin
if (ShowEvents = True) then
ShowMessage('Worksheet was renamed:' + sLineBreak + sLineBreak + 'From:' + #09 + previousFullNameArg + sLineBreak + sLineBreak + 'To:' + #09 + currentFullNameArg);
end;

function TMathcadPrimeEvents.OnWorksheetRequestToUpdateInputs(
  const documentFullNameArg, documentNameArg: WideString;
  const setterArg: IMathcadPrimeValuesSetter): HResult;
begin

end;

function TMathcadPrimeEvents.OnWorksheetSaved(
  const documentFullNameArg: WideString): HResult;
begin
if (ShowEvents = True) then
ShowMessage('Worksheet was saved: ' + documentFullNameArg);
end;

function TMathcadPrimeEvents.OnWorksheetStatesGenerated(const documentFullNameArg,
  documentNameArg: WideString; operationsArg: WorksheetOperations;
  const itemsStatesArg: IMathcadPrimeInputsOutputsStates;
  const conflictsArg: IMathcadPrimeInputsOutputsConflicts): HResult;
begin

end;

function TMathcadPrimeEvents.OnWorksheetStatesGenerating(
  const documentFullNameArg, documentNameArg: WideString;
  operationsArg: WorksheetOperations;
  const itemsStatesArg: IMathcadPrimeInputsOutputsStates;
  const conflictsArg: IMathcadPrimeInputsOutputsConflicts): HResult;
begin

end;

end.
