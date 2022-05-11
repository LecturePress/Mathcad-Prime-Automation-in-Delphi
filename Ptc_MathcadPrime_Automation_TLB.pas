unit Ptc_MathcadPrime_Automation_TLB;

// ************************************************************************ //
// AVERTISSEMENT
// -------
// Les types déclarés dans ce fichier ont été générés à partir de données lues
// depuis la bibliothèque de types. Si cette dernière (via une autre bibliothèque de types
// s'y référant) est explicitement ou indirectement ré-importée, ou la commande "Actualiser"
// de l'éditeur de bibliothèque de types est activée lors de la modification de la bibliothèque
// de types, le contenu de ce fichier sera régénéré et toutes les modifications
// manuellement apportées seront perdues.
// ************************************************************************ //

// $Rev: 98336 $
// Fichier généré le 06/05/2022 19:16:17 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Biblio. types : C:\Program Files\PTC\Mathcad Prime 8.0.0.0\Ptc.MathcadPrime.Automation.tlb (1)
// LIBID : {A24EB614-A183-400F-8207-1E58D61945D6}
// LCID : 0
// Fichier d'aide : 
// Chaîne d'aide : PTC Mathcad Prime COM Automation
// DepndLst : 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.4 mscorlib, (C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb)
// SYS_KIND: SYS_WIN64
// ************************************************************************ //
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans pointeur à type contrôlé.  
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, mscorlib_TLB, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;
  

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés:        
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   Interfaces DISP        : DIID_xxxx                                       
//   Interfaces Non-DISP    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions mineure et majeure de la bibliothèque de types
  Ptc_MathcadPrime_AutomationMajorVersion = 8;
  Ptc_MathcadPrime_AutomationMinorVersion = 0;

  LIBID_Ptc_MathcadPrime_Automation: TGUID = '{A24EB614-A183-400F-8207-1E58D61945D6}';

  IID_IMathcadPrimeApplication: TGUID = '{A027C8B4-F77A-4B1E-BE41-8B76FC865F25}';
  IID_IMathcadPrimeGetValueResult: TGUID = '{A6EB629A-AE39-4902-AFA8-76E1EFE632B6}';
  CLASS_GetValueResult: TGUID = '{AFFC9F46-5CB1-43F6-A477-F4B9A91BD24A}';
  IID_IMathcadPrimeEvents: TGUID = '{A170C6A4-3DEB-43A7-A5C4-9164EF85D1C2}';
  IID_IMathcadPrimeEvents2: TGUID = '{A0ECA09F-83C8-4536-B841-A33D981FBAFA}';
  IID_IMathcadPrimeInputMatrixResult: TGUID = '{A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}';
  IID_IMathcadPrimeInputResult: TGUID = '{A31F2589-4FAD-4CA4-AB33-49672071CF29}';
  IID_IMathcadPrimeInputs: TGUID = '{A317A530-6337-4309-8F1C-C155488EFE96}';
  IID_IMathcadPrimeInputsOutputsConflicts: TGUID = '{AB9C0902-3570-4697-89B4-F3887C6E978F}';
  IID_IMathcadPrimeInputsOutputsStates: TGUID = '{A72DE956-7106-4921-BFE0-90FCBE0DCA5B}';
  IID_IMathcadPrimeMatrix: TGUID = '{A17D73E0-B985-470F-ABA9-D2298799A950}';
  IID_IMathcadPrimeOutputMatrixResult: TGUID = '{A3C56254-B233-48CD-86F7-5D8CE19A7097}';
  IID_IMathcadPrimeOutputMatrixResultAs: TGUID = '{A3E6D703-E828-4EBF-BE27-8AE915CC97C7}';
  IID_IMathcadPrimeOutputResult: TGUID = '{A3740C2B-8863-420B-8AFC-086488EE7DBF}';
  IID_IMathcadPrimeOutputResultAs: TGUID = '{A3D8B0C9-9468-4650-804F-BF6D08297DE4}';
  IID_IMathcadPrimeOutputs: TGUID = '{A354B202-81A1-4EC9-A920-1876641ACF26}';
  IID_IMathcadPrimeSetValueResults: TGUID = '{AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}';
  IID_IMathcadPrimeValuesSetter: TGUID = '{A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}';
  IID_IMathcadPrimeWorksheet: TGUID = '{A17F8C1D-A683-488D-AE43-3B0860FB5B2F}';
  IID_IMathcadPrimeWorksheet2: TGUID = '{A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}';
  IID_IMathcadPrimeWorksheet3: TGUID = '{A27AD87C-6F4E-4B8A-8827-06A22ED16F35}';
  IID_IMathcadPrimeWorksheetReadonlyOptions: TGUID = '{A11258FF-2491-480C-9E3B-EBBF08AF1B72}';
  IID_IMathcadPrimeWorksheets: TGUID = '{A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}';
  IID_IMathcadPrimeApplication2: TGUID = '{A297208B-C701-4A2A-85B9-FCEC8115F0C6}';
  CLASS_InputMatrixResult: TGUID = '{A1FD7C0C-287C-4AA9-AF64-274575A06598}';
  CLASS_InputResult: TGUID = '{A119F7C8-3DCF-4ED6-9BF2-B3BDE3552D25}';
  CLASS_Inputs: TGUID = '{A2E0BC48-0B59-4703-923D-97B59069C511}';
  CLASS_InputsOutputsConflicts: TGUID = '{A2FEB23B-CDAE-4A1A-BFFE-30194E4C618D}';
  CLASS_InputsOutputsStates: TGUID = '{A31B7E02-2FD3-4FEB-907C-505F321E2E03}';
  CLASS_Matrix: TGUID = '{A258774A-1F72-431C-9464-89F831C4CCAF}';
  CLASS_OutputMatrixResult: TGUID = '{A16907E6-0B05-4D9C-A2E3-F47D9D3425A1}';
  CLASS_OutputMatrixResultAs: TGUID = '{A3F99C0E-0529-4208-AA1D-46788D80ACCA}';
  CLASS_OutputResult: TGUID = '{A1ED586B-591A-4DDF-926F-E9778CD3D6CC}';
  CLASS_OutputResultAs: TGUID = '{A3F7211D-6D33-48EC-AFC7-A5471CB0F555}';
  CLASS_Outputs: TGUID = '{A1B3240C-A124-4F3B-AF1A-7D5B5634B3C3}';
  CLASS_SetValueResults: TGUID = '{AFCA092C-04CB-44B4-8FCF-90E6F746B056}';
  CLASS_ValuesSetter: TGUID = '{A666719D-79F4-449B-A4C1-8DEECD5FA18A}';
  CLASS_Worksheet: TGUID = '{A2C7A48C-4B32-495E-AF0B-8357B115A48C}';
  CLASS_WorksheetReadonlyOptions: TGUID = '{A0568616-45DB-4B89-87D6-6769D4AEFFE4}';
  CLASS_Worksheets: TGUID = '{A31DBEE4-5053-4A2C-832F-B0161995FAE8}';
  CLASS_ApplicationObsolete: TGUID = '{A3E4E622-5CFF-4973-96EB-1E3EAB5151C8}';
  IID_IMathcadPrimeApplication3: TGUID = '{A010504B-2FE6-402E-AD27-E24A8DE5C467}';
  CLASS_Application: TGUID = '{A00E8B95-D415-433F-A04E-D298A54A7BB7}';

// *********************************************************************//
// Déclaration d'énumérations définies dans la bibliothèque de types                    
// *********************************************************************//
// Constantes pour enum SaveOption
type
  SaveOption = TOleEnum;
const
  SaveOption_spSaveChanges = $00000000;
  SaveOption_spPromptToSaveChanges = $00000001;
  SaveOption_spDiscardChanges = $00000002;

// Constantes pour enum WorksheetReadonlyOptionNames
type
  WorksheetReadonlyOptionNames = TOleEnum;
const
  WorksheetReadonlyOptionNames_FileLocationHistoryDisabled = $00000000;
  WorksheetReadonlyOptionNames_OperationsWithEnabledStateGeneration = $00000001;
  WorksheetReadonlyOptionNames_RequestToUpdateInputsEnabled = $00000002;
  WorksheetReadonlyOptionNames_CaseInsensitiveAliasComparisonEnabled = $00000003;

// Constantes pour enum MathcadPrimeEvents
type
  MathcadPrimeEvents = TOleEnum;
const
  MathcadPrimeEvents_OnExit = $00000064;
  MathcadPrimeEvents_OnWorksheetSaved = $00000000;
  MathcadPrimeEvents_OnWorksheetClosed = $00000001;
  MathcadPrimeEvents_OnWorksheetModified = $00000002;
  MathcadPrimeEvents_OnWorksheetRenamed = $00000003;
  MathcadPrimeEvents_OnWorksheetInputsOutputsSelected = $00000004;
  MathcadPrimeEvents_OnWorksheetStatesGenerated = $00000005;
  MathcadPrimeEvents_OnWorksheetStatesGenerating = $00000006;
  MathcadPrimeEvents_OnRequestToUpdateInputs = $00000007;

// Constantes pour enum WorksheetOperations
type
  WorksheetOperations = TOleEnum;
const
  WorksheetOperations_None = $00000000;
  WorksheetOperations_Save = $00000001;

// Constantes pour enum ValueResultTypes
type
  ValueResultTypes = TOleEnum;
const
  ValueResultTypes_None = $00000000;
  ValueResultTypes_Real = $00000001;
  ValueResultTypes_String = $00000002;
  ValueResultTypes_Matrix = $00000003;

type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types                     
// *********************************************************************//
  IMathcadPrimeApplication = interface;
  IMathcadPrimeApplicationDisp = dispinterface;
  IMathcadPrimeGetValueResult = interface;
  IMathcadPrimeGetValueResultDisp = dispinterface;
  IMathcadPrimeEvents = interface;
  IMathcadPrimeEvents2 = interface;
  IMathcadPrimeInputMatrixResult = interface;
  IMathcadPrimeInputMatrixResultDisp = dispinterface;
  IMathcadPrimeInputResult = interface;
  IMathcadPrimeInputResultDisp = dispinterface;
  IMathcadPrimeInputs = interface;
  IMathcadPrimeInputsDisp = dispinterface;
  IMathcadPrimeInputsOutputsConflicts = interface;
  IMathcadPrimeInputsOutputsConflictsDisp = dispinterface;
  IMathcadPrimeInputsOutputsStates = interface;
  IMathcadPrimeInputsOutputsStatesDisp = dispinterface;
  IMathcadPrimeMatrix = interface;
  IMathcadPrimeMatrixDisp = dispinterface;
  IMathcadPrimeOutputMatrixResult = interface;
  IMathcadPrimeOutputMatrixResultDisp = dispinterface;
  IMathcadPrimeOutputMatrixResultAs = interface;
  IMathcadPrimeOutputMatrixResultAsDisp = dispinterface;
  IMathcadPrimeOutputResult = interface;
  IMathcadPrimeOutputResultDisp = dispinterface;
  IMathcadPrimeOutputResultAs = interface;
  IMathcadPrimeOutputResultAsDisp = dispinterface;
  IMathcadPrimeOutputs = interface;
  IMathcadPrimeOutputsDisp = dispinterface;
  IMathcadPrimeSetValueResults = interface;
  IMathcadPrimeSetValueResultsDisp = dispinterface;
  IMathcadPrimeValuesSetter = interface;
  IMathcadPrimeValuesSetterDisp = dispinterface;
  IMathcadPrimeWorksheet = interface;
  IMathcadPrimeWorksheetDisp = dispinterface;
  IMathcadPrimeWorksheet2 = interface;
  IMathcadPrimeWorksheet2Disp = dispinterface;
  IMathcadPrimeWorksheet3 = interface;
  IMathcadPrimeWorksheet3Disp = dispinterface;
  IMathcadPrimeWorksheetReadonlyOptions = interface;
  IMathcadPrimeWorksheetReadonlyOptionsDisp = dispinterface;
  IMathcadPrimeWorksheets = interface;
  IMathcadPrimeWorksheetsDisp = dispinterface;
  IMathcadPrimeApplication2 = interface;
  IMathcadPrimeApplication2Disp = dispinterface;
  IMathcadPrimeApplication3 = interface;
  IMathcadPrimeApplication3Disp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types        
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)      
// *********************************************************************//
  GetValueResult = IMathcadPrimeGetValueResult;
  InputMatrixResult = IMathcadPrimeInputMatrixResult;
  InputResult = IMathcadPrimeInputResult;
  Inputs = IMathcadPrimeInputs;
  InputsOutputsConflicts = IMathcadPrimeInputsOutputsConflicts;
  InputsOutputsStates = IMathcadPrimeInputsOutputsStates;
  Matrix = IMathcadPrimeMatrix;
  OutputMatrixResult = IMathcadPrimeOutputMatrixResult;
  OutputMatrixResultAs = IMathcadPrimeOutputMatrixResultAs;
  OutputResult = IMathcadPrimeOutputResult;
  OutputResultAs = IMathcadPrimeOutputResultAs;
  Outputs = IMathcadPrimeOutputs;
  SetValueResults = IMathcadPrimeSetValueResults;
  ValuesSetter = IMathcadPrimeValuesSetter;
  Worksheet = IMathcadPrimeWorksheet3;
  WorksheetReadonlyOptions = IMathcadPrimeWorksheetReadonlyOptions;
  Worksheets = IMathcadPrimeWorksheets;
  ApplicationObsolete = IMathcadPrimeApplication2;
  Application = IMathcadPrimeApplication3;


// *********************************************************************//
// Interface :   IMathcadPrimeApplication
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A027C8B4-F77A-4B1E-BE41-8B76FC865F25}
// *********************************************************************//
  IMathcadPrimeApplication = interface(IDispatch)
    ['{A027C8B4-F77A-4B1E-BE41-8B76FC865F25}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(pRetVal: WordBool); safecall;
    procedure Activate; safecall;
    procedure Quit(saveOptionArg: SaveOption); safecall;
    function Get_ActiveWorksheet: IMathcadPrimeWorksheet; safecall;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; safecall;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property ActiveWorksheet: IMathcadPrimeWorksheet read Get_ActiveWorksheet;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeApplicationDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A027C8B4-F77A-4B1E-BE41-8B76FC865F25}
// *********************************************************************//
  IMathcadPrimeApplicationDisp = dispinterface
    ['{A027C8B4-F77A-4B1E-BE41-8B76FC865F25}']
    property Visible: WordBool dispid 1;
    procedure Activate; dispid 2;
    procedure Quit(saveOptionArg: SaveOption); dispid 3;
    property ActiveWorksheet: IMathcadPrimeWorksheet readonly dispid 4;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; dispid 5;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; dispid 6;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeGetValueResult
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A6EB629A-AE39-4902-AFA8-76E1EFE632B6}
// *********************************************************************//
  IMathcadPrimeGetValueResult = interface(IDispatch)
    ['{A6EB629A-AE39-4902-AFA8-76E1EFE632B6}']
    function Get_ResultType: ValueResultTypes; safecall;
    function Get_ErrorCode: Integer; safecall;
    function Get_Units: WideString; safecall;
    function Get_RealResult: Double; safecall;
    function Get_StringResult: WideString; safecall;
    function Get_MatrixResult: IMathcadPrimeMatrix; safecall;
    property ResultType: ValueResultTypes read Get_ResultType;
    property ErrorCode: Integer read Get_ErrorCode;
    property Units: WideString read Get_Units;
    property RealResult: Double read Get_RealResult;
    property StringResult: WideString read Get_StringResult;
    property MatrixResult: IMathcadPrimeMatrix read Get_MatrixResult;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeGetValueResultDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A6EB629A-AE39-4902-AFA8-76E1EFE632B6}
// *********************************************************************//
  IMathcadPrimeGetValueResultDisp = dispinterface
    ['{A6EB629A-AE39-4902-AFA8-76E1EFE632B6}']
    property ResultType: ValueResultTypes readonly dispid 1;
    property ErrorCode: Integer readonly dispid 2;
    property Units: WideString readonly dispid 3;
    property RealResult: Double readonly dispid 4;
    property StringResult: WideString readonly dispid 5;
    property MatrixResult: IMathcadPrimeMatrix readonly dispid 6;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeEvents
// Indicateurs : (256) OleAutomation
// GUID :        {A170C6A4-3DEB-43A7-A5C4-9164EF85D1C2}
// *********************************************************************//
  IMathcadPrimeEvents = interface(IUnknown)
    ['{A170C6A4-3DEB-43A7-A5C4-9164EF85D1C2}']
    function OnSelect(const inputsOnSelectArg: IMathcadPrimeInputs; 
                      const outputsOnSelectArg: IMathcadPrimeOutputs): HResult; stdcall;
    function OnSave(const documentNameArg: WideString): HResult; stdcall;
    function OnExit: HResult; stdcall;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeEvents2
// Indicateurs : (256) OleAutomation
// GUID :        {A0ECA09F-83C8-4536-B841-A33D981FBAFA}
// *********************************************************************//
  IMathcadPrimeEvents2 = interface(IUnknown)
    ['{A0ECA09F-83C8-4536-B841-A33D981FBAFA}']
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
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeInputMatrixResult
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}
// *********************************************************************//
  IMathcadPrimeInputMatrixResult = interface(IDispatch)
    ['{A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}']
    function Get_ErrorCode: Integer; safecall;
    function Get_MatrixResult: IMathcadPrimeMatrix; safecall;
    function Get_Units: WideString; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property MatrixResult: IMathcadPrimeMatrix read Get_MatrixResult;
    property Units: WideString read Get_Units;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeInputMatrixResultDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}
// *********************************************************************//
  IMathcadPrimeInputMatrixResultDisp = dispinterface
    ['{A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}']
    property ErrorCode: Integer readonly dispid 1;
    property MatrixResult: IMathcadPrimeMatrix readonly dispid 2;
    property Units: WideString readonly dispid 3;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeInputResult
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A31F2589-4FAD-4CA4-AB33-49672071CF29}
// *********************************************************************//
  IMathcadPrimeInputResult = interface(IDispatch)
    ['{A31F2589-4FAD-4CA4-AB33-49672071CF29}']
    function Get_ErrorCode: Integer; safecall;
    function Get_RealResult: Double; safecall;
    function Get_Units: WideString; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property RealResult: Double read Get_RealResult;
    property Units: WideString read Get_Units;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeInputResultDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A31F2589-4FAD-4CA4-AB33-49672071CF29}
// *********************************************************************//
  IMathcadPrimeInputResultDisp = dispinterface
    ['{A31F2589-4FAD-4CA4-AB33-49672071CF29}']
    property ErrorCode: Integer readonly dispid 1;
    property RealResult: Double readonly dispid 2;
    property Units: WideString readonly dispid 3;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeInputs
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A317A530-6337-4309-8F1C-C155488EFE96}
// *********************************************************************//
  IMathcadPrimeInputs = interface(IDispatch)
    ['{A317A530-6337-4309-8F1C-C155488EFE96}']
    function Get_Count: Integer; safecall;
    function GetAliasByIndex(indexArg: Integer): WideString; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeInputsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A317A530-6337-4309-8F1C-C155488EFE96}
// *********************************************************************//
  IMathcadPrimeInputsDisp = dispinterface
    ['{A317A530-6337-4309-8F1C-C155488EFE96}']
    property Count: Integer readonly dispid 1;
    function GetAliasByIndex(indexArg: Integer): WideString; dispid 2;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeInputsOutputsConflicts
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {AB9C0902-3570-4697-89B4-F3887C6E978F}
// *********************************************************************//
  IMathcadPrimeInputsOutputsConflicts = interface(IDispatch)
    ['{AB9C0902-3570-4697-89B4-F3887C6E978F}']
    procedure AddGeneralWarning(const warningArg: WideString); safecall;
    procedure AddGeneralError(const errorArg: WideString); safecall;
    procedure AddItemWarning(const aliasArg: WideString; const warningArg: WideString); safecall;
    procedure AddItemError(const aliasArg: WideString; const errorArg: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeInputsOutputsConflictsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {AB9C0902-3570-4697-89B4-F3887C6E978F}
// *********************************************************************//
  IMathcadPrimeInputsOutputsConflictsDisp = dispinterface
    ['{AB9C0902-3570-4697-89B4-F3887C6E978F}']
    procedure AddGeneralWarning(const warningArg: WideString); dispid 1;
    procedure AddGeneralError(const errorArg: WideString); dispid 2;
    procedure AddItemWarning(const aliasArg: WideString; const warningArg: WideString); dispid 3;
    procedure AddItemError(const aliasArg: WideString; const errorArg: WideString); dispid 4;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeInputsOutputsStates
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A72DE956-7106-4921-BFE0-90FCBE0DCA5B}
// *********************************************************************//
  IMathcadPrimeInputsOutputsStates = interface(IDispatch)
    ['{A72DE956-7106-4921-BFE0-90FCBE0DCA5B}']
    function Get_ScalarInputsCount: Integer; safecall;
    function Get_ScalarOutputsCount: Integer; safecall;
    function GetInputScalarStateByIndex(indexArg: Integer; out aliasArg: WideString; 
                                        out valueArg: Double; out unitsArg: WideString): Integer; safecall;
    function GetOutputScalarStateByIndex(indexArg: Integer; out aliasArg: WideString; 
                                         out valueArg: Double; out unitsArg: WideString): Integer; safecall;
    function GetInputScalarStateByAlias(const aliasArg: WideString; out valueArg: Double; 
                                        out unitsArg: WideString): Integer; safecall;
    function GetOutputScalarStateByAlias(const aliasArg: WideString; out valueArg: Double; 
                                         out unitsArg: WideString): Integer; safecall;
    property ScalarInputsCount: Integer read Get_ScalarInputsCount;
    property ScalarOutputsCount: Integer read Get_ScalarOutputsCount;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeInputsOutputsStatesDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A72DE956-7106-4921-BFE0-90FCBE0DCA5B}
// *********************************************************************//
  IMathcadPrimeInputsOutputsStatesDisp = dispinterface
    ['{A72DE956-7106-4921-BFE0-90FCBE0DCA5B}']
    property ScalarInputsCount: Integer readonly dispid 1;
    property ScalarOutputsCount: Integer readonly dispid 2;
    function GetInputScalarStateByIndex(indexArg: Integer; out aliasArg: WideString; 
                                        out valueArg: Double; out unitsArg: WideString): Integer; dispid 3;
    function GetOutputScalarStateByIndex(indexArg: Integer; out aliasArg: WideString; 
                                         out valueArg: Double; out unitsArg: WideString): Integer; dispid 4;
    function GetInputScalarStateByAlias(const aliasArg: WideString; out valueArg: Double; 
                                        out unitsArg: WideString): Integer; dispid 5;
    function GetOutputScalarStateByAlias(const aliasArg: WideString; out valueArg: Double; 
                                         out unitsArg: WideString): Integer; dispid 6;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeMatrix
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A17D73E0-B985-470F-ABA9-D2298799A950}
// *********************************************************************//
  IMathcadPrimeMatrix = interface(IDispatch)
    ['{A17D73E0-B985-470F-ABA9-D2298799A950}']
    function Get_Rows: Double; safecall;
    function Get_Columns: Double; safecall;
    function SetMatrixElement(rowIndexArg: Integer; columnIndexArg: Integer; valueArg: Double): Integer; safecall;
    function GetMatrixElement(rowIndexArg: Integer; columnIndexArg: Integer; out valueArg: Double): Integer; safecall;
    procedure GhostMethod_IMathcadPrimeMatrix_28_1; safecall;
    property Rows: Double read Get_Rows;
    property Columns: Double read Get_Columns;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeMatrixDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A17D73E0-B985-470F-ABA9-D2298799A950}
// *********************************************************************//
  IMathcadPrimeMatrixDisp = dispinterface
    ['{A17D73E0-B985-470F-ABA9-D2298799A950}']
    property Rows: Double readonly dispid 1;
    property Columns: Double readonly dispid 2;
    function SetMatrixElement(rowIndexArg: Integer; columnIndexArg: Integer; valueArg: Double): Integer; dispid 3;
    function GetMatrixElement(rowIndexArg: Integer; columnIndexArg: Integer; out valueArg: Double): Integer; dispid 4;
    procedure GhostMethod_IMathcadPrimeMatrix_28_1; dispid 1610743812;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeOutputMatrixResult
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3C56254-B233-48CD-86F7-5D8CE19A7097}
// *********************************************************************//
  IMathcadPrimeOutputMatrixResult = interface(IDispatch)
    ['{A3C56254-B233-48CD-86F7-5D8CE19A7097}']
    function Get_ErrorCode: Integer; safecall;
    function Get_MatrixResult: IMathcadPrimeMatrix; safecall;
    function Get_Units: WideString; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property MatrixResult: IMathcadPrimeMatrix read Get_MatrixResult;
    property Units: WideString read Get_Units;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeOutputMatrixResultDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3C56254-B233-48CD-86F7-5D8CE19A7097}
// *********************************************************************//
  IMathcadPrimeOutputMatrixResultDisp = dispinterface
    ['{A3C56254-B233-48CD-86F7-5D8CE19A7097}']
    property ErrorCode: Integer readonly dispid 1;
    property MatrixResult: IMathcadPrimeMatrix readonly dispid 2;
    property Units: WideString readonly dispid 3;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeOutputMatrixResultAs
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3E6D703-E828-4EBF-BE27-8AE915CC97C7}
// *********************************************************************//
  IMathcadPrimeOutputMatrixResultAs = interface(IDispatch)
    ['{A3E6D703-E828-4EBF-BE27-8AE915CC97C7}']
    function Get_ErrorCode: Integer; safecall;
    function Get_MatrixResult: IMathcadPrimeMatrix; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property MatrixResult: IMathcadPrimeMatrix read Get_MatrixResult;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeOutputMatrixResultAsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3E6D703-E828-4EBF-BE27-8AE915CC97C7}
// *********************************************************************//
  IMathcadPrimeOutputMatrixResultAsDisp = dispinterface
    ['{A3E6D703-E828-4EBF-BE27-8AE915CC97C7}']
    property ErrorCode: Integer readonly dispid 1;
    property MatrixResult: IMathcadPrimeMatrix readonly dispid 2;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeOutputResult
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3740C2B-8863-420B-8AFC-086488EE7DBF}
// *********************************************************************//
  IMathcadPrimeOutputResult = interface(IDispatch)
    ['{A3740C2B-8863-420B-8AFC-086488EE7DBF}']
    function Get_ErrorCode: Integer; safecall;
    function Get_RealResult: Double; safecall;
    function Get_Units: WideString; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property RealResult: Double read Get_RealResult;
    property Units: WideString read Get_Units;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeOutputResultDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3740C2B-8863-420B-8AFC-086488EE7DBF}
// *********************************************************************//
  IMathcadPrimeOutputResultDisp = dispinterface
    ['{A3740C2B-8863-420B-8AFC-086488EE7DBF}']
    property ErrorCode: Integer readonly dispid 1;
    property RealResult: Double readonly dispid 2;
    property Units: WideString readonly dispid 3;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeOutputResultAs
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3D8B0C9-9468-4650-804F-BF6D08297DE4}
// *********************************************************************//
  IMathcadPrimeOutputResultAs = interface(IDispatch)
    ['{A3D8B0C9-9468-4650-804F-BF6D08297DE4}']
    function Get_ErrorCode: Integer; safecall;
    function Get_RealResult: Double; safecall;
    property ErrorCode: Integer read Get_ErrorCode;
    property RealResult: Double read Get_RealResult;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeOutputResultAsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A3D8B0C9-9468-4650-804F-BF6D08297DE4}
// *********************************************************************//
  IMathcadPrimeOutputResultAsDisp = dispinterface
    ['{A3D8B0C9-9468-4650-804F-BF6D08297DE4}']
    property ErrorCode: Integer readonly dispid 1;
    property RealResult: Double readonly dispid 2;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeOutputs
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A354B202-81A1-4EC9-A920-1876641ACF26}
// *********************************************************************//
  IMathcadPrimeOutputs = interface(IDispatch)
    ['{A354B202-81A1-4EC9-A920-1876641ACF26}']
    function Get_Count: Integer; safecall;
    function GetAliasByIndex(indexArg: Integer): WideString; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeOutputsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A354B202-81A1-4EC9-A920-1876641ACF26}
// *********************************************************************//
  IMathcadPrimeOutputsDisp = dispinterface
    ['{A354B202-81A1-4EC9-A920-1876641ACF26}']
    property Count: Integer readonly dispid 1;
    function GetAliasByIndex(indexArg: Integer): WideString; dispid 2;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeSetValueResults
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}
// *********************************************************************//
  IMathcadPrimeSetValueResults = interface(IDispatch)
    ['{AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}']
    function Get_Count: Integer; safecall;
    function GetResultByIndex(indexArg: Integer): Integer; safecall;
    function GetResultByAlias(const aliasArg: WideString): Integer; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeSetValueResultsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}
// *********************************************************************//
  IMathcadPrimeSetValueResultsDisp = dispinterface
    ['{AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}']
    property Count: Integer readonly dispid 1;
    function GetResultByIndex(indexArg: Integer): Integer; dispid 2;
    function GetResultByAlias(const aliasArg: WideString): Integer; dispid 3;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeValuesSetter
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}
// *********************************************************************//
  IMathcadPrimeValuesSetter = interface(IDispatch)
    ['{A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}']
    procedure AddScalarValue(const aliasArg: WideString; valueArg: Double; 
                             const unitsArg: WideString); safecall;
    procedure AddMatrixValue(const aliasArg: WideString; valueArg: PSafeArray; 
                             const unitsArg: WideString); safecall;
    procedure AddStringValue(const aliasArg: WideString; const valueArg: WideString); safecall;
    procedure AddSExprValue(const aliasArg: WideString; const sexpressionArg: WideString); safecall;
    function SetValues(secondsArg: Integer): IMathcadPrimeSetValueResults; safecall;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeValuesSetterDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}
// *********************************************************************//
  IMathcadPrimeValuesSetterDisp = dispinterface
    ['{A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}']
    procedure AddScalarValue(const aliasArg: WideString; valueArg: Double; 
                             const unitsArg: WideString); dispid 1;
    procedure AddMatrixValue(const aliasArg: WideString; 
                             valueArg: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                             const unitsArg: WideString); dispid 2;
    procedure AddStringValue(const aliasArg: WideString; const valueArg: WideString); dispid 3;
    procedure AddSExprValue(const aliasArg: WideString; const sexpressionArg: WideString); dispid 4;
    function SetValues(secondsArg: Integer): IMathcadPrimeSetValueResults; dispid 5;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeWorksheet
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A17F8C1D-A683-488D-AE43-3B0860FB5B2F}
// *********************************************************************//
  IMathcadPrimeWorksheet = interface(IDispatch)
    ['{A17F8C1D-A683-488D-AE43-3B0860FB5B2F}']
    function Get_Name: WideString; safecall;
    function Get_FullName: WideString; safecall;
    function Get_IsReadOnly: WordBool; safecall;
    function Get_Modified: WordBool; safecall;
    procedure Set_Modified(pRetVal: WordBool); safecall;
    procedure SetTitle(const titleArg: WideString); safecall;
    procedure Save; safecall;
    procedure SaveAs(const fileName: WideString); safecall;
    procedure Synchronize; safecall;
    procedure PauseCalculation; safecall;
    procedure ResumeCalculation; safecall;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; safecall;
    function Get_Inputs: IMathcadPrimeInputs; safecall;
    function Get_Outputs: IMathcadPrimeOutputs; safecall;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; safecall;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; safecall;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; safecall;
    property Name: WideString read Get_Name;
    property FullName: WideString read Get_FullName;
    property IsReadOnly: WordBool read Get_IsReadOnly;
    property Modified: WordBool read Get_Modified write Set_Modified;
    property Inputs: IMathcadPrimeInputs read Get_Inputs;
    property Outputs: IMathcadPrimeOutputs read Get_Outputs;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeWorksheetDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A17F8C1D-A683-488D-AE43-3B0860FB5B2F}
// *********************************************************************//
  IMathcadPrimeWorksheetDisp = dispinterface
    ['{A17F8C1D-A683-488D-AE43-3B0860FB5B2F}']
    property Name: WideString readonly dispid 1;
    property FullName: WideString readonly dispid 2;
    property IsReadOnly: WordBool readonly dispid 3;
    property Modified: WordBool dispid 4;
    procedure SetTitle(const titleArg: WideString); dispid 5;
    procedure Save; dispid 6;
    procedure SaveAs(const fileName: WideString); dispid 7;
    procedure Synchronize; dispid 8;
    procedure PauseCalculation; dispid 9;
    procedure ResumeCalculation; dispid 10;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; dispid 11;
    property Inputs: IMathcadPrimeInputs readonly dispid 12;
    property Outputs: IMathcadPrimeOutputs readonly dispid 13;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; dispid 14;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; dispid 15;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; dispid 16;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeWorksheet2
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}
// *********************************************************************//
  IMathcadPrimeWorksheet2 = interface(IDispatch)
    ['{A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}']
    function Get_Name: WideString; safecall;
    function Get_FullName: WideString; safecall;
    function Get_IsReadOnly: WordBool; safecall;
    function Get_Modified: WordBool; safecall;
    procedure Set_Modified(pRetVal: WordBool); safecall;
    procedure SetTitle(const titleArg: WideString); safecall;
    procedure Save; safecall;
    procedure SaveAs(const fileName: WideString); safecall;
    procedure Synchronize; safecall;
    procedure PauseCalculation; safecall;
    procedure ResumeCalculation; safecall;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; safecall;
    function Get_Inputs: IMathcadPrimeInputs; safecall;
    function Get_Outputs: IMathcadPrimeOutputs; safecall;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; safecall;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; safecall;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; safecall;
    function CreateMatrix(rowsArg: Integer; columnsArg: Integer): IMathcadPrimeMatrix; safecall;
    function SetMatrixValue(const aliasArg: WideString; const valueArg: IMathcadPrimeMatrix; 
                            const unitsArg: WideString): Integer; safecall;
    function InputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeInputMatrixResult; safecall;
    function OutputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeOutputMatrixResult; safecall;
    function OutputGetMatrixValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputMatrixResultAs; safecall;
    function IsOpen: WordBool; safecall;
    procedure Activate; safecall;
    procedure Close(saveOptionArg: SaveOption); safecall;
    property Name: WideString read Get_Name;
    property FullName: WideString read Get_FullName;
    property IsReadOnly: WordBool read Get_IsReadOnly;
    property Modified: WordBool read Get_Modified write Set_Modified;
    property Inputs: IMathcadPrimeInputs read Get_Inputs;
    property Outputs: IMathcadPrimeOutputs read Get_Outputs;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeWorksheet2Disp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}
// *********************************************************************//
  IMathcadPrimeWorksheet2Disp = dispinterface
    ['{A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}']
    property Name: WideString readonly dispid 1;
    property FullName: WideString readonly dispid 2;
    property IsReadOnly: WordBool readonly dispid 3;
    property Modified: WordBool dispid 4;
    procedure SetTitle(const titleArg: WideString); dispid 5;
    procedure Save; dispid 6;
    procedure SaveAs(const fileName: WideString); dispid 7;
    procedure Synchronize; dispid 8;
    procedure PauseCalculation; dispid 9;
    procedure ResumeCalculation; dispid 10;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; dispid 11;
    property Inputs: IMathcadPrimeInputs readonly dispid 12;
    property Outputs: IMathcadPrimeOutputs readonly dispid 13;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; dispid 14;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; dispid 15;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; dispid 16;
    function CreateMatrix(rowsArg: Integer; columnsArg: Integer): IMathcadPrimeMatrix; dispid 17;
    function SetMatrixValue(const aliasArg: WideString; const valueArg: IMathcadPrimeMatrix; 
                            const unitsArg: WideString): Integer; dispid 18;
    function InputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeInputMatrixResult; dispid 19;
    function OutputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeOutputMatrixResult; dispid 20;
    function OutputGetMatrixValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputMatrixResultAs; dispid 21;
    function IsOpen: WordBool; dispid 22;
    procedure Activate; dispid 23;
    procedure Close(saveOptionArg: SaveOption); dispid 24;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeWorksheet3
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A27AD87C-6F4E-4B8A-8827-06A22ED16F35}
// *********************************************************************//
  IMathcadPrimeWorksheet3 = interface(IDispatch)
    ['{A27AD87C-6F4E-4B8A-8827-06A22ED16F35}']
    function Get_Name: WideString; safecall;
    function Get_FullName: WideString; safecall;
    function Get_IsReadOnly: WordBool; safecall;
    function Get_Modified: WordBool; safecall;
    procedure Set_Modified(pRetVal: WordBool); safecall;
    procedure SetTitle(const titleArg: WideString); safecall;
    procedure Save; safecall;
    procedure SaveAs(const fileName: WideString); safecall;
    procedure Synchronize; safecall;
    procedure PauseCalculation; safecall;
    procedure ResumeCalculation; safecall;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; safecall;
    function Get_Inputs: IMathcadPrimeInputs; safecall;
    function Get_Outputs: IMathcadPrimeOutputs; safecall;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; safecall;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; safecall;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; safecall;
    function CreateMatrix(rowsArg: Integer; columnsArg: Integer): IMathcadPrimeMatrix; safecall;
    function SetMatrixValue(const aliasArg: WideString; const valueArg: IMathcadPrimeMatrix; 
                            const unitsArg: WideString): Integer; safecall;
    function InputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeInputMatrixResult; safecall;
    function OutputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeOutputMatrixResult; safecall;
    function OutputGetMatrixValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputMatrixResultAs; safecall;
    function IsOpen: WordBool; safecall;
    procedure Activate; safecall;
    procedure Close(saveOptionArg: SaveOption); safecall;
    function GetTitle: WideString; safecall;
    function Get_WorksheetTabIcon: WideString; safecall;
    procedure Set_WorksheetTabIcon(const pRetVal: WideString); safecall;
    function Get_WorksheetTabName: WideString; safecall;
    procedure Set_WorksheetTabName(const pRetVal: WideString); safecall;
    function Get_WorksheetClosingPrompt: WideString; safecall;
    procedure Set_WorksheetClosingPrompt(const pRetVal: WideString); safecall;
    function Get_WorksheetDisplayedFilePath: WideString; safecall;
    procedure Set_WorksheetDisplayedFilePath(const pRetVal: WideString); safecall;
    function Get_WorksheetWorkingDirectory: WideString; safecall;
    procedure Set_WorksheetWorkingDirectory(const pRetVal: WideString); safecall;
    function GetWorksheetReadOnlyOptionValue(optionNameArg: WorksheetReadonlyOptionNames): OleVariant; safecall;
    function CreateValuesSetter: IMathcadPrimeValuesSetter; safecall;
    function SetStringValue(const aliasArg: WideString; const valueArg: WideString): Integer; safecall;
    function InputGetStringValue(const aliasArg: WideString): WideString; safecall;
    function OutputGetStringValue(const aliasArg: WideString): WideString; safecall;
    function SetSExprValue(const aliasArg: WideString; const sexpressionArg: WideString): Integer; safecall;
    function InputGetSExprValue(const aliasArg: WideString): WideString; safecall;
    function InputGetValue(const aliasArg: WideString): IMathcadPrimeGetValueResult; safecall;
    function OutputGetValue(const aliasArg: WideString): IMathcadPrimeGetValueResult; safecall;
    function Get_DefaultCalculationTimeout: Integer; safecall;
    procedure Set_DefaultCalculationTimeout(pRetVal: Integer); safecall;
    property Name: WideString read Get_Name;
    property FullName: WideString read Get_FullName;
    property IsReadOnly: WordBool read Get_IsReadOnly;
    property Modified: WordBool read Get_Modified write Set_Modified;
    property Inputs: IMathcadPrimeInputs read Get_Inputs;
    property Outputs: IMathcadPrimeOutputs read Get_Outputs;
    property WorksheetTabIcon: WideString read Get_WorksheetTabIcon write Set_WorksheetTabIcon;
    property WorksheetTabName: WideString read Get_WorksheetTabName write Set_WorksheetTabName;
    property WorksheetClosingPrompt: WideString read Get_WorksheetClosingPrompt write Set_WorksheetClosingPrompt;
    property WorksheetDisplayedFilePath: WideString read Get_WorksheetDisplayedFilePath write Set_WorksheetDisplayedFilePath;
    property WorksheetWorkingDirectory: WideString read Get_WorksheetWorkingDirectory write Set_WorksheetWorkingDirectory;
    property DefaultCalculationTimeout: Integer read Get_DefaultCalculationTimeout write Set_DefaultCalculationTimeout;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeWorksheet3Disp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A27AD87C-6F4E-4B8A-8827-06A22ED16F35}
// *********************************************************************//
  IMathcadPrimeWorksheet3Disp = dispinterface
    ['{A27AD87C-6F4E-4B8A-8827-06A22ED16F35}']
    property Name: WideString readonly dispid 1;
    property FullName: WideString readonly dispid 2;
    property IsReadOnly: WordBool readonly dispid 3;
    property Modified: WordBool dispid 4;
    procedure SetTitle(const titleArg: WideString); dispid 5;
    procedure Save; dispid 6;
    procedure SaveAs(const fileName: WideString); dispid 7;
    procedure Synchronize; dispid 8;
    procedure PauseCalculation; dispid 9;
    procedure ResumeCalculation; dispid 10;
    function SetRealValue(const aliasArg: WideString; valueArg: Double; const unitsArg: WideString): Integer; dispid 11;
    property Inputs: IMathcadPrimeInputs readonly dispid 12;
    property Outputs: IMathcadPrimeOutputs readonly dispid 13;
    function InputGetRealValue(const aliasArg: WideString): IMathcadPrimeInputResult; dispid 14;
    function OutputGetRealValue(const aliasArg: WideString): IMathcadPrimeOutputResult; dispid 15;
    function OutputGetRealValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputResultAs; dispid 16;
    function CreateMatrix(rowsArg: Integer; columnsArg: Integer): IMathcadPrimeMatrix; dispid 17;
    function SetMatrixValue(const aliasArg: WideString; const valueArg: IMathcadPrimeMatrix; 
                            const unitsArg: WideString): Integer; dispid 18;
    function InputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeInputMatrixResult; dispid 19;
    function OutputGetMatrixValue(const aliasArg: WideString): IMathcadPrimeOutputMatrixResult; dispid 20;
    function OutputGetMatrixValueAs(const aliasArg: WideString; const unitsArg: WideString): IMathcadPrimeOutputMatrixResultAs; dispid 21;
    function IsOpen: WordBool; dispid 22;
    procedure Activate; dispid 23;
    procedure Close(saveOptionArg: SaveOption); dispid 24;
    function GetTitle: WideString; dispid 25;
    property WorksheetTabIcon: WideString dispid 26;
    property WorksheetTabName: WideString dispid 27;
    property WorksheetClosingPrompt: WideString dispid 28;
    property WorksheetDisplayedFilePath: WideString dispid 29;
    property WorksheetWorkingDirectory: WideString dispid 30;
    function GetWorksheetReadOnlyOptionValue(optionNameArg: WorksheetReadonlyOptionNames): OleVariant; dispid 31;
    function CreateValuesSetter: IMathcadPrimeValuesSetter; dispid 32;
    function SetStringValue(const aliasArg: WideString; const valueArg: WideString): Integer; dispid 33;
    function InputGetStringValue(const aliasArg: WideString): WideString; dispid 34;
    function OutputGetStringValue(const aliasArg: WideString): WideString; dispid 35;
    function SetSExprValue(const aliasArg: WideString; const sexpressionArg: WideString): Integer; dispid 36;
    function InputGetSExprValue(const aliasArg: WideString): WideString; dispid 37;
    function InputGetValue(const aliasArg: WideString): IMathcadPrimeGetValueResult; dispid 38;
    function OutputGetValue(const aliasArg: WideString): IMathcadPrimeGetValueResult; dispid 39;
    property DefaultCalculationTimeout: Integer dispid 40;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeWorksheetReadonlyOptions
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A11258FF-2491-480C-9E3B-EBBF08AF1B72}
// *********************************************************************//
  IMathcadPrimeWorksheetReadonlyOptions = interface(IDispatch)
    ['{A11258FF-2491-480C-9E3B-EBBF08AF1B72}']
    function SetOptionValue(optionNameArg: WorksheetReadonlyOptionNames; optionValueArg: OleVariant): Integer; safecall;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeWorksheetReadonlyOptionsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A11258FF-2491-480C-9E3B-EBBF08AF1B72}
// *********************************************************************//
  IMathcadPrimeWorksheetReadonlyOptionsDisp = dispinterface
    ['{A11258FF-2491-480C-9E3B-EBBF08AF1B72}']
    function SetOptionValue(optionNameArg: WorksheetReadonlyOptionNames; optionValueArg: OleVariant): Integer; dispid 1;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeWorksheets
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}
// *********************************************************************//
  IMathcadPrimeWorksheets = interface(IDispatch)
    ['{A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}']
    function Get_Count: Integer; safecall;
    function Item(indexArg: Integer): IMathcadPrimeWorksheet2; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeWorksheetsDisp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}
// *********************************************************************//
  IMathcadPrimeWorksheetsDisp = dispinterface
    ['{A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}']
    property Count: Integer readonly dispid 1;
    function Item(indexArg: Integer): IMathcadPrimeWorksheet2; dispid 2;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeApplication2
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A297208B-C701-4A2A-85B9-FCEC8115F0C6}
// *********************************************************************//
  IMathcadPrimeApplication2 = interface(IDispatch)
    ['{A297208B-C701-4A2A-85B9-FCEC8115F0C6}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(pRetVal: WordBool); safecall;
    procedure Activate; safecall;
    procedure Quit(saveOptionArg: SaveOption); safecall;
    function Get_ActiveWorksheet: IMathcadPrimeWorksheet; safecall;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; safecall;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; safecall;
    function Get_Worksheets: IMathcadPrimeWorksheets; safecall;
    procedure CloseAll(saveOptionArg: SaveOption); safecall;
    function GetVersion: WideString; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property ActiveWorksheet: IMathcadPrimeWorksheet read Get_ActiveWorksheet;
    property Worksheets: IMathcadPrimeWorksheets read Get_Worksheets;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeApplication2Disp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A297208B-C701-4A2A-85B9-FCEC8115F0C6}
// *********************************************************************//
  IMathcadPrimeApplication2Disp = dispinterface
    ['{A297208B-C701-4A2A-85B9-FCEC8115F0C6}']
    property Visible: WordBool dispid 1;
    procedure Activate; dispid 2;
    procedure Quit(saveOptionArg: SaveOption); dispid 3;
    property ActiveWorksheet: IMathcadPrimeWorksheet readonly dispid 4;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; dispid 5;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; dispid 6;
    property Worksheets: IMathcadPrimeWorksheets readonly dispid 7;
    procedure CloseAll(saveOptionArg: SaveOption); dispid 8;
    function GetVersion: WideString; dispid 9;
  end;

// *********************************************************************//
// Interface :   IMathcadPrimeApplication3
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A010504B-2FE6-402E-AD27-E24A8DE5C467}
// *********************************************************************//
  IMathcadPrimeApplication3 = interface(IDispatch)
    ['{A010504B-2FE6-402E-AD27-E24A8DE5C467}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(pRetVal: WordBool); safecall;
    procedure Activate; safecall;
    procedure Quit(saveOptionArg: SaveOption); safecall;
    function Get_ActiveWorksheet: IMathcadPrimeWorksheet; safecall;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; safecall;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; safecall;
    function Get_Worksheets: IMathcadPrimeWorksheets; safecall;
    procedure CloseAll(saveOptionArg: SaveOption); safecall;
    function GetVersion: WideString; safecall;
    function InitializeEvents2(const eventsArg: IMathcadPrimeEvents2; subscribeAllArg: WordBool): Integer; safecall;
    function SubscribeEvent(primeEventArg: MathcadPrimeEvents): Integer; safecall;
    function UnsubscribeEvent(primeEventArg: MathcadPrimeEvents): Integer; safecall;
    function CreateWorksheetReadonlyOptions: IMathcadPrimeWorksheetReadonlyOptions; safecall;
    function OpenEx(const documentPathArg: WideString; 
                    const optionsArg: IMathcadPrimeWorksheetReadonlyOptions): IMathcadPrimeWorksheet3; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property ActiveWorksheet: IMathcadPrimeWorksheet read Get_ActiveWorksheet;
    property Worksheets: IMathcadPrimeWorksheets read Get_Worksheets;
  end;

// *********************************************************************//
// DispIntf :    IMathcadPrimeApplication3Disp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID :        {A010504B-2FE6-402E-AD27-E24A8DE5C467}
// *********************************************************************//
  IMathcadPrimeApplication3Disp = dispinterface
    ['{A010504B-2FE6-402E-AD27-E24A8DE5C467}']
    property Visible: WordBool dispid 1;
    procedure Activate; dispid 2;
    procedure Quit(saveOptionArg: SaveOption); dispid 3;
    property ActiveWorksheet: IMathcadPrimeWorksheet readonly dispid 4;
    function Open(const documentPathArg: WideString): IMathcadPrimeWorksheet; dispid 5;
    function InitializeEvents(const eventsArg: IMathcadPrimeEvents): Integer; dispid 6;
    property Worksheets: IMathcadPrimeWorksheets readonly dispid 7;
    procedure CloseAll(saveOptionArg: SaveOption); dispid 8;
    function GetVersion: WideString; dispid 9;
    function InitializeEvents2(const eventsArg: IMathcadPrimeEvents2; subscribeAllArg: WordBool): Integer; dispid 10;
    function SubscribeEvent(primeEventArg: MathcadPrimeEvents): Integer; dispid 11;
    function UnsubscribeEvent(primeEventArg: MathcadPrimeEvents): Integer; dispid 12;
    function CreateWorksheetReadonlyOptions: IMathcadPrimeWorksheetReadonlyOptions; dispid 13;
    function OpenEx(const documentPathArg: WideString; 
                    const optionsArg: IMathcadPrimeWorksheetReadonlyOptions): IMathcadPrimeWorksheet3; dispid 14;
  end;

// *********************************************************************//
// La classe CoGetValueResult fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeGetValueResult exposée
// par la CoClasse GetValueResult. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoGetValueResult = class
    class function Create: IMathcadPrimeGetValueResult;
    class function CreateRemote(const MachineName: string): IMathcadPrimeGetValueResult;
  end;

// *********************************************************************//
// La classe CoInputMatrixResult fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeInputMatrixResult exposée
// par la CoClasse InputMatrixResult. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoInputMatrixResult = class
    class function Create: IMathcadPrimeInputMatrixResult;
    class function CreateRemote(const MachineName: string): IMathcadPrimeInputMatrixResult;
  end;

// *********************************************************************//
// La classe CoInputResult fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeInputResult exposée
// par la CoClasse InputResult. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoInputResult = class
    class function Create: IMathcadPrimeInputResult;
    class function CreateRemote(const MachineName: string): IMathcadPrimeInputResult;
  end;

// *********************************************************************//
// La classe CoInputs fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeInputs exposée
// par la CoClasse Inputs. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoInputs = class
    class function Create: IMathcadPrimeInputs;
    class function CreateRemote(const MachineName: string): IMathcadPrimeInputs;
  end;

// *********************************************************************//
// La classe CoInputsOutputsConflicts fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeInputsOutputsConflicts exposée
// par la CoClasse InputsOutputsConflicts. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoInputsOutputsConflicts = class
    class function Create: IMathcadPrimeInputsOutputsConflicts;
    class function CreateRemote(const MachineName: string): IMathcadPrimeInputsOutputsConflicts;
  end;

// *********************************************************************//
// La classe CoInputsOutputsStates fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeInputsOutputsStates exposée
// par la CoClasse InputsOutputsStates. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoInputsOutputsStates = class
    class function Create: IMathcadPrimeInputsOutputsStates;
    class function CreateRemote(const MachineName: string): IMathcadPrimeInputsOutputsStates;
  end;

// *********************************************************************//
// La classe CoMatrix fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeMatrix exposée
// par la CoClasse Matrix. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoMatrix = class
    class function Create: IMathcadPrimeMatrix;
    class function CreateRemote(const MachineName: string): IMathcadPrimeMatrix;
  end;

// *********************************************************************//
// La classe CoOutputMatrixResult fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeOutputMatrixResult exposée
// par la CoClasse OutputMatrixResult. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoOutputMatrixResult = class
    class function Create: IMathcadPrimeOutputMatrixResult;
    class function CreateRemote(const MachineName: string): IMathcadPrimeOutputMatrixResult;
  end;

// *********************************************************************//
// La classe CoOutputMatrixResultAs fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeOutputMatrixResultAs exposée
// par la CoClasse OutputMatrixResultAs. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoOutputMatrixResultAs = class
    class function Create: IMathcadPrimeOutputMatrixResultAs;
    class function CreateRemote(const MachineName: string): IMathcadPrimeOutputMatrixResultAs;
  end;

// *********************************************************************//
// La classe CoOutputResult fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeOutputResult exposée
// par la CoClasse OutputResult. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoOutputResult = class
    class function Create: IMathcadPrimeOutputResult;
    class function CreateRemote(const MachineName: string): IMathcadPrimeOutputResult;
  end;

// *********************************************************************//
// La classe CoOutputResultAs fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeOutputResultAs exposée
// par la CoClasse OutputResultAs. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoOutputResultAs = class
    class function Create: IMathcadPrimeOutputResultAs;
    class function CreateRemote(const MachineName: string): IMathcadPrimeOutputResultAs;
  end;

// *********************************************************************//
// La classe CoOutputs fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeOutputs exposée
// par la CoClasse Outputs. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoOutputs = class
    class function Create: IMathcadPrimeOutputs;
    class function CreateRemote(const MachineName: string): IMathcadPrimeOutputs;
  end;

// *********************************************************************//
// La classe CoSetValueResults fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeSetValueResults exposée
// par la CoClasse SetValueResults. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoSetValueResults = class
    class function Create: IMathcadPrimeSetValueResults;
    class function CreateRemote(const MachineName: string): IMathcadPrimeSetValueResults;
  end;

// *********************************************************************//
// La classe CoValuesSetter fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeValuesSetter exposée
// par la CoClasse ValuesSetter. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoValuesSetter = class
    class function Create: IMathcadPrimeValuesSetter;
    class function CreateRemote(const MachineName: string): IMathcadPrimeValuesSetter;
  end;

// *********************************************************************//
// La classe CoWorksheet fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeWorksheet3 exposée
// par la CoClasse Worksheet. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoWorksheet = class
    class function Create: IMathcadPrimeWorksheet3;
    class function CreateRemote(const MachineName: string): IMathcadPrimeWorksheet3;
  end;

// *********************************************************************//
// La classe CoWorksheetReadonlyOptions fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeWorksheetReadonlyOptions exposée
// par la CoClasse WorksheetReadonlyOptions. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoWorksheetReadonlyOptions = class
    class function Create: IMathcadPrimeWorksheetReadonlyOptions;
    class function CreateRemote(const MachineName: string): IMathcadPrimeWorksheetReadonlyOptions;
  end;

// *********************************************************************//
// La classe CoWorksheets fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeWorksheets exposée
// par la CoClasse Worksheets. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoWorksheets = class
    class function Create: IMathcadPrimeWorksheets;
    class function CreateRemote(const MachineName: string): IMathcadPrimeWorksheets;
  end;

// *********************************************************************//
// La classe CoApplicationObsolete fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeApplication2 exposée
// par la CoClasse ApplicationObsolete. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoApplicationObsolete = class
    class function Create: IMathcadPrimeApplication2;
    class function CreateRemote(const MachineName: string): IMathcadPrimeApplication2;
  end;

// *********************************************************************//
// La classe CoApplication fournit une méthode Create et CreateRemote pour
// créer des instances de l'interface par défaut IMathcadPrimeApplication3 exposée
// par la CoClasse Application. Les fonctions sont destinées à être utilisées par
// les clients désirant automatiser les objets CoClasse exposés par
// le serveur de cette bibliothèque de types.
// *********************************************************************//
  CoApplication = class
    class function Create: IMathcadPrimeApplication3;
    class function CreateRemote(const MachineName: string): IMathcadPrimeApplication3;
  end;

implementation

uses System.Win.ComObj;

class function CoGetValueResult.Create: IMathcadPrimeGetValueResult;
begin
  Result := CreateComObject(CLASS_GetValueResult) as IMathcadPrimeGetValueResult;
end;

class function CoGetValueResult.CreateRemote(const MachineName: string): IMathcadPrimeGetValueResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_GetValueResult) as IMathcadPrimeGetValueResult;
end;

class function CoInputMatrixResult.Create: IMathcadPrimeInputMatrixResult;
begin
  Result := CreateComObject(CLASS_InputMatrixResult) as IMathcadPrimeInputMatrixResult;
end;

class function CoInputMatrixResult.CreateRemote(const MachineName: string): IMathcadPrimeInputMatrixResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InputMatrixResult) as IMathcadPrimeInputMatrixResult;
end;

class function CoInputResult.Create: IMathcadPrimeInputResult;
begin
  Result := CreateComObject(CLASS_InputResult) as IMathcadPrimeInputResult;
end;

class function CoInputResult.CreateRemote(const MachineName: string): IMathcadPrimeInputResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InputResult) as IMathcadPrimeInputResult;
end;

class function CoInputs.Create: IMathcadPrimeInputs;
begin
  Result := CreateComObject(CLASS_Inputs) as IMathcadPrimeInputs;
end;

class function CoInputs.CreateRemote(const MachineName: string): IMathcadPrimeInputs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Inputs) as IMathcadPrimeInputs;
end;

class function CoInputsOutputsConflicts.Create: IMathcadPrimeInputsOutputsConflicts;
begin
  Result := CreateComObject(CLASS_InputsOutputsConflicts) as IMathcadPrimeInputsOutputsConflicts;
end;

class function CoInputsOutputsConflicts.CreateRemote(const MachineName: string): IMathcadPrimeInputsOutputsConflicts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InputsOutputsConflicts) as IMathcadPrimeInputsOutputsConflicts;
end;

class function CoInputsOutputsStates.Create: IMathcadPrimeInputsOutputsStates;
begin
  Result := CreateComObject(CLASS_InputsOutputsStates) as IMathcadPrimeInputsOutputsStates;
end;

class function CoInputsOutputsStates.CreateRemote(const MachineName: string): IMathcadPrimeInputsOutputsStates;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_InputsOutputsStates) as IMathcadPrimeInputsOutputsStates;
end;

class function CoMatrix.Create: IMathcadPrimeMatrix;
begin
  Result := CreateComObject(CLASS_Matrix) as IMathcadPrimeMatrix;
end;

class function CoMatrix.CreateRemote(const MachineName: string): IMathcadPrimeMatrix;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Matrix) as IMathcadPrimeMatrix;
end;

class function CoOutputMatrixResult.Create: IMathcadPrimeOutputMatrixResult;
begin
  Result := CreateComObject(CLASS_OutputMatrixResult) as IMathcadPrimeOutputMatrixResult;
end;

class function CoOutputMatrixResult.CreateRemote(const MachineName: string): IMathcadPrimeOutputMatrixResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OutputMatrixResult) as IMathcadPrimeOutputMatrixResult;
end;

class function CoOutputMatrixResultAs.Create: IMathcadPrimeOutputMatrixResultAs;
begin
  Result := CreateComObject(CLASS_OutputMatrixResultAs) as IMathcadPrimeOutputMatrixResultAs;
end;

class function CoOutputMatrixResultAs.CreateRemote(const MachineName: string): IMathcadPrimeOutputMatrixResultAs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OutputMatrixResultAs) as IMathcadPrimeOutputMatrixResultAs;
end;

class function CoOutputResult.Create: IMathcadPrimeOutputResult;
begin
  Result := CreateComObject(CLASS_OutputResult) as IMathcadPrimeOutputResult;
end;

class function CoOutputResult.CreateRemote(const MachineName: string): IMathcadPrimeOutputResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OutputResult) as IMathcadPrimeOutputResult;
end;

class function CoOutputResultAs.Create: IMathcadPrimeOutputResultAs;
begin
  Result := CreateComObject(CLASS_OutputResultAs) as IMathcadPrimeOutputResultAs;
end;

class function CoOutputResultAs.CreateRemote(const MachineName: string): IMathcadPrimeOutputResultAs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OutputResultAs) as IMathcadPrimeOutputResultAs;
end;

class function CoOutputs.Create: IMathcadPrimeOutputs;
begin
  Result := CreateComObject(CLASS_Outputs) as IMathcadPrimeOutputs;
end;

class function CoOutputs.CreateRemote(const MachineName: string): IMathcadPrimeOutputs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Outputs) as IMathcadPrimeOutputs;
end;

class function CoSetValueResults.Create: IMathcadPrimeSetValueResults;
begin
  Result := CreateComObject(CLASS_SetValueResults) as IMathcadPrimeSetValueResults;
end;

class function CoSetValueResults.CreateRemote(const MachineName: string): IMathcadPrimeSetValueResults;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SetValueResults) as IMathcadPrimeSetValueResults;
end;

class function CoValuesSetter.Create: IMathcadPrimeValuesSetter;
begin
  Result := CreateComObject(CLASS_ValuesSetter) as IMathcadPrimeValuesSetter;
end;

class function CoValuesSetter.CreateRemote(const MachineName: string): IMathcadPrimeValuesSetter;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ValuesSetter) as IMathcadPrimeValuesSetter;
end;

class function CoWorksheet.Create: IMathcadPrimeWorksheet3;
begin
  Result := CreateComObject(CLASS_Worksheet) as IMathcadPrimeWorksheet3;
end;

class function CoWorksheet.CreateRemote(const MachineName: string): IMathcadPrimeWorksheet3;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Worksheet) as IMathcadPrimeWorksheet3;
end;

class function CoWorksheetReadonlyOptions.Create: IMathcadPrimeWorksheetReadonlyOptions;
begin
  Result := CreateComObject(CLASS_WorksheetReadonlyOptions) as IMathcadPrimeWorksheetReadonlyOptions;
end;

class function CoWorksheetReadonlyOptions.CreateRemote(const MachineName: string): IMathcadPrimeWorksheetReadonlyOptions;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_WorksheetReadonlyOptions) as IMathcadPrimeWorksheetReadonlyOptions;
end;

class function CoWorksheets.Create: IMathcadPrimeWorksheets;
begin
  Result := CreateComObject(CLASS_Worksheets) as IMathcadPrimeWorksheets;
end;

class function CoWorksheets.CreateRemote(const MachineName: string): IMathcadPrimeWorksheets;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Worksheets) as IMathcadPrimeWorksheets;
end;

class function CoApplicationObsolete.Create: IMathcadPrimeApplication2;
begin
  Result := CreateComObject(CLASS_ApplicationObsolete) as IMathcadPrimeApplication2;
end;

class function CoApplicationObsolete.CreateRemote(const MachineName: string): IMathcadPrimeApplication2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ApplicationObsolete) as IMathcadPrimeApplication2;
end;

class function CoApplication.Create: IMathcadPrimeApplication3;
begin
  Result := CreateComObject(CLASS_Application) as IMathcadPrimeApplication3;
end;

class function CoApplication.CreateRemote(const MachineName: string): IMathcadPrimeApplication3;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Application) as IMathcadPrimeApplication3;
end;

end.
