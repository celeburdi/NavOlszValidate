unit mainfrm;

interface

// https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms754649(v=vs.85)

uses
  System.Actions,
  System.Classes,
  System.ImageList,
  System.SysUtils,
  System.Variants,
  System.Win.ComObj,
  Vcl.ActnList,
  Vcl.Buttons,
  Vcl.Controls,
  Vcl.Dialogs,
  Vcl.ExtCtrls,
  Vcl.Forms,
  Vcl.ImgList,
  Vcl.StdCtrls,

  Winapi.msxml,
  Winapi.msxmlIntf,

  Winapi.ActiveX,
  Winapi.CommDlg,
  Winapi.Messages,
  Winapi.Windows;

type
  TXsdRec = record
    Uri: String;
    FileName: String;
  end;

  TfrmMain = class(TForm)
    rgVersion: TRadioGroup;
    ebFileName: TEdit;
    btnValidate: TBitBtn;
    Label1: TLabel;
    aclMain: TActionList;
    ilMain: TImageList;
    btnSelectFile: TButton;
    actValidate: TAction;
    meError: TMemo;
    procedure btnSelectFileClick(Sender: TObject);
    procedure actValidateExecute(Sender: TObject);
    procedure actValidateUpdate(Sender: TObject);
  private
    procedure DoSelectFile;
    procedure DoValidate;
    function DoCreateCache(const AXsdCollection: array of TXsdRec): IXMLDOMSchemaCollection2;
    procedure DoDumpError(AError: IXMLDOMParseError2);
    procedure DoDumpErrors(AErrors: IXMLDOMParseErrorCollection);
  public
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

const
  IID_IClassFactory: TGUID = (
    D1:$00000001;D2:$0000;D3:$0000;D4:($C0,$00,$00,$00,$00,$00,$00,$46));

  XSD_V2: array [0..1] of TXsdRec = (
   (Uri: 'http://schemas.nav.gov.hu/OSA/2.0/data';
    FileName: 'v2/invoiceData.xsd'),
   (Uri: 'http://schemas.nav.gov.hu/OSA/2.0/api';
    FileName: 'v2/invoiceApi.xsd')
  );

  XSD_V3: array [0..3] of TXsdRec = (
   (Uri: 'http://schemas.nav.gov.hu/NTCA/1.0/common';
    FileName: 'v3/common.xsd'),
   (Uri: 'http://schemas.nav.gov.hu/OSA/3.0/base';
    FileName: 'v3/invoiceBase.xsd'),
   (Uri: 'http://schemas.nav.gov.hu/OSA/3.0/data';
    FileName: 'v3/invoiceData.xsd'),
   (Uri: 'http://schemas.nav.gov.hu/OSA/3.0/api';
    FileName: 'v3/invoiceApi.xsd')
  );

var
 MsXmlLib: HMODULE;
 DllGetClassObject: function(const clsid: TGuid;const iid: TIID; out pv): HRESULT; stdcall;

procedure LoadMsXml;
begin
 if MsXmlLib=0 then
 begin
   MsXmlLib:=LoadLibrary('msxml6.dll');
   if MsXmlLib=0 then
     RaiseLastOSError;
 end;
 if not Assigned(DllGetClassObject) then
 begin
   DllGetClassObject:=GetProcAddress(MsXmlLib,'DllGetClassObject');
   if not Assigned(DllGetClassObject) then
    raise Exception.Create('LoadMsXml: DllGetClassObject=nil');
  end;
end;

function CreateXMLSchemaCache60: IXMLDOMSchemaCollection2;
var
 LClassFactory: IClassFactory;
begin
 if not Assigned(DllGetClassObject) then
    raise Exception.Create('CreateXMLSchemaCache60: DllGetClassObject=nil');
 OleCheck(DllGetClassObject(CLASS_XMLSchemaCache60,
   IID_IClassFactory,LClassFactory));
 OleCheck(LClassFactory.CreateInstance(nil,IXMLDOMSchemaCollection2,Result));
end;

function CreateDOMDocument60: IXMLDOMDocument3;
var
 LClassFactory: IClassFactory;
begin
 if not Assigned(DllGetClassObject) then
    raise Exception.Create('CreateDOMDocument60: DllGetClassObject=nil');
 OleCheck(DllGetClassObject(CLASS_DOMDocument60,
   IID_IClassFactory,LClassFactory));
 OleCheck(LClassFactory.CreateInstance(nil,IXMLDOMDocument3,Result));
end;

procedure TfrmMain.DoSelectFile;
var
  ofn: TOpenFileName;
  szFile: array[0..MAX_PATH] of Char;

begin
  FillChar(ofn, SizeOf(TOpenFileName), 0);
  szfile[0]:=#0;
  with ofn do
  begin
    lStructSize := SizeOf(TOpenFileName);
    hwndOwner := Handle;
    lpstrFile := szFile;
    nMaxFile := SizeOf(szFile);
    lpstrTitle := PChar('Select Xml File');
    lpstrInitialDir := nil;
    lpstrFilter := PChar('Xml File (*.xml)'#0'*.xml'#0'All File'#0'*.*'#0);
    Flags := 6144;
  end;
  if GetOpenFileName(ofn) then
    ebFileName.Text:=StrPas(ofn.lpstrFile);
end;


function TfrmMain.DoCreateCache(const AXsdCollection: array of TXsdRec): IXMLDOMSchemaCollection2;
var
  i: integer;
  LXsd: IXMLDOMDocument2;
  LXsdPath: String;
  LFileName: String;

begin
  LXsdPath:=ExtractFilePath(Application.ExeName)+'xsd\';
  Result:= CreateXMLSchemaCache60;
  for i:=Low(AXsdCollection) to High(AXsdCollection) do
  begin
    LXsd:=CreateDOMDocument60;
    LXsd.async:=false;
    LFileName:=LXsdPath+AXsdCollection[i].FileName;
    if not Lxsd.load(LFileName) then
      raise Exception.Create('Xsd Load error: '+LFileName);
    Result.add(AXsdCollection[i].Uri,LXsd);
  end;
end;

procedure TfrmMain.DoDumpError(AError: IXMLDOMParseError2);
begin
  meError.Lines.Add(' Line: '+IntToStr(AError.line)+' Pos: '+IntToStr(AError.linepos));
  meError.Lines.Add(' Reason: '+AError.reason);
  //meError.Lines.Add(' Location: '+AError.errorXPath);
  meError.Lines.Add(' SrcText: '+AError.srcText);
end;

procedure TfrmMain.DoDumpErrors(AErrors: IXMLDOMParseErrorCollection);
var
  i: integer;
begin
  meError.Lines.Add('Error items from the allErrors collection:');
  for i:= 0 to AErrors.length-1 do
  begin
    meError.Lines.Add('ErrorItem['+IntToStr(i)+']:');
    DoDumpError(AErrors.item[i]);
  end;
end;

procedure TfrmMain.DoValidate;
var
  xml: IXMLDOMDocument3;
  error: IXMLDOMParseError;
  cache: IXMLDOMSchemaCollection2;
  LStream: TStream;
  LIStream: IStream;

  i: integer;
  s: String;

begin
  LoadMsXml;

  xml:= CreateDOMDocument60;
  xml.async:=false;
  if rgVersion.ItemIndex=0 then
    cache:=DoCreateCache(Xsd_V2)
  else
    cache:=DoCreateCache(Xsd_V3);
  xml.schemas:=cache;
  xml.validateOnParse:=false;
  xml.setProperty('MultipleErrorMessages', true);
  xml.validateOnParse:=true;

  //xml.resolveExternals:=true;

  LStream:=nil;
  try
    LStream:=TFileStream.Create(ebFileName.Text,fmOpenRead);
    LIStream:=TStreamAdapter.Create(LStream) as IStream;

    meError.Lines.Clear;
    if not xml.load(LIStream) then
    begin
      error:=xml.parseError as IXMLDOMParseError2;
      if error.errorCode<>0 then
        DoDumpErrors((error as IXMLDOMParseError2).allErrors)
      else
        meError.Lines.Add('Xml Load error');
    end
    else
    begin
//      for i:=0 to xml.namespaces.length-1 do
//      begin
//        s:=xml.namespaces[i];
//        if s<>'http://www.w3.org/2001/XMLSchema-instance' then
//          cache.getSchema(s); // raise if not exists namespace xsd
//      end;
      error:=xml.validate as IXMLDOMParseError2;
      if error.errorCode<>0 then
        DoDumpErrors((error as IXMLDOMParseError2).allErrors)
      else
        meError.Lines.Add('Ok');
    end;
  finally
    LStream.Free;
  end;

end;

procedure TfrmMain.actValidateUpdate(Sender: TObject);
begin
  actValidate.Enabled:=Trim(ebFileName.Text)<>'';
end;

procedure TfrmMain.btnSelectFileClick(Sender: TObject);
begin
  DoSelectFile;
end;

procedure TfrmMain.actValidateExecute(Sender: TObject);
begin
  DoValidate;
end;




end.
