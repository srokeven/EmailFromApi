unit uSendEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.ExtCtrls, System.StrUtils, System.Math,
  IdBaseComponent,
  IdComponent,
  IdIOHandler,
  IdIOHandlerSocket,
  IdIOHandlerStack,
  IdSSL,
  IdSSLOpenSSL,
  IdMessage,
  IdText,
  IdSMTP,
  IdExplicitTLSClientServerBase,
  IdMessageBuilder, Vcl.OleCtrls, SHDocVw;

type
  TEmailInfo = record
    eiEmailAddress: string; //Email de destinatario
    eiEmailAddressBccl: string; //Email para copia
    eiSubject: string; //Assunto do email
    eiText: string; //Conteudo da mensagem em array
    eiFilePath: array of string; //Lista com os caminhos dos arquivos de anexo no computador
    eiInformation: string; //Informação pra sair no rodapé do email
    eiTitle: string; //Titulo para aparecer na pagina do email
end;

type
  TEmailSettings = record
    esSSL: integer;
    esSSLSegure: boolean;
    esTLS: integer;
    esTLSSecure: boolean;
    esFromEmail: string;
    esFromName: string;
    esSMTPUser: string;
    esSMTPPassword: string;
    esSMTPHost: string;
    esSMTPPort: string;
    esModelFile: string; //Arquivo com modelo do email
    esModelText: string; //Se a propriedade aModelFile não for preenchida vai verificar a aModelText, que é o modelo em texto puro
end;

type
  TComponentMod = class
    public
      function ExtrairTextoAPartirDePosicao(aTexto, aTrecho: string; aDirecao: integer): string;
      procedure ArredondarCantos(aComponent: TWinControl; Y:String);
      procedure TratarBotaoPanelPersonalizado(var aPanel: TPanel; aCorEntrada,
        aCorEntradaFonte, aCorPadrao, aCorPadraoFonte: integer);
      procedure HandlePanelPersonalizadoOnMouseEnter(Sender: TObject);
      procedure HandlePanelPersonalizadoOnMouseLeave(Sender: TObject);
      procedure TratarBotaoPanelAzul(var aPanel: TPanel);
      procedure TratarBotaoPanelCinza(var aPanel: TPanel);
      procedure TratarBotaoPanelVermelho(var aPanel: TPanel; aTipoCorPadrao: integer);
      procedure TratarBotaoPanelVerde(var aPanel: TPanel);
      procedure HandlePanelAzulOnMouseEnter(Sender: TObject);
      procedure HandlePanelAzulOnMouseLeave(Sender: TObject);
      procedure HandlePanelVermelhoPadraoOnMouseEnter(Sender: TObject);
      procedure HandlePanelVermelhoPadraoOnMouseLeave(Sender: TObject);
      procedure HandlePanelVermelhoOnMouseEnter(Sender: TObject);
      procedure HandlePanelVermelhoOnMouseLeave(Sender: TObject);
      procedure HandlePanelVermelhoBtnFaceOnMouseLeave(Sender: TObject);
      procedure HandlePanelCinzaOnMouseEnter(Sender: TObject);
      procedure HandlePanelCinzaOnMouseLeave(Sender: TObject);
      procedure HandlePanelVerdeOnMouseEnter(Sender: TObject);
      procedure HandlePanelVerdeOnMouseLeave(Sender: TObject);
  end;
  TfmSendEmail = class(TForm)
    pcMain: TPageControl;
    tsSend: TTabSheet;
    tsSettings: TTabSheet;
    edDestinatario: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    edDestinatarioCopia: TEdit;
    Label3: TLabel;
    pnlSearchFile: TPanel;
    mmTextoEmail: TMemo;
    Label4: TLabel;
    pnlSend: TPanel;
    pnlCancel: TPanel;
    lbStatus: TLabel;
    edSubject: TEdit;
    Label11: TLabel;
    OpenFile: TFileOpenDialog;
    tsHTML: TTabSheet;
    mmHTML: TMemo;
    lbModeloHTML: TLabel;
    Panel1: TPanel;
    pnlCarregarHTML: TPanel;
    pnlVisualizarHTML: TPanel;
    cbListaAnexos: TComboBox;
    pnlRemoveFile: TPanel;
    tsVisualizarEmail: TTabSheet;
    wbHtml: TWebBrowser;
    GroupBox1: TGroupBox;
    edPorta: TEdit;
    Label14: TLabel;
    gbConfiguracaoEmail: TGroupBox;
    chkUsarTLS: TCheckBox;
    chkUsarSSL: TCheckBox;
    edServidorSMTP: TEdit;
    edPortaSMTP: TEdit;
    Label10: TLabel;
    Label9: TLabel;
    edUserSMTP: TEdit;
    Label5: TLabel;
    edEmailRemetente: TEdit;
    Label12: TLabel;
    cbSsl: TComboBox;
    Label7: TLabel;
    Label8: TLabel;
    cbUsarTLS: TComboBox;
    Label13: TLabel;
    edNomeDoRemetente: TEdit;
    Label6: TLabel;
    edSenhaSMTP: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure pnlSendClick(Sender: TObject);
    procedure pnlSearchFileClick(Sender: TObject);
    procedure pnlCancelClick(Sender: TObject);
    procedure pnlCarregarHTMLClick(Sender: TObject);
    procedure pnlRemoveFileClick(Sender: TObject);
    procedure pnlVisualizarHTMLClick(Sender: TObject);
  private
    FComponentMod: TComponentMod;
    fFrom, fFromName, fHost, fUserName, fPassword, fPort, fAddAddress: string;
    fSetTLS, fSetSSL: boolean;
    fSSLVersion: TIdSSLVersion;
    fListaAnexos: TStringList;
    //Variaveis da formatação do corpo do email
    fInformation, fFileTextWarn, fText, fSubject, fTitle: string;

    fIdSMTP               : TIdSMTP;
    fIdMsg                : TIdMessage;
    fIdSSLIOHandlerSocket : TIdSSLIOHandlerSocketOpenSSL;
    fIdText               : TIdText;
    procedure AdicionarAnexo(aFilePath, aFileName: string);
    procedure RemoverAnexo(aItem: integer);
    procedure GetStatus(aStatus: string);
    procedure LoadEmailConfig;
    function Validations: boolean;
    function SendEmail: boolean;
    function SetTextParamsHTML(aHTMLLine: string): string;
    procedure ShowError(const aMsg: string; aTitulo: String = 'Erro');
    procedure ShowInformation(const aMsg: string; aTitulo: String = 'Informação');
    procedure ShowWarning(const aMsg: string; aTitulo: String = 'Aviso');
    function ShowQuestion(const aMsg: string; aPreSelect: integer = 0): boolean;
  public
    procedure PreparaEnvioMulti(aEmailSettings: TEmailSettings);
    procedure AdicionaDadosEmail(aEmailInfo: TEmailInfo);
    function ValidaConfigs: boolean;
    procedure PraparaConexao;
    function AbreConexao: boolean;
    procedure AdicionaDestinatarios(aEmail: string);
    function EnviaEmail: boolean;
    procedure LiberarMemoria;
    class function Email(aEmailInfo: TEmailInfo; aEmailSettings: TEmailSettings): boolean; overload;
    class function Email(aEmailInfo: TEmailInfo; aEmailSettings: TEmailSettings; aSilent: boolean): boolean; overload;
  end;

var
  fmSendEmail: TfmSendEmail;

const
  EXTRAIR_PRE_TRECHO = 0;
  EXTRAIR_TRECHO = 1;
  EXTRAIR_POS_TRECHO = 2;

implementation

{$R *.dfm}

function TfmSendEmail.AbreConexao: boolean;
begin
  Result := False;
  fIdSMTP               := TIdSMTP.Create(Self);
  fIdMsg                := TIdMessage.Create(Self);
  fIdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  //Configura os parâmetros necessários para SSL
  case cbSsl.ItemIndex of
    0: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    1: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1_2;
    2: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv3;
    3: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1_1;
    4: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1;
    5: fIdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv2;
  end;
  fIdSSLIOHandlerSocket.SSLOptions.Mode   := sslmClient;
  fIdSMTP.IOHandler := fIdSSLIOHandlerSocket;
  case cbUsarTLS.ItemIndex of
    0: fIdSMTP.UseTLS    := utUseImplicitTLS;
    1: fIdSMTP.UseTLS    := utUseExplicitTLS;
    2: fIdSMTP.UseTLS    := utNoTLSSupport;
    3: fIdSMTP.UseTLS    := utUseRequireTLS;
  end;
  fIdSSLIOHandlerSocket.SSLOptions.Mode   := sslmClient;
  //Variável referente a mensagem
  fIdMsg.CharSet      := 'utf-8';
  fIdMsg.ContentType  := 'text/plain; charset=utf-8';
  fIdMsg.Encoding     := meMIME;
  fIdMsg.From.Name    := fFromName;
  fIdMsg.From.Address := fFrom;
  fIdMsg.Priority     := mpNormal;
  fIdMsg.Subject      := edSubject.Text;

  //Add Destinatário principal
  fIdMsg.Recipients.Add;
  fIdMsg.Recipients.EMailAddresses := edDestinatario.Text;

  //Prepara o Servidor
  fIdSMTP.AuthType  := satDefault;
  fIdSMTP.Host      := fHost;
  fIdSMTP.AuthType  := satDefault;
  fIdSMTP.Port      := StrToIntDef(fPort,0);
  fIdSMTP.Username  := fUserName;
  fIdSMTP.Password  := fPassword;

  try
    //Conecta e Autentica
    fIdSMTP.Connect;
    fIdSMTP.Authenticate;
    //Se a conexão foi bem sucedida, envia a mensagem
    if fIdSMTP.Connected then
    begin
      Result := True;
    end;
  except
    on E: Exception do
    begin
      ShowError('Não foi possivel enviar o email: '+E.Message);
    end;
  end;
end;

procedure TfmSendEmail.AdicionaDadosEmail(aEmailInfo: TEmailInfo);
begin
  edDestinatario.Text := aEmailInfo.eiEmailAddress;
  edSubject.Text := aEmailInfo.eiSubject;

  fListaAnexos := TStringList.Create;
  fmSendEmail.mmTextoEmail.Lines.Text := aEmailInfo.eiText;
  fmSendEmail.fInformation := 'Up2 Soluções, Iguatu-CE, 2022 <br/> '+
                              'Sistemas de automação comercial para atacado e varejo <br/> '+
                              aEmailInfo.eiInformation;
  fmSendEmail.fTitle := aEmailInfo.eiTitle;
end;

procedure TfmSendEmail.AdicionaDestinatarios(aEmail: string);
begin
  fIdMsg.BccList.Add.Address := aEmail;
end;

procedure TfmSendEmail.AdicionarAnexo(aFilePath, aFileName: string);
var
  I: Integer;
begin
  fListaAnexos.Add(aFilePath);
  if aFileName.IsEmpty then
    cbListaAnexos.Items.Add(ExtractFileName(aFilePath))
  else cbListaAnexos.Items.Add(aFileName);
  if not (fListaAnexos.Count = cbListaAnexos.Items.Count) then
  begin
    cbListaAnexos.Items.Clear;
    for I := 0 to fListaAnexos.Count -1 do
      cbListaAnexos.Items.Add(ExtractFileName(fListaAnexos[I]));
  end;
  cbListaAnexos.ItemIndex := cbListaAnexos.Items.Count -1;
end;

class function TfmSendEmail.Email(aEmailInfo: TEmailInfo; aEmailSettings: TEmailSettings): boolean;
var
  I: Integer;
begin
  try
    Application.CreateForm(TfmSendEmail, fmSendEmail);
    fmSendEmail.edDestinatario.Text := aEmailInfo.eiEmailAddress;
    fmSendEmail.edDestinatarioCopia.Text := aEmailInfo.eiEmailAddressBccl;
    fmSendEmail.edSubject.Text := aEmailInfo.eiSubject;

    fmSendEmail.fListaAnexos := TStringList.Create;
    if Length(aEmailInfo.eiFilePath) > 0 then
      for I := Low(aEmailInfo.eiFilePath) to High(aEmailInfo.eiFilePath) do
        fmSendEmail.AdicionarAnexo(aEmailInfo.eiFilePath[I], '');

    fmSendEmail.mmTextoEmail.Lines.Text := aEmailInfo.eiText;
    fmSendEmail.fInformation := 'Up2 Soluções, Iguatu-CE, 2022 <br/> '+
                                'Sistemas de automação comercial para atacado e varejo <br/> '+
                                aEmailInfo.eiInformation;
    fmSendEmail.fTitle := aEmailInfo.eiTitle;

    fmSendEmail.cbSsl.ItemIndex := aEmailSettings.esSSL;
    fmSendEmail.chkUsarSSL.Checked := aEmailSettings.esSSLSegure;
    fmSendEmail.cbUsarTLS.ItemIndex := aEmailSettings.esTLS;
    fmSendEmail.chkUsarTLS.Checked := aEmailSettings.esTLSSecure;
    fmSendEmail.edEmailRemetente.Text := aEmailSettings.esFromEmail;
    fmSendEmail.edNomeDoRemetente.Text := aEmailSettings.esFromName;
    fmSendEmail.edUserSMTP.Text := aEmailSettings.esSMTPUser;
    fmSendEmail.edSenhaSMTP.Text := aEmailSettings.esSMTPPassword;
    fmSendEmail.edServidorSMTP.Text := aEmailSettings.esSMTPHost;
    fmSendEmail.edPortaSMTP.Text := aEmailSettings.esSMTPPort;

    if not (aEmailSettings.esModelFile.IsEmpty) then
    begin
      if FileExists(aEmailSettings.esModelFile) then
        fmSendEmail.mmHTML.Lines.LoadFromFile(aEmailSettings.esModelFile);
    end else
    if not (aEmailSettings.esModelText.IsEmpty) then
      fmSendEmail.mmHTML.Lines.Text := aEmailSettings.esModelText;

    Result := fmSendEmail.ShowModal = mrOk;
  finally
    fmSendEmail.fListaAnexos.Free;
    FreeAndNil(fmSendEmail);
  end;
end;

class function TfmSendEmail.Email(aEmailInfo: TEmailInfo; aEmailSettings: TEmailSettings;
  aSilent: boolean): boolean;
var
  I: Integer;
begin
  try
    Application.CreateForm(TfmSendEmail, fmSendEmail);
    fmSendEmail.edDestinatario.Text := aEmailInfo.eiEmailAddress;
    fmSendEmail.edDestinatarioCopia.Text := aEmailInfo.eiEmailAddressBccl;
    fmSendEmail.edSubject.Text := aEmailInfo.eiSubject;

    fmSendEmail.fListaAnexos := TStringList.Create;
    if Length(aEmailInfo.eiFilePath) > 0 then
      for I := Low(aEmailInfo.eiFilePath) to High(aEmailInfo.eiFilePath) do
        fmSendEmail.AdicionarAnexo(aEmailInfo.eiFilePath[I], '');

    fmSendEmail.mmTextoEmail.Lines.Text := aEmailInfo.eiText;
    fmSendEmail.fInformation := 'Up2 Soluções, Iguatu-CE, 2022 <br/> '+
                                'Sistemas de automação comercial para atacado e varejo <br/> '+
                                aEmailInfo.eiInformation;
    fmSendEmail.fTitle := aEmailInfo.eiTitle;

    fmSendEmail.cbSsl.ItemIndex := aEmailSettings.esSSL;
    fmSendEmail.chkUsarSSL.Checked := aEmailSettings.esSSLSegure;
    fmSendEmail.cbUsarTLS.ItemIndex := aEmailSettings.esTLS;
    fmSendEmail.chkUsarTLS.Checked := aEmailSettings.esTLSSecure;
    fmSendEmail.edEmailRemetente.Text := aEmailSettings.esFromEmail;
    fmSendEmail.edNomeDoRemetente.Text := aEmailSettings.esFromName;
    fmSendEmail.edUserSMTP.Text := aEmailSettings.esSMTPUser;
    fmSendEmail.edSenhaSMTP.Text := aEmailSettings.esSMTPPassword;
    fmSendEmail.edServidorSMTP.Text := aEmailSettings.esSMTPHost;
    fmSendEmail.edPortaSMTP.Text := aEmailSettings.esSMTPPort;

    if not (aEmailSettings.esModelFile.IsEmpty) then
    begin
      if FileExists(aEmailSettings.esModelFile) then
        fmSendEmail.mmHTML.Lines.LoadFromFile(aEmailSettings.esModelFile);
    end else
    if not (aEmailSettings.esModelText.IsEmpty) then
      fmSendEmail.mmHTML.Lines.Text := aEmailSettings.esModelText;

    if fmSendEmail.Validations then
    begin
      fmSendEmail.LoadEmailConfig;
      if fmSendEmail.SendEmail then
      begin
        fmSendEmail.ShowInformation('Enviado com sucesso');
      end else
      begin
        fmSendEmail.ShowWarning('Email não enviado');
        Result := fmSendEmail.ShowModal = mrOk;
      end;
    end else
      Result := fmSendEmail.ShowModal = mrOk;
  finally
    fmSendEmail.fListaAnexos.Free;
    FreeAndNil(fmSendEmail);
  end;
end;

function TfmSendEmail.EnviaEmail: boolean;
var
  I: integer;
begin
  fIdText := TIdText.Create(fIdMsg.MessageParts);
  fIdText.ContentType := 'text/html; text/plain; charset=utf-8';
  fIdText.CharSet := 'utf-8';
  if not (mmHTML.Lines.Text.IsEmpty) then
  begin
    for I := 0 to mmHTML.Lines.Count -1 do
      fIdText.Body.Add(SetTextParamsHTML(mmHTML.Lines[I]));
  end else
  begin
    for I := 0 to mmTextoEmail.Lines.Count -1 do
      fIdText.Body.Add(UTF8Decode(mmTextoEmail.Lines[I])+' <br/>');
  end;
  fIdSMTP.Send(fIdMsg);
  Result := True;
end;

procedure TfmSendEmail.FormCreate(Sender: TObject);
begin
  FComponentMod.ArredondarCantos(pnlSend, '4');
  FComponentMod.ArredondarCantos(pnlCancel, '4');
  FComponentMod.ArredondarCantos(pnlCarregarHTML, '4');
  FComponentMod.ArredondarCantos(pnlVisualizarHTML, '4');
  FComponentMod.TratarBotaoPanelAzul(pnlSend);
  FComponentMod.TratarBotaoPanelVermelho(pnlCancel, 0);
  FComponentMod.TratarBotaoPanelCinza(pnlCarregarHTML);
  FComponentMod.TratarBotaoPanelCinza(pnlVisualizarHTML);
  GetStatus('Preencha os dados do destinatário...');
  pcMain.ActivePage := tsSend;
end;

procedure TfmSendEmail.GetStatus(aStatus: string);
begin
  lbStatus.Caption := aStatus;
  Application.ProcessMessages;
end;

procedure TfmSendEmail.LiberarMemoria;
begin
  UnLoadOpenSSLLibrary;

  FreeAndNil(fIdMsg);
  FreeAndNil(fIdSMTP);
end;

procedure TfmSendEmail.LoadEmailConfig;
var
  I: Integer;
begin
  GetStatus('Carregando configuração.');

  fFrom      := edEmailRemetente.Text;
  fFromName  := edNomeDoRemetente.Text;
  fHost      := edServidorSMTP.Text;
  fUserName  := edUserSMTP.Text;
  fPassword  := edSenhaSMTP.Text;
  fPort      := edPortaSMTP.Text;

  fAddAddress:= edDestinatarioCopia.Text;
  fSetTLS    := chkUsarTLS.Checked;
  fSetSSL    := chkUsarSSL.Checked;

  fSubject := edSubject.Text;
  if not (fListaAnexos.Count = 0) then
    fFileTextWarn := 'Existem arquivos anexados para esse email';
end;

procedure TfmSendEmail.pnlCancelClick(Sender: TObject);
begin
  ModalResult := mrClose;
end;

procedure TfmSendEmail.pnlCarregarHTMLClick(Sender: TObject);
var
  SL: TStringList;
  I: Integer;
begin
  OpenFile.DefaultExtension := '*.html';
  OpenFile.Options := [];
  OpenFile.Files.Clear;
  if OpenFile.Execute then
  begin
    mmHTML.Clear;
    SL := TStringList.Create;
    SL.LoadFromFile(OpenFile.FileName);
    for I := 0 to SL.Count - 1 do
      mmHTML.Lines.Add(SL[I]);
    SL.Free;
  end;
end;

procedure TfmSendEmail.pnlRemoveFileClick(Sender: TObject);
begin
  if cbListaAnexos.Items.Count > 0 then
  begin
    if ShowQuestion('Deseja remover o anexo selecionado?') then
      RemoverAnexo(cbListaAnexos.ItemIndex);
  end else ShowWarning('Nenhum arquivo anexado');
end;

procedure TfmSendEmail.pnlSearchFileClick(Sender: TObject);
var
  I: integer;
begin
  OpenFile.DefaultExtension := '';
  OpenFile.Files.Clear;
  OpenFile.Options := [fdoAllowMultiSelect];
  if OpenFile.Execute then
  begin
    for I := 0 to OpenFile.Files.Count - 1 do
      AdicionarAnexo(OpenFile.Files[I], '');
  end;
end;

procedure TfmSendEmail.pnlSendClick(Sender: TObject);
begin
  if Validations then
  begin
    LoadEmailConfig;
    if SendEmail then
    begin
      ShowInformation('Enviado com sucesso');
      ModalResult := mrOk;
    end;
  end;
end;

procedure TfmSendEmail.pnlVisualizarHTMLClick(Sender: TObject);
var
  I: integer;
  arq: TextFile;
  vPreviewHTML: string;
begin
  if mmHTML.Lines.Text.IsEmpty then
  begin
    ShowWarning('Nenhum conteúdo para exibir');
    Exit;
  end;
  vPreviewHTML := ExtractFilePath(ParamStr(0))+ 'PreviewEmail.html';
  if FileExists(vPreviewHTML) then
    DeleteFile(vPreviewHTML);
  AssignFile(arq, vPreviewHTML);
  Rewrite(arq);

  for I := 0 to mmHTML.Lines.Count -1 do
    Writeln(arq, SetTextParamsHTML(mmHTML.Lines[I]));
  CloseFile(arq);
  wbHtml.Navigate(vPreviewHTML);
  pcMain.ActivePage := tsVisualizarEmail;
end;

procedure TfmSendEmail.PraparaConexao;
begin
  LoadEmailConfig;
end;

procedure TfmSendEmail.PreparaEnvioMulti(aEmailSettings: TEmailSettings);
begin
  fmSendEmail.cbSsl.ItemIndex := aEmailSettings.esSSL;
  fmSendEmail.chkUsarSSL.Checked := aEmailSettings.esSSLSegure;
  fmSendEmail.cbUsarTLS.ItemIndex := aEmailSettings.esTLS;
  fmSendEmail.chkUsarTLS.Checked := aEmailSettings.esTLSSecure;
  fmSendEmail.edEmailRemetente.Text := aEmailSettings.esFromEmail;
  fmSendEmail.edNomeDoRemetente.Text := aEmailSettings.esFromName;
  fmSendEmail.edUserSMTP.Text := aEmailSettings.esSMTPUser;
  fmSendEmail.edSenhaSMTP.Text := aEmailSettings.esSMTPPassword;
  fmSendEmail.edServidorSMTP.Text := aEmailSettings.esSMTPHost;
  fmSendEmail.edPortaSMTP.Text := aEmailSettings.esSMTPPort;

  if not (aEmailSettings.esModelFile.IsEmpty) then
  begin
    if FileExists(aEmailSettings.esModelFile) then
      fmSendEmail.mmHTML.Lines.LoadFromFile(aEmailSettings.esModelFile);
  end else
  if not (aEmailSettings.esModelText.IsEmpty) then
    fmSendEmail.mmHTML.Lines.Text := aEmailSettings.esModelText;
end;

procedure TfmSendEmail.RemoverAnexo(aItem: integer);
var
  vListaAuxiliar: TStringList;
  I: Integer;
begin
  vListaAuxiliar := TStringList.Create;
  for I := 0 to fListaAnexos.Count -1 do
    if not (aItem = I) then
      vListaAuxiliar.Add(fListaAnexos[I]);
  fListaAnexos.Clear;
  for I := 0 to vListaAuxiliar.Count -1 do
    fListaAnexos.Add(vListaAuxiliar[I]);
  cbListaAnexos.Items.Delete(aItem);
  cbListaAnexos.ItemIndex := cbListaAnexos.Items.Count -1;
  vListaAuxiliar.Free;
end;

function TfmSendEmail.SendEmail: boolean;
var
  IdSSLIOHandlerSocket : TIdSSLIOHandlerSocketOpenSSL;
  idMsg                : TIdMessage;
  IdText               : TIdText;
  idSMTP               : TIdSMTP;
  MsgBuilder           : TIdMessageBuilderPlain;
  I: integer;
  procedure InitializeISO88591(var VHeaderEncoding: Char;
  var VCharSet: string);
  begin
    VCharSet := 'ISO-8859-1';
  end;
begin
  Result := False;
  GetStatus('Preparando envio. Por favor aguarde...');
  try
    IdSMTP               := TIdSMTP.Create(Self);
    idMsg                := TIdMessage.Create(Self);
    IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
    //Configura os parâmetros necessários para SSL
    case cbSsl.ItemIndex of
      0: IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
      1: IdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1_2;
      2: IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv3;
      3: IdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1_1;
      4: IdSSLIOHandlerSocket.SSLOptions.Method := sslvTLSv1;
      5: IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv2;
    end;
    IdSSLIOHandlerSocket.SSLOptions.Mode   := sslmClient;
    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    case cbUsarTLS.ItemIndex of
      0: IdSMTP.UseTLS    := utUseImplicitTLS;
      1: IdSMTP.UseTLS    := utUseExplicitTLS;
      2: IdSMTP.UseTLS    := utNoTLSSupport;
      3: IdSMTP.UseTLS    := utUseRequireTLS;
    end;
    IdSSLIOHandlerSocket.SSLOptions.Mode   := sslmClient;
    //Variável referente a mensagem
    idMsg.CharSet      := 'utf-8';
    idMsg.ContentType  := 'text/plain; charset=utf-8';
    idMsg.Encoding     := meMIME;
    idMsg.From.Name    := fFromName;
    idMsg.From.Address := fFrom;
    idMsg.Priority     := mpNormal;
    idMsg.Subject      := edSubject.Text;

    //Add Destinatário(s)
    idMsg.Recipients.Add;
    idMsg.Recipients.EMailAddresses := edDestinatario.Text;
    idMsg.BccList.EMailAddresses    := edDestinatarioCopia.Text;

    //Verifica se tem anexo, comportamento estranho nos caracteres quando email não tem anexo
    if not (fListaAnexos.Count = 0) then
    begin
      with TIdMessageBuilderHtml.Create do
      try
        PlainTextCharset := 'utf-8';
        if not (mmHTML.Lines.Text.IsEmpty) then
        begin
          for I := 0 to mmHTML.Lines.Count -1 do
            Html.Add(SetTextParamsHTML(mmHTML.Lines[I]));
        end else
        begin
          for I := 0 to mmTextoEmail.Lines.Count -1 do
            Html.Add(UTF8Decode(mmTextoEmail.Lines[I])+' <br/>');
        end;
        //Html.Text := 'HTML goes here';
        if not (fListaAnexos.Count = 0) then
        begin
          for I := 0 to fListaAnexos.Count -1 do
          begin
            if FileExists(fListaAnexos[I]) then
             HtmlFiles.Add(fListaAnexos[I]);
          end;
        end;
        FillMessage(IdMsg);
      finally
        Free;
      end;
    end
    else
    begin
      idText := TIdText.Create(idMsg.MessageParts);
      idText.ContentType := 'text/html; text/plain; charset=utf-8';
      idText.CharSet := 'utf-8';
      if not (mmHTML.Lines.Text.IsEmpty) then
      begin
        for I := 0 to mmHTML.Lines.Count -1 do
          idText.Body.Add(SetTextParamsHTML(mmHTML.Lines[I]));
      end else
      begin
        for I := 0 to mmTextoEmail.Lines.Count -1 do
          idText.Body.Add(UTF8Decode(mmTextoEmail.Lines[I])+' <br/>');
      end;
    end;

    //Prepara o Servidor
    IdSMTP.AuthType  := satDefault;
    IdSMTP.Host      := fHost;
    IdSMTP.AuthType  := satDefault;
    IdSMTP.Port      := StrToIntDef(fPort,0);
    IdSMTP.Username  := fUserName;
    IdSMTP.Password  := fPassword;

    try
      //Conecta e Autentica
      GetStatus('Conectando com o serviço. Por favor aguarde...');
      IdSMTP.Connect;
      GetStatus('Autenticando credenciais...');
      IdSMTP.Authenticate;
      //Se a conexão foi bem sucedida, envia a mensagem
      if IdSMTP.Connected then
      begin
        try
          GetStatus('Enviando email. Por favor aguarde...');

          IdSMTP.Send(idMsg);
          Result := True;
          GetStatus('Email enviado!');
        except on E:Exception do
          begin
            GetStatus('Email não enviado!');
            ShowError('Erro ao tentar enviar email: ' + E.Message);
          end;
        end;
      end;

      //Depois de tudo pronto, desconecta do servidor SMTP
      if IdSMTP.Connected then
        IdSMTP.Disconnect;
    except
      on E: Exception do
      begin
        ShowError('Não foi possivel enviar o email: '+E.Message);
        GetStatus('Email não enviado');
      end;
    end;
  finally
    UnLoadOpenSSLLibrary;

    FreeAndNil(idMsg);
    FreeAndNil(IdSSLIOHandlerSocket);
    FreeAndNil(idSMTP);
  end;
end;

function TfmSendEmail.SetTextParamsHTML(aHTMLLine: string): string;
var
  vOutputText: string;
  vTexto: string;
  I: integer;
begin
  vOutputText := aHTMLLine;
  vOutputText := vOutputText.Replace('$TITULO', fTitle, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ TITULO', fTitle, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ASSUNTO', fSubject, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ ASSUNTO', fSubject, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$CLIENTE', fFromName, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ CLIENTE', fFromName, [rfReplaceAll, rfIgnoreCase]);
  for I := 0 to mmTextoEmail.Lines.Count -1 do
    vTexto := vTexto + mmTextoEmail.Lines[I]+' <br/>';
  vOutputText := vOutputText.Replace('$TEXTO', vTexto, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ TEXTO', vTexto, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ANEXO', fFileTextWarn, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ ANEXO', fFileTextWarn, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$INFO', fInformation, [rfReplaceAll, rfIgnoreCase]);
  vOutputText := vOutputText.Replace('$ INFO', fInformation, [rfReplaceAll, rfIgnoreCase]);
  Result := vOutputText;
end;

procedure TfmSendEmail.ShowError(const aMsg: string; aTitulo:String = 'Erro');
begin
  Application.MessageBox(PWideChar(aMsg), PWideChar(aTitulo), MB_ICONERROR + MB_OK);
end;

procedure TfmSendEmail.ShowInformation(const aMsg: string; aTitulo:String = 'Informação');
begin
  Application.MessageBox(PWideChar(aMsg), PWideChar(aTitulo), MB_ICONINFORMATION + MB_OK);
end;

procedure TfmSendEmail.ShowWarning(const aMsg: string; aTitulo:String = 'Aviso');
begin
  Application.MessageBox(PWideChar(aMsg), PWideChar(aTitulo), MB_ICONWARNING + MB_OK);
end;

function TfmSendEmail.ShowQuestion(const aMsg: string; aPreSelect: integer = 0): boolean;
begin
  Result := (Application.MessageBox(PWideChar(aMsg), PWideChar('Envio de email'),
  MB_ICONQUESTION + MB_YESNO + IfThen(aPreSelect = 0, MB_DEFBUTTON2, MB_DEFBUTTON1)) = IDYES);
end;

function TfmSendEmail.ValidaConfigs: boolean;
begin
  Result := Validations;
end;

function TfmSendEmail.Validations: boolean;
begin
  GetStatus('Validando informações necessárias');
  if edDestinatario.Text = EmptyStr then
  begin
    ShowWarning('Informe o email do destinatário');
    pcMain.ActivePage := tsSend;
    edDestinatario.SetFocus;
    Exit(False);
  end;
  if edSubject.Text = EmptyStr then
  begin
    ShowWarning('Informe o assunto do email');
    pcMain.ActivePage := tsSend;
    edSubject.SetFocus;
    Exit(False);
  end;
  if mmTextoEmail.Lines.Text = EmptyStr then
  begin
    ShowWarning('Informe o conteúdo do email');
    pcMain.ActivePage := tsSend;
    mmTextoEmail.SetFocus;
    Exit(False);
  end;

  if edEmailRemetente.Text = EmptyStr then
  begin
    ShowWarning('Informe o email do remetente');
    pcMain.ActivePage := tsSettings;
    edEmailRemetente.SetFocus;
    Exit(False);
  end;
  if edNomeDoRemetente.Text = EmptyStr then
  begin
    ShowWarning('Informe o nome do remetente');
    pcMain.ActivePage := tsSettings;
    edNomeDoRemetente.SetFocus;
    Exit(False);
  end;
  if edUserSMTP.Text = EmptyStr then
  begin
    ShowWarning('Informe o usuário do serviço de SMTP');
    pcMain.ActivePage := tsSettings;
    edUserSMTP.SetFocus;
    Exit(False);
  end;
  if edSenhaSMTP.Text = EmptyStr then
  begin
    ShowWarning('Informe a senha do serviço de SMTP');
    pcMain.ActivePage := tsSettings;
    edSenhaSMTP.SetFocus;
    Exit(False);
  end;
  if edServidorSMTP.Text = EmptyStr then
  begin
    ShowWarning('Informe o servidor do serviço de SMTP');
    pcMain.ActivePage := tsSettings;
    edServidorSMTP.SetFocus;
    Exit(False);
  end;
  if edPortaSMTP.Text = EmptyStr then
  begin
    ShowWarning('Informe a porta do serviço de SMTP');
    pcMain.ActivePage := tsSettings;
    edPortaSMTP.SetFocus;
    Exit(False);
  end;
  Result := True;
end;

{ TComponentMod }

function TComponentMod.ExtrairTextoAPartirDePosicao(aTexto, aTrecho: string; aDirecao: integer): string;
begin
  Result := '';
  if not (aTexto.IsEmpty) then
  begin
    if pos(aTrecho, aTexto) > 0 then
    begin
      case aDirecao of
        EXTRAIR_PRE_TRECHO: Result := copy(aTexto, 1, pos(aTrecho, aTexto)-1);
        EXTRAIR_TRECHO: Result := copy(aTexto, 1, pos(aTrecho, aTexto)) + copy(aTexto, pos(aTrecho, aTexto)+Length(aTrecho), Length(aTexto));
        EXTRAIR_POS_TRECHO: Result := copy(aTexto, pos(aTrecho, aTexto)+Length(aTrecho), Length(aTexto));
      end;
    end;
  end;
end;

procedure TComponentMod.ArredondarCantos(aComponent: TWinControl; Y: String);
var
  BX: TRect;
  mdo: HRGN;
begin
  with aComponent do
  begin
    BX := ClientRect;
    mdo := CreateRoundRectRgn(BX.Left, BX.Top, BX.Right,
    BX.Bottom, StrToInt(Y), StrToInt(Y)) ;
    Perform(EM_GETRECT, 0, lParam(@BX)) ;
    InflateRect(BX, - 4, - 4) ;
    Perform(EM_SETRECTNP, 0, lParam(@BX)) ;
    SetWindowRgn(Handle, mdo, True) ;
    Invalidate;
	TabStop := True;
  end;
end;

procedure TComponentMod.TratarBotaoPanelAzul(var aPanel: TPanel);
begin
  aPanel.Color := clHighlight;
  aPanel.Cursor := crHandPoint;
  aPanel.Font.Color := clWhite;
  aPanel.Font.Style := [fsBold];
  aPanel.OnMouseEnter := HandlePanelAzulOnMouseEnter;
  aPanel.OnMouseLeave := HandlePanelAzulOnMouseLeave;
end;

procedure TComponentMod.TratarBotaoPanelCinza(var aPanel: TPanel);
begin
  aPanel.Color := clBtnShadow;
  aPanel.Cursor := crHandPoint;
  aPanel.Font.Color := clWhite;
  aPanel.Font.Style := [fsBold];
  aPanel.OnMouseEnter := HandlePanelCinzaOnMouseEnter;
  aPanel.OnMouseLeave := HandlePanelCinzaOnMouseLeave;
end;

procedure TComponentMod.TratarBotaoPanelPersonalizado(var aPanel: TPanel;
  aCorEntrada, aCorEntradaFonte, aCorPadrao, aCorPadraoFonte: integer);
begin
  aPanel.Cursor := crHandPoint;
  aPanel.Color := aCorPadrao;
  aPanel.Font.Color := aCorPadraoFonte;
  aPanel.Font.Style := [fsBold];
  aPanel.Hint := aCorEntrada.ToString+'|'+aCorEntradaFonte.ToString;
  aPanel.OnMouseEnter := HandlePanelPersonalizadoOnMouseEnter;
  aPanel.OnMouseLeave := HandlePanelPersonalizadoOnMouseLeave;
end;

procedure TComponentMod.HandlePanelPersonalizadoOnMouseEnter(Sender: TObject);
var
  _Cores: string;
begin
  _Cores := TPanel(Sender).Hint;
  TPanel(Sender).Hint := ColorToString(TPanel(Sender).Color)+'|'+ColorToString(TPanel(Sender).Font.Color);
  TPanel(Sender).Color := StringToColor(ExtrairTextoAPartirDePosicao(_Cores, '|', EXTRAIR_PRE_TRECHO));
  TPanel(Sender).Font.Color := StringToColor(ExtrairTextoAPartirDePosicao(_Cores, '|', EXTRAIR_POS_TRECHO));
end;

procedure TComponentMod.HandlePanelPersonalizadoOnMouseLeave(Sender: TObject);
var
  _Cores: string;
begin
  _Cores := TPanel(Sender).Hint;
  TPanel(Sender).Hint := ColorToString(TPanel(Sender).Color)+'|'+ColorToString(TPanel(Sender).Font.Color);
  TPanel(Sender).Color := StringToColor(ExtrairTextoAPartirDePosicao(_Cores, '|', EXTRAIR_PRE_TRECHO));
  TPanel(Sender).Font.Color := StringToColor(ExtrairTextoAPartirDePosicao(_Cores, '|', EXTRAIR_POS_TRECHO));
end;

procedure TComponentMod.HandlePanelAzulOnMouseEnter(Sender: TObject);
begin
  TPanel(Sender).Color := $00AA5E00;
end;

procedure TComponentMod.HandlePanelAzulOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := clHighlight;
end;

procedure TComponentMod.HandlePanelCinzaOnMouseEnter(Sender: TObject);
begin
  TPanel(Sender).Color := clGray;
end;

procedure TComponentMod.HandlePanelCinzaOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := clBtnShadow;
end;

procedure TComponentMod.TratarBotaoPanelVerde(var aPanel: TPanel);
begin
  aPanel.Color := clGreen;
  aPanel.Cursor := crHandPoint;
  aPanel.Font.Color := clWhite;
  aPanel.Font.Style := [fsBold];
  aPanel.OnMouseEnter := HandlePanelVerdeOnMouseEnter;
  aPanel.OnMouseLeave := HandlePanelVerdeOnMouseLeave;
end;

procedure TComponentMod.TratarBotaoPanelVermelho(var aPanel: TPanel; aTipoCorPadrao: integer);
begin
  aPanel.Cursor := crHandPoint;
  aPanel.Font.Style := [fsBold];
  case aTipoCorPadrao of
    0: begin //Padrão, sem o mouse: vermelho; com o mouse vermelho escuro
      aPanel.Color := $000000D7;
      aPanel.Font.Color := clWhite;
      aPanel.OnMouseEnter := HandlePanelVermelhoPadraoOnMouseEnter;
      aPanel.OnMouseLeave := HandlePanelVermelhoPadraoOnMouseLeave;
    end;
    1: begin //Novo, sem o mouse: branco; Com o mouse: vermelho
      aPanel.Color := clWhite;
      aPanel.Font.Color := clBlack;
      aPanel.OnMouseEnter := HandlePanelVermelhoOnMouseEnter;
      aPanel.OnMouseLeave := HandlePanelVermelhoOnMouseLeave;
    end;
    2: begin //Novo, sem o mouse: BtnFace; Com o mouse: vermelho
      aPanel.Color := clBtnFace;
      aPanel.Font.Color := clBlack;
      aPanel.OnMouseEnter := HandlePanelVermelhoOnMouseEnter;
      aPanel.OnMouseLeave := HandlePanelVermelhoBtnFaceOnMouseLeave;
    end;
  end;
end;

procedure TComponentMod.HandlePanelVermelhoPadraoOnMouseEnter(Sender: TObject);
begin
  TPanel(Sender).Color := $000000AA;
end;

procedure TComponentMod.HandlePanelVermelhoPadraoOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := $000000D7;
end;

procedure TComponentMod.HandlePanelVermelhoOnMouseEnter(Sender: TObject);
begin
  TPanel(Sender).Color := $000000D7; //Estilizado
  TPanel(Sender).Font.Color := clWhite;
end;

procedure TComponentMod.HandlePanelVermelhoOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := clWhite; //Estilizado
  TPanel(Sender).Font.Color := clBlack;
end;

procedure TComponentMod.HandlePanelVerdeOnMouseEnter(Sender: TObject);
begin
  TPanel(Sender).Color := $00004000; //Estilizado
  TPanel(Sender).Font.Color := clWhite;
end;

procedure TComponentMod.HandlePanelVerdeOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := clGreen; //Estilizado
  TPanel(Sender).Font.Color := clWhite;
end;

procedure TComponentMod.HandlePanelVermelhoBtnFaceOnMouseLeave(Sender: TObject);
begin
  TPanel(Sender).Color := clBtnFace; //Estilizado
  TPanel(Sender).Font.Color := clBlack;
end;

end.
