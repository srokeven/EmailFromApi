object fmSendEmail: TfmSendEmail
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Envio de Email'
  ClientHeight = 556
  ClientWidth = 737
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poMainFormCenter
  Scaled = False
  OnCreate = FormCreate
  TextHeight = 16
  object pcMain: TPageControl
    Left = 0
    Top = 0
    Width = 737
    Height = 556
    ActivePage = tsSettings
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 733
    ExplicitHeight = 555
    object tsSend: TTabSheet
      Caption = 'Envio'
      DesignSize = (
        729
        525)
      object Label1: TLabel
        Left = 16
        Top = 16
        Width = 125
        Height = 16
        Caption = 'Email do destinat'#225'rio:'
      end
      object Label2: TLabel
        Left = 16
        Top = 68
        Width = 104
        Height = 16
        Caption = 'Enviar c'#243'pia para:'
      end
      object Label3: TLabel
        Left = 16
        Top = 172
        Width = 40
        Height = 16
        Caption = 'Anexo:'
      end
      object Label4: TLabel
        Left = 16
        Top = 224
        Width = 85
        Height = 16
        Caption = 'Texto do email'
      end
      object lbStatus: TLabel
        Left = 16
        Top = 484
        Width = 421
        Height = 16
        Anchors = [akLeft, akRight, akBottom]
        AutoSize = False
        Caption = 'Preencha os dados do destinat'#225'rio...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMenuText
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ExplicitTop = 350
        ExplicitWidth = 363
      end
      object Label11: TLabel
        Left = 16
        Top = 120
        Width = 108
        Height = 16
        Caption = 'Assunto do e-mail:'
      end
      object edDestinatario: TEdit
        Left = 16
        Top = 38
        Width = 691
        Height = 24
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 0
        Text = 'email@email.com'
      end
      object edDestinatarioCopia: TEdit
        Left = 16
        Top = 90
        Width = 691
        Height = 24
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 1
        Text = 'emailcopia@email.com'
      end
      object pnlSearchFile: TPanel
        Left = 538
        Top = 194
        Width = 81
        Height = 24
        Cursor = crHandPoint
        Anchors = [akTop, akRight]
        BevelKind = bkFlat
        BevelOuter = bvNone
        Caption = 'Adicionar'
        ParentBackground = False
        TabOrder = 3
        StyleElements = []
        OnClick = pnlSearchFileClick
      end
      object mmTextoEmail: TMemo
        Left = 16
        Top = 246
        Width = 691
        Height = 209
        Anchors = [akLeft, akTop, akRight, akBottom]
        Color = clInfoBk
        TabOrder = 4
      end
      object pnlSend: TPanel
        Left = 578
        Top = 474
        Width = 129
        Height = 37
        Anchors = [akRight, akBottom]
        BevelOuter = bvNone
        Caption = 'Enviar'
        Color = clHighlight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindow
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 5
        StyleElements = []
        OnClick = pnlSendClick
      end
      object pnlCancel: TPanel
        Left = 443
        Top = 474
        Width = 129
        Height = 37
        Anchors = [akRight, akBottom]
        BevelOuter = bvNone
        Caption = 'Fechar'
        Color = 215
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindow
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 6
        StyleElements = []
        OnClick = pnlCancelClick
      end
      object edSubject: TEdit
        Left = 16
        Top = 142
        Width = 691
        Height = 24
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 2
        Text = 'Assunto'
      end
      object cbListaAnexos: TComboBox
        Left = 16
        Top = 195
        Width = 513
        Height = 22
        Style = csOwnerDrawFixed
        Color = cl3DLight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
      end
      object pnlRemoveFile: TPanel
        Left = 626
        Top = 194
        Width = 81
        Height = 24
        Cursor = crHandPoint
        Anchors = [akTop, akRight]
        BevelKind = bkFlat
        BevelOuter = bvNone
        Caption = 'Remover'
        Color = 215
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentBackground = False
        ParentFont = False
        TabOrder = 8
        StyleElements = []
        OnClick = pnlRemoveFileClick
      end
    end
    object tsSettings: TTabSheet
      Caption = 'Configura'#231#245'es'
      ImageIndex = 1
      DesignSize = (
        729
        525)
      object GroupBox1: TGroupBox
        Left = 3
        Top = 303
        Width = 723
        Height = 90
        Anchors = [akLeft, akTop, akRight]
        Caption = 'Configura'#231#227'o da Api'
        TabOrder = 1
        object Label14: TLabel
          Left = 19
          Top = 25
          Width = 92
          Height = 16
          Caption = 'Porta de servi'#231'o'
        end
        object edPorta: TEdit
          Left = 19
          Top = 47
          Width = 121
          Height = 24
          TabOrder = 0
          Text = '9050'
        end
      end
      object gbConfiguracaoEmail: TGroupBox
        Left = 3
        Top = 0
        Width = 723
        Height = 297
        Anchors = [akLeft, akTop, akRight]
        Caption = 'Configura'#231#227'o de email'
        TabOrder = 0
        object Label10: TLabel
          Left = 339
          Top = 192
          Width = 67
          Height = 16
          Caption = 'Porta SMTP'
        end
        object Label9: TLabel
          Left = 19
          Top = 192
          Width = 85
          Height = 16
          Caption = 'Servidor SMTP'
        end
        object Label5: TLabel
          Left = 19
          Top = 136
          Width = 80
          Height = 16
          Caption = 'Usu'#225'rio SMTP'
        end
        object Label12: TLabel
          Left = 19
          Top = 79
          Width = 97
          Height = 16
          Caption = 'Email Remetente'
        end
        object Label7: TLabel
          Left = 19
          Top = 27
          Width = 101
          Height = 16
          Caption = 'Configura'#231#227'o SSL'
        end
        object Label8: TLabel
          Left = 194
          Top = 27
          Width = 52
          Height = 16
          Caption = 'Usar TLS'
        end
        object Label13: TLabel
          Left = 339
          Top = 79
          Width = 117
          Height = 16
          Caption = 'Nome do Remetente'
        end
        object Label6: TLabel
          Left = 339
          Top = 136
          Width = 73
          Height = 16
          Caption = 'Senha SMTP'
        end
        object chkUsarTLS: TCheckBox
          Left = 339
          Top = 256
          Width = 320
          Height = 17
          Caption = 'Requer conex'#227'o segura (TLS)'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
        object chkUsarSSL: TCheckBox
          Left = 19
          Top = 256
          Width = 293
          Height = 17
          Caption = 'Requer uma conex'#227'o criptografada (SSL)'
          TabOrder = 1
        end
        object edServidorSMTP: TEdit
          Left = 19
          Top = 214
          Width = 293
          Height = 24
          TabOrder = 2
        end
        object edPortaSMTP: TEdit
          Left = 339
          Top = 214
          Width = 320
          Height = 24
          TabOrder = 3
        end
        object edUserSMTP: TEdit
          Left = 19
          Top = 158
          Width = 293
          Height = 24
          TabOrder = 4
        end
        object edEmailRemetente: TEdit
          Left = 19
          Top = 101
          Width = 293
          Height = 24
          TabOrder = 5
        end
        object cbSsl: TComboBox
          Left = 19
          Top = 49
          Width = 145
          Height = 24
          Style = csDropDownList
          ItemIndex = 3
          TabOrder = 6
          Text = 'sslvTLSv1_1'
          Items.Strings = (
            'sslvSSLv23'
            'sslvTLSv1_2'
            'sslvSSLv3'
            'sslvTLSv1_1'
            'sslvTLSv1'
            'sslvSSLv2')
        end
        object cbUsarTLS: TComboBox
          Left = 194
          Top = 49
          Width = 145
          Height = 24
          Style = csDropDownList
          ItemIndex = 1
          TabOrder = 7
          Text = 'utUseExplicitTLS'
          Items.Strings = (
            'utUseImplicitTLS'
            'utUseExplicitTLS'
            'utNoTLSSupport'
            'utUseRequireTLS')
        end
        object edNomeDoRemetente: TEdit
          Left = 339
          Top = 101
          Width = 320
          Height = 24
          TabOrder = 8
        end
        object edSenhaSMTP: TEdit
          Left = 339
          Top = 158
          Width = 320
          Height = 24
          PasswordChar = '*'
          TabOrder = 9
        end
      end
    end
    object tsHTML: TTabSheet
      Caption = 'Modelo do Email'
      ImageIndex = 2
      object lbModeloHTML: TLabel
        Left = 0
        Top = 0
        Width = 729
        Height = 16
        Align = alTop
        Caption = 'HTML de modelo para envio de email'
        ExplicitWidth = 213
      end
      object mmHTML: TMemo
        Left = 0
        Top = 16
        Width = 729
        Height = 462
        Align = alClient
        Lines.Strings = (
          
            '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "' +
            'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
          '<html xmlns="http://www.w3.org/1999/xhtml">'
          '   <head>'
          
            '      <meta http-equiv="Content-Type" content="text/html; charse' +
            't=UTF-8" />'
          '      <title>$TITULO</title>'
          
            '      <meta name="viewport" content="width=device-width, initial' +
            '-scale=1.0"/>'
          '   </head>'
          
            '   <body style="margin: 0; padding: 0; background-color:#e6e6e6;' +
            '">'
          '      <table cellpadding="0" cellspacing="0" width="100%">'
          '         <tr>'
          '            <td>'
          
            '               <table align="center" border="1" cellpadding="0" ' +
            'cellspacing="0" width="600">'
          '                  <tr>'
          
            '                     <td align="center" bgcolor="#f6f6f6" style=' +
            '"padding: 10px 0 10px 0;">'
          
            '                        <div style="background:url(https://wstec' +
            'hsolucoes.net/View/imgs/logoup2.png) no-repeat center; backgroun' +
            'd-size: contain; width:700px; height: 180px;"></div>'
          '                     </td>'
          '                  </tr>'
          '                  <tr>'
          
            '                     <td bgcolor="#ffffff" style="padding: 40px ' +
            '30px 40px 30px;">'
          
            '                        <table cellpadding="0" cellspacing="0" w' +
            'idth="100%">'
          '                           <tr >'
          '                              <td>'
          '                                 <h3>$ASSUNTO</h3>'
          
            #9#9#9#9' <p style="margin-top: -10px;">Remetente: <strong>$CLIENTE</' +
            'strong></p>'
          '                              </td>'
          '                           </tr>'
          '                           <tr>'
          '                              <td>'
          '                                 <p>$TEXTO</p>'
          '                              </td>'
          '                           </tr>'
          '                           <tr>'
          '                              <td height="40">'
          
            '                                 <strong><p style="margin-top: 4' +
            '0px;">$ANEXO</p></strong>'
          '                              </td>'
          '                           </tr>'
          '                        </table>'
          '                     </td>'
          '                  </tr>'
          '                  <tr>'
          
            '                     <td bgcolor="#f55702" style="padding: 30px ' +
            '30px 30px 30px;">'
          
            '                        <table cellpadding="0" cellspacing="0" w' +
            'idth="100%">'
          '                           <tr>'
          '                              <td width="75%">'
          
            '                                 <p style="color: white;">$INFO<' +
            '/p>'
          '                              </td>'
          '                           </tr>'
          '                        </table>'
          '                     </td>'
          '                  </tr>'
          '               </table>'
          '            </td>'
          '         </tr>'
          '      </table>'
          '   </body>'
          '</html>')
        ScrollBars = ssVertical
        TabOrder = 0
        WordWrap = False
        ExplicitWidth = 725
        ExplicitHeight = 461
      end
      object Panel1: TPanel
        Left = 0
        Top = 478
        Width = 729
        Height = 47
        Align = alBottom
        TabOrder = 1
        ExplicitTop = 477
        ExplicitWidth = 725
        object pnlCarregarHTML: TPanel
          Left = 8
          Top = 6
          Width = 129
          Height = 37
          BevelOuter = bvNone
          Caption = 'Carregar Modelo'
          Color = clBtnShadow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindow
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 0
          StyleElements = []
          OnClick = pnlCarregarHTMLClick
        end
        object pnlVisualizarHTML: TPanel
          Left = 143
          Top = 6
          Width = 129
          Height = 37
          BevelOuter = bvNone
          Caption = 'Visualizar'
          Color = clBtnShadow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindow
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 1
          StyleElements = []
          OnClick = pnlVisualizarHTMLClick
        end
      end
    end
    object tsVisualizarEmail: TTabSheet
      Caption = 'Visualizar Email'
      ImageIndex = 3
      object wbHtml: TWebBrowser
        Left = 0
        Top = 0
        Width = 729
        Height = 525
        Align = alClient
        TabOrder = 0
        ExplicitWidth = 722
        ExplicitHeight = 515
        ControlData = {
          4C000000584B0000433600000000000000000000000000000000000000000000
          000000004C000000000000000000000001000000E0D057007335CF11AE690800
          2B2E126208000000000000004C0000000114020000000000C000000000000046
          8000000000000000000000000000000000000000000000000000000000000000
          00000000000000000100000000000000000000000000000000000000}
      end
    end
  end
  object OpenFile: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <>
    Options = [fdoAllowMultiSelect]
    Left = 588
    Top = 123
  end
end
