object FrmTest: TFrmTest
  Left = 0
  Top = 0
  Caption = 'RTF2HTML Tester'
  ClientHeight = 451
  ClientWidth = 735
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 351
    Width = 735
    Height = 3
    Cursor = crVSplit
    Align = alBottom
    ExplicitTop = 41
    ExplicitWidth = 164
  end
  object LbLog: TListBox
    Left = 0
    Top = 354
    Width = 735
    Height = 97
    Align = alBottom
    ItemHeight = 13
    TabOrder = 0
    ExplicitTop = 252
    ExplicitWidth = 576
  end
  object PnlToolbar: TPanel
    Left = 0
    Top = 0
    Width = 735
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    ExplicitWidth = 576
    object Button3: TButton
      Left = 8
      Top = 9
      Width = 97
      Height = 25
      Action = CmdFileOpen
      ImageMargins.Left = 4
      Images = ImageList1
      TabOrder = 0
    end
    object Button2: TButton
      Left = 112
      Top = 9
      Width = 114
      Height = 25
      Action = CmdConvert
      DisabledImageIndex = 2
      ImageMargins.Left = 4
      Images = ImageList1
      TabOrder = 1
    end
  end
  object Editor: TRichEdit
    Left = 0
    Top = 41
    Width = 735
    Height = 310
    Align = alClient
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 2
    Zoom = 100
    ExplicitWidth = 576
    ExplicitHeight = 208
  end
  object ActionList1: TActionList
    Images = ImageList1
    Left = 216
    Top = 92
    object CmdConvert: TAction
      Caption = 'Save as HTML...'
      ImageIndex = 1
      OnExecute = CmdConvertExecute
      OnUpdate = CmdConvertUpdate
    end
    object CmdFileOpen: TAction
      Caption = 'Open RTF...'
      ImageIndex = 0
      OnExecute = CmdFileOpenExecute
    end
  end
  object ImageList1: TImageList
    ColorDepth = cd32Bit
    Left = 276
    Top = 92
    Bitmap = {
      494C010103000800580010001000FFFFFFFF2110FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000001000000001002000000000000010
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000003000000033000000330000
      0033000000330000003300000033000000330000003300000033000000330000
      0033000000330000002F00000000000000000000000000000000000000510000
      00E70101018A0101018A0101018A0101018A0101018A0101018A000000E70000
      00E7000000E7000000E7000000AD000000000000000000000000000000290000
      0074000000450000004500000045000000450000004500000045000000740000
      0074000000740000007400000057000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003D89BCF24298D2FF3F94D0FF3D92
      CFFF3D92CEFF3E92CEFF3E92CEFF3E92CEFF3E92CEFF3E92CEFF3E92CEFF3E92
      CEFF3E93CFFF3983B6F00000000E00000000000000000000004C000000DB2B2B
      2BF7B9ABABFF424242FF424242FFB5A7A7FFB5A7A7FFB9ABABFF424242FF5454
      54FF515151FF676767FF000000DB0000000000000000000000260000006E0A0A
      0A7C2E2B2B8011111180111111802D2A2A802D2A2A802E2B2B80111111801515
      1580141414801A1A1A800000006E000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000004399D2FF3E94D0FFABFBFFFF9BF3
      FFFF92F1FFFF93F1FFFF93F1FFFF93F1FFFF93F1FFFF93F1FFFF93F1FFFF93F1
      FFFFA6F8FFFF64B8E3FF060F155F0000000000000000000000CE656565FF3E3E
      3EFFBAB1B1FF3E3E3EFF3E3E3EFFB1A8A8FFB1A8A8FFBAB1B1FF3E3E3EFF5454
      54FF494949FF696969FF000000CE000000000000000000000067191919801010
      10802E2C2C8010101080101010802C2A2A802C2A2A802E2C2C80101010801515
      1580121212801A1A1A8000000067000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000004298D2FF4EA6D9FF8EDAF5FFA2EE
      FFFF82E5FEFF84E5FEFF84E5FEFF85E6FEFF85E6FEFF85E6FEFF85E6FEFF84E6
      FEFF96EBFFFF8CD8F5FF1D4562B80000000000000000000000C9636363FF3A3A
      3AFFC0BBBBFF262626FF262626FFB7B2B2FFB7B2B2FFC0BBBBFF3A3A3AFF5454
      54FF414141FF6B6B6BFF000000C9000000000000000000000065191919800F0F
      0F80302F2F800A0A0A800A0A0A802E2C2C802E2C2C80302F2F800F0F0F801515
      1580101010801B1B1B8000000065000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000004196D1FF6ABEE8FF6CBDE6FFBBF2
      FFFF74DEFDFF76DEFCFF77DEFCFF7ADFFCFF7CDFFCFF7CDFFCFF7CDFFCFF7BDF
      FCFF80E0FDFFADF0FFFF4C9DD3FF0000000E00000000000000C5666666FF3636
      36FFC8C7C7FFC3C2C2FFC3C2C2FFC3C2C2FFC3C2C2FFC8C7C7FF363636FF5454
      54FF393939FF6E6E6EFF000000C50000000000000000000000631A1A1A800E0E
      0E803232328031303080313030803130308031303080323232800E0E0E801515
      15800E0E0E801C1C1C8000000063000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003F95D0FF8AD7F5FF43A1D8FFDDFD
      FFFFDAFAFFFFDBFAFFFFDEFAFFFF73DCFCFF75DBFAFF74DAFAFF73DAFAFF73DA
      FAFF71D9FAFFA1E8FFFF7BBFE6FF060F155E00000000000000C16A6A6AFF3333
      33FF323232FF323232FF323232FF323232FF323232FF323232FF333333FF3333
      33FF333333FF727272FF000000C10000000000000000000000611B1B1B800D0D
      0D800D0D0D800D0D0D800D0D0D800D0D0D800D0D0D800D0D0D800D0D0D800D0D
      0D800D0D0D801D1D1D8000000061000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003D94D0FFABF0FFFF439DD6FF358C
      CBFF358CCBFF358CCBFF368BCBFF5BBEEAFF6ED9FBFF69D6FAFF67D5F9FF66D4
      F9FF65D4F9FF82DEFCFFAAE0F6FF1D4563B900000000000000BE6D6D6DFF6464
      64FF646464FF646464FF646464FF646464FF646464FF646464FF646464FF6464
      64FF646464FF6D6D6DFF000000BE00000000000000000000005F1B1B1B801919
      1980191919801919198019191980191919801919198019191980191919801919
      1980191919801B1B1B800000005F000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003C92CFFFB9F4FFFF72DBFBFF6ACC
      F2FF6BCDF3FF6BCEF3FF6CCEF3FF469CD4FF55BAE9FFDAF8FFFFD7F6FFFFD6F6
      FFFFD5F6FFFFD5F7FFFFDBFCFFFF3D94D0FF00000000000000BB717171FFD4D4
      C9FFF4F4E4FFF4F4E4FFF4F4E4FFF4F4E4FFF4F4E4FFF4F4E4FFF4F4E4FFF4F4
      E4FFD4D4C9FF717171FF000000BB00000000000000000000005E1C1C1C803535
      32803D3D39803D3D39803D3D39803D3D39803D3D39803D3D39803D3D39803D3D
      3980353532801C1C1C800000005E000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003B92CFFFC0F3FFFF70DAFBFF73DB
      FBFF74DBFCFF74DBFCFF75DCFCFF72DAFAFF439CD4FF368CCBFF358CCBFF348C
      CCFF338DCCFF3790CEFF3C94D0FF3881B3EB00000000000000B8747474FFF6F6
      E9FFECECDFFFECECDFFFECECDFFFECECDFFFECECDFFFECECDFFFECECDFFFECEC
      DFFFF6F6E9FF747474FF000000B800000000000000000000005C1D1D1D803D3D
      3A803B3B38803B3B38803B3B38803B3B38803B3B38803B3B38803B3B38803B3B
      38803D3D3A801D1D1D800000005C000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003A92CFFFCAF6FFFF68D5F9FF6BD5
      F9FF6AD5F9FF68D5F9FF68D5FAFF69D7FBFF67D4FAFF5DC7F1FF5DC7F2FF5CC8
      F2FFB4E3F8FF3C94D0FF0A1821690000000000000000000000B5787878FFF8F8
      EFFFF1F1E7FFF1F1E7FFF1F1E7FFF1F1E7FFF1F1E7FFF1F1E7FFF1F1E7FFF1F1
      E7FFF8F8EFFF787878FF000000B500000000000000000000005B1E1E1E803E3E
      3C803C3C3A803C3C3A803C3C3A803C3C3A803C3C3A803C3C3A803C3C3A803C3C
      3A803E3E3C801E1E1E800000005B000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003A92CFFFD5F7FFFF5FD1F9FF60D0
      F8FFB4EBFDFFD9F6FFFFDAF8FFFFDAF8FFFFDBF9FFFFDCFAFFFFDCFAFFFFDCFB
      FFFFE0FFFFFF3D95D0FF020608330000000000000000000000B27B7B7BFFFBFB
      F5FFF6F6F0FFF6F6F0FFF6F6F0FFF6F6F0FFF6F6F0FFF6F6F0FFF6F6F0FFF6F6
      F0FFFBFBF5FF7B7B7BFF000000B20000000000000000000000591F1F1F803F3F
      3D803D3D3C803D3D3C803D3D3C803D3D3C803D3D3C803D3D3C803D3D3C803D3D
      3C803F3F3D801F1F1F8000000059000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000003C94D0FFDCFCFFFFD8F7FFFFD8F7
      FFFFDBFAFFFF348ECDFF3891CEFF3992CFFF3992CFFF3992CFFF3992CFFF3A92
      CFFF3C94D0FF2F6C95D7000000000000000000000000000000B07D7D7DFFFEFE
      FBFFFBFBF8FFFBFBF8FFFBFBF8FFFBFBF8FFFBFBF8FFFBFBF8FFFBFBF8FFFBFB
      F8FFFEFEFBFF7D7D7DFF000000B00000000000000000000000581F1F1F803F3F
      3F803F3F3E803F3F3E803F3F3E803F3F3E803F3F3E803F3F3E803F3F3E803F3F
      3E803F3F3F801F1F1F8000000058000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000001F4864B03C94D0FF3992CFFF3992
      CFFF3C94D0FF2C658DD200000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000AE848484FFFFFF
      FFFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFF
      FEFFFFFFFFFF848484FF000000AE000000000000000000000057212121804040
      408040403F8040403F8040403F8040403F8040403F8040403F8040403F804040
      3F80404040802121218000000057000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000081000000AC1414
      0D6614140D6614140D6614140D6614140D6614140D6614140D6614140D661414
      0D6614140D66000000AC00000081000000000000000000000041000000560404
      0333040403330404033304040333040403330404033304040333040403330404
      0333040403330000005600000041000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000100000000100010000000000800000000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000}
  end
  object DlgOpenRTF: TOpenDialog
    DefaultExt = '.rtf'
    Filter = 'Rich text format files (*.rtf)|*.rtf'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Left = 68
    Top = 92
  end
  object DlgSaveHtml: TSaveDialog
    DefaultExt = '.html'
    Filter = 'HTML files (*.htm;*.html)|*.htm;*.html'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 140
    Top = 92
  end
end
