object IWUserSession: TIWUserSession
  OldCreateOrder = False
  Height = 206
  Width = 377
  object ZQuery1: TZQuery
    Connection = ZConnection1
    Params = <>
    Left = 40
    Top = 40
  end
  object ZConnection1: TZConnection
    ControlsCodePage = cCP_UTF16
    HostName = 'localhost'
    Port = 3306
    Database = 'ppl_input'
    User = 'root'
    Password = '123'
    Protocol = 'mysql'
    Left = 144
    Top = 56
  end
end
