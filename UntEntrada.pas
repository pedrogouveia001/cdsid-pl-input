unit UntEntrada;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompButton,
  IWCompListbox, IWCompEdit, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWCompLabel, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, IWGrids, ComObj, Dialogs,
  Variants, Math, UserSessionUnit, IWExtCtrls, jpeg, windows, ActiveX,
  IWWebGrid, IWAdvWebGrid, IWCompTabControl, IWAdvWebGridExcel, IWCompText,
  IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent, IWBaseHTML40Component,
  chart, series, pngimage, Grids,
  IWCompProgressBar, ExtCtrls, IWHTMLControls, TeeGDIPlus, TeEngine, TeeProcs,
  IWTMSPopup, IWHelpTip, Menus,

  BubbleCh, TeePolar,
  IWCompRadioButton, IWTMSCheckList, TeeLisB,
  TeeSeriesTextEd, Buttons, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Messages, Graphics, ComCtrls, StdCtrls, ErrorBar,
  IWProgressIndicator, IWDBGrids,
  IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, GanttCh, IWCompRectangle,
  TeeDoubleHorizBar;

type
  TIWForm1 = class(TIWAppForm)
    rgnmatcons: TIWRegion;
    IWImage31: TIWImage;
    IWImage32: TIWImage;
    IWLabel20: TIWLabel;
    IWRegion18: TIWRegion;
    IWRegion30: TIWRegion;
    IWRegion1: TIWRegion;
    IWLabel3: TIWLabel;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    btnupdate: TIWButton;
    rgnmatrizdeconsequencias: TIWRegion;
    IWLabel4: TIWLabel;
    btnshowresults: TIWButton;
    grdmatrizconseq: TTIWAdvWebGrid;
    cmbbxauxiliar: TIWComboBox;
    rgnresultado: TIWRegion;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWLabel7: TIWLabel;
    IWRegion3: TIWRegion;
    IWRegion4: TIWRegion;
    chrtResultado: TChart;
    IWRegion5: TIWRegion;
    grdresultados: TIWGrid;
    IWRegion6: TIWRegion;
    grdcases: TIWGrid;
    IWLabel8: TIWLabel;
    IWLabel9: TIWLabel;
    Chart1: TChart;
    Series2: TBarSeries;
    edtnumalt: TIWEdit;
    grdImport: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWLabel10: TIWLabel;
    FileArquivo: TIWFile;
    btnimport: TIWButton;
    IWLabel11: TIWLabel;
    IWLabel12: TIWLabel;
    IWLink1: TIWLink;
    rgndetalhamentoalt: TIWRegion;
    IWImage7: TIWImage;
    IWImage8: TIWImage;
    IWLabel24: TIWLabel;
    IWLink5: TIWLink;
    IWRegion14: TIWRegion;
    IWRegion15: TIWRegion;
    IWRegion16: TIWRegion;
    grdestatisticaalternativa: TIWGrid;
    IWText2: TIWText;
    IWRegion2: TIWRegion;
    grdprobvitoria: TIWGrid;
    Series1: TBarSeries;
    rgnAnaliseEstatisticaPromethee: TIWRegion;
    IWImage11: TIWImage;
    IWImage12: TIWImage;
    IWLabel32: TIWLabel;
    IWLink10: TIWLink;
    IWText3: TIWText;
    IWText4: TIWText;
    IWText6: TIWText;
    IWRegion23: TIWRegion;
    IWRegion24: TIWRegion;
    IWRegion26: TIWRegion;
    grdestatisticaPromethee: TIWGrid;
    IWRegion27: TIWRegion;
    grdmaxminPromethee: TIWGrid;
    IWRegion28: TIWRegion;
    grdexportacao: TTIWAdvWebGrid;
    lnkExportar: TIWLink;
    lnkexportacaopromethee: TIWLink;
    cmbbxcriteria: TIWComboBox;
    rgnregradedecisao: TIWRegion;
    IWImage15: TIWImage;
    IWImage16: TIWImage;
    IWLabel42: TIWLabel;
    IWLink12: TIWLink;
    IWRegion37: TIWRegion;
    IWRegion38: TIWRegion;
    IWRegion39: TIWRegion;
    grdregradedecisao: TIWGrid;
    txtregradedecisao: TIWText;
    rgndescricao: TIWRegion;
    IWText7: TIWText;
    IWLink15: TIWLink;
    rgnalternativa: TIWRegion;
    IWImage19: TIWImage;
    IWImage20: TIWImage;
    IWLabel49: TIWLabel;
    IWLabel50: TIWLabel;
    IWLabel51: TIWLabel;
    IWLink2: TIWLink;
    IWRegion45: TIWRegion;
    IWRegion46: TIWRegion;
    IWRegion47: TIWRegion;
    grdalternativa: TIWGrid;
    IWRegion48: TIWRegion;
    grdcasesalternativa: TIWGrid;
    grdrecomendacao: TIWGrid;
    IWLabel27: TIWLabel;
    IWLabel28: TIWLabel;
    IWLabel29: TIWLabel;
    lnkdescricaolimiar: TIWLink;
    rgnlimiares: TIWRegion;
    grdlimiares: TTIWAdvWebGrid;
    btnatualizarlimiares: TIWButton;
    txtdescricaoregra: TIWText;
    IWLabel30: TIWLabel;
    rgnregradedecisaoPromethee: TIWRegion;
    IWImage9: TIWImage;
    IWImage10: TIWImage;
    IWLabel31: TIWLabel;
    IWLink6: TIWLink;
    txtregradedecisaopromethee: TIWText;
    grdrecomendacaopromethee: TIWGrid;
    IWLabel46: TIWLabel;
    IWLabel47: TIWLabel;
    lnkdescricaolimiarPromethee: TIWLink;
    IWRegion19: TIWRegion;
    IWRegion20: TIWRegion;
    IWRegion21: TIWRegion;
    grdregradedecisaopromethee: TIWGrid;
    rgndescricaoPromethee: TIWRegion;
    IWText9: TIWText;
    IWLabel55: TIWLabel;
    rgnlimiaresPromethee: TIWRegion;
    btnatualizarlimiaresPromethee: TIWButton;
    txtdescricaoregraPromethee: TIWText;
    IWLabel56: TIWLabel;
    grdlimiarespromethee: TTIWAdvWebGrid;
    lnkregrasdedecisao: TIWLink;
    grdperdamediaemaxima: TIWGrid;
    lblp1: TIWLabel;
    lblp2: TIWLabel;
    lblp3: TIWLabel;
    lblp4: TIWLabel;
    cmbbxp1: TIWComboBox;
    cmbbxp2: TIWComboBox;
    cmbbxp3: TIWComboBox;
    cmbbxp4: TIWComboBox;
    cmbbxp1promethee: TIWComboBox;
    lblp1promethee: TIWLabel;
    cmbbxp2promethee: TIWComboBox;
    cmbbxp3promethee: TIWComboBox;
    cmbbxp4promethee: TIWComboBox;
    lblp2promethee: TIWLabel;
    lblp3promethee: TIWLabel;
    lblp4promethee: TIWLabel;
    lnkregrasdedecisaoPromethee: TIWLink;
    grdperdamediaemaximapromethee: TIWGrid;
    IWLabel13: TIWLabel;
    IWLabel14: TIWLabel;
    IWLabel25: TIWLabel;
    IWLabel26: TIWLabel;
    IWLabel43: TIWLabel;
    IWLabel44: TIWLabel;
    IWLabel36: TIWLabel;
    IWLabel45: TIWLabel;
    IWLabel33: TIWLabel;
    IWLabel34: TIWLabel;
    tbctrlResultados: TIWTabControl;
    pgAlternativas: TIWTabPage;
    pgPOA: TIWTabPage;
    IWRegion17: TIWRegion;
    imgPOA: TIWImage;
    IWRegion22: TIWRegion;
    imgAlternative: TIWImage;
    chrtPOA: TChart;
    BarSeries1: TBarSeries;
    pgComparacao: TIWTabPage;
    IWRegion25: TIWRegion;
    imgcomparacao: TIWImage;
    chrtComparacao: TChart;
    Series3: TLineSeries;
    Series4: TLineSeries;
    rgncomparacao: TIWRegion;
    IWLabel57: TIWLabel;
    grdcomparacao: TIWGrid;
    IWLink7: TIWLink;
    IWLabel37: TIWLabel;
    cmbbxAlternativa: TIWComboBox;
    imgalternativa: TIWImage;
    chrtAlternativa: TChart;
    Series5: TPieSeries;
    IWLabel38: TIWLabel;
    IWLabel39: TIWLabel;
    IWLabel40: TIWLabel;
    IWLabel41: TIWLabel;
    IWLabel48: TIWLabel;
    rgnintro: TIWRegion;
    lblWarning: TIWLabel;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    IWButton3: TIWButton;
    IWLabel52: TIWLabel;
    IWImage13: TIWImage;
    IWLabel53: TIWLabel;
    IWLabel70: TIWLabel;
    IWImage14: TIWImage;
    IWImage27: TIWImage;
    imgSolve: TIWImage;
    txtwarnning: TIWText;
    iwrgnCaixaDialogoRU: TIWRegion;
    lblWg: TIWLabel;
    IWRegionru: TIWRegion;
    txtCaixaDialogoRU: TIWText;
    btnBackRU: TIWButton;
    btncontinueRU: TIWButton;
    IWRegion29: TIWRegion;
    IWRegion31: TIWRegion;
    IWRegion32: TIWRegion;
    IWLabel54: TIWLabel;
    IWLabel58: TIWLabel;
    rgnEnterLogin: TIWRegion;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    btnEnter: TIWButton;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    IWLabel59: TIWLabel;
    IWLabel60: TIWLabel;
    btnrecover: TIWButton;
    txtpassrecover: TIWText;
    txtnewcode: TIWText;
    IWImage18: TIWImage;
    IWImage33: TIWImage;
    lnkesqueceusenha: TIWLink;
    lnkbacklogin: TIWLink;
    IWRegion33: TIWRegion;
    IWRegion34: TIWRegion;
    iwrgnCaixaDialogoRU2: TIWRegion;
    lblWg2: TIWLabel;
    IWRegionru2: TIWRegion;
    txtCaixaDialogoRU2: TIWText;
    btnBackRU2: TIWButton;
    btncontinueRU2: TIWButton;
    IWRegion35: TIWRegion;
    IWLabel62: TIWLabel;
    IWLabel63: TIWLabel;
    IWLabel64: TIWLabel;
    rgnOpcaoInput: TIWRegion;
    lblpleasechoose: TIWLabel;
    btnRegisternew: TIWButton;
    btnContinueregist: TIWButton;
    btnLogoutRgnOpcaoInput: TIWButton;
    IWImage37: TIWImage;
    IWImage38: TIWImage;
    IWLink11: TIWLink;
    IWImage47: TIWImage;
    IWImage48: TIWImage;
    IWRegion36: TIWRegion;
    IWRegion40: TIWRegion;
    IWRegion41: TIWRegion;
    IWLabel61: TIWLabel;
    IWLabel65: TIWLabel;
    IWLabel66: TIWLabel;
    rgnLoginproject: TIWRegion;
    lblChooseproj: TIWLabel;
    lstbxProjects: TIWListbox;
    btnReset3: TIWButton;
    btnLogoutRgnLogin: TIWButton;
    lblMexplanation: TIWLabel;
    btnauxchoose: TIWButton;
    IWImage25: TIWImage;
    IWImage34: TIWImage;
    IWRegion42: TIWRegion;
    IWRegion43: TIWRegion;
    IWRegion44: TIWRegion;
    IWRegion49: TIWRegion;
    IWLabel67: TIWLabel;
    IWLabel69: TIWLabel;
    IWLabel71: TIWLabel;
    IWLabel84: TIWLabel;
    edtnameproblem: TIWEdit;
    IWLink8: TIWLink;
    IWLink9: TIWLink;
    IWButton1: TIWButton;
    IWButton2: TIWButton;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWLabel15: TIWLabel;
    rgnanaliseestatistica: TIWRegion;
    IWImage3: TIWImage;
    IWImage4: TIWImage;
    IWLabel16: TIWLabel;
    IWLink3: TIWLink;
    IWText1: TIWText;
    IWText5: TIWText;
    IWText8: TIWText;
    IWLink4: TIWLink;
    IWLink13: TIWLink;
    IWLabel19: TIWLabel;
    IWRegion8: TIWRegion;
    IWRegion9: TIWRegion;
    IWRegion10: TIWRegion;
    grdestatistica: TIWGrid;
    IWRegion11: TIWRegion;
    grdmaxmin: TIWGrid;
    IWLabel17: TIWLabel;
    IWLabel18: TIWLabel;
    IWTimer1: TIWTimer;
    procedure rgnmatconsCreate(Sender: TObject);
    procedure btnupdateClick(Sender: TObject);
    procedure cmbbxauxiliarAsyncChange(Sender: TObject;
      EventParams: TStringList);
    procedure cmbbxauxiliarChange(Sender: TObject);
    procedure btnshowresultsClick(Sender: TObject);
    procedure GerarCases;
    procedure rgnresultadoCreate(Sender: TObject);
    procedure SetarCores;
    procedure btnimportClick(Sender: TObject);
    procedure LimpaVariaveis;
    function TrocaPV(Num: String): Double;
    function TrocaVP(Num: Double): String;
    function trocavp1(Num: String): Double;
    procedure rgnregradedecisaoCreate(Sender: TObject);
    procedure rgnregradedecisaoprometheeCreate(Sender: TObject);
    procedure grdlimiaresComboChange(Sender: TObject;
      RowIndex, ColumnIndex: Integer; AValue: string);
    procedure grdlimiaresEdit(Sender: TObject);
    procedure EscalaRazao;
    procedure PLEscolha(auxinst: Integer);
    procedure PlOrdenacao(auxinst: Integer);
    procedure DSiTrimWorkingSet;
    procedure tbctrlResultadosChange(Sender: TObject);
    procedure btnupdatevisualizacaoClick(Sender: TObject);
    procedure cmbbxAlternativaChange(Sender: TObject);
    procedure rgnalternativaCreate(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink2Click(Sender: TObject);
    procedure rgnintroCreate(Sender: TObject);
    procedure btnRegisterClick(Sender: TObject);
    procedure btnBackRUClick(Sender: TObject);
    procedure btncontinueRUClick(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure lnkesqueceusenhaClick(Sender: TObject);
    procedure btnBackRU2Click(Sender: TObject);
    procedure btncontinueRU2Click(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure rgnEnterLoginCreate(Sender: TObject);
    procedure rgnOpcaoInputCreate(Sender: TObject);
    procedure btnRegisternewClick(Sender: TObject);
    procedure btnContinueregistClick(Sender: TObject);
    procedure IWLink11Click(Sender: TObject);
    procedure lnkbackloginClick(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure IWLink9Click(Sender: TObject);
    procedure rgnLoginprojectCreate(Sender: TObject);
    procedure SaveProblem;
    procedure btnauxchooseClick(Sender: TObject);
    procedure IWButton1Click(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure rgndetalhamentoaltCreate(Sender: TObject);
    procedure IWLink5Click(Sender: TObject);
    procedure ProgramarComponentes;
    function Probabilidade_AMaiorQueB(AInf, ASup, BInf, BSup: Double): Double;
    procedure IWLink1Click(Sender: TObject);
    procedure lnkExportarClick(Sender: TObject);
    procedure rgnanaliseestatisticaCreate(Sender: TObject);
  public
  end;

implementation

uses lpsolve51, uAPI_SolveLP, ServerController;
{$R *.dfm}

Function TIWForm1.TrocaPV(Num: String): Double;
var
  S: string;
  i: Integer;
begin
  S := Num;
  for i := 2 to length(S) - 1 do
    if S[i] = '.' then
      S[i] := ',';

  Result := StrToFloat(S);
end;

Function TIWForm1.TrocaVP(Num: Double): String;
var
  S: string;
  i: Integer;
begin
  S := FloatToStr(Num);
  for i := 2 to length(S) - 1 do
    if S[i] = ',' then
      S[i] := '.';

  Result := S;
end;

Function TIWForm1.trocavp1(Num: String): Double;
var
  S: string;
  i: Integer;
begin
  S := Num;
  for i := 2 to length(S) - 1 do
    if S[i] = ',' then
      S[i] := '.';

  Result := StrToFloat(S);
end;

function TIWForm1.Probabilidade_AMaiorQueB(AInf, ASup, BInf, BSup: Double)
  : Double;
var
  fA, fB, densidadeConjunta: Double;
  integral1, integral2: Double;
  b1, b2: Double;
  a0, b0: Double;
begin
  // Verificações básicas: modelo degenerado permitido quando AInf = ASup ou BInf = BSup
  if (AInf > ASup) or (BInf > BSup) then
  begin
    Result := -1; // Erro: intervalos inválidos (inf > sup)
    Exit;
  end;

  // Caso degenerado: A fixo
  if AInf = ASup then
  begin
    a0 := AInf;
    // Se B também fixo
    if BInf = BSup then
    begin
      b0 := BInf;
      if a0 > b0 then
        Result := 1.0
      else
        Result := 0.0;
      Exit;
    end;
    // B contínuo
    // P(a0 > B) = P(B < a0) = (min(a0, BSup) - BInf) / (BSup - BInf)
    if a0 <= BInf then
      Result := 0.0
    else if a0 >= BSup then
      Result := 1.0
    else
      Result := (a0 - BInf) / (BSup - BInf);
    Exit;
  end;

  // Caso degenerado: B fixo
  if BInf = BSup then
  begin
    b0 := BInf;
    // A contínuo
    // P(A > b0) = (ASup - max(AInf, b0)) / (ASup - AInf)
    if b0 >= ASup then
      Result := 0.0
    else if b0 <= AInf then
      Result := 1.0
    else
      Result := (ASup - b0) / (ASup - AInf);
    Exit;
  end;

  // Caso geral: A e B contínuos
  fA := 1 / (ASup - AInf);
  fB := 1 / (BSup - BInf);
  densidadeConjunta := fA * fB;

  // Caso 1: A sempre > B
  if AInf >= BSup then
  begin
    Result := 1.0;
    Exit;
  end;

  // Caso 2: A nunca > B
  if ASup <= BInf then
  begin
    Result := 0.0;
    Exit;
  end;

  // Região 1: b de BInf até min(AInf, BSup)
  b1 := BInf;
  b2 := Min(AInf, BSup);
  if b2 > b1 then
    integral1 := (b2 - b1) * (ASup - AInf) * densidadeConjunta
  else
    integral1 := 0.0;

  // Região 2: b de max(AInf, BInf) até min(ASup, BSup)
  b1 := Max(AInf, BInf);
  b2 := Min(ASup, BSup);
  if b2 > b1 then
  begin
    // ∫ (ASup - b) db = [ASup*b - b^2/2] de b1 a b2
    integral2 := (ASup * b2 - 0.5 * b2 * b2) - (ASup * b1 - 0.5 * b1 * b1);
    integral2 := integral2 * densidadeConjunta;
  end
  else
    integral2 := 0.0;

  Result := integral1 + integral2;
end;

procedure TIWForm1.PlOrdenacao(auxinst: Integer);
var
  i, j, k, Nrest, a, b, c, z, II, JJ, DNrel, NumDV: Integer; // Contadores
  m1, m2, m3: Integer;
  dominance: boolean; // Var auxiliar que indica se uma relação de dominância foi estabelecida
  NumCri, NumAlt, Numl, NAH: Integer; // Número de critérios,
  // num de alternativas, num de linhasna Matriz do Modelo de PL e número de restrições
  // provenientes de Avaliações holísticas, número de alternativas potencialmente ótimas
  MatA, MatCons: Array of Array of Double; // Matriz do Modelo de PL e Matriz de Consequências Normalizada
  Criorder: Array of Integer; // Vetor que armazena a ordenação dos critérios
  LimMin, LimMax: Array of Double; // Vetores que armazenam os limites inferior e
  // Superior da relação entre os critérios
  VetAH: Array of Array of Integer; // Matriz que armazena as informações das AH
  OPAH: Array of string; // Vetor que armazena o operador definido na relação estabelecida na AH
  G, ep: Double;
  // VARIÁVEIS LP SOLVE
  // ResultVector armazena os resultados obtidos pelo LP Solve (a função máxima e
  // coeficientes das variáveis de decisão).
  // FO recebe os coeficientes das variáveis de decisão na função objetivo e será
  // usado no LP Solve.
  // Limit recebe os valores das constantes das restrições e será usado no LP Solve.
  // X armazena
  ResultVector: array of real;
  FO: array of string;
  // typeprob determina se o problema é contínuo (typeprob = 0) ou binário (typeprob = 1).
  // No caso do fitradeoff, é sempre contínuo pois as constantes de escala são contínuas.
  // dirprob determina se o problema é de minimização (dirprob = 0) ou de maximização
  // (dirprob = 1), SF nos diz se houve solução viável.
  typeprob, dirprob, SF: Integer;
  Op: array of string;
  limit: array of real;
  Rmod: array of AnsiString;
  Lp: THandle;
  S1: string;
  MatrizParaPar: Array of Array of Integer; // Armazena as relações de dominÂncia par a par
  MaxDIJ, MaxDJI: Double; // Armazena a máxima diferença entre duas alternativas
  Aux1, Aux2: Double;
  // VARIÁVEIS PARA A DLL DE C++

  icase: Pinteger; // ponteiro do c++ para indicar se existe ou não solução possível
  MinVK: Double;
  var_int, var_bin: Array of Integer;
  MSSG: Integer;
begin

  /// ///////// Recebendo o valor das variáveis salvas na usersession /////////////
  MinVK := TIWUserSession(WebApplication.Data).MinVK;
  NumCri := TIWUserSession(WebApplication.Data).numcrit;
  NumAlt := TIWUserSession(WebApplication.Data).NumAlt;

  ep := TIWUserSession(WebApplication.Data).epsilon2;
  /// O número de linhas na matriz do PPL equivale ao número de critérios - 1 (número ///////////
  /// de restrições do espaço de pesos, mais o número de restrições da avaliação //////////////
  /// holística, mais um da restrição da normalização dos pesos, mais um da função objetivo,//
  /// mais o número de restrições de potencial otimalidade.//////////////////////////////////

  Numl := 2 * (NumCri - 1) + 5;
  Setlength(MatrizParaPar, NumAlt, NumAlt);
  Setlength(TIWUserSession(WebApplication.Data).maxDiff, auxinst + 1, NumAlt, NumAlt);
  setlength(TIWUserSession(WebApplication.Data).mediamaxdiff, numalt, numalt);
  setlength(TIWUserSession(WebApplication.Data).maxmaxdiff, numalt, numalt);
  Setlength(MatCons, NumAlt, NumCri);
  Setlength(Criorder, NumCri);
  Setlength(MatA, Numl, NumCri + 2);
  Setlength(Op, Numl - 1);
  Setlength(limit, Numl - 1);
  Setlength(Rmod, Numl - 1);
  Setlength(LimMax, NumCri - 1);
  Setlength(LimMin, NumCri - 1);
  Setlength(VetAH, NAH);
  Setlength(OPAH, NAH);
  Setlength(FO, NumCri);
  Setlength(ResultVector, NumCri + 1);
  /// Var de decisão são contínuas ///
  for i := 0 to NumAlt - 1 do
  begin
    for j := 0 to NumCri - 1 do
      MatCons[i, j] := TIWUserSession(WebApplication.Data)
        .MatrizConseqNorm[i, j];
  end;
  for i := 0 to NumCri - 1 do
    Criorder[i] := TIWUserSession(WebApplication.Data)
      .CasesOrdemcrit[auxinst, i] - 1;
  for i := 0 to NumCri - 2 do
  begin
    // setlength(LimMin, i + 1);
    LimMin[i] := 0;
    LimMax[i] := 1;
  end;

  /// ///////////////////////////////////////////////////////////////////////////
  /// Limpando a Matriz do problema de possíveis valores provindos de sua inicialização ///
  for i := 0 to Numl - 1 do
    for j := 0 to NumCri do
      MatA[i, j] := 0;

  m1 := 0;
  m2 := 0;
  m3 := 0;

  /// //////////// Programação Linear - MODELO PROB DE ORDENAÇÃO ///////////////

  /// Preenchendo as Restrições referentes ao limite superior da relação dos critérios ///
  /// transformando-a em uma restrição do tipo '<=' para seguir o padrão adotado no SAD ///
  for i := 0 to NumCri - 2 do
  begin
    MatA[i + 1, i + 1] := (-1) * LimMax[i];
    MatA[i + 1, i + 2] := 1;
    Op[i] := '<=';
    INC(Nrest);
    INC(m1);
  end;

  a := Nrest;

  /// Preenchendo as Restrições referentes ao limite inferior da relação dos critérios ///
  /// Essa restrição já é do tipo '<=' que segue o padrão adotado no SAD ///
  for i := a to a + NumCri - 2 do
  begin
    MatA[i + 1, i - a + 1] := LimMin[i - a];
    MatA[i + 1, i - a + 2] := (-1) * 1;
    Op[i] := '<=';
    INC(Nrest);
    INC(m1);
  end;

  a := Nrest;

  /// Preenchendo aa nova restrição Kn > limiar///
  MatA[a + 1, NumCri] := 1;
  MatA[a + 1, 0] := MinVK;
  Op[a] := '>=';
  INC(m2);
  INC(Nrest);

  a := Nrest;

  /// Preenchendo as Restrições referentes à Normalização dos Pesos///
  for j := 0 to NumCri - 1 do
    MatA[a + 1, j + 1] := 1;
  MatA[a + 1, 0] := 1;
  Op[a] := '=';
  INC(Nrest);
  INC(m3);

  a := Nrest;

  NumDV := NumCri;

  /// ////////////////// Verifica as relações de dominância /////////////////////
  for II := 0 to NumAlt - 1 do
    for JJ := 0 to NumAlt - 1 do
    begin
      if (II = JJ) then
       continue;
      if (TIWUserSession(WebApplication.Data).matrizaltsel[auxinst, jj] = 1) and (TIWUserSession(WebApplication.Data).matrizaltsel[auxinst, ii] = 0)  then
      continue;

      for j := 0 to NumCri - 1 do
      begin
        MatA[Nrest + 1, j + 1] := MatCons[jj, Criorder[j]] - MatCons
          [ii, Criorder[j]];
      end; // j
      MatA[Nrest + 1, 0] := -0.00001;
      Op[Nrest] := '<=';

      typeprob := 0;
      /// Var de decisão são contínuas ///
      dirprob := 1;
      /// Problema de maximização ///

      /// Preenchendo o lado direito da restrição no vetor utilizado no LP_Solve ///
      for i := 0 to Numl - 2 do
        limit[i] := MatA[i + 1, 0];

      /// Preenchendo as restrições no vetor utilizado no LP_Solve ///
      for i := 0 to Numl - 2 do
      begin
        Rmod[i] := TrocaVP(MatA[i + 1, 1]);
        for j := 2 to NumCri do
          Rmod[i] := Rmod[i] + ' ' + TrocaVP(MatA[i + 1, j]);
      end;

      /// Preenchendo a função objetivo, onde as var de decisão são as constantes de ///
      /// escala e os coeficientes são as diferenças entre as consequências das//
      /// alternativa IJ normalizadas ///
      for i := 0 to NumCri - 1 do
      begin
        MatA[0, i + 1] := MatCons[II, Criorder[i]] - MatCons[JJ, Criorder[i]];
        FO[i] := TrocaVP(MatCons[II, Criorder[i]] - MatCons[JJ, Criorder[i]]);
      end;
      /// Resolvendo o PPL ///
      SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int, var_bin,
        Rmod, ResultVector, Lp, MSSG);

      SF := MSSG;
      if SF = 0 then
      begin
        /// Verificando se há algum K não nulo. ////
        b := 0;
        for i := 1 to NumCri do
          if ResultVector[i] <> 0 then
            break
          else
            INC(b);

        MaxDIJ := 0;
        // Se obtem o valor da maxdif.

        MaxDIJ := ResultVector[0];

      end; // end do if viavel = 1
      With TIWUserSession(WebApplication.Data) do
      begin
      if maxdij > 0 then
      begin
      TIWUserSession(WebApplication.Data).maxDiff[auxinst, II, JJ] := MaxDIJ;

      if maxDiff[auxinst, II, JJ] > TIWUserSession(WebApplication.Data).maxmaxdiff[ii,jj] then
      maxmaxdiff[ii,jj] :=  maxDiff[auxinst, II, JJ];


      end;
      TIWUserSession(WebApplication.Data).mediamaxdiff[ii,jj] := TIWUserSession(WebApplication.Data).mediamaxdiff[ii,jj] +  TIWUserSession(WebApplication.Data).maxDiff[auxinst, II, JJ];

      end;
    end; // For II

end;

procedure TIWForm1.PLEscolha(auxinst: Integer);
var
  i, j, k, Nrest, a, b, z, II, JJ, c: Integer; // Contadores
  NumCri, NumAlt, Numl, NAH, NumPOA: Integer; // Número de critérios,
  // num de alternativas, num de linhasna Matriz do Modelo de PL e número de restrições
  // provenientes de Avaliações holísticas, número de alternativas potencialmente ótimas
  MatA, MatCons: Array of Array of Double; // Matriz do Modelo de PL e Matriz de Consequências Normalizada
  VetPOA, Criorder: Array of Integer; // Vetor que armazena os índices das alternativas
  // potencialmente ótimas do problema, vetor que armazena as alternativas que passaram por uma
  // avaliação holística e vetor que armazena a ordenação dos critérios
  VetAH: Array of Array of Integer; // Matriz que armazena as informações das AH
  LimMin, LimMax: Array of Double; // Vetores que armazenam os limites inferior e
  // Superior da relação entre os critérios
  OPAH: Array of string; // Vetor que armazena o operador definido na relação estabelecida na AH
  G, ep: Double;
  // VARIÁVEIS LP SOLVE
  // ResultVector armazena os resultados obtidos pelo LP Solve (a função máxima e
  // coeficientes das variáveis de decisão).
  // FO recebe os coeficientes das variáveis de decisão na função objetivo e será
  // usado no LP Solve.
  // Limit recebe os valores das constantes das restrições e será usado no LP Solve.
  // X armazena
  NumDV, NRA, NRB: Integer; // número de var de decisão do modelo, inteiros auxiliares para a AH
  ResultVector: array of real;
  // typeprob determina se o problema é contínuo (typeprob = 0) ou binário (typeprob = 1).
  // No caso do fitradeoff, é sempre contínuo pois as constantes de escala são contínuas.
  // dirprob determina se o problema é de minimização (dirprob = 0) ou de maximização
  // (dirprob = 1), SF nos diz se houve solução viável.
  typeprob, dirprob, SF: Integer;
  FO, Op: array of string;
  limit: array of real;
  Rmod: array of AnsiString;
  var_int, var_bin: Array of Integer;
  /// /vetores que carregam var de decisão inteiras e binárias
  Lp: THandle;
  S1: string;
  MAXDI, MAXDJ: array [0 .. 2] of Double;
  AuxHEz: array of Integer; // vetor auxiliar para facilitar a elaboração de res_
  // trições com var binária da HE
  MSSG: Integer;
  MinVK: Double;
  pular, novo: boolean;
begin

  /// ///////// Recebendo o valor das variáveis salvas na usersession /////////////

  NumCri := TIWUserSession(WebApplication.Data).numcrit;
  NumAlt := TIWUserSession(WebApplication.Data).NumAlt;
  TIWUserSession(WebApplication.Data).NDA := NumAlt;
  NumPOA := TIWUserSession(WebApplication.Data).NDA;
  ep := TIWUserSession(WebApplication.Data).epsilon2;
  MinVK := TIWUserSession(WebApplication.Data).MinVK;
  /// Se inicia o número de linhas da Matriz como um valor suficiete para as restrições
  /// da PL, novas linhas serão acrescentadas de acordo ao número de AH realizadas///
  Numl := 2 * (NumCri - 1) + 4 + (NumPOA - 1);
  NumDV := NumCri;
  Setlength(MatCons, NumAlt, NumCri);
  Setlength(Criorder, NumCri);
  Setlength(MatA, Numl, NumCri + 1);
  Setlength(Op, Numl - 1);
  Setlength(limit, Numl - 1);
  Setlength(Rmod, Numl - 1);
  Setlength(VetPOA, NumPOA);
  Setlength(LimMax, NumCri - 1);
  Setlength(LimMin, NumCri - 1);
  Setlength(VetAH, NAH);
  Setlength(OPAH, NAH);
  Setlength(FO, NumCri);
  Setlength(ResultVector, NumCri + 1);

  typeprob := 0;
  /// Var de decisão são contínuas ///
  for i := 0 to NumAlt - 1 do
  begin
    for j := 0 to NumCri - 1 do
      MatCons[i, j] := TIWUserSession(WebApplication.Data)
        .MatrizConseqNorm[i, j];
  end;
  for i := 0 to NumCri - 1 do
    Criorder[i] := TIWUserSession(WebApplication.Data)
      .CasesOrdemcrit[auxinst, i] - 1;
  for i := 0 to NumCri - 2 do
  begin
    // setlength(LimMin, i + 1);
    LimMin[i] := 0;
    LimMax[i] := 1;
  end;

  for i := 0 to NumPOA - 1 do
  begin
    VetPOA[i] := i;
  end;

  /// ///////////////////////////////////////////////////////////////////////////
  /// Limpando a Matriz do problema de possíveis valores provindos de sua inicialização ///
  for i := 0 to Numl - 1 do
    for j := 0 to NumCri do
      MatA[i, j] := 0;

  /// //////////// Programação Linear - MODELO PROB DE ESCOLHA //////////////////

  /// Preenchendo as Restrições referentes ao limite superior da relação dos critérios ///
  /// transformando-a em uma restrição do tipo '<=' para seguir o padrão adotado no SAD ///
  for i := 0 to NumCri - 2 do
  begin
    MatA[i + 1, i + 1] := (-1) * LimMax[i];
    MatA[i + 1, i + 2] := 1;
    Op[i] := '<=';
    INC(Nrest);
  end;

  a := Nrest;

  /// Preenchendo as Restrições referentes ao limite inferior da relação dos critérios ///
  /// Essa restrição já é do tipo '<=' que segue o padrão adotado no SAD ///
  for i := a to a + NumCri - 2 do
  begin
    MatA[i + 1, i - a + 1] := LimMin[i - a];
    MatA[i + 1, i - a + 2] := (-1) * 1;
    Op[i] := '<=';
    INC(Nrest);
  end;

  a := Nrest;

  /// / Adiciona variáveis binárias ao problema caso necessário ////////////////
  Setlength(FO, NumDV);
  Setlength(var_bin, (NumDV - NumCri));
  Setlength(ResultVector, NumDV + 1);
  for j := 0 to length(var_bin) - 1 do
    var_bin[j] := NumCri + j;
  /// //////////////////////////////////////////////////////////////////////////
  a := Nrest;

  a := Nrest;

  /// Preenchendo aa nova restrição Kn > limiar///
  MatA[a + 1, NumCri] := -1;
  MatA[a + 1, 0] := -MinVK;
  Op[a] := '<=';
  INC(Nrest);

  a := Nrest;

  /// Preenchendo as Restrições referentes à Normalização dos Pesos///
  for j := 0 to NumCri - 1 do
    MatA[a + 1, j + 1] := 1;
  MatA[a + 1, 0] := 1;
  Op[a] := '=';
  INC(Nrest);

  a := Nrest;

  a := Nrest;
  NumPOA := 0;

  for z := 0 to TIWUserSession(WebApplication.Data).NDA - 1 do
  begin
    k := VetPOA[z];
    b := a;

    /// Preenchendo as Restrições de Potencial Otimalidade ///
    /// transformando-a em uma restrição do tipo '<=' para seguir o padrão adotado no SAD ///
    for i := a to a + TIWUserSession(WebApplication.Data).NDA - 1 do
    begin
      if VetPOA[i - a] <> k then
      begin
        for j := 0 to NumCri - 1 do
          MatA[b + 1, j + 1] := MatCons[VetPOA[i - a], Criorder[j]] - MatCons
            [k, Criorder[j]];
        Op[b] := '<=';
        MatA[b + 1, 0] := -TIWUserSession(WebApplication.Data).Epson;
        INC(Nrest);
        INC(b);
      end;
    end;

    /// Preenchendo a função objetivo, onde as var de decisão são as constantes de ///
    /// escala e os coeficientes são as consequências da alternativa normalizada ///
    for i := 0 to NumCri - 1 do
    begin
      MatA[0, i + 1] := MatCons[k, Criorder[i]];
      FO[i] := TrocaVP(MatCons[k, Criorder[i]]);
    end;
    /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
    for i := NumCri to NumDV - 1 do
    begin
      MatA[0, i + 1] := 0;
      FO[i] := '0';
    end;

    /// Preenchendo o lado direito da restrição no vetor utilizado no LP_Solve ///
    for i := 0 to Numl - 2 do
      limit[i] := MatA[i + 1, 0];

    /// Preenchendo as restrições no vetor utilizado no LP_Solve ///
    for i := 0 to Numl - 2 do
    begin
      Rmod[i] := TrocaVP(MatA[i + 1, 1]);
      for j := 2 to NumDV do
        Rmod[i] := Rmod[i] + ' ' + TrocaVP(MatA[i + 1, j]);
    end;

    /// Problema de maximização ///
    dirprob := 1;

    /// Resolvendo o PPL ///
    SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int, var_bin,
      Rmod, ResultVector, Lp, MSSG);

    TIWUserSession(WebApplication.Data).Vgmax[auxinst, z] := ResultVector[0];

    SF := MSSG;
    if SF = 0 then
    begin
      /// Verificando se há algum K não nulo. ///
      Setlength(TIWUserSession(WebApplication.Data).Sconst, NumCri);
      b := 0;
      for i := 1 to NumCri do
        if ResultVector[i] <> 0 then
          break
        else
          INC(b);

      if b <> NumCri then
      begin
        /// Atualização do Vetor de Alternativas Potencialmente Ótimas ///
        INC(NumPOA);
        Setlength(TIWUserSession(WebApplication.Data).VetPOA, NumPOA);
        Setlength(TIWUserSession(WebApplication.Data).VetPesos, NumPOA, NumCri);

        TIWUserSession(WebApplication.Data).VetPOA[NumPOA - 1] := k;
        TIWUserSession(WebApplication.Data).Vgmax[auxinst, NumPOA - 1] :=
          ResultVector[0];
        for i := 0 to NumCri - 1 do
        begin
          TIWUserSession(WebApplication.Data).Sconst[i] := ResultVector[i + 1];
          TIWUserSession(WebApplication.Data).VetPesos[NumPOA - 1, i] :=
            ResultVector[i + 1];
        end;
      end;

      // ** Só é preciso solucionar o problema de minimização <=> a alternativa é POA **//
      // ** Além disso, as restrições e a FO é a mesma **//
      /// Problema de maximização ///
      dirprob := 0;
      Setlength(ResultVector, 0);
      Setlength(ResultVector, NumDV + 1);
      /// Resolvendo o PPL ///
      SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int, var_bin,
        Rmod, ResultVector, Lp, MSSG);
      if MSSG = 0 then
        TIWUserSession(WebApplication.Data).VGMin[auxinst, z] := ResultVector
          [0];

    end;

    /// Limpando as restrições de Otimalidade da Matriz.
    for j := 0 to NumCri do
    begin
      MatA[0, j] := 0;
      for i := a + 1 to Nrest do
        MatA[i, j] := 0;
    end;
    Nrest := a;
    Setlength(ResultVector, 0);
    Setlength(ResultVector, NumDV + 1);
  end;

  // ** Caso nenhuma POA seja encontrada, o problema pode estar no epson **//
  if NumPOA = 0 then
  begin
    TIWUserSession(WebApplication.Data).Epson := 0;
    PLEscolha(auxinst);
    Exit;
  end;

  TIWUserSession(WebApplication.Data).NDA := NumPOA;

  Setlength(VetPOA, NumPOA);
  for i := 0 to NumPOA - 1 do
    VetPOA[i] := TIWUserSession(WebApplication.Data).VetPOA[i];

  for j := 0 to NumAlt - 1 do
    TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := -1;

  for i := 0 to length(VetPOA) - 1 do
  begin
    TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := VetPOA[i];


    TIWUserSession(WebApplication.Data).probabilidadealt[vetpoa[i]] := TIWUserSession(WebApplication.Data).probabilidadealt[vetpoa[i]] + (1/numpoa);

    TIWUserSession(WebApplication.Data).AltPOA[VetPOA[i]] := 1 + TIWUserSession
      (WebApplication.Data).AltPOA[VetPOA[i]];

    TIWUserSession(WebApplication.Data).MatrizAltSel[auxinst, VetPOA[i]] := 1;

    if auxinst = 0 then
      TIWUserSession(WebApplication.Data).matrizsol[auxinst, VetPOA[i]] := 1;

    INC(TIWUserSession(WebApplication.Data).AltTamanhoPOA[VetPOA[i], length
        (VetPOA) - 1]);

  end;

  With TIWUserSession(WebApplication.Data) do
  begin
    novo := true;
    if auxinst = 0 then
      novo := false;
    if novo = true then
    begin
      for j := 0 to length(TIWUserSession(WebApplication.Data).matrizsol) - 1 do
      begin
        k := 0;
        pular := false;
        for i := 0 to NumAlt - 1 do
        begin

          if (TIWUserSession(WebApplication.Data).matrizsol[j, i]
              = TIWUserSession(WebApplication.Data).MatrizAltSel[auxinst, i])
            and (pular = false) then
            INC(k)
          else
            pular := true;
        end;
        if k = NumAlt then
        begin
          resultsol[j] := resultsol[j] + 1;
          novo := false;
          break;
        end;
      end;
    end;
    if novo = true then
    begin
      Setlength(matrizsol, length(matrizsol) + 1, NumAlt);
      Setlength(resultsol, length(resultsol) + 1);
      k := length(matrizsol) - 1;
      resultsol[k] := 1;
      for i := 0 to length(VetPOA) - 1 do
        matrizsol[k, VetPOA[i]] := MatrizAltSel[auxinst, VetPOA[i]];
    end;
    VetNumPOA[auxinst] := NumPOA;
    INC(numsolinst[NumPOA - 1]);
  end;

  /// ///////////////////////////////////////////////////////////////////////////
  /// //////////////////////Max diferença////////////////////////////////////////
  if (NumPOA > 1) and (NumPOA <= 3) and (ep <> 0) then
  begin

    /// Preenchendo o lado direito da restrição no vetor utilizado no LP_Solve //
    Setlength(limit, Nrest);
    for i := 0 to Nrest - 1 do
      limit[i] := MatA[i + 1, 0];

    /// Preenchendo as restrições no vetor utilizado no LP_Solve ///
    Setlength(Rmod, Nrest);
    for i := 0 to Nrest - 1 do
    begin
      Rmod[i] := TrocaVP(MatA[i + 1, 1]);
      for j := 2 to NumDV do
        Rmod[i] := Rmod[i] + ' ' + TrocaVP(MatA[i + 1, j]);
    end;
    Setlength(Op, Nrest);
    a := 1;

    c := 0;
    for II := 0 to NumPOA - 1 do
      for JJ := II + 1 to NumPOA - 1 do
      begin
        /// ///////// Verifica a max diferença entre II - JJ////////////////////

        /// / Preenche a função objetivo ///////////////////////////////////////
        for k := 0 to NumCri - 1 do
          FO[k] := TrocaVP
            (MatCons[VetPOA[II], Criorder[k]] - MatCons[VetPOA[JJ], Criorder[k]]
            );
        /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
        for i := NumCri to NumDV - 1 do
        begin
          FO[i] := '0';
        end;

        /// Resolvendo o PPL ///
        SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int,
          var_bin, Rmod, ResultVector, Lp, MSSG);

        SF := MSSG;

        if SF = 0 then
        begin
          /// Verificando se há algum K não nulo. ///
          b := 0;
          for i := 1 to NumCri do
            if ResultVector[i] <> 0 then
              break
            else
              INC(b);

          MAXDI[II] := ResultVector[0];
        end; // If SF

        /// ///////// Verifica a max diferença entre JJ - II////////////////////
        /// / Preenche a função objetivo ///////////////////////////////////////
        for k := 0 to NumCri - 1 do
          FO[k] := TrocaVP
            (MatCons[VetPOA[JJ], Criorder[k]] - MatCons[VetPOA[II], Criorder[k]]
            );
        /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
        for i := NumCri to NumDV - 1 do
        begin
          FO[i] := '0';
        end;

        /// Resolvendo o PPL ///
        SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int,
          var_bin, Rmod, ResultVector, Lp, MSSG);

        SF := MSSG;
        if SF = 0 then
        begin
          /// Verificando se há algum K não nulo. ///
          b := 0;
          for i := 1 to NumCri do
            if ResultVector[i] <> 0 then
              break
            else
              INC(b);

          MAXDJ[II] := ResultVector[0];
        end; // If SF
        {
          for j := 0 to numalt - 1 do
          TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := -1;

          for I := 0 to length(vetPOA) - 1 do
          begin
          TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := vetPOA[i];

          TIWUserSession(WebApplication.Data).AltPOA[vetpoa[i]] := 1 + TIWUserSession(WebApplication.Data).AltPOA[vetpoa[i]];
          end;
          }
        /// / Deve-se considerar o módulo da máx diferença entre as alternativas ////�
        if MAXDI[II] < 0 then
          MAXDI[II] := -MAXDI[II];
        if MAXDJ[II] < 0 then
          MAXDJ[II] := -MAXDJ[II];

        if (MAXDI[II] <= ep) and (MAXDJ[II] <= ep) then
          INC(c);

      end; // for JJ

  end; // if numpoa
  // delete_lp(lp);

  // free_lp(lp);

  DSiTrimWorkingSet;
end;

procedure TIWForm1.DSiTrimWorkingSet;
var
  hProcess: THandle;
begin
  hProcess := OpenProcess(PROCESS_SET_QUOTA, false, GetCurrentProcessId);
  try
    SetProcessWorkingSetSize(hProcess, $FFFFFFFF, $FFFFFFFF);
  finally
    CloseHandle(hProcess);
  end;
end; { DSiTrimWorkingSet }

procedure TIWForm1.tbctrlResultadosChange(Sender: TObject);
begin
  if tbctrlResultados.activepage = 2 then
  begin
    rgncomparacao.Visible := true;
    IWRegion5.Visible := false;
    IWLabel8.Visible := false;
    IWLabel9.Visible := false;
    IWRegion6.Visible := false;
  end
  else
  begin
    rgncomparacao.Visible := false;
    IWRegion5.Visible := true;
    IWLabel8.Visible := true;
    IWLabel9.Visible := true;
    IWRegion6.Visible := true;
  end;

end;

procedure TIWForm1.GerarCases;
var
  Vetor, Indices: array of Integer;
  Fatorial, Linha, i, j, k, kk, Temp: Integer;
  N: Integer;
  Matriz: array of array of Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    // Calcular o fatorial de N (número total de permutações)
    Fatorial := 1;
    N := numcrit;
    for i := 2 to N do
      Fatorial := Fatorial * i;

    // Redimensionar a matriz para armazenar todas as permutações
    Setlength(Matriz, Fatorial, N);

    // Inicializar o vetor de entrada e índices
    Setlength(Vetor, N);
    for i := 0 to N - 1 do
      Vetor[i] := i + 1;

    Setlength(Indices, N);
    for i := 0 to N - 1 do
      Indices[i] := 0;

    // Copiar a primeira permutação
    Linha := 0;
    for i := 0 to N - 1 do
      Matriz[Linha][i] := Vetor[i];
    INC(Linha);

    // Algoritmo Heap's para gerar as permutações
    i := 0;
    while i < N do
    begin
      if Indices[i] < i then
      begin
        // Trocar elementos
        if i mod 2 = 0 then
          Temp := Vetor[0]
        else
          Temp := Vetor[Indices[i]];

        if i mod 2 = 0 then
          Vetor[0] := Vetor[i]
        else
          Vetor[Indices[i]] := Vetor[i];

        Vetor[i] := Temp;

        // Armazenar a nova permutação na matriz
        for j := 0 to N - 1 do
          Matriz[Linha][j] := Vetor[j];
        INC(Linha);

        // Incrementar o índice
        INC(Indices[i]);
        i := 0;
      end
      else
      begin
        Indices[i] := 0;
        INC(i);
      end;
    end;

    if numcrit <= 7 then
    begin
      // Ordenar a matriz de permutações em ordem lexicográfica
      for i := 0 to Fatorial - 2 do
        for j := i + 1 to Fatorial - 1 do
        begin
          for k := 0 to N - 1 do
          begin
            if Matriz[i][k] < Matriz[j][k] then
              break
            else if Matriz[i][k] > Matriz[j][k] then
            begin
              // Trocar as linhas
              for kk := 0 to N - 1 do
              begin
                Temp := Matriz[i][kk];
                Matriz[i][kk] := Matriz[j][kk];
                Matriz[j][kk] := Temp;
              end;
              break;
            end;
          end;
        end;
    end;

    Setlength(CasesOrdemcrit, Fatorial, numcrit);
    for i := 0 to Fatorial - 1 do
    begin
      for j := 0 to numcrit - 1 do
        CasesOrdemcrit[i, j] := Matriz[i, j];

    end;
  end;
end;

procedure TIWForm1.grdlimiaresComboChange(Sender: TObject;
  RowIndex, ColumnIndex: Integer; AValue: string);
var
  i: Integer;
  aux: string;
begin
  if StrToFloat(grdlimiares.DynEdits[1, 0]) < 0.03 then
    grdlimiares.DynEdits[1, 0] := '0.03';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) < 0.02 then
    grdlimiares.DynEdits[1, 1] := '0.02';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) < 0.01 then
    grdlimiares.DynEdits[1, 2] := '0.01';

  if StrToFloat(grdlimiares.DynEdits[1, 3]) > 0.97 then
    grdlimiares.DynEdits[1, 3] := '0.97';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) > 0.98 then
    grdlimiares.DynEdits[1, 2] := '0.98';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) > 0.99 then
    grdlimiares.DynEdits[1, 1] := '0.99';

  for i := grdlimiares.rowcount - 1 downto 1 do
  begin
    if StrToFloat(grdlimiares.DynEdits[1, i]) > StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      aux := grdlimiares.DynEdits[1, i];
      grdlimiares.DynEdits[1, i] := grdlimiares.DynEdits[1, i - 1];
      grdlimiares.DynEdits[1, i - 1] := aux;
    end;
    if StrToFloat(grdlimiares.DynEdits[1, i]) = StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      grdlimiares.DynEdits[i, i - 1] := formatfloat
        ('0.00', StrToFloat(grdlimiares.DynEdits[i, i]) + 0.01);
    end;
  end;
end;

procedure TIWForm1.grdlimiaresEdit(Sender: TObject);
var
  i: Integer;
  aux: string;
begin
  if StrToFloat(grdlimiares.DynEdits[1, 0]) < 0.03 then
    grdlimiares.DynEdits[1, 0] := '0.03';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) < 0.02 then
    grdlimiares.DynEdits[1, 1] := '0.02';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) < 0.01 then
    grdlimiares.DynEdits[1, 2] := '0.01';

  if StrToFloat(grdlimiares.DynEdits[1, 3]) > 0.97 then
    grdlimiares.DynEdits[1, 3] := '0.97';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) > 0.98 then
    grdlimiares.DynEdits[1, 2] := '0.98';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) > 0.99 then
    grdlimiares.DynEdits[1, 1] := '0.99';

  for i := grdlimiares.rowcount - 1 downto 1 do
  begin
    if StrToFloat(grdlimiares.DynEdits[1, i]) > StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      aux := grdlimiares.DynEdits[1, i];
      grdlimiares.DynEdits[1, i] := grdlimiares.DynEdits[1, i - 1];
      grdlimiares.DynEdits[1, i - 1] := aux;
    end;
    if StrToFloat(grdlimiares.DynEdits[1, i]) = StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      grdlimiares.DynEdits[i, i - 1] := formatfloat
        ('0.00', StrToFloat(grdlimiares.DynEdits[i, i]) + 0.01);
    end;
  end;
end;

procedure TIWForm1.IWButton1Click(Sender: TObject);
begin
  rgnalternativa.Visible := true;
  rgnresultado.Visible := false;
end;

procedure TIWForm1.IWButton2Click(Sender: TObject);
begin
  rgnanaliseestatistica.Visible := true;
  rgnresultado.Visible := false;
end;

procedure TIWForm1.IWLink11Click(Sender: TObject);
begin
  rgnOpcaoInput.Visible := false;
  rgnEnterLogin.Visible := true;
end;

procedure TIWForm1.IWLink1Click(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnOpcaoInput.Visible := true;
end;

procedure TIWForm1.IWLink2Click(Sender: TObject);
begin
  rgnalternativa.Visible := false;
  rgnresultado.Visible := true;
end;

procedure TIWForm1.IWLink5Click(Sender: TObject);
begin
  rgndetalhamentoalt.Visible := false;
  rgnresultado.Visible := true;
end;

procedure TIWForm1.IWLink7Click(Sender: TObject);
begin
  rgnalternativa.Visible := true;
  rgnresultado.Visible := false;
end;

procedure TIWForm1.IWLink8Click(Sender: TObject);
begin
  rgnOpcaoInput.Visible := true;
  rgnmatcons.Visible := false;
end;

procedure TIWForm1.IWLink9Click(Sender: TObject);
begin
  rgnOpcaoInput.Visible := true;
  rgnLoginproject.Visible := false;
end;

procedure TIWForm1.LimpaVariaveis;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    // Limpeza de arrays dinâmicos
    Finalize(NomeCrit);
    Finalize(NomeAlt);
    Finalize(TipoCrit);
    Finalize(Niveis);
    Finalize(MatrizConseq);
    Finalize(MatrizConseqNorm);
    Finalize(PesoCrit);
    Finalize(CasesOrdemcrit);
    Finalize(Max);
    Finalize(Min);
    Finalize(MatrizPOA);
    Finalize(MatrizPOAPromethee);
    Finalize(ResultadoROC);
    Finalize(ResultadoPromethee);
    Finalize(matrizsol);
    Finalize(MatrizSolPromethee);
    Finalize(ResultAlt);
    Finalize(ResultAltPromethee);
    Finalize(resultsol);
    Finalize(ResultSolPromethee);
    Finalize(coresalt);
    Finalize(coresalt1);
    Finalize(CaseSol);
    Finalize(CaseSolPromethee);
    Finalize(Crit10);
    Finalize(Crit12);
    Finalize(Crit13);
    Finalize(Crit14);
    Finalize(Crit20);
    Finalize(Crit21);
    Finalize(Crit23);
    Finalize(Crit24);
    Finalize(MatrizDifVG);
    Finalize(MatrizDifVGPromethee);
    Finalize(MediaDifSol);
    Finalize(MaxDifSol);
    Finalize(DesvioPadraoDifSol);
    Finalize(MinDifSol);
    Finalize(MediaDifSolPromethee);
    Finalize(MaxDifSolPromethee);
    Finalize(DesvioPadraoDifSolPromethee);
    Finalize(MinDifSolPromethee);
    Finalize(SobClassMatrix);
    Finalize(PositiveFlow);
    Finalize(NegativeFlow);
    Finalize(NetFlow);
    Finalize(MatrizParaPar);
    Finalize(OrdemSol);
    Finalize(OrdemSolPromethee);
    Finalize(ordemalt);
    Finalize(ordemaltpromethee);
    Finalize(SimilaridadeAlt);
    Finalize(SomenteAditivo);
    Finalize(SomenteSobreclassificacao);
    Finalize(VetPOA);
    Finalize(MatrizPOA2);
    Finalize(Sconst);
    Finalize(VetPesos);
    Finalize(AltPOA);
    Finalize(PergNAdj);
    Finalize(MatrizAltSel);
    Finalize(VetNumPOA);
    Finalize(numsolinst);
    Finalize(AltTamanhoPOA);
    Finalize(MatrizComparacaoAlt);
    Finalize(AmbasAltNaPOA);
    Finalize(NenhumaAltNaPOA);
    Finalize(CritName);
    Finalize(RestName);
    Finalize(AltName);
    Finalize(CritSigla);
    Finalize(AltSigla);
    Finalize(critcode);
    Finalize(altcode);
    Finalize(PrefDir);
    Finalize(Vgmax);
    Finalize(VGMin);
    Finalize(Mediavgmax);
    Finalize(MediaVgMin);
    Finalize(MaxVgMax);
    Finalize(MaxVGMin);
    Finalize(MinVgMax);
    Finalize(MinVGMin);
    Finalize(AmplitudeVG);
    Finalize(difvg);
    Finalize(MediaAmplitudeVG);
    Finalize(MaxAmplitudeVG);
    Finalize(MinAmplitudeVG);
    Finalize(AmplitudeParAParinst);
    Finalize(ProbVitoriaParapar);
    Finalize(mediaprobvitoriaparapar);
    Finalize(maximoprobvitoriaparapar);
    Finalize(amplitudeparapar);
    Finalize(MediaGeralProbVitoriaAlt);
    Finalize(MaximoGeralProbVitoriaAlt);
    Finalize(MediaGeralProbDerrotaAlt);
    Finalize(MaximoGeralProbDerrotaAlt);
    Finalize(mediadifvg);
    Finalize(maximodifvg);
    Finalize(mediageralprobalt);

    // Reinicialização de variáveis escalares
    numcrit := 0;
    NumAlt := 0;
    excelverif := false;
    Total := 0.0;
    comp1 := 0.0;
    comp2 := 0.0;
    ind := 0.0;
    pref := 0.0;
    gauss := 0.0;
    SimilaridadeGeral := 0.0;
    MinMaxMaxNaoSol := 0.0;
    MinMaxMaxNaoSolPromethee := 0.0;
    NDA := 0;
    epsilon2 := 0.0;
    MinVK := 0.0;
    Epson := 0.0;

    // Limpar todos os TChart
    // chrtResultado.SeriesList.Clear;
    // Chart1.SeriesList.Clear;
    chrtPOA.SeriesList.Clear;
    // chrtComparacao.SeriesList.Clear;
    // chrtAlternativa.SeriesList.Clear;

    // Limpar todos os TIWGrid
    grdresultados.Clear;
    grdcases.Clear;
    grdestatisticaalternativa.Clear;
    grdprobvitoria.Clear;
    grdestatisticapromethee.Clear;
    grdmaxminPromethee.Clear;
    grdestatistica.Clear;
    grdmaxmin.Clear;
    grdregradedecisao.Clear;
    grdalternativa.Clear;
    grdcasesalternativa.Clear;
    grdrecomendacao.Clear;
    grdregradedecisaopromethee.Clear;
    grdrecomendacaopromethee.Clear;
    grdperdamediaemaxima.Clear;
    grdperdamediaemaximapromethee.Clear;
    grdcomparacao.Clear;

  end;
end;

procedure TIWForm1.lnkbackloginClick(Sender: TObject);
begin
  rgnEnterLogin.Visible := false;
  rgnintro.Visible := true;
end;

procedure TIWForm1.lnkesqueceusenhaClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU2.Visible := true;
end;

procedure TIWForm1.lnkExportarClick(Sender: TObject);
var
  i, j, k: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    grdexportacao.Columns.Clear;
    grdexportacao.totalrows := 9 + grdestatisticaalternativa.rowcount +
      grdcomparacao.rowcount + grdresultados.rowcount;
    if grdestatisticaalternativa.ColumnCount >= grdcomparacao.ColumnCount then
    begin
      for i := 0 to grdestatisticaalternativa.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end
    else
    begin
      for i := 0 to grdcomparacao.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end;

    grdexportacao.Cells[0, 0] := 'Result Report - PPL Input';
    grdexportacao.Cells[0, 2] := 'Percentage of each different solution:';
    k := 2;
    for i := 0 to grdresultados.rowcount - 1 do
      for j := 0 to grdresultados.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdresultados.Cell[i, j].Text;
      end;
    k := k + 1 + grdresultados.rowcount;

    grdexportacao.Cells[0, k + 1] := 'Alternative Comparison:';
    k := k + 1;

    // grdexportacao.Cells[0,k] := 'Consequences Variation percentage values:';
    for i := 0 to grdcomparacao.rowcount - 1 do
      for j := 0 to grdcomparacao.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdcomparacao.Cell[i, j].Text;
      end;
    k := k + 1 + grdcomparacao.rowcount;
    grdexportacao.Cells[0, k + 1] := 'Statistical Indexes of alternatives:';
    k := k + 1;
    // grdexportacao.TotalRows := K + 1 + grdcrivarWeight.ColumnCount;
    // grdexportacao.Cells[0,k] := 'Weights Variation percentage values:';
    for i := 0 to grdestatisticaalternativa.rowcount - 1 do
      for j := 0 to grdestatisticaalternativa.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdestatisticaalternativa.Cell[i, j].Text;
      end;
    k := k + 2 + grdestatisticaalternativa.rowcount;

    TIWAdvWebGridExcelIO1.AdvWebGrid := grdexportacao;
    TIWAdvWebGridExcelIO1.XLSExport('Result Report' + '.xls', WebApplication);

  end;
end;

// Transforma os valores da consMatrix em uma escala de razão
procedure TIWForm1.EscalaRazao;
var
  i, j: Integer;
  Maxl, Minl, MaxA, MinA, MaxB, MinB: Array of Double;
  AuxVetPOA: array of Integer;
  // ConsRazao: array of array of Double;
begin

  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(MatrizConseqNorm, NumAlt, numcrit);
    Setlength(Maxl, numcrit);
    Setlength(Minl, numcrit);

    Setlength(MaxA, numcrit);
    Setlength(MinA, numcrit);
    Setlength(MaxB, numcrit);
    Setlength(MinB, numcrit);

    // ** Buscando maximos e minimos locais (Escala de Razão) **//
    for j := 0 to numcrit - 1 do
    begin
      Maxl[j] := MatrizConseq[0, j];
      Minl[j] := MatrizConseq[0, j];
    end;
    for j := 0 to numcrit - 1 do
      for i := 0 to NumAlt - 1 do
      begin
        if MatrizConseq[i, j] > Maxl[j] then
          Maxl[j] := MatrizConseq[i, j];
        if MatrizConseq[i, j] < Minl[j] then
          Minl[j] := MatrizConseq[i, j];
      end;

    for j := 0 to numcrit - 1 do
      for i := 0 to NumAlt - 1 do
      begin
        if (TipoCrit[j] = 1) or (TipoCrit[j] = 3) or (TipoCrit[j] = 5) then
        begin // dos criterios de maximização

          // * ConsequenciaPositivas ou zero *//
          if (Maxl[j] > 0) and (Minl[j] >= 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              (TIWUserSession(WebApplication.Data).MatrizConseq[i, j] / Maxl[j]
              );
          // fim caso para consequencias são diferente de zero

          // * Consequencias mistas (+, -, 0)*//
          if (Maxl[j] >= 0) and (Minl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] /
              (Maxl[j] - Minl[j]) + 1 - Maxl[j] / (Maxl[j] - Minl[j]);

          // * Consequencias negativas *//
          if (Minl[j] < 0) and (Maxl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] := Maxl
              [j] / TIWUserSession(WebApplication.Data).MatrizConseq[i, j];

        end; // Fimdoscritérios de maximização

        if (TipoCrit[j] = 0) or (TipoCrit[j] = 2) or (TipoCrit[j] = 4) then
        begin // inicio dos critérios de minimização

          // * ConsequenciaPositivas *//
          if (MatrizConseq[i, j] <> 0) and (Minl[j] > 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              (Minl[j] / TIWUserSession(WebApplication.Data)
                .MatrizConseq[i, j]);
          // para consequencias diferentes de zero

          // * Consequencias mistas (+, -, 0)*//
          if (Minl[j] <= 0) and (Maxl[j] >= 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] /
              (Minl[j] - Maxl[j]) + 1 - Minl[j] / (Minl[j] - Maxl[j]);

          // * Consequencias negativas *//
          if (Minl[j] < 0) and (Maxl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] / Minl[j];

        end; // fim dos critérios de minimização
      end; // i

  end; // with
end; // fim do procedimento

procedure TIWForm1.SetarCores;
Var
  i, j: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Setlength(coresalt, 130);
    Setlength(coresalt1, TIWUserSession(WebApplication.Data).NumAlt);
    coresalt[0] := $000000CC; // vermelho
    coresalt[1] := $0033BBFF; // amarelo
    coresalt[2] := $00CC9900; // azul
    coresalt[3] := $00BBBBFF; // rosa
    coresalt[4] := $005AE6C8; // verde
    coresalt[5] := $00FF00FF;
    coresalt[6] := $00808000;
    coresalt[7] := $00800000;
    coresalt[8] := $00000080;
    coresalt[9] := $0000FF00;
    coresalt[10] := $00008080;
    coresalt[11] := $00800080;
    coresalt[12] := $00FFFF00;
    coresalt[13] := $00C0DCC0;
    coresalt[14] := $00A4A0A0;
    coresalt[15] := $00E049BB;
    coresalt[16] := $005353FF;
    coresalt[17] := $000080FF;
    coresalt[18] := $00800040;
    coresalt[19] := $00434321;
    coresalt[20] := $00189AD8;
    coresalt[21] := $000000C1;
    coresalt[22] := $00DC0E65;
    coresalt[23] := $0000AE00;
    coresalt[24] := $002676CE;
    coresalt[25] := $008000FF; // índigo
    coresalt[26] := $00808080; // cinza claro
    coresalt[27] := $00FF4500; // laranja avermelhado
    coresalt[28] := $00FFD700; // ouro claro
    coresalt[29] := $00640000; // marrom escuro
    coresalt[30] := $008B4513; // marrom de sela
    coresalt[31] := $00A52A2A; // marrom de fogo
    coresalt[32] := $00D2691E; // chocolate
    coresalt[33] := $008BB000; // marrom escuro
    coresalt[34] := $00B22222; // tijolo
    coresalt[35] := $00FF6347; // tomate
    coresalt[36] := $00FF8C00; // laranja escuro
    coresalt[37] := $00FFA07A; // salmão claro
    coresalt[38] := $00FF7F50; // coral claro
    coresalt[39] := $00FF0000; // vermelho claro
    coresalt[40] := $00CD5C5C; // marrom claro
    coresalt[41] := $008B0000; // marrom escuro
    coresalt[42] := $00A52A2A; // marrom de fogo
    coresalt[43] := $00B22222; // tijolo
    coresalt[44] := $008B0000; // marrom escuro
    coresalt[45] := $00A52A2A; // marrom de fogo
    coresalt[46] := $00B22222; // tijolo
    coresalt[47] := $008B0000; // marrom escuro
    coresalt[48] := $00A52A2A; // marrom de fogo
    coresalt[49] := $00B22222; // tijolo
    coresalt[50] := $008B0000; // marrom escuro
    coresalt[51] := $00A52A2A; // marrom de fogo
    coresalt[52] := $00B22222; // tijolo
    coresalt[53] := $008B0000; // marrom escuro
    coresalt[54] := $00A52A2A; // marrom de fogo
    coresalt[55] := $00B22222; // tijolo
    coresalt[56] := $00FF69B4; // rosa quente
    coresalt[57] := $00FFE4E1; // rosa claro
    coresalt[58] := $00FF1493; // rosa profundo
    coresalt[59] := $00C71585; // vermelho violeta
    coresalt[60] := $00DB7093; // roxo claro
    coresalt[61] := $00DA70D6; // orquídea
    coresalt[62] := $009932CC; // índigo escuro
    coresalt[63] := $009ACD32; // verde amarelo
    coresalt[64] := $00C0FF3E; // verde limão brilhante
    coresalt[65] := $00AEEEEE; // azul celeste
    coresalt[66] := $000000CC; // vermelho
    coresalt[67] := $0033BBFF; // amarelo
    coresalt[68] := $00CC9900; // azul
    coresalt[69] := $00BBBBFF; // rosa
    coresalt[70] := $005AE6C8; // verde
    coresalt[71] := $00FF00FF;
    coresalt[72] := $00808000;
    coresalt[73] := $00800000;
    coresalt[74] := $00000080;
    coresalt[75] := $0000FF00;
    coresalt[76] := $00008080;
    coresalt[77] := $00800080;
    coresalt[78] := $00FFFF00;
    coresalt[79] := $00C0DCC0;
    coresalt[80] := $00A4A0A0;
    coresalt[81] := $00E049BB;
    coresalt[82] := $005353FF;
    coresalt[83] := $000080FF;
    coresalt[84] := $00800040;
    coresalt[85] := $00434321;
    coresalt[86] := $00189AD8;
    coresalt[87] := $000000C1;
    coresalt[88] := $00DC0E65;
    coresalt[89] := $0000AE00;
    coresalt[90] := $002676CE;
    coresalt[91] := $008000FF; // índigo
    coresalt[92] := $00808080; // cinza claro
    coresalt[93] := $00FF4500; // laranja avermelhado
    coresalt[94] := $00FFD700; // ouro claro
    coresalt[95] := $00640000; // marrom escuro
    coresalt[96] := $008B4513; // marrom de sela
    coresalt[97] := $00A52A2A; // marrom de fogo
    coresalt[98] := $00D2691E; // chocolate
    coresalt[99] := $008B0000; // marrom escuro
    coresalt[100] := $00B22222; // tijolo
    coresalt[101] := $00FF6347; // tomate
    coresalt[102] := $00FF8C00; // laranja escuro
    coresalt[103] := $00FFA07A; // salmão claro
    coresalt[104] := $00FF7F50; // coral claro
    coresalt[105] := $00FF0000; // vermelho claro
    coresalt[106] := $00CD5C5C; // marrom claro
    coresalt[107] := $008B0000; // marrom escuro
    coresalt[108] := $00A52A2A; // marrom de fogo
    coresalt[109] := $00B22222; // tijolo
    coresalt[110] := $008B0000; // marrom escuro
    coresalt[111] := $00A52A2A; // marrom de fogo
    coresalt[112] := $00B22222; // tijolo
    coresalt[113] := $008B0000; // marrom escuro
    coresalt[114] := $00A52A2A; // marrom de fogo
    coresalt[115] := $00B22222; // tijolo
    coresalt[116] := $008B0000; // marrom escuro
    coresalt[117] := $00A52A2A; // marrom de fogo
    coresalt[118] := $00B22222; // tijolo
    coresalt[119] := $008B0000; // marrom escuro
    coresalt[120] := $00A52A2A; // marrom de fogo
    coresalt[121] := $00B22222; // tijolo
    coresalt[122] := $00FF69B4; // rosa quente
    coresalt[123] := $00FFE4E1; // rosa claro
    coresalt[124] := $00FF1493; // rosa profundo
    coresalt[125] := $00C71585; // vermelho violeta
    coresalt[126] := $00DB7093; // roxo claro
    coresalt[127] := $00DA70D6; // orquídea
    coresalt[128] := $009932CC; // índigo escuro
    coresalt[129] := $009ACD32; // verde amarelo
  end;

end;

procedure TIWForm1.ProgramarComponentes;
var
  i, j, JJ, k, qtdpareto, S, r: Integer;
  auxregra, Aux1, Aux2, aux3: real;
  bar, barpoa: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Naltsel: Integer;
  Px: real;
  faixaprob: array of real;
  pular: array of boolean;
  pular2: boolean;
  auxresultado: real;
  Linha: array of TBarSeries;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  SetLength(GanhoMedio, numalt);
  SetLength(GanhoMaximo, Numalt);
  SetLength(PerdaMedia, Numalt);
  SetLength(PerdaMaxima, Numalt);
  Setlength(probabilidadealt, Numalt);

    for j := 0 to NumAlt - 1 do
    begin
      Mediavgmax[j] := 0;
      MediaVgMin[j] := 0;
      MaxVgMax[j] := 0;
      MaxVGMin[j] := 0;
      MinVgMax[j] := 9999999999;
      MinVGMin[j] := 9999999999;
      MediaAmplitudeVG[j] := 0;
      MaxAmplitudeVG[j] := 0;
      MinAmplitudeVG[j] := 999999999;
    end;
    resultsol[0] := 1;
    for i := 0 to length(CasesOrdemcrit) - 1 do
    begin
      PLEscolha(i);
      PlOrdenacao(i);
      for j := 0 to NumAlt - 1 do
      begin
        AmplitudeVG[i, j] := Vgmax[i, j] - VGMin[i, j];
        TIWUserSession(WebApplication.Data).Mediavgmax[j] := Mediavgmax[j]
          + Vgmax[i, j];
        TIWUserSession(WebApplication.Data).MediaVgMin[j] := MediaVgMin[j]
          + VGMin[i, j];
        if MaxVgMax[j] < Vgmax[i, j] then
          TIWUserSession(WebApplication.Data).MaxVgMax[j] := Vgmax[i, j];
        if MaxVGMin[j] < VGMin[i, j] then
          TIWUserSession(WebApplication.Data).MaxVGMin[j] := VGMin[i, j];

        if MinVgMax[j] > Vgmax[i, j] then
          TIWUserSession(WebApplication.Data).MinVgMax[j] := Vgmax[i, j];
        if MinVGMin[j] > VGMin[i, j] then
          TIWUserSession(WebApplication.Data).MinVGMin[j] := VGMin[i, j];

        MediaAmplitudeVG[j] := MediaAmplitudeVG[j] + AmplitudeVG[i, j];

        if MaxAmplitudeVG[j] < AmplitudeVG[i, j] then
          TIWUserSession(WebApplication.Data).MaxAmplitudeVG[j] := AmplitudeVG
            [i, j];
        if MinAmplitudeVG[j] > AmplitudeVG[i, j] then
          TIWUserSession(WebApplication.Data).MinAmplitudeVG[j] := AmplitudeVG
            [i, j];

        if MinAmplitudeVG[j] < 0 then
          MinAmplitudeVG[j] := 0;

      end;

      for j := 0 to NumAlt - 2 do
        for k := j + 1 to NumAlt - 1 do
        begin
          if j = k then
            continue;

          { if (vgmax[i,j] >= vgmax[i,k]) and (vgmin[i,j] >= vgmin[i,k]) then
            begin
            amplitudeparaparinst[i,j,k] := vgmax[i,j] + vgmin[i,j];
            //amplitudeparaparinst[i,k,j] := vgmax[i,j] + vgmin[i,j];
            end;
            if (vgmax[i,j] >= vgmax[i,k]) and (vgmin[i,j] <= vgmin[i,k]) then
            begin
            amplitudeparaparinst[i,j,k] := vgmax[i,j] + vgmin[i,k];
            //amplitudeparaparinst[i,k,j] := vgmax[i,j] + vgmin[i,k];
            end;
            if (vgmax[i,j] <= vgmax[i,k]) and (vgmin[i,j] >= vgmin[i,k]) then
            begin
            amplitudeparaparinst[i,j,k] := vgmax[i,k] + vgmin[i,j];
            //amplitudeparaparinst[i,k,j] := vgmax[i,k] + vgmin[i,j];
            end;
            if (vgmax[i,j] <= vgmax[i,k]) and (vgmin[i,j] <= vgmin[i,k]) then
            begin
            amplitudeparaparinst[i,j,k] := vgmax[i,k] + vgmin[i,k];
            //amplitudeparaparinst[i,k,j] := vgmax[i,k] + vgmin[i,k];
            end;
            {if vgmax[i,j] >= vgmax[i,k] then
            probvitoriaparapar[i,j,k] := probvitoriaparapar[i,j,k] + vgmax[i,j] - vgmax[i,k];
            if vgmin[i,j] >= vgmin[i,k] then
            probvitoriaparapar[i,j,k] := probvitoriaparapar[i,j,k] + vgmin[i,j] - vgmin[i,k];
            }

          for JJ := 0 to 1000 - 1 do
          begin
            difvg[i, j, k] := difvg[i, j, k] +
              (VGMin[i, j] + random * (Vgmax[i, j] - VGMin[i, j])) -
              (VGMin[i, k] + random * (Vgmax[i, k] - VGMin[i, k]));
          end;
          TIWUserSession(WebApplication.Data).difvg[i, j, k] := difvg[i, j, k]
            / 1000;
          if difvg[i, j, k] >= 0 then
            mediadifvg[j, k] := mediadifvg[j, k] + difvg[i, j, k];
          if difvg[i, j, k] < 0 then
            mediadifvg[k, j] := mediadifvg[k, j] - difvg[i, j, k];

          if difvg[i, j, k] > maximodifvg[j, k] then
            maximodifvg[j, k] := difvg[i, j, k];

          if (-difvg[i, j, k]) > maximodifvg[k, j] then
            TIWUserSession(WebApplication.Data).maximodifvg[k, j] := -difvg
              [i, j, k];

          TIWUserSession(WebApplication.Data).ProbVitoriaParapar[i, j, k] :=
            Probabilidade_AMaiorQueB(VGMin[i, j], Vgmax[i, j], VGMin[i, k],
            Vgmax[i, k]);
          TIWUserSession(WebApplication.Data).ProbVitoriaParapar[i, k, j] :=
            1 - TIWUserSession(WebApplication.Data).ProbVitoriaParapar[i, j, k];

          mediaprobvitoriaparapar[j, k] := mediaprobvitoriaparapar[j, k]
            + ProbVitoriaParapar[i, j, k];
          mediaprobvitoriaparapar[k, j] := mediaprobvitoriaparapar[k, j]
            + ProbVitoriaParapar[i, k, j];

          // amplitudeparapar[j,k] := amplitudeparapar[j,k] + amplitudeparaparinst[i,j,k];
          if ProbVitoriaParapar[i, j, k] > maximoprobvitoriaparapar[j, k] then
            maximoprobvitoriaparapar[j, k] := ProbVitoriaParapar[i, j, k];

          if ProbVitoriaParapar[i, k, j] > maximoprobvitoriaparapar[k, j] then
            maximoprobvitoriaparapar[k, j] := ProbVitoriaParapar[i, k, j];
        end;

    end;

     Naltsel := 0;
    for i := 0 to length(AltPOA) - 1 do
    begin
      if AltPOA[i] <> 0 then
        INC(Naltsel);
    end;

    Setlength(ordemalt, NumAlt);
    Setlength(pular, NumAlt);
    for i := 0 to NumAlt - 1 do
      pular[i] := false;

    for i := 0 to NumAlt - 1 do
    begin
      auxresultado := 0;
      for j := 0 to NumAlt - 1 do
      begin
        if pular[j] = true then
          continue;
        if TIWUserSession(WebApplication.Data).AltPOA[j] > auxresultado then
          auxresultado := AltPOA[j];
      end;
      pular2 := false;
      for j := 0 to NumAlt - 1 do
      begin
        if (pular2 = true) or (pular[j] = true) then
          continue;
        if auxresultado = AltPOA[j] then
        begin
          pular[j] := true;

          TIWUserSession(WebApplication.Data).ordemalt[i] := j;
          pular2 := true;
        end;

      end;
    end;




    for i := 0 to naltsel - 1 do
    begin
      TIWUserSession(WebApplication.Data).probabilidadealt[ordemalt[i]] := TIWUserSession(WebApplication.Data).probabilidadealt[ordemalt[i]] / length(CasesOrdemcrit);
      for j := 0 to naltsel - 1 do
      begin
        //TIWUserSession(WebApplication.Data).GanhoMedio[i] := GanhoMedio[i] + mediamaxdiff[i,j];
        //TIWUserSession(WebApplication.Data).Perdamedia[i] := PerdaMedia[i] + mediamaxdiff[j,i];

        TIWUserSession(WebApplication.Data).mediamaxdiff[ordemalt[i],ordemalt[j]] := mediamaxdiff[ordemalt[i],ordemalt[j]] / length(CasesOrdemcrit);

        if maxmaxdiff[ordemalt[i],ordemalt[j]] > GanhoMaximo[ordemalt[i]] then
        TIWUserSession(WebApplication.Data).GanhoMaximo[ordemalt[i]] := maxmaxdiff[ordemalt[i],ordemalt[j]];

        if maxmaxdiff[ordemalt[j],ordemalt[i]] > PerdaMaxima[ordemalt[i]] then
        TIWUserSession(WebApplication.Data).PerdaMaxima[ordemalt[i]] := maxmaxdiff[ordemalt[j],ordemalt[i]];

      end;
      //TIWUserSession(WebApplication.Data).GanhoMedio[i] := GanhoMedio[i] / (length(CasesOrdemcrit) * (Numalt-1));
      //TIWUserSession(WebApplication.Data).PerdaMedia[i] := PerdaMedia[i] / (length(CasesOrdemcrit) * (Numalt-1));
    end;

    for i := 0 to naltsel - 1 do
    begin
      for j := 0 to naltsel - 1 do
      begin
       if i = j then
       continue;

       GanhoMedio[ordemalt[i]] := GanhoMedio[ordemalt[i]] + mediamaxdiff[ordemalt[i],ordemalt[j]];
       PerdaMedia[ordemalt[i]] := PerdaMedia[ordemalt[i]] + mediamaxdiff[ordemalt[j],ordemalt[i]];
      end;
      GanhoMedio[ordemalt[i]] := GanhoMedio[ordemalt[i]] / (Naltsel-1);
      PerdaMedia[ordemalt[i]] := PerdaMedia[ordemalt[i]] / (Naltsel-1);
    end;

    for j := 0 to NumAlt - 1 do
    begin
      TIWUserSession(WebApplication.Data).Mediavgmax[j] := Mediavgmax[j]
        / length(CasesOrdemcrit);
      TIWUserSession(WebApplication.Data).MediaVgMin[j] := MediaVgMin[j]
        / length(CasesOrdemcrit);
      TIWUserSession(WebApplication.Data).MediaAmplitudeVG[j] :=
        MediaAmplitudeVG[j] / length(CasesOrdemcrit);
    end;
    for i := 0 to length(CasesOrdemcrit) - 1 do
      for j := 0 to NumAlt - 2 do
        for k := j + 1 to NumAlt - 1 do
        begin
          if j = k then
            continue;

          if MatrizAltSel[i, j] > MatrizAltSel[i, k] then
            MatrizComparacaoAlt[j, k] := MatrizComparacaoAlt[j, k] + 1;
          if MatrizAltSel[i, j] < MatrizAltSel[i, k] then
            MatrizComparacaoAlt[k, j] := MatrizComparacaoAlt[k, j] + 1;

          if (MatrizAltSel[i, j] = 1) and (MatrizAltSel[i, k] = 1) then
          begin
            AmbasAltNaPOA[j, k] := AmbasAltNaPOA[j, k] + 1;
            AmbasAltNaPOA[k, j] := AmbasAltNaPOA[k, j] + 1;
          end;
          if (MatrizAltSel[i, j] = 0) and (MatrizAltSel[i, k] = 0) then
          begin
            NenhumaAltNaPOA[j, k] := NenhumaAltNaPOA[j, k] + 1;
            NenhumaAltNaPOA[k, j] := NenhumaAltNaPOA[k, j] + 1;
          end;
        end;

    Setlength(mediageralprobalt, NumAlt);
    auxresultado := 0;
    for j := 0 to NumAlt - 1 do
    begin

      for k := 0 to NumAlt - 1 do
      begin

        if mediaprobvitoriaparapar[j, k] <> 0 then
          TIWUserSession(WebApplication.Data).mediadifvg[j, k] := mediadifvg
            [j, k] / (mediaprobvitoriaparapar[j, k])
        else
          TIWUserSession(WebApplication.Data).mediadifvg[j, k] := 0;

        TIWUserSession(WebApplication.Data).mediaprobvitoriaparapar[j, k] :=
          mediaprobvitoriaparapar[j, k] / length(CasesOrdemcrit);

        if MaximoGeralProbVitoriaAlt[j] < TIWUserSession(WebApplication.Data)
          .maximoprobvitoriaparapar[j, k] then
          TIWUserSession(WebApplication.Data).MaximoGeralProbVitoriaAlt[j] :=
            maximoprobvitoriaparapar[j, k];

      end;
      // auxresultado := auxresultado + mediageralprobalt[j];
    end;

    Setlength(OrdemSol, length(matrizsol));
    Setlength(pular, length(matrizsol));
    for i := 0 to length(matrizsol) - 1 do
      pular[i] := false;

    for i := 0 to length(matrizsol) - 1 do
    begin
      auxresultado := 0;
      for j := 0 to length(matrizsol) - 1 do
      begin
        if pular[j] = true then
          continue;
        if TIWUserSession(WebApplication.Data).resultsol[j] > auxresultado then
          auxresultado := resultsol[j];
      end;
      pular2 := false;
      for j := 0 to length(matrizsol) - 1 do
      begin
        if (pular2 = true) or (pular[j] = true) then
          continue;
        if auxresultado = resultsol[j] then
        begin
          pular[j] := true;
          TIWUserSession(WebApplication.Data).OrdemSol[i] := j;
          pular2 := true;
        end;

      end;
    end;




    auxresultado := 0;
    for j := 0 to Naltsel - 1 do
    begin
      mediageralprobalt[ordemalt[j]] := 1;
      for k := 0 to Naltsel - 1 do
      begin
        if TIWUserSession(WebApplication.Data).ordemalt[j] = ordemalt[k] then
          continue;
        mediageralprobalt[ordemalt[j]] := mediageralprobalt[ordemalt[j]]
          * TIWUserSession(WebApplication.Data).mediaprobvitoriaparapar
          [ordemalt[j], ordemalt[k]];
      end;
      auxresultado := auxresultado + mediageralprobalt[ordemalt[j]];
    end;
    for j := 0 to NumAlt - 1 do
      TIWUserSession(WebApplication.Data).mediageralprobalt[ordemalt[j]] :=
        mediageralprobalt[ordemalt[j]] / auxresultado;

    // Surrogate;
    SetarCores;
    // ResultadosPromethee;
    // ComparacaoMetodos;
    chrtResultado.RemoveAllSeries;
    chrtResultado.Legend.Items.Clear;
    // chrtresultado.ClearChart;
    chrtResultado.BottomAxis.Items.Clear;
    chrtResultado.View3D := false;
    Setlength(bar, NumAlt);
    i := 0;
    Total := length(CasesOrdemcrit);
    totalacumulado := 0;
    for j := 0 to NumAlt - 1 do
    begin
      if AltPOA[ordemalt[j]] <> 0 then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[i]);
        bar[i].AddXY(0, AltPOA[ordemalt[j]], formatfloat
            ('0', AltPOA[ordemalt[j]]), coresalt[i]);
        // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
        chrtResultado.Legend.Item[i].Text := NomeAlt[ordemalt[j]];
        bar[i].Color := coresalt[i];
        // totalacumulado := totalacumulado + AltPOA[Ordemalt[j]];
        INC(i);
      end;
    end;
    // Total := 0;
    bar[1] := TBarSeries.Create(self);
    chrtResultado.AddSeries(bar[1]);
    bar[1].AddXY(0, Total, '', clWhite);
    bar[1].Visible := false;
    // chrtResultado.LeftAxis.Increment := 0.2; // ?
    // chrtResultado.Legend.Visible := false;
    chrtResultado.BottomAxis.LabelsAlternate;
    chrtResultado.BottomAxis.Items.Automatic := false;
    // chrtResultado.LeftAxis.Maximum := Total;
    chrtResultado.LeftAxis.AutomaticMaximum := true;
    /// chrtResultado.LeftAxis.MaximumOffset := 16;
    // chrtResultado.LeftAxis.Minimum := 0;
    // chrtresultado.leftaxis.Items.Automatic := false;
    // chrtResultado.LeftAxis.Increment := Total;

    imgAlternative.Picture.Bitmap.Assign(chrtResultado.TeeCreateBitmap
        (clWebPURPLE, rect(0, 0, imgAlternative.Width, imgAlternative.Height)));

    Total := length(CasesOrdemcrit);
    totalacumulado := 0;
    Setlength(barpoa, 2);
    i := 0;
    for j := 0 to length(resultsol) - 1 do
    begin
      if (totalacumulado > 0.8 * Total) or (i > 10) then
      begin
        barpoa[i] := TBarSeries.Create(self);
        chrtPOA.AddSeries(barpoa[i]);
        barpoa[i].AddXY(0, Total - totalacumulado, formatfloat
            ('0', Total - totalacumulado));
        // chrtpoa.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

        chrtPOA.Legend.Item[i].Text := 'Other POAs';
        break;
      end;
      if resultsol[OrdemSol[j]] <> 0 then
      begin
        barpoa[i] := TBarSeries.Create(self);
        chrtPOA.AddSeries(barpoa[i]);
        barpoa[i].AddXY(0, resultsol[OrdemSol[j]], formatfloat
            ('0', resultsol[OrdemSol[j]]));
        // chrtpoa.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
        chrtPOA.Legend.Item[i].Text := 'POA Subset ' + inttostr(i + 1);
        totalacumulado := totalacumulado + resultsol[OrdemSol[j]];
        INC(i);
      end;
    end;

    // chrtpoa.LeftAxis.Increment := 0.2; // ?
    // chrtpoa.Legend.Visible := false;
    chrtPOA.BottomAxis.LabelsAlternate;
    chrtPOA.BottomAxis.Items.Automatic := false;
    // chrtpoa.LeftAxis.Maximum := Total;
    chrtPOA.LeftAxis.AutomaticMaximum := true;
    /// chrtpoa.LeftAxis.MaximumOffset := 16;
    // chrtpoa.LeftAxis.Minimum := 0;
    // chrtpoa.leftaxis.Items.Automatic := false;
    // chrtpoa.LeftAxis.Increment := Total;

    imgPOA.Picture.Bitmap.Assign
      (chrtPOA.TeeCreateBitmap(clWebPURPLE, rect(0, 0, imgPOA.Width,
          imgPOA.Height)));

    chrtComparacao.RemoveAllSeries;
    chrtComparacao.Legend.Items.Clear;
    // chrtcomparacao.ClearChart;
    chrtComparacao.BottomAxis.Items.Clear;
    chrtComparacao.View3D := false;
    Setlength(Linha, NumAlt + 1);
    for i := 0 to NumAlt - 1 do
    begin
      if AltPOA[ordemalt[i]] <> 0 then
      begin
        Linha[i] := TBarSeries.Create(self);
        // linha[i].LinePen.Width := 3;
        Linha[i].Legend.Text := NomeAlt[ordemalt[i]];
        chrtComparacao.AddSeries(Linha[i]);
        // linha[i].AddXY(j, AltTamanhoPOA[OrdemSol[i],j], formatfloat('0', AltTamanhoPOA[OrdemSol[i],j]);
      end;
    end;
    Linha[NumAlt] := TBarSeries.Create(self);
    // linha[numalt].LinePen.Width := 3;
    // linha[numalt].linepen.Color := clwebblack;
    Linha[NumAlt].Legend.Text := 'Total';
    // linha[numalt].Legend.Color := clwebblack;
    chrtComparacao.AddSeries(Linha[NumAlt]);

    Linha[NumAlt].Marks.Visible := false;

    JJ := 0;
    for j := 0 to NumAlt - 2 do
    begin
      if numsolinst[j] <> 0 then
      begin
        Linha[NumAlt].AddXY(JJ, numsolinst[j], formatfloat('0', numsolinst[j]));
        for i := 0 to NumAlt - 1 do
        begin
          if AltPOA[ordemalt[i]] <> 0 then
          begin
            Linha[i].AddXY(JJ, AltTamanhoPOA[ordemalt[i], j]);
            Linha[i].Marks.Visible := false;

          end;
        end;
        chrtComparacao.BottomAxis.Items.Add(JJ, inttostr(j + 1) + ' Alt.');
        INC(JJ);
      end;
    end;
    // chrtcomparacao.LeftAxis.Increment := 0.2; // ?
    // chrtcomparacao.Legend.Visible := false;
    chrtComparacao.BottomAxis.LabelsAlternate;
    chrtComparacao.BottomAxis.Items.Automatic := false;
    // chrtcomparacao.LeftAxis.Maximum := Total;
    chrtComparacao.LeftAxis.AutomaticMaximum := true;
    /// chrtcomparacao.LeftAxis.MaximumOffset := 16;
    // chrtcomparacao.LeftAxis.Minimum := 0;
    // chrtcomparacao.leftaxis.Items.Automatic := false;
    // chrtcomparacao.LeftAxis.Increment := Total;

    imgcomparacao.Picture.Bitmap.Assign(chrtComparacao.TeeCreateBitmap
        (clWebPURPLE, rect(0, 0, imgcomparacao.Width, imgcomparacao.Height)));

    grdresultados.Clear;

    grdresultados.rowcount := 3 + length(matrizsol);
    grdresultados.ColumnCount := 3 + Naltsel;
    { if grdresultados.ColumnCount * 80 > 832 then
      grdresultados.Width := grdresultados.ColumnCount * 80
      else
      grdresultados.Width := 832; }
    grdresultados.Cell[0, 0].Text := 'POA Subsets';
    for i := 0 to length(matrizsol) - 1 do
    begin
      grdresultados.Cell[i + 1, 0].Text := 'POA Subset ' + inttostr(i + 1)
        + ':';
      grdresultados.Cell[i + 1, Naltsel + 1].Text := formatfloat
        ('0', resultsol[OrdemSol[i]]);
      grdresultados.Cell[i + 1, Naltsel + 2].Text := formatfloat
        ('0', resultsol[OrdemSol[i]] / Total * 100) + '%';

      for j := 0 to Naltsel - 1 do
      begin
        if matrizsol[i, ordemalt[j]] = 1 then
          grdresultados.Cell[i + 1, j + 1].Text := 'X';
      end;
    end;

    for j := 0 to Naltsel - 1 do
    begin
      grdresultados.Cell[0, j + 1].Text := NomeAlt[ordemalt[j]];
      grdresultados.Cell[length(matrizsol) + 1, j + 1].Text := formatfloat
        ('0', AltPOA[ordemalt[j]]);
      grdresultados.Cell[length(matrizsol) + 2, j + 1].Text := formatfloat
        ('0', AltPOA[ordemalt[j]] / Total * 100) + '%';
    end;

    grdresultados.Cell[0, Naltsel + 1].Text := 'Frequency';
    grdresultados.Cell[0, Naltsel + 2].Text := '%';

    grdresultados.Cell[length(matrizsol) + 1, 0].Text := 'Frequency';
    grdresultados.Cell[length(matrizsol) + 2, 0].Text := '%';
    i := 0;

    for j := 0 to grdresultados.ColumnCount - 1 do
      for i := 0 to grdresultados.rowcount - 1 do
      begin
        grdresultados.Cell[i, j].Alignment := taCenter;
        grdresultados.Cell[i, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;

    for j := 0 to grdresultados.ColumnCount - 1 do
    begin
      grdresultados.Cell[0, j].Font.Style := [fsbold];
      grdresultados.Cell[1 + length(matrizsol), j].Font.Color := clwebgray;
      grdresultados.Cell[1 + length(matrizsol), j].BGColor := clWebWhiteSmoke;

      grdresultados.Cell[2 + length(matrizsol), j].Font.Color := clwebgray;
      grdresultados.Cell[2 + length(matrizsol), j].BGColor := clWebWhiteSmoke;
    end;
    grdresultados.Cell[1 + length(matrizsol), 0].Font.Style := [fsbold];
    grdresultados.Cell[2 + length(matrizsol), 0].Font.Style := [fsbold];

    for i := 0 to grdresultados.rowcount - 1 do
    begin
      grdresultados.Cell[i, Naltsel + 1].Font.Color := clwebgray;
      grdresultados.Cell[i, Naltsel + 2].BGColor := clWebWhiteSmoke;
    end;

    grdcomparacao.rowcount := Naltsel + 2;
    grdcomparacao.ColumnCount := 1 + NumAlt;

    for j := 0 to grdcomparacao.ColumnCount - 1 do
    begin
      grdcomparacao.Cell[0, j].Font.Style := [fsbold];
      for i := 0 to grdcomparacao.rowcount - 1 do
      begin
        grdcomparacao.Cell[i, j].Alignment := taCenter;
        grdcomparacao.Cell[i, j].wrap := true;
        //
      end;
    end;
    for i := 0 to Naltsel - 1 do
    begin
      grdcomparacao.Cell[i + 1, 0].Text := NomeAlt[ordemalt[i]];
      grdcomparacao.Cell[i + 1, NumAlt].Text := formatfloat
        ('0', AltPOA[ordemalt[i]]);
      JJ := 0;
      for j := 0 to NumAlt - 2 do
      begin
        if numsolinst[j] <> 0 then
        begin
          grdcomparacao.Cell[i + 1, JJ + 1].Text := formatfloat
            ('0', AltTamanhoPOA[ordemalt[i], j]);
          INC(JJ);
        end;
      end;
    end;

    JJ := 0;
    for j := 0 to NumAlt - 2 do
    begin
      if numsolinst[j] <> 0 then
      begin
        if j = 0 then
          grdcomparacao.Cell[0, JJ + 1].Text := 'Included alone in POA';
        if j = 1 then
          grdcomparacao.Cell[0, JJ + 1].Text :=
            'Included with 1 other alt in POA';
        if j > 1 then
          grdcomparacao.Cell[0, JJ + 1].Text := 'Included with ' + inttostr
            (j + 1) + ' other alts in POA';

        grdcomparacao.Cell[Naltsel + 1, j + 1].Text := formatfloat
          ('0', numsolinst[j]);
        INC(JJ);
      end;
    end;
    grdcomparacao.Cell[0, 0].Text := 'Alternatives';
    grdcomparacao.Cell[Naltsel + 1, 0].Text := 'Total:';
    grdcomparacao.Cell[0, jj+1].Text := 'Total:';

    grdcomparacao.Cell[Naltsel + 1, 0].Font.Style := [fsbold];

    grdcomparacao.ColumnCount := 1 + JJ + 1;

    grdcases.rowcount := Round(Total) + 1;
    grdcases.ColumnCount := 2 + numcrit;

    for j := 0 to grdcases.ColumnCount - 1 do
      for i := 0 to grdcases.rowcount - 1 do
      begin
        grdcases.Cell[i, j].Alignment := taCenter;
        grdcases.Cell[i, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcases.ColumnCount - 1 do
      grdcases.Cell[0, j].Font.Style := [fsbold];

    grdcases.Cell[0, 0].Text := 'Order';
    for i := 0 to numcrit - 1 do
      grdcases.Cell[0, i + 1].Text := NomeCrit[i];

    grdcases.Cell[0, numcrit + 1].Text := 'POA Subset';

    for j := 0 to length(CasesOrdemcrit) - 1 do
    begin
      grdcases.Cell[j + 1, 0].Text := 'O' + inttostr(j + 1);
      for i := 0 to NumAlt - 1 do
      begin

        if TIWUserSession(WebApplication.Data).MatrizAltSel[j, i] = 1 then
        begin
          if grdcases.Cell[j + 1, numcrit + 1].Text = '' then
            grdcases.Cell[j + 1, numcrit + 1].Text := NomeAlt[i]
          else
            grdcases.Cell[j + 1, numcrit + 1].Text := grdcases.Cell
              [j + 1, numcrit + 1].Text + ', ' + NomeAlt[i];
        end;
      end;
      for i := 0 to numcrit - 1 do
        grdcases.Cell[j + 1, i + 1].Text := inttostr(CasesOrdemcrit[j, i]) + 'º'
    end;

    cmbbxAlternativa.Clear;
    for i := 0 to Naltsel - 1 do
    begin
      cmbbxAlternativa.Items.Add(NomeAlt[ordemalt[i]]);
    end;
    cmbbxAlternativa.ItemIndex := 0;

    cmbbxAlternativaChange(self);

    grdestatisticaalternativa.Clear;
    grdestatisticaalternativa.ColumnCount := 11;
    grdestatisticaalternativa.rowcount := 1 + Naltsel;

    grdestatisticaalternativa.Cell[0, 0].Text := 'Alternative';
    grdestatisticaalternativa.Cell[0, 1].Text := 'Probability';
    grdestatisticaalternativa.Cell[0, 2].Text := 'μ Vmax';
    grdestatisticaalternativa.Cell[0, 3].Text := 'μ Vmin';
    grdestatisticaalternativa.Cell[0, 4].Text := 'μ range';
    grdestatisticaalternativa.Cell[0, 5].Text := 'Max Vmax';
    grdestatisticaalternativa.Cell[0, 6].Text := 'Max Vmin';
    grdestatisticaalternativa.Cell[0, 7].Text := 'Max range';
    grdestatisticaalternativa.Cell[0, 8].Text := 'Min Vmax';
    grdestatisticaalternativa.Cell[0, 9].Text := 'Min Vmin';
    grdestatisticaalternativa.Cell[0, 10].Text := 'Min range';

    for i := 0 to Naltsel - 1 do
    begin
      grdestatisticaalternativa.Cell[i + 1, 0].Text := NomeAlt[ordemalt[i]];
      grdestatisticaalternativa.Cell[i + 1, 1].Text := formatfloat
        ('0', probabilidadealt[ordemalt[i]] * 100) + '%';
      grdestatisticaalternativa.Cell[i + 1, 2].Text := formatfloat
        ('0.00', Mediavgmax[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 3].Text := formatfloat
        ('0.00', MediaVgMin[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 4].Text := formatfloat
        ('0.00', MediaAmplitudeVG[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 5].Text := formatfloat
        ('0.00', MaxVgMax[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 6].Text := formatfloat
        ('0.00', MaxVGMin[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 7].Text := formatfloat
        ('0.00', MaxAmplitudeVG[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 8].Text := formatfloat
        ('0.00', MinVgMax[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 9].Text := formatfloat
        ('0.00', MinVGMin[ordemalt[i]]);
      grdestatisticaalternativa.Cell[i + 1, 10].Text := formatfloat
        ('0.00', MinAmplitudeVG[ordemalt[i]]);
    end;

    for j := 0 to grdestatisticaalternativa.ColumnCount - 1 do
      for i := 0 to grdestatisticaalternativa.rowcount - 1 do
      begin
        grdestatisticaalternativa.Cell[i, j].Alignment := taCenter;
        grdestatisticaalternativa.Cell[i, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdestatisticaalternativa.ColumnCount - 1 do
      grdestatisticaalternativa.Cell[0, j].Font.Style := [fsbold];

    grdprobvitoria.Clear;
    grdprobvitoria.rowcount := 1 + Naltsel;
    grdprobvitoria.ColumnCount := 2 + Naltsel;
    grdprobvitoria.Cell[0, Naltsel + 1].Text := 'Global Probability';
    grdprobvitoria.Cell[0, Naltsel + 1].Font.Style := [fsbold];
    grdprobvitoria.Cell[0, Naltsel + 1].Font.Color := ClWebRed;
    for i := 0 to Naltsel - 1 do
    begin
      grdprobvitoria.Cell[0, i + 1].Text := NomeAlt[ordemalt[i]];
      grdprobvitoria.Cell[i + 1, 0].Text := NomeAlt[ordemalt[i]];
      grdprobvitoria.Cell[i + 1, Naltsel + 1].Text := formatfloat
        ('0.00', mediageralprobalt[ordemalt[i]] * 100) + '%';

      grdprobvitoria.Cell[0, i + 1].Font.Style := [fsbold];
      grdprobvitoria.Cell[i + 1, 0].Font.Style := [fsbold];
      grdprobvitoria.Cell[i + 1, Naltsel + 1].Font.Color := ClWebRed;
    end;
    for j := 0 to grdprobvitoria.ColumnCount - 1 do
      for i := 0 to grdprobvitoria.rowcount - 1 do
      begin
        grdprobvitoria.Cell[i, j].Alignment := taCenter;
        grdprobvitoria.Cell[i, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for i := 0 to Naltsel - 1 do
      for j := 0 to Naltsel - 1 do
      begin
        if ordemalt[i] = ordemalt[j] then
          continue;
        grdprobvitoria.Cell[i + 1, j + 1].Text := formatfloat
          ('0.00', mediaprobvitoriaparapar[ordemalt[i], ordemalt[j]] * 100)
          + '%';
      end;



    Total := length(casesordemcrit);
    {for i := 0 to length(ResultSol) - 1 do
      Total := Total + ResultSol[i];
     }
    grdestatistica.Clear;
    grdestatistica.rowcount := 3 + naltsel;
    grdestatistica.ColumnCount := 3 + naltsel;

    for i := 0 to naltsel - 1 do
    begin
      grdestatistica.Cell[0, i + 3].Alignment := taCenter;
      // grdestatistica.Cell[0, i+1].wrap := True;
      grdestatistica.Cell[0, i + 3].Font.Style := [fsbold];
      grdestatistica.Cell[1, i + 3].Alignment := taCenter;
      // grdestatistica.Cell[1, i+1].wrap := True;
      grdestatistica.Cell[1, i + 3].Font.Style := [fsbold];
      // grdestatistica.Cell[0, i + 4].Text := 'Sol ' + inttostr(i + 1);

      grdestatistica.Cell[0, i + 3].Font.Color := clwebgray;
      grdestatistica.Cell[1, i + 3].Text := ' μ';

      grdestatistica.Cell[i + 2, 2].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 2].wrap := True;
      grdestatistica.Cell[i + 2, 2].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 2].Font.Color := clwebgray;

      grdestatistica.Cell[i + 2, 2].Text := nomealt[ordemalt[i]];
      grdestatistica.cell[0, i+3].text := Nomealt[ordemalt[i]];



      grdestatistica.Cell[i + 2, 1].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 1].wrap := True;
      grdestatistica.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 1].Text := formatfloat('0.00', GanhoMedio[Ordemalt[i]]);


      grdestatistica.Cell[i + 2, 0].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 0].Text := formatfloat
        ('0.00', probabilidadealt[ordemalt[i]] * 100) + '%';
    end;

    grdestatistica.Cell[1, 1].Alignment := taCenter;
    grdestatistica.Cell[1, 1].wrap := True;
    grdestatistica.Cell[1, 1].Font.Style := [fsbold];
    grdestatistica.Cell[1, 1].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 1].Text := 'Average Gain';



    grdestatistica.Cell[1, 0].Alignment := taCenter;
    grdestatistica.Cell[1, 0].wrap := True;
    grdestatistica.Cell[1, 0].Font.Style := [fsbold];
    grdestatistica.Cell[1, 0].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 0].Text := '%';

    grdestatistica.Cell[2 + naltsel, 2].Alignment := taCenter;
    grdestatistica.Cell[2 + naltsel, 2].wrap := True;
    grdestatistica.Cell[2 + naltsel, 2].Font.Style := [fsbold];
    grdestatistica.Cell[2 + naltsel, 2].Font.Color := clWebBlue;
    grdestatistica.Cell[2 + naltsel, 2].Text := 'Average Loss:';


    for i := 0 to naltsel - 1 do
    begin
      grdestatistica.Cell[2 + naltsel, 3 + i].Text := formatfloat
        ('0.00', PerdaMedia[Ordemalt[i]]);
      grdestatistica.Cell[2 + naltsel, 3 + i].Alignment := taCenter;
      grdestatistica.Cell[2 + naltsel, 3 + i].wrap := True;
      grdestatistica.Cell[2 + naltsel, 3 + i].Font.Color :=
        clWebBlue;


    end;

    for i := 0 to naltsel - 1 do
      for j := 0 to naltsel - 1 do
      begin
        grdestatistica.Cell[i + 2, j + 3].Alignment := taCenter;
        grdestatistica.Cell[i + 2, j + 3].wrap := True;
        if i = j then
          grdestatistica.Cell[i + 2, j + 3].Text := ''
        else
        begin
          grdestatistica.Cell[i + 2, j + 3].Text := formatfloat
            ('0.00', mediamaxdiff[ordemalt[i], ordemalt[j]]);
        end;
      end;


    grdmaxmin.Clear;
    grdmaxmin.rowcount := 3 + naltsel;
    grdmaxmin.ColumnCount := 2 + naltsel;

    for i := 0 to naltsel - 1 do
    begin
      grdmaxmin.Cell[0, i + 2].Alignment := taCenter;
      // grdmaxmin.Cell[0, i+1].wrap := True;
      grdmaxmin.Cell[0, i + 2].Font.Style := [fsbold];
      grdmaxmin.Cell[1, i + 2].Alignment := taCenter;
      // grdmaxmin.Cell[1, i+1].wrap := True;
      grdmaxmin.Cell[1, i + 2].Font.Style := [fsbold];
      // grdmaxmin.Cell[0, i + 5].Text := 'Sol ' + inttostr(i + 1);
      grdmaxmin.Cell[0, i + 2].Font.Color := clwebgray;
      grdmaxmin.Cell[1, i + 2].Text := 'Max';

      grdmaxmin.Cell[i + 2, 1].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 1].wrap := True;
      grdmaxmin.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 1].Font.Color := clwebgray;

      grdmaxmin.Cell[0, 2 + i].text := NomeAlt[ordemalt[i]];
      grdmaxmin.cell[2 + i, 1].text := NomeAlt[ordemalt[i]];


      // grdmaxmin.Cell[i + 2, 4].Text := 'Sol ' + inttostr(i + 1);


      grdmaxmin.Cell[i + 2, 0].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 0].Text := formatfloat
        ('0.00', GanhoMaximo[Ordemalt[i]]);

      grdmaxmin.Cell[2 + naltsel, 2 + i].Text := formatfloat
      ('0.00', PerdaMaxima[Ordemalt[i]]);
      grdmaxmin.Cell[2 + naltsel, 2 + i].Alignment := taCenter;
      grdmaxmin.Cell[2 + naltsel, 2 + i].wrap := True;
      grdmaxmin.Cell[2 + naltsel, 2 + i].Font.Color := clWebBlue;

    end;

      grdmaxmin.Cell[1, 0].Alignment := taCenter;
      grdmaxmin.Cell[1, 0].wrap := True;
      grdmaxmin.Cell[1, 0].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 0].Text := 'Max Gain';



    grdmaxmin.Cell[2 + naltsel, 1].Alignment := taCenter;
    grdmaxmin.Cell[2 + naltsel, 1].wrap := True;
    grdmaxmin.Cell[2 + naltsel, 1].Font.Style := [fsbold];
    grdmaxmin.Cell[2 + naltsel, 1].Font.Color := clWebBlue;
    grdmaxmin.Cell[2 + naltsel, 1].Text := 'Max Loss:';




    for i := 0 to naltsel - 1 do
      for j := 0 to naltsel - 1 do
      begin
        grdmaxmin.Cell[i + 2, j + 2].Alignment := taCenter;
        grdmaxmin.Cell[i + 2, j + 2].wrap := True;
        if i = j then
          grdmaxmin.Cell[i + 2, j + 2].Text := ''
        else
        begin
          grdmaxmin.Cell[i + 2, j + 2].Text := formatfloat
            ('0.00', maxmaxdiff[ordemalt[i], ordemalt[j]]);
        end;
      end;


  end;
end;

procedure TIWForm1.btnauxchooseClick(Sender: TObject);
var
  i, j, JJ, k, qtdpareto, S, r: Integer;
  auxregra, Aux1, Aux2, aux3: real;
  bar, barpoa: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Naltsel: Integer;
  Px: real;
  faixaprob: array of real;
  pular: array of boolean;
  pular2: boolean;
  auxresultado: real;
  Linha: array of TBarSeries;

begin
  With TIWUserSession(WebApplication.Data) do
  begin
    LimpaVariaveis;
    if lstbxprojects.ItemIndex = -1 then
        begin
          webapplication.showmessage('Select a problem to continue!');
          exit;
        end;
    with UserSession.ZQuery1 do
    begin
      TIWUserSession(WebApplication.Data).NameProblem := TIWUserSession
        (WebApplication.Data).problems[lstbxProjects.ItemIndex];
      TIWUserSession(WebApplication.Data).CurrentID_problem := StrToInt
        (TIWUserSession(WebApplication.Data).id_problem[lstbxProjects.ItemIndex]
        );
      SQL.Text :=
        'SELECT ID_alternativa FROM alternativa WHERE ID_problema=' + inttostr
        (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      TIWUserSession(WebApplication.Data).NumAlt := RecordCount;
      Setlength(TIWUserSession(WebApplication.Data).altcode, TIWUserSession
          (WebApplication.Data).NumAlt);
      for S := 1 to TIWUserSession(WebApplication.Data).NumAlt do
      begin
        TIWUserSession(WebApplication.Data).altcode[S - 1] := FieldByName
          ('ID_alternativa').AsInteger;
        Next;
      end;
      Close;
      SQL.Text :=
        'SELECT ID_criterio FROM criterio WHERE ID_problema=' + inttostr
        (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      TIWUserSession(WebApplication.Data).numcrit := RecordCount;
      Setlength(TIWUserSession(WebApplication.Data).critcode, TIWUserSession
          (WebApplication.Data).numcrit);
      for r := 1 to TIWUserSession(WebApplication.Data).numcrit do
      begin
        TIWUserSession(WebApplication.Data).critcode[r - 1] := FieldByName
          ('ID_criterio').AsInteger;
        Next;
      end;
      Close;
    end;
    // >>>>
    // Gravação de dados dos pesos, limiares e tipos dos critérios

    Setlength(TIWUserSession(WebApplication.Data).NomeCrit, TIWUserSession
        (WebApplication.Data).numcrit);
    Setlength(TIWUserSession(WebApplication.Data).CritSigla, TIWUserSession
        (WebApplication.Data).numcrit);
    Setlength(TIWUserSession(WebApplication.Data).TipoCrit, TIWUserSession
        (WebApplication.Data).numcrit);
    Setlength(TIWUserSession(WebApplication.Data).Niveis, TIWUserSession
        (WebApplication.Data).numcrit);

    with UserSession.ZQuery1 do
    begin
      for i := 1 to TIWUserSession(WebApplication.Data).numcrit do
      begin
        SQL.Text := 'SELECT * FROM criterio WHERE ID_criterio=' + inttostr
          (TIWUserSession(WebApplication.Data).critcode[i - 1]);
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).CritSigla[i - 1] :=
          'Crit' + inttostr(i);
        TIWUserSession(WebApplication.Data).NomeCrit[i - 1] := FieldByName
          ('nome_criterio').AsString;

        if FieldByName('tipo_criterio').AsString <> '' then
          TIWUserSession(WebApplication.Data).TipoCrit[i - 1] := StrToIntDef
            (FieldByName('tipo_criterio').AsString, 2)
        else
          TIWUserSession(WebApplication.Data).TipoCrit[i - 1] := -1;

        if FieldByName('niveis').AsString <> '' then
          TIWUserSession(WebApplication.Data).Niveis[i - 1] := StrToIntDef
            (FieldByName('niveis').AsString, 2)
        else
          TIWUserSession(WebApplication.Data).Niveis[i - 1] := -1;

        Close;
      end;
      Setlength(TIWUserSession(WebApplication.Data).NomeAlt, TIWUserSession
          (WebApplication.Data).NumAlt);

      for j := 1 to TIWUserSession(WebApplication.Data).NumAlt do
      begin
        SQL.Text := 'SELECT * FROM alternativa WHERE ID_alternativa=' + inttostr
          (TIWUserSession(WebApplication.Data).altcode[j - 1]);
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).NomeAlt[j - 1] := FieldByName
          ('nome_alternativa').AsString;
        Close;
      end;

    end;

    // Gravação de dados da matriz consequência
    Setlength(TIWUserSession(WebApplication.Data).MatrizConseq, TIWUserSession
        (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
        .numcrit);

    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'SELECT valor_performance FROM matrizconsequencia WHERE ID_problema=' +
        inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      for i := 1 to TIWUserSession(WebApplication.Data).numcrit do
        for j := 1 to TIWUserSession(WebApplication.Data).NumAlt do
        begin
          TIWUserSession(WebApplication.Data).MatrizConseq[j - 1, i - 1] :=
            StrToFloat(FieldByName('valor_performance').AsString);
          Next;
        end;
      Close;
    end;
    GerarCases;
    EscalaRazao;
    epsilon2 := 0.01;
    Epson := 0.01;
    MinVK := 0.000000001;
    Setlength(TIWUserSession(WebApplication.Data).MatrizPOA2, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).MatrizAltSel, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).AltPOA, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).matrizsol, 1, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).resultsol, 1);
    Setlength(TIWUserSession(WebApplication.Data).VetNumPOA, length
        (CasesOrdemcrit));
    Setlength(TIWUserSession(WebApplication.Data).numsolinst, NumAlt - 1);
    Setlength(TIWUserSession(WebApplication.Data).AltTamanhoPOA, NumAlt,
      NumAlt - 1);
    Setlength(TIWUserSession(WebApplication.Data).MatrizComparacaoAlt, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).AmbasAltNaPOA, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).NenhumaAltNaPOA, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).Vgmax, length(CasesOrdemcrit)
        , TIWUserSession(WebApplication.Data).NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).VGMin, length(CasesOrdemcrit)
        , TIWUserSession(WebApplication.Data).NumAlt);
    Setlength(AmplitudeVG, length(CasesOrdemcrit), NumAlt);
    Setlength(Mediavgmax, NumAlt);
    Setlength(MediaVgMin, NumAlt);
    Setlength(MaxVgMax, NumAlt);
    Setlength(MaxVGMin, NumAlt);
    Setlength(MinVgMax, NumAlt);
    Setlength(MinVGMin, NumAlt);
    Setlength(MediaAmplitudeVG, NumAlt);
    Setlength(MaxAmplitudeVG, NumAlt);
    Setlength(MinAmplitudeVG, NumAlt);
    Setlength(AmplitudeParAParinst, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(ProbVitoriaParapar, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(mediaprobvitoriaparapar, NumAlt, NumAlt);
    Setlength(amplitudeparapar, NumAlt, NumAlt);
    Setlength(mediaprobvitoriaparapar, NumAlt, NumAlt);
    Setlength(maximoprobvitoriaparapar, NumAlt, NumAlt);
    Setlength(amplitudeparapar, NumAlt, NumAlt);
    Setlength(MaximoGeralProbVitoriaAlt, NumAlt);
    Setlength(MediaGeralProbVitoriaAlt, NumAlt);
    Setlength(difvg, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(mediadifvg, NumAlt, NumAlt);
    Setlength(maximodifvg, NumAlt, NumAlt);

    ProgramarComponentes;

    rgnLoginproject.Visible := false;
    rgnresultado.Visible := true;

    // SaveProblem;
    // btnshowresults.Enabled := true;

  end;
end;

procedure TIWForm1.btnBackRU2Click(Sender: TObject);
begin
  iwrgnCaixaDialogoRU2.Visible := false;
end;

procedure TIWForm1.btnBackRUClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU.Visible := false;
end;

procedure TIWForm1.btnContinueregistClick(Sender: TObject);
var
  numproblemas, i, j: Integer;
begin
  rgnOpcaoInput.Visible := false;
  rgnLoginproject.Visible := true;

  lstbxProjects.Clear;

  With TIWUserSession(WebApplication.Data) do
  begin
  novoproblema := false;
  retomada := true;
    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'SELECT COUNT(nome_problema) FROM problema WHERE ID_usuario = "' +
        (TIWUserSession(WebApplication.Data).ID_user) + '";';
      ExecSQL;
      Open;
      numproblemas := FieldByName('COUNT(nome_problema)').AsInteger;
      Close;

      Setlength(problems, numproblemas);
      Setlength(id_problem, numproblemas);
      Setlength(date_problem, numproblemas);

      SQL.Text := 'SELECT * FROM problema WHERE ID_usuario = "' +
        (TIWUserSession(WebApplication.Data).ID_user) + '";';
      ExecSQL;
      Open;
      First;
      i := 0;
      while i < numproblemas do
      begin
        problems[i] := FieldByName('nome_problema').AsString;
        id_problem[i] := FieldByName('ID_problema').AsString;
        date_problem[i] := FieldByName('data_problema').AsString;
        lstbxProjects.Items.Add(problems[i] + '(' + date_problem[i] + ')');
        Next;
        INC(i);
      end;
      Close;

    end;
  end;

end;

procedure TIWForm1.btncontinueRU2Click(Sender: TObject);
begin
  WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm1.btncontinueRUClick(Sender: TObject);
begin
  WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm1.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  email, id_email: string;
  senha, validation: string;
  quantregis: Integer;

begin
  IWLabel9.Visible := false;
  IWLabel8.Visible := false;
  TIWUserSession(WebApplication.Data).primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM usuario WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    Open;
    quantregis := FieldByName('count(*)').AsInteger;
    Close;
  end;
  if quantregis = 0 then
  begin
    IWLabel9.Visible := true;
    lnkbacklogin.Visible := true;
  end
  else
  begin
    TIWUserSession(WebApplication.Data).email := edtemail.Text;
    with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM usuario WHERE email ="' +
        (edtemail.Text) + '"';
      ExecSQL;
      Open;
      validation := FieldByName('validacao').AsString;
      Close;

      SQL.Text := 'SELECT password FROM usuario WHERE email ="' +
        (edtemail.Text) + '"';
      ExecSQL;
      Open;
      senha := FieldByName('password').AsString;
      Close;
    end;
    if senha = edtSenha.Text then
    begin
      with UserSession.ZQuery1 do
      begin
        SQL.Text := 'SELECT id FROM usuario WHERE email="' + (edtemail.Text)
          + '"';
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).ID_user := FieldByName('id')
          .AsString;
        Close;

        SQL.Text := 'SELECT analyst FROM usuario WHERE email="' +
          (edtemail.Text) + '"';
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).analyst := 0;
        Close;

        if TIWUserSession(WebApplication.Data).analyst = 0 then
        begin
          TIWUserSession(WebApplication.Data).email := edtemail.Text;
          TIWUserSession(WebApplication.Data).senha := edtSenha.Text;
        end;

      end;

      // * Caso o usuário esteja trocando a senha de cadastro *//
      { if validation = 'password' then
        begin
        WebApplication.ShowMessage('Temporary password has been found. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with changing your password.');
        edtSenha.Text := '';
        //Exit;
        end; }

      // * Caso seja o primeiro acesso *//
      if validation <> 'validado' then
      begin
        WebApplication.ShowMessage(
          'User has not been validated. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with user validation.');
        edtSenha.Clear;
        Exit;
      end;

      rgnEnterLogin.Visible := false;
      rgnOpcaoInput.Visible := true;
      // rgnOpcaoInput.Visible := false;
    end
    else
    begin
      IWLabel8.Visible := true;
      lnkbacklogin.Visible := true;
    end;

  end;

end;

procedure TIWForm1.btnimportClick(Sender: TObject);
var
  i, j: Integer;
  cred: boolean;
  vetchar: array [0 .. 9] of string;
  Linha, coluna, ncrit, nalt: Integer;
  tipoarquivo : string;
begin

  if (FileArquivo.Filename = '') then
  begin
    if TIWUserSession(WebApplication.Data).excelverif = false then
      WebApplication.ShowMessage('You must choose an excel file');
    Exit;
  end
  else if FileArquivo.Filename <> '' then
  begin
    { vetchar[0] := '!';
      vetchar[1] := '@';
      vetchar[2] := '#';
      vetchar[3] := '$';
      vetchar[4] := '%';
      vetchar[5] := '&';
      vetchar[6] := '*';
      vetchar[7] := '^';
      vetchar[8] := '-';
      vetchar[9] := '+'; }
    cred := false;
    { for i := 0 to 9 do
      for j := 0 to length(edteditarnome.Text) do
      begin
      if vetchar[i] = edteditarnome.Text[j] then
      cred := true;
      end; }

    if cred = true then
    begin
      WebApplication.ShowMessage(
        'Special characters are not allowed. You must choose a different name for your problem.');
      Exit;
    end
    else
    begin
      tipoarquivo := FileArquivo.ContentType;
      if (tipoarquivo = ' application/vnd.ms-excel') then
      // verifica se o arquivo é compatível
      begin
      FileArquivo.SaveToFile('C:\Users\Public\Documents' +
          (FileArquivo.Filename));
      TIWAdvWebGridExcelIO1.XLSImport
        ('C:\Users\Public\Documents' + (FileArquivo.Filename));

      TIWAdvWebGridExcelIO1.AdvWebGrid := grdImport;
      end
      else
      begin
        webapplication.ShowMessage('Select an excel file to continue!');
        exit;
      end;

      ncrit := grdImport.Columns.Count;
      nalt := grdImport.totalrows;

      for i := ncrit - 1 downto 1 do
      begin
        if grdImport.Cells[i, 0] = '' then
          grdImport.Columns.delete(i);
      end;
      coluna := 0;
      for i := 1 to grdImport.Columns.Count - 1 do
      begin
        if grdImport.Cells[i, 0] <> '' then
          INC(coluna);
      end;
      Linha := 8;
      for j := 8 to grdImport.totalrows - 1 do
      begin
        if grdImport.Cells[0, j] <> '' then
          INC(Linha);
      end;

      if coluna = 3 then
        cmbbxcriteria.ItemIndex := 0;
      if coluna = 4 then
        cmbbxcriteria.ItemIndex := 1;
      if coluna = 5 then
        cmbbxcriteria.ItemIndex := 2;
      if coluna = 6 then
        cmbbxcriteria.ItemIndex := 3;
      if coluna >= 7 then
        cmbbxcriteria.ItemIndex := 4;

      edtnumalt.Text := inttostr(Linha - 8);
      btnupdateClick(self);

      With TIWUserSession(WebApplication.Data) do
      begin
        for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
        begin
          if grdImport.Cells[i + 1, 1] = '0' then
            cmbbxTipoCrit[i].ItemIndex := 0;
          if grdImport.Cells[i + 1, 1] = '1' then
            cmbbxTipoCrit[i].ItemIndex := 1;
          if grdImport.Cells[i + 1, 1] = '2' then
            cmbbxTipoCrit[i].ItemIndex := 2;
          if grdImport.Cells[i + 1, 1] = '3' then
            cmbbxTipoCrit[i].ItemIndex := 3;
          if grdImport.Cells[i + 1, 1] = '4' then
            cmbbxTipoCrit[i].ItemIndex := 4;
          if grdImport.Cells[i + 1, 1] = '5' then
            cmbbxTipoCrit[i].ItemIndex := 5;

          cmbbxtipocrit[i].Enabled := false;




          cmbbxauxiliarChange(self);

          if (TIWUserSession(WebApplication.Data).cmbbxTipoCrit[i]
              .ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3) then
          begin
            if grdImport.Cells[i + 1, 6] = '2' then
              cmbbxNiveisCrit[i].ItemIndex := 0;
            if grdImport.Cells[i + 1, 6] = '3' then
              cmbbxNiveisCrit[i].ItemIndex := 1;
            if grdImport.Cells[i + 1, 6] = '4' then
              cmbbxNiveisCrit[i].ItemIndex := 2;
            if grdImport.Cells[i + 1, 6] = '5' then
              cmbbxNiveisCrit[i].ItemIndex := 3;
            if grdImport.Cells[i + 1, 6] = '6' then
              cmbbxNiveisCrit[i].ItemIndex := 4;
            if grdImport.Cells[i + 1, 6] = '7' then
              cmbbxNiveisCrit[i].ItemIndex := 5;
          end;



          cmbbxNiveisCrit[i].Enabled := false;
          for j := 0 to NumAlt - 1 do
          begin
            edtMatrizConseq[j, i].Text := grdImport.Cells[1 + i, 8 + j];
            edtMatrizConseq[j, i].Enabled := false;
          end;
        end;
        for i := 0 to numcrit - 1 do
        begin
          edtNomeCrit[i].Text := grdImport.Cells[i + 1, 0];
          NomeCrit[i] := grdImport.Cells[i + 1, 0];

        end;
        for j := 0 to NumAlt - 1 do
        begin
          edtnomealt[j].Text := grdImport.Cells[0, j + 8];
          NomeAlt[j] := grdImport.Cells[0, j + 8];
        end;
      end;

    end;
    TIWUserSession(WebApplication.Data).excelverif := true;
  end;
end;

procedure TIWForm1.btnLoginClick(Sender: TObject);
begin
  rgnintro.Visible := false;
  rgnEnterLogin.Visible := true;
end;

procedure TIWForm1.btnRegisterClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU.Visible := true;
end;

procedure TIWForm1.btnRegisternewClick(Sender: TObject);
var
i,j : integer;
begin
rgnopcaoinput.Visible := false;
rgnmatcons.Visible := true;

 LimpaVariaveis;
 edtnumalt.Text := '';
 cmbbxcriteria.ItemIndex := -1;

 edtnameproblem.Clear;
 btnshowresults.Enabled := false;
 With TIWUserSession(WebApplication.Data) do
  begin
    novoproblema := true;
    retomada := false;
    //Numalt := strtoint(edtnumalt.Text);
    //Numcrit := strtoint(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    if length(edtNomeCrit) <> 0 then
    begin
      for i := 0 to length(edtNomeCrit) - 1 do
        edtNomeCrit[i].Destroy;
      Finalize(edtNomeCrit);
    end;
    if length(edtnomealt) <> 0 then
    begin
      for i := 0 to length(edtnomealt) - 1 do
        edtnomealt[i].Destroy;
      Finalize(edtnomealt);
    end;
    if length(cmbbxTipoCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxTipoCrit) - 1 do
        cmbbxTipoCrit[i].Destroy;
      Finalize(cmbbxTipoCrit);
    end;
    if length(cmbbxNiveisCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxNiveisCrit) - 1 do
        cmbbxNiveisCrit[i].Destroy;
      Finalize(cmbbxNiveisCrit);
    end;
    if length(edtMatrizConseq) <> 0 then
    begin
      for i := 0 to length(edtMatrizConseq) - 1 do
        for j := 0 to length(edtMatrizConseq[i]) - 1 do
          edtMatrizConseq[i, j].Destroy;
      Finalize(edtMatrizConseq);
    end;
  end;


end;

procedure TIWForm1.btnshowresultsClick(Sender: TObject);
var
  i, j, JJ, k, qtdpareto: Integer;
  auxregra, Aux1, Aux2, aux3: real;
  bar, barpoa: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Naltsel: Integer;
  Px: real;
  faixaprob: array of real;
  pular: array of boolean;
  pular2: boolean;
  auxresultado: real;
  Linha: array of TBarSeries;
  valor : double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    if edtnameproblem.Text = '' then // verifica se foi inserido um nome pro arquivo
    begin // se não tiver nome do problema
      WebApplication.ShowMessage('Please, enter a problem''s name!');
      Exit;
    end; // se não tiver nome do problema

    // LimpaVariaveis;

    NumAlt := StrToInt(edtnumalt.Text);
    numcrit := StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);

    NameProblem := edtnameproblem.Text;

    for i := 0 to numcrit - 1 do
     begin
       if ((cmbbxtipocrit[i].itemindex = 2) or (cmbbxtipocrit[i].itemindex = 3)) then
       begin
        if (cmbbxniveiscrit[i].itemindex = -1) then
        begin
          WebApplication.showmessage('Fill the fields correctly!');
          exit;
        end;
        for j := 0 to numalt - 1 do
        begin
          if cmbbxniveiscrit[i].itemindex = 0 then
          begin
            if (strtofloat(edtmatrizconseq[j,i].text) <> 0) and (strtofloat(edtmatrizconseq[j,i].text) <> 1) then
            begin
              WebApplication.showmessage('Fill the fields correctly!');
              exit;
            end;
          end
          else
          begin
            if (strtofloat(edtmatrizconseq[j,i].text) <= 0) or (strtofloat(edtmatrizconseq[j,i].text) > 2+cmbbxniveiscrit[i].itemindex) then
            begin
              WebApplication.showmessage('Fill the fields correctly!');
              exit;
            end;
          end;

        end;
       end;
     end;

     for j := 0 to NumAlt - 1 do
      for i := 0 to numcrit - 1 do
        begin
          if trystrtofloat(edtMatrizConseq[j, i].Text, valor) = false then
          begin
          WebApplication.showmessage('Fill the fields correctly!');
          exit;
          end;

        TIWUserSession(WebApplication.Data).MatrizConseq[j, i] := TrocaPV
          (edtMatrizConseq[j, i].Text);
        end;

    for i := 0 to numcrit - 1 do
    begin
      TIWUserSession(WebApplication.Data).TipoCrit[i] := cmbbxTipoCrit[i]
        .ItemIndex;
      if (TipoCrit[i] = 2) or (TipoCrit[i] = 3) then
        TIWUserSession(WebApplication.Data).Niveis[i] := StrToInt
          (cmbbxNiveisCrit[i].Items[cmbbxNiveisCrit[i].ItemIndex]);
    end;





     GerarCases;

    { for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] <> 2) and (tipocrit[i] <> 3) then
      begin
      Max[i] := -9999999999;
      Min[i] := 9999999999;
      for j := 0 to Numalt - 1 do
      begin
      if Max[i] < MatrizConseq[j, i] then
      Max[i] := MatrizConseq[j, i];
      if Min[i] > MatrizConseq[j, i] then
      Min[i] := MatrizConseq[j, i];
      end; // j
      end // if
      else
      begin
      if niveis[i] = 2 then
      begin
      Max[i] := 1;
      Min[i] := 0;
      end;
      if niveis[i] > 2 then
      begin
      TIWUserSession(WebApplication.Data).Max[i] := niveis[i];
      TIWUserSession(WebApplication.Data).Min[i] := 1;
      end;
      end;
      end; // i
      for j := 0 to Numalt - 1 do
      for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] = 1) or (tipocrit[i] = 3) or (tipocrit[i] = 5) then
      MatrizConseqNorm[j, i] := (MatrizConseq[j, i] - Min[i]) /
      (Max[i] - Min[i]);
      /// Normalização para os critérios de minimização //////////////////////
      if (tipocrit[i] = 0) or (tipocrit[i] = 2) or (tipocrit[i] = 4) then
      TIWUserSession(WebApplication.Data).MatrizConseqNorm[j, i] :=
      (MatrizConseq[j, i] - Max[i]) / (-Max[i] + Min[i]);
      end; }

    EscalaRazao;
    epsilon2 := 0.01;
    Epson := 0.01;
    MinVK := 0.000000001;
    Setlength(TIWUserSession(WebApplication.Data).MatrizPOA2, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).MatrizAltSel, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).AltPOA, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).matrizsol, 1, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).resultsol, 1);
    Setlength(TIWUserSession(WebApplication.Data).VetNumPOA, length
        (CasesOrdemcrit));
    Setlength(TIWUserSession(WebApplication.Data).numsolinst, NumAlt - 1);
    Setlength(TIWUserSession(WebApplication.Data).AltTamanhoPOA, NumAlt,
      NumAlt - 1);
    Setlength(TIWUserSession(WebApplication.Data).MatrizComparacaoAlt, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).AmbasAltNaPOA, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).NenhumaAltNaPOA, NumAlt,
      NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).Vgmax, length(CasesOrdemcrit)
        , TIWUserSession(WebApplication.Data).NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).VGMin, length(CasesOrdemcrit)
        , TIWUserSession(WebApplication.Data).NumAlt);
    Setlength(AmplitudeVG, length(CasesOrdemcrit), NumAlt);
    Setlength(Mediavgmax, NumAlt);
    Setlength(MediaVgMin, NumAlt);
    Setlength(MaxVgMax, NumAlt);
    Setlength(MaxVGMin, NumAlt);
    Setlength(MinVgMax, NumAlt);
    Setlength(MinVGMin, NumAlt);
    Setlength(MediaAmplitudeVG, NumAlt);
    Setlength(MaxAmplitudeVG, NumAlt);
    Setlength(MinAmplitudeVG, NumAlt);
    Setlength(AmplitudeParAParinst, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(ProbVitoriaParapar, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(mediaprobvitoriaparapar, NumAlt, NumAlt);
    Setlength(maximoprobvitoriaparapar, NumAlt, NumAlt);
    Setlength(amplitudeparapar, NumAlt, NumAlt);
    Setlength(MaximoGeralProbVitoriaAlt, NumAlt);
    Setlength(MediaGeralProbVitoriaAlt, NumAlt);
    Setlength(difvg, length(CasesOrdemcrit), NumAlt, NumAlt);
    Setlength(mediadifvg, NumAlt, NumAlt);
    Setlength(maximodifvg, NumAlt, NumAlt);
    ProgramarComponentes;
    rgnmatcons.Visible := false;
    rgnresultado.Visible := true;

    SaveProblem;
    // btnshowresults.Enabled := true;

  end;
end;

procedure TIWForm1.btnupdateClick(Sender: TObject);
var
  i, j, valor: Integer;
begin
  if not trystrtoint(edtnumalt.text, valor) then begin
  edtnumalt.text := '';
  webapplication.ShowMessage('Fill the number of alternatives correctly!');
  exit;
  end;
  if cmbbxcriteria.ItemIndex = -1 then begin
  webapplication.ShowMessage('Fill the number of criteria correctly!');
  exit;
  end;
  if edtnumalt.Text = '1' then begin
  edtnumalt.text := '';
  webapplication.ShowMessage('The number of alternative must be greater than 1!');
  exit;
  end;
  LimpaVariaveis;
  for j := 1 to grdmatrizconseq.Columns.Count - 1 do
    grdmatrizconseq.Columns.delete(1);

  grdmatrizconseq.rowcount := StrToInt(edtnumalt.Text) + 4;

  for i := 0 to StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]) - 1 do
  begin
    grdmatrizconseq.Columns.Add;
    grdmatrizconseq.Columns.Items[i + 1].Width := 200;
  end;
  With TIWUserSession(WebApplication.Data) do
  begin
    NumAlt := StrToInt(edtnumalt.Text);
    numcrit := StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    if length(edtNomeCrit) <> 0 then
    begin
      for i := 0 to length(edtNomeCrit) - 1 do
        edtNomeCrit[i].Destroy;
      Finalize(edtNomeCrit);
    end;
    if length(edtnomealt) <> 0 then
    begin
      for i := 0 to length(edtnomealt) - 1 do
        edtnomealt[i].Destroy;
      Finalize(edtnomealt);
    end;
    if length(cmbbxTipoCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxTipoCrit) - 1 do
        cmbbxTipoCrit[i].Destroy;
      Finalize(cmbbxTipoCrit);
    end;
    if length(cmbbxNiveisCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxNiveisCrit) - 1 do
        cmbbxNiveisCrit[i].Destroy;
      Finalize(cmbbxNiveisCrit);
    end;
    if length(edtMatrizConseq) <> 0 then
    begin
      for i := 0 to length(edtMatrizConseq) - 1 do
        for j := 0 to length(edtMatrizConseq[i]) - 1 do
          edtMatrizConseq[i, j].Destroy;
      Finalize(edtMatrizConseq);
    end;

    Setlength(edtNomeCrit, numcrit);
    Setlength(edtnomealt, NumAlt);
    Setlength(cmbbxTipoCrit, numcrit);
    Setlength(cmbbxNiveisCrit, numcrit);
    Setlength(edtMatrizConseq, NumAlt, numcrit);

    Setlength(NomeCrit, numcrit);
    Setlength(NomeAlt, NumAlt);
    Setlength(TipoCrit, numcrit);
    Setlength(Niveis, numcrit);
    Setlength(MatrizConseq, NumAlt, numcrit);
    Setlength(MatrizConseqNorm, NumAlt, numcrit);
    Setlength(Max, numcrit);
    Setlength(Min, numcrit);

    for i := 0 to numcrit - 1 do
    begin
      edtNomeCrit[i] := TIWEdit.Create(self);
      edtNomeCrit[i].Parent := rgnmatrizdeconsequencias;
      edtNomeCrit[i].Top := 55;
      edtNomeCrit[i].Left := 17 + 201 * (i + 1);
      edtNomeCrit[i].Width := 200;
      edtNomeCrit[i].Height := 23;
      edtNomeCrit[i].Text := 'Crit ' + inttostr(i + 1);
      NomeCrit[i] := 'Crit ' + inttostr(i + 1);
      edtNomeCrit[i].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      edtNomeCrit[i].Enabled := false;

      cmbbxTipoCrit[i] := TIWComboBox.Create(self);
      cmbbxTipoCrit[i].Parent := rgnmatrizdeconsequencias;
      cmbbxTipoCrit[i].Top := 78;
      cmbbxTipoCrit[i].Left := 17 + 201 * (i + 1);
      cmbbxTipoCrit[i].Width := 200;
      cmbbxTipoCrit[i].Height := 23;
      cmbbxTipoCrit[i].Items.Add('0: Continuous(Min)');
      cmbbxTipoCrit[i].Items.Add('1: Continuous(Max)');
      cmbbxTipoCrit[i].Items.Add('2: Discreet(Min)');
      cmbbxTipoCrit[i].Items.Add('3: Discreet(Max)');
      cmbbxTipoCrit[i].Items.Add('4: Integer(Min)');
      cmbbxTipoCrit[i].Items.Add('5: Integer(Max)');
      cmbbxTipoCrit[i].ZIndex := 1000;
      cmbbxTipoCrit[i].OnChange := cmbbxauxiliar.OnChange;

      cmbbxNiveisCrit[i] := TIWComboBox.Create(self);
      cmbbxNiveisCrit[i].Parent := rgnmatrizdeconsequencias;
      cmbbxNiveisCrit[i].Top := 101;
      cmbbxNiveisCrit[i].Left := 17 + 201 * (i + 1);
      cmbbxNiveisCrit[i].Width := 200;
      cmbbxNiveisCrit[i].Height := 33;
      cmbbxNiveisCrit[i].Items.Add('2');
      cmbbxNiveisCrit[i].Items.Add('3');
      cmbbxNiveisCrit[i].Items.Add('4');
      cmbbxNiveisCrit[i].Items.Add('5');
      cmbbxNiveisCrit[i].Items.Add('6');
      cmbbxNiveisCrit[i].Items.Add('7');
      cmbbxNiveisCrit[i].ZIndex := 1000;
      cmbbxNiveisCrit[i].Visible := false;

      // cmbbxTipocrit[i] := TiwComboBox.Create(self);

      DSiTrimWorkingSet;
    end;
    for j := 0 to NumAlt - 1 do
    begin
      edtnomealt[j] := TIWEdit.Create(self);
      edtnomealt[j].Parent := rgnmatrizdeconsequencias;
      edtnomealt[j].Top := 66 + 23 * (j + 4);
      edtnomealt[j].Left := 17;
      edtnomealt[j].Width := 200;
      edtnomealt[j].Height := 23;
      edtnomealt[j].Text := 'Alt ' + inttostr(j + 1);
      NomeAlt[j] := 'Alt ' + inttostr(j + 1);
      edtnomealt[j].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      edtnomealt[j].Enabled := false;

      DSiTrimWorkingSet;
    end;

    for i := 0 to numcrit - 1 do
      for j := 0 to NumAlt - 1 do
      begin
        edtMatrizConseq[j, i] := TIWEdit.Create(self);
        edtMatrizConseq[j, i].Parent := rgnmatrizdeconsequencias;
        edtMatrizConseq[j, i].Top := 66 + 23 * (j + 4);
        edtMatrizConseq[j, i].Left := 17 + 201 * (i + 1);
        edtMatrizConseq[j, i].Width := 200;
        edtMatrizConseq[j, i].Height := 23;
        edtMatrizConseq[j, i].Text := '';
        edtMatrizConseq[j, i].Font.FontFamily :=
          'Verdana, Helvetica, Sans-Serif';
        edtMatrizConseq[j, i].Enabled := true;
        edtMatrizConseq[j, i].ZIndex := 1000;

        DSiTrimWorkingSet;
      end;
  end;
  btnshowresults.Enabled := true;
end;

procedure TIWForm1.btnupdatevisualizacaoClick(Sender: TObject);
var
  i, j, k: Integer;
  bar: array of TBarSeries;
  totalacumulado: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    //if rdgVisualizacao.ItemIndex = 0 then
    //begin
      chrtResultado.RemoveAllSeries;
      chrtResultado.Legend.Items.Clear;
      // chrtresultado.ClearChart;
      chrtResultado.BottomAxis.Items.Clear;
      chrtResultado.View3D := false;
      Setlength(bar, NumAlt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to NumAlt - 1 do
      begin
        if AltPOA[ordemalt[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, AltPOA[ordemalt[j]], formatfloat
              ('0', AltPOA[ordemalt[j]]), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtResultado.Legend.Item[i].Text := NomeAlt[ordemalt[j]];
          bar[i].Color := coresalt[i];
          // totalacumulado := totalacumulado + AltPOA[Ordemalt[j]];
          INC(i);
        end;
      end;
      // Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtResultado.AddSeries(bar[1]);
      { for i := 0 to length(ResultSol) - 1 do
        Total := Total + ResultSol[i];
        } bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.Legend.Visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      // chrtResultado.LeftAxis.Maximum := Total;
      chrtResultado.LeftAxis.AutomaticMaximum := true;
      /// chrtResultado.LeftAxis.MaximumOffset := 16;
      // chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.leftaxis.Items.Automatic := false;
      // chrtResultado.LeftAxis.Increment := Total;

    {end
    else
    begin
      chrtResultado.RemoveAllSeries;
      chrtResultado.Legend.Items.Clear;
      // chrtresultadoPromethee.ClearChart;
      chrtResultado.BottomAxis.Items.Clear;
      chrtResultado.View3D := false;
      Setlength(bar, NumAlt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to NumAlt - 1 do
      begin
        if AltPOA[ordemalt[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, 100 * AltPOA[ordemalt[j]] / Total, formatfloat
              ('0', 100 * AltPOA[ordemalt[j]] / Total), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtResultado.Legend.Item[i].Text := NomeAlt[ordemalt[j]];
          bar[i].Color := coresalt[i];
          // totalacumulado := totalacumulado + ResultAlt[Ordemalt[j]];
          INC(i);
        end;
      end;
      Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtResultado.AddSeries(bar[1]);
      for i := 0 to length(resultsol) - 1 do
        Total := Total + resultsol[i];
      bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.Legend.Visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      // chrtResultado.LeftAxis.Maximum := 100;
      chrtResultado.LeftAxis.AutomaticMaximum := true;
      /// chrtResultado.LeftAxis.MaximumOffset := 16;
      // chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.leftaxis.Items.Automatic := false;
      // chrtResultado.LeftAxis.Increment := 100;
    end;}
    imgAlternative.Picture.Bitmap.Assign(chrtResultado.TeeCreateBitmap
        (clWebPURPLE, rect(0, 0, imgAlternative.Width, imgAlternative.Height)));
  end;
end;

procedure TIWForm1.cmbbxAlternativaChange(Sender: TObject);
var
  i, j, II, JJ: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    grdalternativa.Clear;
    grdcasesalternativa.Clear;

    grdalternativa.rowcount := 1;
    grdalternativa.ColumnCount := 3;
    chrtAlternativa.series[0].Clear;

    for i := 0 to cmbbxAlternativa.Items.Count - 1 do
    begin
      if TIWUserSession(WebApplication.Data).AltTamanhoPOA
        [TIWUserSession(WebApplication.Data)
        .ordemalt[cmbbxAlternativa.ItemIndex], i] <> 0 then
      begin
        grdalternativa.rowcount := 1 + grdalternativa.rowcount;
        grdalternativa.Cell[grdalternativa.rowcount - 1, 0].Text := inttostr
          (i + 1) + ' alternatives in POA';
        grdalternativa.Cell[grdalternativa.rowcount - 1, 1].Text := formatfloat
          ('0', AltTamanhoPOA[ordemalt[cmbbxAlternativa.ItemIndex], i]);
        grdalternativa.Cell[grdalternativa.rowcount - 1, 2].Text := formatfloat
          ('0.00', 100 * AltTamanhoPOA[ordemalt[cmbbxAlternativa.ItemIndex], i]
            / AltPOA[ordemalt[cmbbxAlternativa.ItemIndex]]) + '%';

        chrtAlternativa.series[0].Add
          (100 * AltTamanhoPOA[ordemalt[cmbbxAlternativa.ItemIndex], i] / AltPOA
            [ordemalt[cmbbxAlternativa.ItemIndex]], inttostr(i + 1)
            + ' alts in POA');
      end;
    end;
    chrtAlternativa.Title.Caption := 'Distribution of ' + NomeAlt
      [ordemalt[cmbbxAlternativa.ItemIndex]] + ' in POA';
    grdalternativa.Cell[0, 0].Text := 'Alts Included in POA';
    grdalternativa.Cell[0, 1].Text := 'Frequency';
    grdalternativa.Cell[0, 2].Text := '%';
    for j := 0 to grdalternativa.ColumnCount - 1 do
    begin
      grdalternativa.Cell[0, j].Font.Style := [fsbold];
      for i := 0 to grdalternativa.rowcount - 1 do
      begin
        grdalternativa.Cell[i, j].Alignment := taCenter;
        grdalternativa.Cell[i, j].wrap := true;
        //
      end;
    end;

    imgalternativa.Picture.Bitmap.Assign(chrtAlternativa.TeeCreateBitmap
        (clWebPURPLE, rect(0, 0, imgalternativa.Width, imgalternativa.Height)));

    grdcasesalternativa.Clear;
    grdcasesalternativa.rowcount := 2;
    grdcasesalternativa.ColumnCount := 2 + numcrit;

    for j := 0 to grdcasesalternativa.ColumnCount - 1 do
      grdcasesalternativa.Cell[0, j].Font.Style := [fsbold];

    grdcasesalternativa.Cell[0, 0].Text := 'Order';
    for i := 0 to numcrit - 1 do
      grdcasesalternativa.Cell[0, i + 1].Text := NomeCrit[i];

    grdcasesalternativa.Cell[0, numcrit + 1].Text := 'POA Subset';

    JJ := 0;
    for j := 0 to length(CasesOrdemcrit) - 1 do
    begin
      if MatrizAltSel[j, ordemalt[cmbbxAlternativa.ItemIndex]] = 1 then
      begin
        grdcasesalternativa.Cell[JJ + 1, 0].Text := 'O' + inttostr(j + 1);
        for i := 0 to NumAlt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).MatrizAltSel[j, i] = 1 then
          begin
            if grdcasesalternativa.Cell[JJ + 1, numcrit + 1].Text = '' then
              grdcasesalternativa.Cell[JJ + 1, numcrit + 1].Text := NomeAlt[i]
            else
              grdcasesalternativa.Cell[JJ + 1, numcrit + 1].Text :=
                grdcasesalternativa.Cell[JJ + 1, numcrit + 1]
                .Text + ', ' + NomeAlt[i];
          end;
        end;
        for i := 0 to numcrit - 1 do
          grdcasesalternativa.Cell[JJ + 1, i + 1].Text := inttostr
            (CasesOrdemcrit[j, i]) + 'º';

        INC(JJ);
        grdcasesalternativa.rowcount := grdcasesalternativa.rowcount + 1;
      end;
    end;

    for j := 0 to grdcasesalternativa.ColumnCount - 1 do
      for i := 0 to grdcasesalternativa.rowcount - 1 do
      begin
        grdcasesalternativa.Cell[i, j].Alignment := taCenter;
        grdcasesalternativa.Cell[i, j].wrap := true;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;

    DSiTrimWorkingSet;
  end;

end;

procedure TIWForm1.cmbbxauxiliarAsyncChange(Sender: TObject;
  EventParams: TStringList);
var
  i: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[i].ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3)
        then
        cmbbxNiveisCrit[i].Visible := true
      else
        cmbbxNiveisCrit[i].Visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxauxiliarChange(Sender: TObject);
var
  i: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[i].ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3)
        then
        cmbbxNiveisCrit[i].Visible := true
      else
        cmbbxNiveisCrit[i].Visible := false;
    end;
  end;
end;

procedure TIWForm1.rgnalternativaCreate(Sender: TObject);
begin
  rgnalternativa.Align := alclient;
end;

procedure TIWForm1.rgnanaliseestatisticaCreate(Sender: TObject);
begin
rgnanaliseestatistica.align := AlClient;
end;

procedure TIWForm1.rgndetalhamentoaltCreate(Sender: TObject);
begin
  rgndetalhamentoalt.Align := alclient;
end;

procedure TIWForm1.rgnEnterLoginCreate(Sender: TObject);
begin
  rgnEnterLogin.Align := alclient;
end;

procedure TIWForm1.rgnintroCreate(Sender: TObject);
begin
  rgnintro.Align := alclient;
end;

procedure TIWForm1.rgnLoginprojectCreate(Sender: TObject);
begin
  rgnLoginproject.Align := alclient;
end;

procedure TIWForm1.rgnmatconsCreate(Sender: TObject);
begin
  rgnmatcons.Align := alclient;
  grdmatrizconseq.Cells[0, 0] := 'Criteria:';
  grdmatrizconseq.Cells[0, 1] := 'Criterion Type:';
  grdmatrizconseq.Cells[0, 2] := 'Number of levels of discrete criteria:';
  grdmatrizconseq.Cells[0, 3] := 'Alternatives:';
  grdmatrizconseq.Cells[1, 3] := 'Consequence Matrix:';
end;

procedure TIWForm1.rgnOpcaoInputCreate(Sender: TObject);
begin
  rgnOpcaoInput.Align := alclient;
end;

procedure TIWForm1.rgnregradedecisaoCreate(Sender: TObject);
begin
  rgnregradedecisao.Align := alclient;
end;

procedure TIWForm1.rgnregradedecisaoprometheeCreate(Sender: TObject);
begin
  rgnregradedecisaoPromethee.Align := alclient;
end;

procedure TIWForm1.rgnresultadoCreate(Sender: TObject);
begin
  rgnresultado.Align := alclient;
end;

procedure TIWForm1.SaveProblem;
var
  i, j, x, k: Integer;
  numero: Double;
begin



  // IWButton16.Cursor := crWait;

  With TIWUserSession(WebApplication.Data) do
  begin
    Setlength(TIWUserSession(WebApplication.Data).critcode, TIWUserSession
        (WebApplication.Data).numcrit);
    Setlength(TIWUserSession(WebApplication.Data).altcode, TIWUserSession
        (WebApplication.Data).NumAlt);
    // Setando a matriz de consequencias
    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'INSERT INTO problema (nome_problema, data_problema,hora_problema, ID_usuario) VALUES ("' + TIWUserSession(WebApplication.Data).NameProblem + '", "' + datetostr(date) + '", "' + datetostr(now) + '","' + ID_user + '")';
      ExecSQL;
      SQL.Text :=
        'SELECT ID_problema from problema WHERE nome_problema = "'
        + TIWUserSession(WebApplication.Data).NameProblem + '"';
      ExecSQL;
      Open;
      Last;
      TIWUserSession(WebApplication.Data).CurrentID_problem := FieldByName
        ('ID_problema').AsInteger;
      Close;
      SQL.Text :=
        'INSERT INTO usuario_problema (ID_usuario, ID_problema) VALUES ( ' +
        ID_user + ' ,' + inttostr(TIWUserSession(WebApplication.Data)
          .CurrentID_problem) + ')';
      ExecSQL;

    end;

    with UserSession.ZQuery1 do
    begin
      for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'INSERT INTO criterio (nome_criterio, tipo_criterio, niveis, ID_problema) VALUES ("' + NomeCrit[i] + '","' + inttostr(TipoCrit[i]) + '","' + inttostr(Niveis[i]) + '","' + inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem) + '")';
          ExecSQL;
          // prefdir[i] := 1 - prefdir[i];
        end; // with usersession
      end; // for i
      for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'INSERT INTO alternativa (nome_alternativa, ID_problema) VALUES ("'
            + NomeAlt[j] + '","' + inttostr(TIWUserSession(WebApplication.Data)
              .CurrentID_problem) + '")';
          ExecSQL;
        end; // with usersession
      end; // for j
    end; // with zquery

    with UserSession.ZQuery1 do
    begin
      for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'SELECT ID_criterio FROM criterio WHERE nome_criterio="' + NomeCrit
            [i] + '" AND ID_problema= "' + inttostr
            (TIWUserSession(WebApplication.Data).CurrentID_problem) + '"';
          ExecSQL;
          Open;
          critcode[i] := UserSession.ZQuery1.FieldByName('ID_criterio')
            .AsInteger;
          Close;
        end; // with usersession
      end; // for i
      for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'SELECT ID_alternativa FROM alternativa WHERE nome_alternativa="' +
            NomeAlt[j] + '" AND ID_problema="' + inttostr
            (TIWUserSession(WebApplication.Data).CurrentID_problem) + '"';
          ExecSQL;
          Open;
          altcode[j] := UserSession.ZQuery1.FieldByName('ID_alternativa')
            .AsInteger;
          Close;
        end; // with usersession
      end; // for j
    end; // with zquery

    // Essa parte do código tem que ir para o botão de submeter a matriz
    UserSession.ZQuery1.SQL.Text :=
      'SELECT ID_alternativa FROM alternativa WHERE ID_problema=' + inttostr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
    UserSession.ZQuery1.ExecSQL;
    UserSession.ZQuery1.Open;
    for j := 1 to TIWUserSession(WebApplication.Data).NumAlt do
    begin
      TIWUserSession(WebApplication.Data).altcode[j - 1] :=
        UserSession.ZQuery1.FieldByName('ID_alternativa').AsInteger;
      UserSession.ZQuery1.Next;
    end;
    UserSession.ZQuery1.Close;
    //

    // Essa parte do código tem que ir para o botão de submeter a matriz
    UserSession.ZQuery1.SQL.Text :=
      'SELECT ID_criterio FROM criterio WHERE ID_problema=' + inttostr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
    UserSession.ZQuery1.ExecSQL;
    UserSession.ZQuery1.Open;
    for k := 1 to TIWUserSession(WebApplication.Data).numcrit do
    begin
      TIWUserSession(WebApplication.Data).critcode[k - 1] :=
        UserSession.ZQuery1.FieldByName('ID_criterio').AsInteger;
      UserSession.ZQuery1.Next;
    end;
    UserSession.ZQuery1.Close;
    //
    x := 0;
    { with UserSession.ZQuery1 do
      begin
      SQL.Text := 'SELECT * from matrizconsequencia WHERE ID_problema=' + IntToStr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      x := RecordCount;
      Close;
      end; }
    if (x = 0) then
    begin
      with UserSession.ZQuery1 do
      begin

        for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
          begin

            SQL.Text :=
              'INSERT INTO matrizconsequencia (ID_alternativa, ID_criterio, valor_performance, ID_problema) VALUES (' + inttostr(TIWUserSession(WebApplication.Data).altcode[j]) + ',' + inttostr(TIWUserSession(WebApplication.Data).critcode[i]) + ',"' + FloatToStr(MatrizConseq[j, i]) + '",' + inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem) + ')';
            ExecSQL;
          end;
        end;

      end;
    end;
  end;

  WebApplication.ShowMessage('Data collection submitted', smAlert);
end;

initialization

TIWForm1.SetAsMainForm;

end.
