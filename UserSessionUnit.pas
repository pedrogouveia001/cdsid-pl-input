unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompLabel,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, TeeGDIPlus, IWCompButton, TeEngine, ExtCtrls,
  TeeProcs, Chart, IWExtCtrls, IWWebGrid, IWAdvWebGrid, BubbleCh, TeePolar,
  IWCompListbox, IWCompRadioButton, IWTMSCheckList, series, IWGrids, TeeLisB,
  TeeSeriesTextEd, Buttons, Math, pngimage, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns, ActnList, ImgList,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Windows, Messages, Variants, Graphics, Dialogs, ComCtrls, Menus, StdCtrls,
  jpeg, ComObj, Grids, IWCompText, IWHTMLControls, ErrorBar,
  IWAdvWebGridExcel, IWProgressIndicator, IWDBGrids,
  IWCompProgressBar, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWCompTabControl, IWHelpTip, IWTMSPopup, IWTMSMenus, GanttCh, IWCompMemo,
  IWCompCheckbox, ZAbstractConnection, ZConnection, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset;
type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery1: TZQuery;
    ZConnection1: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
    edtNomeCrit : array of TiwEdit;
    edtNomeAlt : array of TiwEdit;
    cmbbxTipoCrit : array of TiwComboBox;
    cmbbxNiveisCrit : array of TiwComboBox;
    edtMatrizConseq : array of array of TiwEdit;

    NomeCrit, NomeAlt : array of string;
    NumCrit, Numalt : integer;
    Tipocrit : array of integer;
    Niveis : array of integer;
    MatrizConseq, MatrizConseqNorm : array of array of double;
    PesoCrit : array of double;
    CasesOrdemCrit : array of array of integer;
    //OrdemCrit : array of integer;
    Max, Min : array of double;
    MatrizPOA, MatrizPOAPromethee : array of array of integer;
    ResultadoROC, ResultadoPromethee : array of array of double;
    MatrizSol, MatrizSolPromethee : array of array of double;
    ResultAlt, ResultAltPromethee : array of double;
    ResultSol, ResultSolPromethee : array of double;
    coresalt, coresalt1: array of Tcolor;
    CaseSol, CaseSolPromethee : array of integer;
    excelverif : boolean;
    Total: double;
    Crit10, Crit12, Crit13,  Crit14 : array of real;
    Crit20, Crit21, Crit23, Crit24 : array of real;
    MatrizDifVG, MatrizDifVGPromethee : array of array of double;
    MediaDifSol, MaxDifSol, DesvioPadraoDifSol, MinDifSol : array of array of double;
    MediaDifSolPromethee, MaxDifSolPromethee, DesvioPadraoDifSolPromethee, MinDifSolPromethee : array of array of double;
    SobClassMatrix: array of array of real;
    PositiveFlow, NegativeFlow, NetFlow : array of real;
    matrizParaPar: array of array of array of real;
    comp1, comp2, ind, pref, gauss: real;
    OrdemSol, OrdemSolPromethee, ordemalt, ordemaltpromethee : array of integer;
    SimilaridadeGeral : real;
    SimilaridadeAlt, SomenteAditivo, SomenteSobreclassificacao : array of real;
    MinMaxMaxNaoSol,MinMaxMaxNaoSolPromethee : real;
    VetPOA : array of integer;
    MatrizPOA2 : array of array of integer;
    Sconst : array of double;
    NDA : integer;
    epsilon2, MinVK, Epson : double;
    VetPesos : array of array of double;
    AltPOA, PergNAdj : array of double;
    MatrizAltSel : array of array of integer;
    VetNumPOA : array of integer;
    NumSolInst : array of integer;
    AltTamanhoPoa : array of array of integer;
    MatrizComparacaoAlt : array of array of double;
    AmbasAltNaPOA : array of array of double;
    NenhumaAltNaPOA : array of array of double;
    primeiroacesso : boolean;
    email, ID_user, senha : string;
    analyst : integer;
    CritName, RestName, AltName, CritSigla, AltSigla: array of string;
    critcode, altcode, PrefDir: array of Integer;
    CurrentID_problem : integer;
    NameProblem : string;
    problems, id_problem, date_problem : array of string;
    vgmax, vgmin : array of array of double;
    Mediavgmax, MediaVgMin : array of double;
    MaxVgMax, MaxVGMin : array of double;
    MinVgMax, MinVGMin : array of double;
    AmplitudeVG : array of array of double;
    difvg : array of array of array of double;
    MediaAmplitudeVG, MaxAmplitudeVG, MinAmplitudeVG : array of double;
    AmplitudeParAParinst, ProbVitoriaParapar : array of array of array of double;
    mediaprobvitoriaparapar, maximoprobvitoriaparapar, amplitudeparapar : array of array of double;
    MediaGeralProbVitoriaAlt, MaximoGeralProbVitoriaAlt : array of double;
    MediaGeralProbDerrotaAlt, MaximoGeralProbDerrotaAlt : array of double;
    mediadifvg, maximodifvg : array of array of double;
    mediageralprobalt : array of double;
    maxdiff : array of array of array of double;
    mediaMaxDiff : array of array of double;
    MaxMaxdiff : array of array of double;
    GanhoMedio : array of double;
    PerdaMedia : array of double;
    GanhoMaximo : array of double;
    PerdaMaxima : array of double;
    probabilidadealt : array of double;
    novoproblema, retomada : boolean;
    //naltsel : integer;


  end;

implementation

uses lpsolve51, uAPI_SolveLP, UntEntrada;

{$R *.dfm}

end.