; IMTSolver 2.0 - instalador Inno Setup
;
; Compilar:
;   ISCC.exe IMTSolver.iss                     -> pacote completo
;   ISCC.exe /DLEVE IMTSolver.iss              -> pacote leve (sem SCIP nem Gurobi)
;
; Nao pede administrador: instala no perfil do usuario, na pasta de
; suplementos do Excel, que ja e local confiavel (nenhum aviso de macro).

#ifdef LEVE
  #define MotorPasta "imtsolver_leve"
  #define Sufixo     "_leve"
  #define Sabor      "leve (CBC, HiGHS e NEOS)"
#else
  #define MotorPasta "imtsolver"
  #define Sufixo     ""
  #define Sabor      "completo (CBC, HiGHS, SCIP e NEOS)"
#endif

[Setup]
AppId={{9E3C1A54-7C2E-4B1E-9E1F-2A6D4B8F0C31}
AppName=IMTSolver
AppVersion=2.0
AppVerName=IMTSolver 2.0
AppPublisher=Pedro da Silva Bezerra - Instituto Maua de Tecnologia
DefaultDirName={userappdata}\Microsoft\AddIns
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=no
CreateAppDir=yes
Uninstallable=yes
UninstallDisplayName=IMTSolver 2.0 (suplemento do Excel)
UninstallDisplayIcon={app}\{#MotorPasta}\imtsolver.exe
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=IMTSolver_2.0_Setup{#Sufixo}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
AppComments=Otimizacao linear e inteira no Excel - pacote {#Sabor}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[InstallDelete]
; O motor tem que ser substituido inteiro, nao mesclado. Sem isto, quem
; instala o pacote leve por cima do completo fica com o SCIP e o Gurobi
; antigos na pasta: 175 MB de sobra, e a lista de solvers mente.
Type: filesandordirs; Name: "{app}\imtsolver"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\imtsolver"

[Files]
Source: "IMTSolver.xlam";           DestDir: "{app}"; Flags: ignoreversion
Source: "IMTSolver_Exemplos.xlsx";  DestDir: "{app}"; Flags: ignoreversion
Source: "LEIAME.md";                DestDir: "{app}"; Flags: ignoreversion
Source: "{#MotorPasta}\*";          DestDir: "{app}\imtsolver"; Flags: ignoreversion recursesubdirs createallsubdirs

[Messages]
brazilianportuguese.WelcomeLabel2=Isto vai instalar o IMTSolver 2.0, pacote {#Sabor}.%n%nO IMTSolver e um suplemento de otimizacao para o Excel: voce monta o modelo na planilha e escolhe o solver. Feche o Excel antes de continuar.
brazilianportuguese.FinishedLabel=O IMTSolver foi instalado. Abra o Excel: a aba IMTSolver aparece na faixa de opcoes.

[Code]

function ExcelAberto(): Boolean;
begin
  Result := FindWindowByClassName('XLMAIN') <> 0;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if ExcelAberto() then
  begin
    // numa instalacao silenciosa nao ha ninguem para clicar em OK:
    // sem esta guarda o instalador ficaria parado para sempre
    if not WizardSilent() then
      MsgBox('O Excel esta aberto.' + #13#10#13#10 +
             'Feche o Excel e rode este instalador de novo.',
             mbError, MB_OK);
    Result := False;
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  if ExcelAberto() then
  begin
    if not UninstallSilent() then
      MsgBox('O Excel esta aberto.' + #13#10#13#10 +
             'Feche o Excel e desinstale de novo.',
             mbError, MB_OK);
    Result := False;
  end;
end;

// As versoes de Office que usam este registro:
//   16.0 = 2016/2019/2021/365    15.0 = 2013    14.0 = 2010
function ChaveDaVersao(i: Integer): String;
begin
  case i of
    0: Result := '16.0';
    1: Result := '15.0';
    2: Result := '14.0';
  else
    Result := '';
  end;
end;

// Registra o suplemento na primeira chave OPEN / OPEN1 / OPEN2 livre,
// para nao atropelar o Solver do Excel nem o OpenSolver de quem ja usa.
procedure RegistrarNoExcel();
var
  i, n: Integer;
  chave, nome, valor, lido, livre: String;
  jaTem: Boolean;
begin
  valor := '/R "IMTSolver.xlam"';
  for i := 0 to 2 do
  begin
    if not RegKeyExists(HKEY_CURRENT_USER,
                        'Software\Microsoft\Office\' + ChaveDaVersao(i) + '\Excel') then
      Continue;

    chave := 'Software\Microsoft\Office\' + ChaveDaVersao(i) + '\Excel\Options';

    // Varremos TODAS as chaves OPEN, sem parar na primeira que falta.
    // Parar cedo criava um registro duplicado quando o Excel deixava um
    // buraco na numeracao (OPEN, OPEN1, OPEN3 sem OPEN2, por exemplo).
    jaTem := False;
    livre := '';
    for n := 0 to 29 do
    begin
      if n = 0 then nome := 'OPEN' else nome := 'OPEN' + IntToStr(n);
      if RegQueryStringValue(HKEY_CURRENT_USER, chave, nome, lido) then
      begin
        if Pos('IMTSolver.xlam', lido) > 0 then
          jaTem := True;
      end
      else
        if livre = '' then livre := nome;
    end;

    if (not jaTem) and (livre <> '') then
      RegWriteStringValue(HKEY_CURRENT_USER, chave, livre, valor);
  end;
end;

procedure LimparRegistroDoExcel();
var
  i, n: Integer;
  chave, nome, lido, base: String;
begin
  for i := 0 to 2 do
  begin
    chave := 'Software\Microsoft\Office\' + ChaveDaVersao(i) + '\Excel\Options';
    if not RegKeyExists(HKEY_CURRENT_USER, chave) then
      Continue;
    for n := 0 to 29 do
    begin
      if n = 0 then nome := 'OPEN' else nome := 'OPEN' + IntToStr(n);
      if RegQueryStringValue(HKEY_CURRENT_USER, chave, nome, lido) then
        if Pos('IMTSolver.xlam', lido) > 0 then
          RegDeleteValue(HKEY_CURRENT_USER, chave, nome);
    end;

    // O Excel guarda a marcacao da caixinha da janela "Suplementos" em
    // outro lugar, fora das chaves OPEN. Sem limpar aqui, quem desinstalar
    // leva um aviso "nao foi possivel encontrar IMTSolver.xlam" em toda
    // abertura do Excel - com o registro OPEN ja limpo, o que torna o
    // problema dificil de achar.
    base := 'Software\Microsoft\Office\' + ChaveDaVersao(i) + '\Excel\';
    RegDeleteValue(HKEY_CURRENT_USER, base + 'Add-in Manager', 'IMTSolver.xlam');
    RegDeleteValue(HKEY_CURRENT_USER, base + 'Add-in Manager', ExpandConstant('{app}\IMTSolver.xlam'));
    RegDeleteValue(HKEY_CURRENT_USER, base + 'AddInLoadTimes', ExpandConstant('{app}\IMTSolver.xlam'));
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegistrarNoExcel();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    LimparRegistroDoExcel();
end;
