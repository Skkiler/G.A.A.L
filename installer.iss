; -- installer.iss --

[Setup]
; Identificador único da sua aplicação (substitua pelo seu GUID)
AppId={{12345678-90AB-CDEF-1234-567890ABCDEF}}
AppName=G.A.A.L
AppVersion=1.0
; Diretório padrão de instalação
DefaultDirName={pf}\G.A.A.L
DefaultGroupName=G.A.A.L
Uninstallable=yes
CreateUninstallRegKey=yes
OutputBaseFilename=G.A.A.L_Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Files]
; Executável principal + splash.png
Source: "G.A.A.L.exe";       DestDir: "{app}"; Flags: ignoreversion restartreplace
Source: "splash.png";        DestDir: "{app}"; Flags: ignoreversion
; Toda a pasta _internal
Source: "_internal\*";       DestDir: "{app}\_internal"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\G.A.A.L";       Filename: "{app}\G.A.A.L.exe"
Name: "{userdesktop}\G.A.A.L"; Filename: "{app}\G.A.A.L.exe"; Tasks: desktopicon

[Tasks]
Name: desktopicon; Description: "Criar atalho na área de trabalho"; GroupDescription: "Tarefas adicionais:"; Flags: unchecked

[Run]
Filename: "{app}\G.A.A.L.exe"; Description: "Iniciar G.A.A.L"; Flags: nowait postinstall skipifsilent

; ================= opcional: remoção de pasta vazia após uninstall =================
[UninstallDelete]
Type: filesandordirs; Name: "{app}"

; ================== lógica pré-instalação em [Code] ==================
[Code]
// Intercepta o clique em “Avançar” na página de escolha de pasta
function NextButtonClick(CurPageID: Integer): Boolean;
var
  Msg: String;
begin
  Result := True;
  if CurPageID = wpSelectDir then
  begin
    // monta a mensagem concatenando o caminho já expandido
    Msg := 'Já existe uma instalação de G.A.A.L em:' + #13#10 +
           ExpandConstant('{app}') + #13#10#13#10 +
           'Deseja substituir?';
    // se o usuário clicar em “Não”, cancela o avanço
    if MsgBox(Msg, mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;
