; installer.iss — skrypt Inno Setup budujący instalator „Suma Wpłat”
;
; Budowanie instalatora:
;   1) Najpierw uruchomić build_exe.ps1 — powstanie dist\Suma Wplat.exe
;   2) Następnie skompilować ten plik kompilatorem Inno Setup (ISCC.exe installer.iss)
;
; Wynik: dist\Suma Wplat Setup.exe — jeden plik do przekazywania użytkownikom.
;
; Instalacja przebiega w folderze profilu użytkownika (bez uprawnień
; administratora), tworzony jest skrót w menu Start i (opcjonalnie) na
; pulpicie, program rejestruje się w „Dodaj/usuń programy” z pełną
; obsługą odinstalowania.

#define MyAppName "Suma Wpłat"
; Wersje podaje build_all.ps1 przez /DMyAppVersion, czytajac ja z src/app.py —
; dzieki temu numer w programie i w instalatorze nie moga sie rozjechac.
; Wartosc ponizej to zapas, gdy ktos kompiluje ten plik recznie.
#ifndef MyAppVersion
  #define MyAppVersion "1.3"
#endif
; Neutralna nazwa wydawcy — żadnych danych osobistych w instalatorze
; (program ma być swobodnie przekazywalny, bez niczyich danych osobowych)
#define MyAppPublisher "Suma Wpłat"
; Nazwy plików bez polskich znaków. Plik instalatora trafia na stronę wydań
; GitHuba, a „ł” zamieniłoby się w adresie na %C5%82 — link wyglądałby
; podejrzanie i gorzej się go przekleja.
#define MyAppExeName "Suma Wplat.exe"
; Poprzednia nazwa programu — potrzebna tylko do posprzątania po aktualizacji
#define MyAppLegacyName "mBank Analyzer"

[Setup]
AppId={{B6F2B6B0-2F0A-4C7E-9C9A-7E6B6E1A9D3A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
; Instalujemy bez uprawnień administratora — w folderze bieżącego użytkownika,
; aby niedoświadczony użytkownik nie natknął się na monit UAC
PrivilegesRequired=lowest
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Przy aktualizacji Inno domyślnie odtwarza nazwę grupy zapamiętaną w rejestrze,
; więc po zmianie nazwy programu folder w menu Start nadal nazywałby się
; „mBank Analyzer”, choć skróty w środku byłyby już nowe. Ścieżkę instalacji
; zostawiamy zapamiętaną (inaczej stary folder z plikami zostałby sierotą),
; ale nazwę grupy bierzemy zawsze z DefaultGroupName.
UsePreviousGroup=no
OutputDir=dist
OutputBaseFilename=Suma Wplat Setup
SetupIconFile=assets\app.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
; Wyłączamy Restart Manager: program to jeden .exe z PyInstallera, który
; działa jako dwa procesy (bez okna i z oknem). Windows umie zamknąć tylko
; ten z oknem, plik zostaje zajęty i instalator kończył się błędem
; „nie był w stanie automatycznie zamknąć aplikacji” (17.11.2026).
; Zamiast tego sami sprawdzamy i zamykamy program, patrz PrepareToInstall.
CloseApplications=no

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
polish.AppRunningAsk=Program Suma Wpłat jest teraz otwarty. Aby zainstalować aktualizację, trzeba go zamknąć.%n%nZamknąć go teraz?
polish.AppRunningStop=Zamknij program Suma Wpłat i uruchom instalator ponownie.
polish.AppRunningUninstallAsk=Program Suma Wpłat jest teraz otwarty. Aby go odinstalować, trzeba go zamknąć.%n%nZamknąć go teraz?
english.AppRunningAsk=Suma Wpłat is currently running. It has to be closed to install the update.%n%nClose it now?
english.AppRunningStop=Close Suma Wpłat and run the installer again.
english.AppRunningUninstallAsk=Suma Wpłat is currently running. It has to be closed to uninstall it.%n%nClose it now?

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

; Aktualizacja z wersji o poprzedniej nazwie: AppId się nie zmienia, więc
; Inno instaluje w tym samym folderze co wcześniej. Zostałby w nim stary plik
; .exe, a w menu Start — druga, nieaktualna grupa skrótów.
[InstallDelete]
Type: files;          Name: "{app}\{#MyAppLegacyName}.exe"
Type: filesandordirs; Name: "{autoprograms}\{#MyAppLegacyName}"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; Nazwa skrótu na pulpicie wyliczana jest dynamicznie (patrz [Code] niżej) —
; aby nie nadpisać już istniejącego skrótu o tej samej nazwie, który mógł
; zostać utworzony ręcznie i wskazywać na coś innego (np. na wersję
; deweloperską programu uruchamianą bezpośrednio ze źródeł).
Name: "{autodesktop}\{code:GetDesktopIconName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[Code]
// Czy program jest uruchomiony (którykolwiek z jego dwóch procesów).
// WMI zamiast FindWindow: proces bez okna też trzyma plik .exe.
function IsAppRunning(): Boolean;
var
  Locator, Service, Found: Variant;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('.', 'root\CIMV2');
    Found := Service.ExecQuery(
      'SELECT ProcessId FROM Win32_Process WHERE Name = ''{#MyAppExeName}''');
    Result := Found.Count > 0;
  except
    Result := False;   // bez WMI nie blokujemy instalacji
  end;
end;

// Zamyka program razem z procesem-matką (/T) i czeka, aż plik się zwolni.
// Program zapisuje ustawienia od razu przy zmianie, więc nic się nie gubi.
procedure CloseRunningApp();
var
  ResultCode, I: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM "{#MyAppExeName}"', '',
       SW_HIDE, ewWaitUntilTerminated, ResultCode);
  for I := 1 to 20 do
  begin
    if not IsAppRunning() then
      Break;
    Sleep(250);
  end;
  Sleep(300);
end;

// Przed kopiowaniem plików: jeśli program jest otwarty, pytamy i zamykamy.
// Niepusty wynik zatrzymuje instalację z tym komunikatem, nic nie psując.
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if IsAppRunning() then
  begin
    if SuppressibleMsgBox(CustomMessage('AppRunningAsk'),
                          mbConfirmation, MB_YESNO, IDYES) = IDYES then
    begin
      CloseRunningApp();
      if IsAppRunning() then
        Result := CustomMessage('AppRunningStop');
    end
    else
      Result := CustomMessage('AppRunningStop');
  end;
end;

// To samo przy odinstalowaniu: otwarty program zostawiłby swój .exe.
function InitializeUninstall(): Boolean;
begin
  Result := True;
  if IsAppRunning() then
  begin
    if SuppressibleMsgBox(CustomMessage('AppRunningUninstallAsk'),
                          mbConfirmation, MB_YESNO, IDYES) = IDYES then
    begin
      CloseRunningApp();
      Result := not IsAppRunning();
      if not Result then
        SuppressibleMsgBox(CustomMessage('AppRunningStop'), mbError, MB_OK, IDOK);
    end
    else
      Result := False;
  end;
end;

// Zwraca ścieżkę, na którą wskazuje istniejący plik .lnk (przez COM
// WScript.Shell — ten sam mechanizm, którym PowerShell czyta skróty).
// Gdy odczyt się nie powiedzie — zwraca pusty tekst.
function GetShortcutTarget(LinkFile: String): String;
var
  WshShell: Variant;
  Shortcut: Variant;
begin
  Result := '';
  try
    WshShell := CreateOleObject('WScript.Shell');
    Shortcut := WshShell.CreateShortcut(LinkFile);
    Result := Shortcut.TargetPath;
  except
    Result := '';
  end;
end;

// Dobiera nazwę dla skrótu na pulpicie (bez rozszerzenia .lnk):
//   - "Suma Wpłat", jeśli pliku o tej nazwie jeszcze nie ma;
//   - ta sama nazwa, jeśli istniejący skrót wskazuje do wnętrza folderu
//     instalacji tego programu — to znaczy reinstalacja/aktualizacja
//     i można go bezpiecznie nadpisać;
//   - "Suma Wpłat (1)", "(2)", ... jeśli pod tą nazwą leży już
//     OBCY skrót (wskazujący gdzie indziej) — aby go nie nadpisać.
function GetDesktopIconName(Param: String): String;
var
  BaseName, CandidateName, LinkPath, Target, AppDir: String;
  Desktop: String;
  I: Integer;
begin
  Desktop  := ExpandConstant('{autodesktop}');
  AppDir   := ExpandConstant('{app}');
  BaseName := '{#MyAppName}';
  CandidateName := BaseName;
  I := 1;
  while True do
  begin
    LinkPath := Desktop + '\' + CandidateName + '.lnk';
    if not FileExists(LinkPath) then
      Break;

    Target := GetShortcutTarget(LinkPath);
    if (Target <> '') and (Pos(Lowercase(AppDir), Lowercase(Target)) > 0) then
      Break; // to nasz własny skrót z poprzedniej instalacji — używamy tej nazwy

    CandidateName := BaseName + ' (' + IntToStr(I) + ')';
    I := I + 1;
  end;
  Result := CandidateName;
end;

// Po aktualizacji z poprzedniej nazwy usuwamy stary skrót z pulpitu —
// wskazuje na plik .exe, którego już nie ma. Kasujemy wyłącznie skróty
// prowadzące do wnętrza naszego folderu instalacji; obcy plik o tej samej
// nazwie zostawiamy nietknięty (ta sama ostrożność, co w GetDesktopIconName).
procedure RemoveLegacyDesktopIcons();
var
  I: Integer;
  LinkPath, Target, AppDir, Desktop, BaseName: String;
begin
  Desktop  := ExpandConstant('{autodesktop}');
  AppDir   := ExpandConstant('{app}');
  BaseName := '{#MyAppLegacyName}';
  for I := 0 to 9 do
  begin
    if I = 0 then
      LinkPath := Desktop + '\' + BaseName + '.lnk'
    else
      LinkPath := Desktop + '\' + BaseName + ' (' + IntToStr(I) + ').lnk';
    if FileExists(LinkPath) then
    begin
      Target := GetShortcutTarget(LinkPath);
      if (Target <> '') and (Pos(Lowercase(AppDir), Lowercase(Target)) > 0) then
        DeleteFile(LinkPath);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RemoveLegacyDesktopIcons();
end;

// Po usunięciu plików programu proponujemy także wyczyszczenie folderu ustawień.
// Folder %APPDATA%\Suma Wpłat zawiera config.json z danymi sprzedawcy
// i innymi ustawieniami użytkownika — deinstalator domyślnie go nie rusza,
// aby ustawienia przetrwały reinstalację/aktualizację.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataDir, LegacyDataDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    AppDataDir    := ExpandConstant('{userappdata}\{#MyAppName}');
    LegacyDataDir := ExpandConstant('{userappdata}\{#MyAppLegacyName}');
    if DirExists(AppDataDir) then
    begin
      // SuppressibleMsgBox, a nie MsgBox: zwykły MsgBox ignoruje
      // /SUPPRESSMSGBOXES i przy cichej deinstalacji (/VERYSILENT) program
      // zawiesza się na tym pytaniu, czekając na kliknięcie, którego nie ma.
      // Domyślna odpowiedź to IDNO — bez pytania ustawienia zostają.
      if SuppressibleMsgBox(
        'Czy usunąć również folder z ustawieniami programu?' + #13#10 +
        '(dane sprzedawcy, konfiguracja)' + #13#10#13#10 +
        'Wybierz "Nie", aby zachować ustawienia — przydadzą się' + #13#10 +
        'przy ewentualnej ponownej instalacji.',
        mbConfirmation, MB_YESNO, IDNO) = IDYES then
      begin
        DelTree(AppDataDir, True, True, True);
        // Ustawienia zostały skopiowane ze starego folderu, więc kasujemy
        // także jego — inaczej zostałby po nas nieużywany katalog
        if DirExists(LegacyDataDir) then
          DelTree(LegacyDataDir, True, True, True);
      end;
    end;
  end;
end;
