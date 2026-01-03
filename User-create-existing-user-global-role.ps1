$modules = @("Microsoft.Graph.Authentication","Microsoft.Graph.Users","Microsoft.Graph.Identity.DirectoryManagement")
foreach ($m in $modules) {
    $available = Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $available) {
        Write-Host "Installing $m..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
        $available = Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1
    }
    if ($available) {
        Write-Host "Importing $($available.Name) v$($available.Version) from $($available.Path)" -ForegroundColor Yellow
        Import-Module $available.Path -Force
    } else {
        Write-Host "Failed to find or install $m" -ForegroundColor Red
    }
}

$script:isPaused = $false
$script:isStopped = $false
$script:processedCount = 0
$script:successCount = 0
$script:failedCount = 0
$script:skippedCount = 0
$script:logFile = "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$script:keyListenerJob = $null

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
    Add-Content -Path $script:logFile -Value $logMessage
}

function Start-KeyboardListener {
    $script:keyListenerJob = Start-Job -ScriptBlock {
        while ($true) {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                switch ($key.Key) {
                    'P' { return 'PAUSE' }
                    'R' { return 'RESUME' }
                    'S' { return 'STOP' }
                }
            }
            Start-Sleep -Milliseconds 100
        }
    }
}

function Check-KeyPress {
    if ($script:keyListenerJob) {
        $result = Receive-Job -Job $script:keyListenerJob -ErrorAction SilentlyContinue
        if ($result) {
            switch ($result) {
                'PAUSE' { $script:isPaused = $true; Write-Log "Paused - Press R to resume" -Level "WARNING" }
                'RESUME' { $script:isPaused = $false; Write-Log "Resumed" -Level "INFO" }
                'STOP' { $script:isStopped = $true; Write-Log "Stopped" -Level "ERROR" }
            }
            Stop-Job -Job $script:keyListenerJob
            Remove-Job -Job $script:keyListenerJob
            Start-KeyboardListener
        }
    }
}

function Show-ProgressBar {
    param([int]$Current, [int]$Total, [string]$CurrentUser)
    $pct = [math]::Round(($Current / $Total) * 100, 2)
    $filled = [math]::Floor(($pct / 100) * 50)
    $empty = 50 - $filled
    $bar = "#" * $filled + "-" * $empty
    Write-Host "`r[$bar] $pct% | $Current/$Total | $CurrentUser" -NoNewline -ForegroundColor Cyan
}

function Connect-ToMicrosoftGraph {
    try {
        Write-Log "Connecting to Microsoft Graph (interactive browser)..." -Level "INFO"
        Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║  A browser popup will appear — complete full sign-in      ║" -ForegroundColor Yellow
        Write-Host "║  (This may require MFA/2FA)                              ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        
        $scopes = @(
            "User.ReadWrite.All",
            "Directory.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory"
        )

        # Use Interactive Browser Flow (popup) - default behavior when not using device code
        Connect-MgGraph -Scopes $scopes -NoWelcome
        
        Write-Log "Connected successfully!" -Level "SUCCESS"
        
        $context = Get-MgContext
        if ($context -and $context.Account) {
            Write-Log "Logged in as: $($context.Account)" -Level "INFO"
            Write-Log "Tenant: $($context.TenantId)" -Level "INFO"
            return $true
        } else {
            Write-Log "Context is null - authentication may have failed" -Level "ERROR"
            return $false
        }
    } catch {
        Write-Log "Connection failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-GlobalAdminRole {
    try {
        Write-Log "Getting Global Administrator role..." -Level "INFO"
        
        # Check connection first
        $context = Get-MgContext
        if (-not $context) {
            Write-Log "Not connected to Graph API" -Level "ERROR"
            return $null
        }
        
        Write-Log "Querying existing directory roles..." -Level "INFO"
        $role = Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'" -ErrorAction SilentlyContinue

        if (-not $role) {
            Write-Log "Role not found. Preparing to enable..." -Level "WARNING"

            Write-Log "Querying role templates..." -Level "INFO"
            $template = Get-MgDirectoryRoleTemplate -Filter "DisplayName eq 'Global Administrator'" -ErrorAction SilentlyContinue

            if (-not $template) {
                Write-Log "Role template not found; cannot enable role" -Level "ERROR"
                return $null
            }

            Write-Log "Template found: $($template.Id). Enabling role..." -Level "INFO"
            try {
                Enable-MgDirectoryRole -RoleTemplateId $template.Id -ErrorAction Stop -Verbose:$false
                Start-Sleep -Seconds 2
                Write-Log "Re-querying directory roles after enable..." -Level "INFO"
                $role = Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'" -ErrorAction SilentlyContinue
                if ($role) { Write-Log "Role enabled successfully: $($role.Id)" -Level "SUCCESS" } else { Write-Log "Role still not present after enable" -Level "ERROR" }
            } catch {
                Write-Log "Enable-MgDirectoryRole failed: $($_.Exception.Message)" -Level "ERROR"
                return $null
            }
        } else {
            Write-Log "Role found: $($role.Id)" -Level "SUCCESS"
        }
        
        return $role
    } catch {
        Write-Log "Error getting role:  $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Level "ERROR"
        return $null
    }
}

function Assign-GlobalAdminToUser {
    param([string]$UserPrincipalName, [string]$RoleId)
    try {
        $user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
        
        if (-not $user) {
            Write-Log "User not found: $UserPrincipalName" -Level "WARNING"
            $script:failedCount++
            return $false
        }
        
        $existing = Get-MgDirectoryRoleMember -DirectoryRoleId $RoleId -ErrorAction SilentlyContinue | Where-Object { $_.Id -eq $user.Id }
        
        if ($existing) {
            Write-Log "Already has role: $UserPrincipalName" -Level "INFO"
            $script:skippedCount++
            return $true
        }
        
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" }
        
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $RoleId -BodyParameter $body -ErrorAction Stop
        
        Write-Log "Assigned:  $UserPrincipalName" -Level "SUCCESS"
        $script:successCount++
        return $true
    } catch {
        Write-Log "Failed ($UserPrincipalName): $($_.Exception.Message)" -Level "ERROR"
        $script:failedCount++
        return $false
    }
}

function Get-RoleMemberCounts {
    param([string]$RoleId)
    try {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $RoleId -All -ErrorAction Stop
        
        $adminCount = 0
        $userCount = 0
        
        foreach ($member in $members) {
            $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
            if ($user) {
                if ($user.UserPrincipalName -like "*admin*") { $adminCount++ } else { $userCount++ }
            }
        }
        
        return @{ TotalCount = $members.Count; AdminCount = $adminCount; UserCount = $userCount }
    } catch {
        Write-Log "Error getting counts: $($_.Exception.Message)" -Level "WARNING"
        return @{ TotalCount = 0; AdminCount = 0; UserCount = 0 }
    }
}

Write-Host "`n=============================================================" -ForegroundColor Cyan
Write-Host "  Global Administrator Role Assignment Tool" -ForegroundColor Cyan
Write-Host "  (Interactive Authentication - browser popup)" -ForegroundColor Yellow
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host "`nControls: [P] Pause | [R] Resume | [S] Stop`n" -ForegroundColor Yellow

Write-Log "Script started" -Level "INFO"

if (-not (Connect-ToMicrosoftGraph)) { 
    Write-Log "Failed to connect.  Exiting..." -Level "ERROR"
    exit 
}

$context = Get-MgContext
$adminUpn = $null
if ($context -and $context.Account) { $adminUpn = $context.Account }

if (-not $adminUpn) {
    Write-Log "Unable to determine signed-in account UPN; please ensure you signed in interactively" -Level "ERROR"
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit
}

$domain = ($adminUpn -split '@')[-1]
Write-Log "Signed in as: $adminUpn; using domain: $domain" -Level "INFO"

$role = Get-GlobalAdminRole
if (-not $role) { 
    Write-Log "Failed to get Global Admin role. Exiting..." -Level "ERROR"
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 
}

# Mode selection: 1 = create new users then assign; 2 = assign existing users
$mode = Read-Host "Choose action: 1 = Create new users+assign Global Admin, 2 = Assign Global Admin to existing users (enter 1 or 2)"
if ($mode -ne '1' -and $mode -ne '2') {
    Write-Log "Invalid mode selected: $mode" -Level "ERROR"
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit
}


# Built-in names list (one base name or UPN per item). Edit this list as needed.
$rawNames = @(
    "AmaliaAiken",
    "AtlasShook",
    "ClementineElliott",
    "DeborahThompson",
    "KinsleyHolmes",
    "BriarHaralson",
    "ShaneGilbert",
    "KaraMacmillan",
    "AndresSimpson",
    "AidenGardner",
    "OscarVital",
    "MikaylaGama",
    "YareliBiggs",
    "MiloTowers",
    "JourneeGuevara",
    "RobinLee",
    "EstherLewis",
    "HarlemMoon",
    "LeaHooker",
    "BearList",
    "DeclanWoods",
    "VihaanSims",
    "EmoryHenson",
    "ElsaButz",
    "VioletaGarcia",
    "AlijahFarley",
    "ScarletteDeeter",
    "RoyaltyDiaz",
    "JonathanWilkinson",
    "DanteMoss",
    "AlaiyaHernandez",
    "AlmaTeed",
    "FinneganGolden",
    "OmariTrevino",
    "EmerieChavez",
    "VicenteDavis",
    "SaraPointer",
    "SebastianMartin",
    "MilenaSandin",
    "ClarissaRobinson",
    "DuncanWilson",
    "CarmenBurns",
    "ThiagoStamm",
    "QuinnLake",
    "JaniyahLammers",
    "WayneReal",
    "BellamyHaupt",
    "CharleighEmerson",
    "AileenSwain",
    "RayanQuandt",
    "JolieHenkel",
    "BrixtonLyons",
    "CharlotteCampbell",
    "GunnarDavis",
    "DakariSaucedo",
    "KinleyConrad",
    "DilanPearson",
    "MohammadCrockett",
    "KeylaBrown",
    "ItzayanaEcheverria",
    "DariusSalisbury",
    "JadeWalker",
    "LilahMcBride",
    "CooperClodfelter",
    "AshlynStanley",
    "EthanPotter",
    "HelenSmith",
    "DiegoMcDonald",
    "JackMiller",
    "BryanRounds",
    "KinsleyBrill",
    "KarenMcDonald",
    "MariaJohnson",
    "AubreeReed",
    "NolanWarner",
    "LianaJones",
    "KeatonDukes",
    "LukeCarnegie",
    "JaxNagel",
    "EllaOuellette",
    "ForestJohnson",
    "RaeganKelly",
    "BrunoLong",
    "CharleeKimbro",
    "MauriceGaither",
    "KaylaLee",
    "EdenBrandon",
    "HelenJones",
    "KyraSmallwood",
    "JudeAustin",
    "AugustineNorwood",
    "MacyAvila",
    "EverestBoyer",
    "AlissonThomas",
    "BrodieThompson",
    "MariamSweeney",
    "AzaleaCornett",
    "JonesDillard",
    "EvangelineHenderson",
    "SethOgle",
    "HaloPaquette",
    "RaphaelBangs",
    "DavidRex",
    "LillianAdams",
    "DeaconMeade",
    "BellaAtwood",
    "HectorHarris",
    "ValeriaShippy",
    "AubreyMoody",
    "WillieMorales",
    "ValeriaGuinyard",
    "HectorGartman",
    "NavyBarnard",
    "SawyerStout",
    "HarryJohnson",
    "MadelynTaylor",
    "SamSmith",
    "ZekeMarks",
    "MabelRoberge",
    "LayneSussman",
    "WyattHand",
    "AdelePack",
    "MckenzieBlair",
    "MikaylaBlind",
    "AlfredoMartin",
    "ArmandoJones",
    "FrancescaLloyd",
    "EdgarMay",
    "DeaconGibbs",
    "ReignLira",
    "CarlyVargas",
    "WaverlyPiper",
    "KeeganCrandell",
    "PaisleighBrown",
    "MyaLandry",
    "YaretziSeaton",
    "CarsonEntwistle",
    "KameronLevy",
    "ShawnSevigny",
    "ZoraShelton",
    "SawyerBailey",
    "MateoMcMillan",
    "MilanaYanez",
    "KatalinaAdams",
    "SashaMartinez",
    "JazielBrooks",
    "MercyShepardson",
    "EttaReeder",
    "AugustDix",
    "ZaneBarna",
    "VihaanWorthy",
    "RainaSandoval",
    "LouisRadford",
    "JazielMorgan",
    "GuadalupeRoy",
    "RoselynFry",
    "HarleyWilliams",
    "MarcosKitchens",
    "AveryAhmed",
    "ZariaHolt",
    "LaytonLee",
    "TristanPetrin",
    "VincenzoHolland",
    "MinaStrickland",
    "CaydenBlair",
    "MarcelObrien",
    "MilenaWeir",
    "JoelleLavoie",
    "AdamJones",
    "SaulKirkland",
    "IvannaLopez",
    "FelixMarini",
    "SavannahStamps",
    "JasperBraun",
    "EmoryRipley",
    "RaquelViola",
    "LelandRobertson",
    "PaisleeWhite",
    "ZavierAdams",
    "PaisleeMann",
    "RhettSanchez",
    "KeenanMoore",
    "CleoKnox",
    "EithanFox",
    "PorterTurner",
    "AnnabelleBarth",
    "KenjiKelleher",
    "ColetteBartholomew",
    "HadleyLeon",
    "JadeMyer",
    "WalterAdams",
    "RaeganAnderson",
    "QuinnPardo",
    "AlonzoDowning",
    "FayeRodriguez",
    "HenrikCargill",
    "ZayleeGonzalez",
    "ForestMayer",
    "CarolineRogers",
    "AvalynnLyall",
    "LeslieErickson",
    "MordechaiPoore",
    "JakeCook",
    "FreyjaBragg",
    "PedroSummer",
    "MarcusOstlund",
    "GeraldHansen",
    "JulissaHamilton",
    "AllisonSharp",
    "SophiaTorrence",
    "MarcellusMason",
    "MarloweJackson",
    "SylviaTrammell",
    "KyroSanchez",
    "MalikRios",
    "ZanderHartman",
    "KamrynCassano",
    "YahirAustin",
    "ChrisMackie",
    "OctaviaShaw",
    "MckennaDaugherty",
    "FrancisBeck",
    "TeaganKeller",
    "JiannaWilliamson",
    "KataleyaClaar",
    "MatteoJack",
    "ElianBrooks",
    "MarenBrown",
    "KristianPrather",
    "GaelMurray",
    "RayBarnes",
    "AlondraPonton",
    "EliannaGreenspan",
    "AyleenTucker",
    "MarceloThomas",
    "DrewBrasher",
    "RemiTurner",
    "FelixMiller",
    "WaylonJuan",
    "AriesReno",
    "AliciaGreenfield",
    "IzaiahBrown",
    "IzaiahFields",
    "CoraMcConnell",
    "PaulaMoldenhauer",
    "CalMarshall",
    "AdleyShead",
    "SylviaCarder",
    "SevynWatson",
    "NikolasBowers",
    "FernandoBreazeale",
    "AriyaArrellano",
    "TristenWorley",
    "LandryHamilton",
    "KaiserHall",
    "MeghanCrawford",
    "SierraMedford",
    "AyaanFerranti",
    "ScoutHathorn",
    "GordonDion",
    "DahliaBurton",
    "AmelieTucker",
    "MaisonKaspar",
    "JoleneChavira",
    "LylahMabe",
    "DamariLitchfield",
    "AmandaCummings",
    "NathanaelBeam",
    "GracelynnMcKeever",
    "KyloGammon",
    "WalterBrown",
    "ElsieHarris",
    "WinstonRarick",
    "BakerRoberson",
    "ColterMiller",
    "MariaEdgar",
    "AlayahOakman",
    "MyraMorales",
    "MohammedHinson",
    "RoryMcGee",
    "JusticePrins",
    "GemmaAlvarez",
    "JulissaDavis",
    "OakleeLee",
    "HavenMcCabe",
    "FrancesNeal",
    "NikoClayton",
    "DenverLuckett",
    "EricGomez",
    "CarlRichardson",
    "LoganKenyon",
    "JillianHarvey",
    "FlorenceStewart",
    "TatumHill",
    "JoannaFugate",
    "EllisonCraft",
    "ZakaiHite",
    "MaximoKnuth",
    "DeanBennett",
    "TitanDay"
)


Write-Log "Using embedded names list: $($rawNames.Count) entries" -Level "INFO"

$userList = @()

if ($mode -eq '1') {
    $createCountInput = Read-Host "How many users to create? Enter a number"
    if (-not [int]::TryParse($createCountInput, [ref]$null)) {
        $createCount = [int]$createCountInput
    } else {
        $createCount = [int]$createCountInput
    }
    if ($createCount -le 0) {
        Write-Log "Invalid create count: $createCount" -Level "ERROR"
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        exit
    }

    $password = '1Zahangir@8899'
    $createdCreds = @()
    $baseCount = $rawNames.Count

    for ($i = 0; $i -lt $createCount; $i++) {
        $base = $rawNames[$i % $baseCount]
        $suffix = [math]::Floor($i / $baseCount)
        if ($suffix -gt 0) { $usernameBase = "$base$suffix"; $displayName = "$base $suffix" } else { $usernameBase = $base; $displayName = $base }
        $mailNick = ($usernameBase -replace '\s','') -replace '[^A-Za-z0-9]',''
        $upn = "$mailNick@$domain"

        # Ensure uniqueness by checking existing user
        $attempt = 0
        $finalUpn = $upn
        while ((Get-MgUser -Filter "userPrincipalName eq '$finalUpn'" -ErrorAction SilentlyContinue) -and $attempt -lt 20) {
            $attempt++
            $finalUpn = "$mailNick$attempt@$domain"
        }

        try {
            New-MgUser -AccountEnabled:$true -DisplayName $displayName -MailNickname $mailNick -UserPrincipalName $finalUpn -PasswordProfile @{ ForceChangePasswordNextSignIn = $false; Password = $password } -ErrorAction Stop
            Write-Log "Created user: $finalUpn" -Level "SUCCESS"
            $createdCreds += [PSCustomObject]@{ UserPrincipalName = $finalUpn; Password = $password }

            # Assign role immediately
            Assign-GlobalAdminToUser -UserPrincipalName $finalUpn -RoleId $role.Id

            # Add to userList for later reporting
            $userList += [PSCustomObject]@{ UserPrincipalName = $finalUpn }
        } catch {
            Write-Log "Failed to create $($finalUpn): $($_.Exception.Message)" -Level "ERROR"
        }

        Start-Sleep -Milliseconds 200
    }

    # Write created credentials to CSV
    $outPath = "CreatedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $createdCreds | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
    Write-Log "Created credentials exported: $outPath" -Level "INFO"
}

else {
    # mode 2: assign Global Admin to all users in the signed-in domain
    Write-Log "Mode 2 selected: querying all users in domain $domain" -Level "INFO"
    try {
        $filter = "endsWith(userPrincipalName,'@$domain')"
        $allUsers = Get-MgUser -Filter $filter -ConsistencyLevel eventual -All -ErrorAction Stop
        foreach ($u in $allUsers) {
            $userList += [PSCustomObject]@{ UserPrincipalName = $u.UserPrincipalName }
        }
    } catch {
        Write-Log "Failed to retrieve domain users: $($_.Exception.Message)" -Level "ERROR"
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        exit
    }
}

Write-Log "Users to process: $($userList.Count)" -Level "INFO"

Start-KeyboardListener
Write-Host "`n"
Write-Log "Starting role assignment..." -Level "INFO"
Write-Host ""

$total = $userList.Count
foreach ($user in $userList) {
    while ($script:isPaused -and -not $script:isStopped) {
        Check-KeyPress
        Start-Sleep -Milliseconds 500
    }
    
    if ($script:isStopped) {
        Write-Host "`n"
        Write-Log "Process stopped by user" -Level "WARNING"
        break
    }
    
    Check-KeyPress
    
    $script:processedCount++
    $upn = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { $user.username }
    Show-ProgressBar -Current $script:processedCount -Total $total -CurrentUser $upn

    Assign-GlobalAdminToUser -UserPrincipalName $upn -RoleId $role.Id
    
    Start-Sleep -Milliseconds 300
}

Write-Host "`n`n"
Write-Log "Getting final role member counts..." -Level "INFO"
$counts = Get-RoleMemberCounts -RoleId $role.Id

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "                   SUMMARY REPORT" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Total Processed:      $($script:processedCount)" -ForegroundColor White
Write-Host "Successfully Assigned: $($script:successCount)" -ForegroundColor Green
Write-Host "Already Had Role:    $($script:skippedCount)" -ForegroundColor Yellow
Write-Host "Failed:              $($script:failedCount)" -ForegroundColor Red
Write-Host "------------------------------------------------------------" -ForegroundColor Green
Write-Host "Global Admin Role Members:" -ForegroundColor Cyan
Write-Host "  Total Members: $($counts.TotalCount)" -ForegroundColor White
Write-Host "  Admin Accounts: $($counts.AdminCount)" -ForegroundColor Cyan
Write-Host "  User Accounts: $($counts.UserCount)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Green

Write-Log "Summary - Processed: $($script:processedCount) | Success: $($script:successCount) | Skipped: $($script:skippedCount) | Failed: $($script:failedCount)" -Level "INFO"

if ($script:keyListenerJob) {
    Stop-Job -Job $script:keyListenerJob -ErrorAction SilentlyContinue
    Remove-Job -Job $script:keyListenerJob -ErrorAction SilentlyContinue
}

Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "Script completed.  Log file: $($script:logFile)" -Level "SUCCESS"