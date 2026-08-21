<#
.SYNOPSIS
NASまたはローカルの資料フォルダから統合議事録ベースを生成します。

.DESCRIPTION
資料フォルダをローカルジョブへコピーし、Windows PowerPointによるPDF変換、
PodmanコンテナによるDOCX生成、指定先への完成ファイル返却を順に実行します。
PowerPointファイルは入力フォルダ直下のPPTX/PPTMを対象とします。
-ExportPdfを指定した場合は、生成したPDFも出力先のPDFサブフォルダへ保存します。
ローカルジョブは成功・失敗を問わず処理終了時に削除します。

.EXAMPLE
.\build_minutes.ps1 -Source "\\nas\share\materials" -Destination "\\nas\share\minutes"

.EXAMPLE
.\build_minutes.ps1 -Source "C:\work\materials" -Force -ExportPdf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Source,

    [Parameter(Position = 1)]
    [string]$Destination,

    [ValidateNotNullOrEmpty()]
    [string]$OutputName = "統合議事録ベース.docx",

    [ValidateNotNullOrEmpty()]
    [string]$ImageName = "ppt2word",

    [string]$JobRoot,

    [ValidateRange(1, 1200)]
    [int]$Dpi = 300,

    [ValidateRange(1, 95)]
    [int]$JpegQuality = 85,

    [switch]$Force,

    [switch]$ExportPdf
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$SupportedExtensions = @(".pptx", ".pptm")
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$PdfExporter = Join-Path $ScriptDirectory "pptx_to_pdf.js"
$LogPath = $null
$JobDirectory = $null
$Completed = $false
$CurrentStage = "初期化"
$ExitCode = 1

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$MustExist
    )

    if ($MustExist) {
        return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    if ($null -ne $script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function Invoke-LoggedNative {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @(),

        [int[]]$SuccessExitCodes = @(0),

        [System.Text.Encoding]$NativeOutputEncoding = $null
    )

    Write-Log ("実行: {0} {1}" -f $FilePath, ($ArgumentList -join " "))

    $previousErrorActionPreference = $ErrorActionPreference
    $previousConsoleOutputEncoding = $null
    $consoleEncodingChanged = $false
    try {
        # Windows PowerShell decodes native stdout/stderr using the console output
        # encoding. Override it only for commands whose output encoding is known.
        if ($null -ne $NativeOutputEncoding) {
            $previousConsoleOutputEncoding = [Console]::OutputEncoding
            [Console]::OutputEncoding = $NativeOutputEncoding
            $consoleEncodingChanged = $true
        }

        # Native stderr is captured as output. The process exit code is checked
        # explicitly so normal diagnostic messages do not become terminating errors.
        $ErrorActionPreference = "Continue"
        $commandOutput = & $FilePath @ArgumentList 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        if ($consoleEncodingChanged) {
            [Console]::OutputEncoding = $previousConsoleOutputEncoding
        }
        $ErrorActionPreference = $previousErrorActionPreference
    }

    foreach ($item in @($commandOutput)) {
        $outputLine = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($outputLine)) {
            Write-Log $outputLine
        }
    }

    if ($SuccessExitCodes -notcontains $exitCode) {
        throw "コマンドが終了コード $exitCode で失敗しました: $FilePath"
    }

    return $exitCode
}

function Assert-CommandAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "必要なコマンドが見つかりません: $Name"
    }
}

function Assert-SafeMountPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Path.Contains(",")) {
        throw "Podmanの--mountで扱えないため、カンマを含まない作業パスを指定してください: $Path"
    }
}

function Get-PowerPointFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    return @(
        Get-ChildItem -LiteralPath $Directory -File -Force |
            Where-Object {
                ($SupportedExtensions -contains $_.Extension.ToLowerInvariant()) -and
                (-not $_.Name.StartsWith("~$"))
            } |
            Sort-Object Name
    )
}

function Assert-UniquePowerPointStems {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    $duplicates = @(
        $Files |
            Group-Object { $_.BaseName.ToLowerInvariant() } |
            Where-Object { $_.Count -gt 1 }
    )

    if ($duplicates.Count -gt 0) {
        $names = $duplicates | ForEach-Object { ($_.Group.Name -join ", ") }
        throw "同名PDFが衝突するPowerPointファイルがあります: $($names -join "; ")"
    }
}

try {
    if (-not (Test-Path -LiteralPath $PdfExporter -PathType Leaf)) {
        throw "PDF変換スクリプトが見つかりません: $PdfExporter"
    }

    if ([System.IO.Path]::GetFileName($OutputName) -ne $OutputName) {
        throw "-OutputName にはファイル名だけを指定してください: $OutputName"
    }
    if ($OutputName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw "-OutputName に使用できない文字が含まれています: $OutputName"
    }
    if ([System.IO.Path]::GetExtension($OutputName).ToLowerInvariant() -ne ".docx") {
        throw "-OutputName は.docxファイル名で指定してください: $OutputName"
    }

    $SourceDirectory = Get-AbsolutePath -Path $Source -MustExist
    if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
        throw "-Source にはフォルダを指定してください: $SourceDirectory"
    }

    [System.IO.FileInfo[]]$SourcePowerPointFiles = @(Get-PowerPointFiles -Directory $SourceDirectory)
    if ($SourcePowerPointFiles.Count -eq 0) {
        throw "入力フォルダ直下にPPTX/PPTMがありません: $SourceDirectory"
    }
    Assert-UniquePowerPointStems -Files $SourcePowerPointFiles

    if ([string]::IsNullOrWhiteSpace($Destination)) {
        $DestinationDirectory = $SourceDirectory
    }
    else {
        $DestinationDirectory = Get-AbsolutePath -Path $Destination
    }
    New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    $DestinationDirectory = Get-AbsolutePath -Path $DestinationDirectory -MustExist

    $DestinationOutput = Join-Path $DestinationDirectory $OutputName
    if ((Test-Path -LiteralPath $DestinationOutput) -and (-not $Force)) {
        throw "出力先が既に存在します。上書きする場合は-Forceを指定してください: $DestinationOutput"
    }

    $PdfDestinationDirectory = $null
    if ($ExportPdf) {
        $PdfDestinationDirectory = Join-Path $DestinationDirectory "PDF"
        if (Test-Path -LiteralPath $PdfDestinationDirectory -PathType Leaf) {
            throw "PDF出力先と同名のファイルが存在します: $PdfDestinationDirectory"
        }

        foreach ($sourceFile in $SourcePowerPointFiles) {
            $pdfDestination = Join-Path $PdfDestinationDirectory ($sourceFile.BaseName + ".pdf")
            if ((Test-Path -LiteralPath $pdfDestination) -and (-not $Force)) {
                throw "PDF出力先が既に存在します。上書きする場合は-Forceを指定してください: $pdfDestination"
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($JobRoot)) {
        if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
            $JobRoot = Join-Path $env:LOCALAPPDATA "ppt2word\jobs"
        }
        else {
            $JobRoot = Join-Path $env:TEMP "ppt2word-jobs"
        }
    }
    $JobRoot = Get-AbsolutePath -Path $JobRoot
    $sourcePrefix = $SourceDirectory.TrimEnd([char[]]"\/") + [System.IO.Path]::DirectorySeparatorChar
    $jobRootPrefix = $JobRoot.TrimEnd([char[]]"\/") + [System.IO.Path]::DirectorySeparatorChar
    if ($jobRootPrefix.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "-JobRoot は入力元フォルダの外側に指定してください: $JobRoot"
    }
    New-Item -ItemType Directory -Path $JobRoot -Force | Out-Null

    $trimmedSourceDirectory = $SourceDirectory.TrimEnd([char[]]"\/")
    $sourceLeaf = Split-Path -Leaf $trimmedSourceDirectory
    if ([string]::IsNullOrWhiteSpace($sourceLeaf)) {
        $sourceLeaf = "source"
    }
    $invalidNameCharacters = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($character in $invalidNameCharacters) {
        $sourceLeaf = $sourceLeaf.Replace([string]$character, "_")
    }

    $jobName = "{0}_{1}_{2}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), $PID, $sourceLeaf
    $JobDirectory = Join-Path $JobRoot $jobName
    $InputDirectory = Join-Path $JobDirectory "input"
    $OutputDirectory = Join-Path $JobDirectory "output"
    New-Item -ItemType Directory -Path $InputDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $LogPath = Join-Path $JobDirectory "process.log"
    New-Item -ItemType File -Path $LogPath -Force | Out-Null

    Write-Log "処理を開始します。"
    Write-Log "入力元: $SourceDirectory"
    Write-Log "出力先: $DestinationOutput"
    Write-Log "ローカルジョブ: $JobDirectory"

    Assert-CommandAvailable -Name "robocopy.exe"
    Assert-CommandAvailable -Name "cscript.exe"
    Assert-CommandAvailable -Name "podman.exe"

    Assert-SafeMountPath -Path $InputDirectory
    Assert-SafeMountPath -Path $OutputDirectory

    $imageStatus = Invoke-LoggedNative `
        -FilePath "podman.exe" `
        -ArgumentList @("image", "exists", $ImageName) `
        -SuccessExitCodes @(0, 1)
    if ($imageStatus -ne 0) {
        throw "Podmanイメージが見つかりません。先にビルドしてください: podman build -t $ImageName ."
    }

    Write-Log "NASまたは元フォルダからローカルジョブへ資料一式をコピーします。"
    $robocopyArguments = @(
        $SourceDirectory,
        $InputDirectory,
        "/E",
        "/COPY:DAT",
        "/DCOPY:DAT",
        "/R:2",
        "/W:2",
        "/XJ",
        "/NFL",
        "/NDL",
        "/NP"
    )
    $CurrentStage = "入力資料のコピー"
    Invoke-LoggedNative -FilePath "robocopy.exe" -ArgumentList $robocopyArguments -SuccessExitCodes (0..7) | Out-Null
    Write-Log "入力資料のコピーが完了しました。"

    $CurrentStage = "コピー後のPowerPointファイル検査"
    Write-Log "コピー後のPowerPointファイル一覧を検査します。"
    [System.IO.FileInfo[]]$PowerPointFiles = @(Get-PowerPointFiles -Directory $InputDirectory)
    if ($PowerPointFiles.Count -eq 0) {
        throw "入力フォルダ直下にPPTX/PPTMがありません: $SourceDirectory"
    }
    Assert-UniquePowerPointStems -Files $PowerPointFiles
    $missingCopiedFiles = @(
        foreach ($sourceFile in $SourcePowerPointFiles) {
            $copiedPath = Join-Path $InputDirectory $sourceFile.Name
            if (-not (Test-Path -LiteralPath $copiedPath -PathType Leaf)) {
                $sourceFile.Name
            }
        }
    )
    if ($missingCopiedFiles.Count -gt 0) {
        throw "ローカルコピーに不足しているPowerPointファイルがあります: $($missingCopiedFiles -join ', ')"
    }
    Write-Log ("PowerPointファイル: {0}件" -f $PowerPointFiles.Count)

    $CurrentStage = "PowerPoint PDF変換"
    Write-Log "Windows版PowerPointでPDFを生成します。"
    $pdfArguments = @("//nologo", $PdfExporter) + @($PowerPointFiles.FullName)
    Invoke-LoggedNative -FilePath "cscript.exe" -ArgumentList $pdfArguments | Out-Null

    foreach ($presentation in $PowerPointFiles) {
        $pdfPath = Join-Path $InputDirectory ($presentation.BaseName + ".pdf")
        if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf)) {
            throw "PDFが生成されていません: $pdfPath"
        }
    }
    Write-Log "全PowerPointファイルの同名PDFを確認しました。"

    $CurrentStage = "PodmanによるDOCX生成"
    Write-Log "Podmanコンテナで統合Word文書を生成します。"
    $containerOutputPath = "/output/$OutputName"
    $podmanArguments = @(
        "run",
        "--rm",
        "--mount", "type=bind,source=$InputDirectory,target=/workspace,readonly",
        "--mount", "type=bind,source=$OutputDirectory,target=/output",
        $ImageName,
        "--dpi", [string]$Dpi,
        "--jpeg-quality", [string]$JpegQuality,
        "-o", $containerOutputPath
    )
    $podmanArguments += @($PowerPointFiles.Name)

    Invoke-LoggedNative `
        -FilePath "podman.exe" `
        -ArgumentList $podmanArguments `
        -NativeOutputEncoding ([System.Text.Encoding]::UTF8) | Out-Null

    $LocalOutput = Join-Path $OutputDirectory $OutputName
    if (-not (Test-Path -LiteralPath $LocalOutput -PathType Leaf)) {
        throw "コンテナの出力DOCXが見つかりません: $LocalOutput"
    }

    $CurrentStage = "完成DOCXの返却"
    Write-Log "完成DOCXを指定先へコピーします。"
    if ((Test-Path -LiteralPath $DestinationOutput) -and (-not $Force)) {
        throw "処理中に出力先ファイルが作成されました。上書きせず終了します: $DestinationOutput"
    }
    $temporaryDestination = Join-Path $DestinationDirectory (
        "{0}.{1}.tmp" -f $OutputName, [Guid]::NewGuid().ToString("N")
    )
    $backupDestination = $null
    try {
        Copy-Item -LiteralPath $LocalOutput -Destination $temporaryDestination -Force

        $localLength = (Get-Item -LiteralPath $LocalOutput).Length
        $temporaryLength = (Get-Item -LiteralPath $temporaryDestination).Length
        if ($localLength -ne $temporaryLength) {
            throw "一時コピーのファイルサイズが一致しません: $temporaryDestination"
        }

        if (Test-Path -LiteralPath $DestinationOutput) {
            if (-not $Force) {
                throw "処理中に出力先ファイルが作成されました。上書きせず終了します: $DestinationOutput"
            }

            $backupDestination = Join-Path $DestinationDirectory (
                ".{0}.{1}.bak" -f $OutputName, [Guid]::NewGuid().ToString("N")
            )
            Move-Item -LiteralPath $DestinationOutput -Destination $backupDestination
        }

        try {
            Move-Item -LiteralPath $temporaryDestination -Destination $DestinationOutput
        }
        catch {
            $publishError = $_
            if (($null -ne $backupDestination) -and
                (Test-Path -LiteralPath $backupDestination) -and
                (-not (Test-Path -LiteralPath $DestinationOutput))) {
                try {
                    Move-Item -LiteralPath $backupDestination -Destination $DestinationOutput
                    $backupDestination = $null
                }
                catch {
                    Write-Log "警告: 既存DOCXの自動復元に失敗しました。バックアップを保持します: $backupDestination"
                }
            }
            throw $publishError
        }

        if (($null -ne $backupDestination) -and (Test-Path -LiteralPath $backupDestination)) {
            try {
                Remove-Item -LiteralPath $backupDestination -Force
                $backupDestination = $null
            }
            catch {
                Write-Log "警告: 置換前DOCXのバックアップを削除できませんでした: $backupDestination"
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryDestination) {
            Remove-Item -LiteralPath $temporaryDestination -Force -ErrorAction SilentlyContinue
        }
        if (($null -ne $backupDestination) -and (Test-Path -LiteralPath $backupDestination)) {
            Write-Log "置換前DOCXのバックアップを保持しています: $backupDestination"
        }
    }

    if ($ExportPdf) {
        $CurrentStage = "PDFの返却"
        Write-Log "生成PDFを出力先のPDFフォルダへコピーします。"
        New-Item -ItemType Directory -Path $PdfDestinationDirectory -Force | Out-Null

        foreach ($presentation in $PowerPointFiles) {
            $localPdf = Join-Path $InputDirectory ($presentation.BaseName + ".pdf")
            $destinationPdf = Join-Path $PdfDestinationDirectory ($presentation.BaseName + ".pdf")

            if ((Test-Path -LiteralPath $destinationPdf) -and (-not $Force)) {
                throw "処理中にPDF出力先が作成されました。上書きせず終了します: $destinationPdf"
            }

            $temporaryPdf = Join-Path $PdfDestinationDirectory (
                "{0}.{1}.tmp" -f ($presentation.BaseName + ".pdf"), [Guid]::NewGuid().ToString("N")
            )
            $backupPdf = $null
            try {
                Copy-Item -LiteralPath $localPdf -Destination $temporaryPdf -Force

                $localPdfLength = (Get-Item -LiteralPath $localPdf).Length
                $temporaryPdfLength = (Get-Item -LiteralPath $temporaryPdf).Length
                if ($localPdfLength -ne $temporaryPdfLength) {
                    throw "PDF一時コピーのファイルサイズが一致しません: $temporaryPdf"
                }

                if (Test-Path -LiteralPath $destinationPdf) {
                    if (-not $Force) {
                        throw "処理中にPDF出力先が作成されました。上書きせず終了します: $destinationPdf"
                    }

                    $backupPdf = Join-Path $PdfDestinationDirectory (
                        "{0}.{1}.bak" -f ($presentation.BaseName + ".pdf"), [Guid]::NewGuid().ToString("N")
                    )
                    Move-Item -LiteralPath $destinationPdf -Destination $backupPdf
                }

                try {
                    Move-Item -LiteralPath $temporaryPdf -Destination $destinationPdf
                }
                catch {
                    $publishPdfError = $_
                    if (($null -ne $backupPdf) -and
                        (Test-Path -LiteralPath $backupPdf) -and
                        (-not (Test-Path -LiteralPath $destinationPdf))) {
                        try {
                            Move-Item -LiteralPath $backupPdf -Destination $destinationPdf
                            $backupPdf = $null
                        }
                        catch {
                            Write-Log "警告: 既存PDFの自動復元に失敗しました。バックアップを保持します: $backupPdf"
                        }
                    }
                    throw $publishPdfError
                }

                if (($null -ne $backupPdf) -and (Test-Path -LiteralPath $backupPdf)) {
                    try {
                        Remove-Item -LiteralPath $backupPdf -Force
                        $backupPdf = $null
                    }
                    catch {
                        Write-Log "警告: 置換前PDFのバックアップを削除できませんでした: $backupPdf"
                    }
                }
            }
            finally {
                if (Test-Path -LiteralPath $temporaryPdf) {
                    Remove-Item -LiteralPath $temporaryPdf -Force -ErrorAction SilentlyContinue
                }
                if (($null -ne $backupPdf) -and (Test-Path -LiteralPath $backupPdf)) {
                    Write-Log "置換前PDFのバックアップを保持しています: $backupPdf"
                }
            }

            Write-Log "PDF保存: $destinationPdf"
        }
    }

    $Completed = $true
    $ExitCode = 0
    Write-Log "完了: $DestinationOutput"
    Write-Host ("完了: {0}" -f $DestinationOutput)
    if ($ExportPdf) {
        Write-Host ("PDF保存先: {0}" -f $PdfDestinationDirectory)
    }
}
catch {
    $errorRecord = $_
    $message = [string]$errorRecord.Exception.Message
    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = [string]$errorRecord
    }
    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = "詳細メッセージのない例外が発生しました。"
    }

    $exceptionType = $errorRecord.Exception.GetType().FullName
    $positionMessage = [string]$errorRecord.InvocationInfo.PositionMessage
    $fullyQualifiedErrorId = [string]$errorRecord.FullyQualifiedErrorId
    $scriptStackTrace = [string]$errorRecord.ScriptStackTrace

    if ($null -ne $LogPath) {
        Write-Log "失敗した段階: $CurrentStage"
        Write-Log "失敗: $message"
        Write-Log "例外型: $exceptionType"
        if (-not [string]::IsNullOrWhiteSpace($fullyQualifiedErrorId)) {
            Write-Log "エラーID: $fullyQualifiedErrorId"
        }
        if (-not [string]::IsNullOrWhiteSpace($positionMessage)) {
            Write-Log ("発生位置: " + ($positionMessage -replace "[\r\n]+", " "))
        }
        if (-not [string]::IsNullOrWhiteSpace($scriptStackTrace)) {
            Write-Log ("スタック: " + ($scriptStackTrace -replace "[\r\n]+", " | "))
        }
        Write-Log "ローカルジョブは終了処理で削除します: $JobDirectory"
    }

    [Console]::Error.WriteLine("失敗した段階: {0}" -f $CurrentStage)
    [Console]::Error.WriteLine("エラー: {0}" -f $message)
    $ExitCode = 1
}
finally {
    if (($null -ne $JobDirectory) -and
        (Test-Path -LiteralPath $JobDirectory -PathType Container)) {
        if ($null -ne $LogPath) {
            Write-Log "ローカル一時ジョブを削除します。"
        }

        # process.logもジョブ内にあるため、削除開始後はファイルログへ書き込まない。
        $LogPath = $null
        try {
            Remove-Item -LiteralPath $JobDirectory -Recurse -Force -ErrorAction Stop
            Write-Host ("ローカル一時ファイルを削除しました: {0}" -f $JobDirectory)
        }
        catch {
            [Console]::Error.WriteLine("警告: ローカル一時ファイルの削除に失敗しました: {0}" -f $JobDirectory)
            [Console]::Error.WriteLine("手動で削除してください: {0}" -f $_.Exception.Message)
            $ExitCode = 1
        }
    }
}

exit $ExitCode
