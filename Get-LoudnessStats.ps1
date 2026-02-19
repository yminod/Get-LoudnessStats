#Requires -Version 7.4

function Get-LoudnessStats {
<#
.SYNOPSIS
Collects loudness and level statistics (LUFS / LRA / Peak / RMS, etc.) from audio files using FFmpeg.

.DESCRIPTION
This function invokes FFmpeg audio filters and parses their textual output to extract
a small set of commonly used loudness / level metrics.

It uses:
- astats: overall Peak level / RMS level / Noise floor
- ebur128: Integrated Loudness (I) / Loudness Range (LRA) / True Peak (Peak)

Input is file paths (wildcards supported). Multiple paths and pipeline input are supported.
By default it runs in parallel; use -Serial to process sequentially.

This function requires ffmpeg to be available on PATH.

.PARAMETER Path
One or more input file paths (wildcards supported).
Accepts pipeline input.

.PARAMETER AstatsWindowSize
The astats "length" in seconds (default: 1.0).
This value is passed to ffmpeg as part of the astats filter arguments.

.PARAMETER Serial
Process files sequentially (disable parallel processing).

.PARAMETER ThrottleLimit
Maximum number of concurrent ffmpeg processes when running in parallel (default: 5).
Ignored when -Serial is specified.

.PARAMETER InitialFileCount
Initial capacity for the internal file list (default: 128).
You can increase this when you expect a large number of input files to reduce list resizing.

.INPUTS
System.String
You can pipe file paths (strings) into this function.

.OUTPUTS
System.Management.Automation.PSCustomObject

The output object contains the following properties:

- Name (string)
  The input file name (leaf).

- Peak (double)
  "Peak level dB" reported by astats (typically interpreted as dBFS).

- RMS (double)
  "RMS level dB" reported by astats (typically interpreted as dBFS).

- NoiseFloor (double|string)
  "Noise floor dB" reported by astats.
  If FFmpeg reports -inf, this function preserves the literal string "-inf".
  Note: Excel (and many CSV-oriented tools) cannot represent negative infinity as a numeric value.
  Treat "-inf" as a special value meaning "below numeric range" and handle it explicitly if you need calculations.

- TruePeak (double)
  "Peak dBFS" reported by ebur128 (treated as true peak, dBFS).

- IntegratedLoudness (double)
  "I" (Integrated Loudness, LUFS) from ebur128.

- LoudnessRange (double)
  "LRA" (Loudness Range, LU) from ebur128.

- LRALow (double)
  "LRA low" (LUFS) from ebur128.

- LRAHigh (double)
  "LRA high" (LUFS) from ebur128.

.EXAMPLE
Get-LoudnessStats -Path .\*.wav -Serial

Collect statistics for all WAV files sequentially (e.g., when you want to avoid multiple concurrent ffmpeg instances).

.EXAMPLE
Get-ChildItem *.wav | Get-LoudnessStats -ThrottleLimit 12

Enumerate all WAV files and analyze them via pipeline input increasing parallelism level.

.NOTES
Requirements:
- PowerShell 7.4+
- ffmpeg available on PATH

.LINK
https://ffmpeg.org/ffmpeg-filters.html#astats
https://ffmpeg.org/ffmpeg-filters.html#ebur128
#>

    [CmdletBinding()]
    param (
        [Parameter(
             Mandatory,
             ValueFromPipeline,
             ValueFromPipelineByPropertyName)]
        [SupportsWildcards()]
        [string[]]$Path,

        [double]$AstatsWindowSize = 1.0,

        [switch]$Serial,
        [int]$ThrottleLimit = 5,

        [int]$InitialFileCount = 128
    )
    begin {
        try { $null = Get-Command ffmpeg -ErrorAction Stop }
        catch { throw "ffmpeg not found in PATH." }

        # Defaults to $null if the variable is missing.
        $avLogForceNoColor = $Env:AV_LOG_FORCE_NOCOLOR

        $files = [Collections.Generic.List[string]]::new($InitialFileCount)
    }
    process {
        foreach ($p in $Path) {
            foreach ($r in Resolve-Path -Path $p -ErrorAction Stop) {
                $files.Add($r.Path)
            }
        }
    }
    end {
        $sbTemplate = @'
$source = $_
$lines = & ffmpeg `
  -hide_banner `
  -nostats `
  -i $source `
  -af ('astats=' `
  + "length={0}:" `
  + 'measure_perchannel=none:' `
  + 'measure_overall=Peak_level+RMS_level+Noise_floor,' `
  + 'ebur128=' `
  + 'peak=true:' `
  + 'dualmono=1:' `
  + 'framelog=quiet') `
  -f null - 2>&1

$peak = $null
$rms = $null
$noiseFloor = $null
$truePeak = $null
$loudness = $null
$loudnessRange = $null
$loudnessRangeLow = $null
$loudnessRangeHigh = $null
switch -CaseSensitive -Regex ($lines) {{
    # af:astats
    'Peak level dB: (-?[.0-9]+)$' {{
        $peak = [math]::Round([double]$Matches[1], 1)
    }}
    'RMS level dB: (-[.0-9]+)$' {{
        $rms = [math]::Round([double]$Matches[1], 1)
    }}
    'Noise floor dB: (-(?:[.0-9]+|inf))$' {{
        $value = $Matches[1]
        $noiseFloor = if ($value -match 'inf') {{
            $value
        }}
        else {{
            [math]::Round([double]$value, 1)
        }}
    }}
    # af:ebur128
    'I: +(-[.0-9]+) LUFS$' {{
        $loudness = [double]$Matches[1]
    }}
    'LRA: +([.0-9]+) LU$' {{
        $loudnessRange = [double]$Matches[1]
    }}
    'LRA low: +(-[.0-9]+) LUFS$' {{
        $loudnessRangeLow = [double]$Matches[1]
    }}
    'LRA high: +(-[.0-9]+) LUFS$' {{
        $loudnessRangeHigh = [double]$Matches[1]
    }}
    'Peak: +(-?[.0-9]+) dBFS$' {{
        $truePeak = [double]$Matches[1]
    }}
}}
[pscustomobject]@{{
    'Name' = Split-Path -Path $source -Leaf
    'Peak' = $peak
    'RMS' = $rms
    'NoiseFloor' = $noiseFloor
    'TruePeak' = $truePeak
    'IntegratedLoudness' = $loudness
    'LoudnessRange' = $loudnessRange
    'LRALow' = $loudnessRangeLow
    'LRAHigh' = $loudnessRangeHigh
}}
'@

        try {
            # Disable ANSI color codes for clean output.
            $Env:AV_LOG_FORCE_NOCOLOR = 1

            # Apply [Globalization.CultureInfo]::InvariantCulture.
            $WindowSizeStr = [string]$AstatsWindowSize

            if ($Serial) {
                $files | ForEach-Object `
                  -Process ([scriptblock]::Create(($sbTemplate -f '${WindowSizeStr}')))
            }
            else {
                $files | ForEach-Object -ThrottleLimit $ThrottleLimit `
                  -Parallel ([scriptblock]::Create(($sbTemplate -f '${using:WindowSizeStr}')))
            }
        }
        finally {
            if ($avLogForceNoColor -eq $null) {
                Remove-Item Env:AV_LOG_FORCE_NOCOLOR -ErrorAction SilentlyContinue
            }
            else {
                $Env:AV_LOG_FORCE_NOCOLOR = $avLogForceNoColor
            }
        }
    }
}
