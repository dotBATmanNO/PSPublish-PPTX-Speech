<#
    .SYNOPSIS
    This script automates the running of a slide show, 
    including text-to-speech and (pending) generation of .SRT subtitle files.
        
    .DESCRIPTION
	Update the Publish-PPTX.Speech.XLSX worksheet to generate .CSV file.
    Alternatively update the .CSV file directly if you do not have Excel available.
  
    .LINK
    https://github.com/dotBATmanNO/PSPublish-PPTX-Speech/
#>

[CmdletBinding(PositionalBinding=$false)]

 param (
    # The full path of PPTX source file.
    # Note: The corresponding .CSV file must exist as well.
    [Parameter(Position=0)][string]$Path,
    # By default the script will generate Slidexxx.wav files in Record Folder
    # Use -SaveFile 0 if you want to run the slideshow on read-only media.
    $SaveFile = $true,
    # Only one gender is supported for Slidexxx.wav files.
    # Pro-tip: Purchase a professional voice or enable Cortana for Speech-to-Text.
    $SaveFileGender = "Female")

Function Out-Speech 
{
    # Based on Out-Speech created by Guido Oliveira
   	
	[CmdletBinding()]
	param(
	  [String[]]$Message="Test Message.",
	  [String[]]$Gender="Female")
	begin 
    {
	  try   { Add-Type -Assembly System.Speech -ErrorAction Stop }
	  catch { Write-Error -Message "Error loading the required assemblies" }
	}
	process
    {

     $voice = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop
            
     $voice.SelectVoiceByHints($Gender)
            			
     $voice.Speak($message) | Out-Null
				
	}
	end
    {
		
	}
}

Function fnSaveWAVFile
{
    # Based on Out-Speech created by Guido Oliveira
   	
    [CmdletBinding()]
	param(
	  [String[]]$Message="Test Message.",
	  [String]$WAVFileName)
    
    try   { Add-Type -Assembly System.Speech -ErrorAction Stop }
    catch { Write-Error -Message "Error loading the required assemblies" }

    $voicesave = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop
    
    $voicesave.SelectVoiceByHints($SaveFileGender)
    $Voicesave.SetOutputToWaveFile($WAVFileOut)
    ForEach ($strmsg in $Message) { $Voicesave.Speak($strmsg) }
    $voicesave.Dispose()

} # End function fnSaveWAVFile

# Start of main script. Check input first.
If ( $Path -eq "" -or (Test-Path $Path) -eq $False ) 
{
  Write-Host "The PowerPoint file '$($Path)' was not found!"
}
else
{

    $PPTXPath = Split-Path -parent $Path                                  # Retrieve PATH of PPTX file
    $PPTXFileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)  # Retrieve filename of PPTX file (no extension)
    $SlideCSVFile = "$($PPTXPath)\$($PPTXFileName).CSV"                   # Build variable holding name of CSV file
    If ( (Test-Path $SlideCSVFile) -eq $False ) 
    {
      Write-Host "The PowerPoint file needs to be supported by a CSV file named '$($SlideCSVFile)!"
      Write-Host "Generate your .CSV file using the Publish-PPTX-Speech.XLSX file."
      Write-Host ".. or copy and edit Publish-PPTX-Speech.csv if you do not have Excel."
      Break
    }
    
    If ($SaveFile) # User wants to save Slidexxx.wav files
    {
        $RecordPath = Join-Path -Path $PPTXPath -ChildPath "Record"           # Check for and prepare folder for .WAV files
        If ( (Test-Path $RecordPath) -eq $False ) 
        {
            New-Item -Path $PPTXPath -Name "Record" -ItemType "Directory" | Out-Null
        }
        
    }
    
    $ppAdvanceOnClick = 0
    $ppSlideShowRehearseNewTimings = 3
    $ppSlideShowPointerAlwaysHidden = 3

    # $ppShowTypeKiosk = 3
    # $ppSlideShowDone = 5

    $objPPT = New-Object -ComObject "PowerPoint.Application"
    $objPPT.Visible = 1

    Try
    {
        # Using SaveFile parameter to decide if new Slide Advance timings shall be saved.
        $objPresentation = $objPPT.Presentations.Open($Path, !$SaveFile,$false,$false)
    }
    Catch 
    {
        write-host "File '$($path)' failed to open in PowerPoint."
        Break   
    }

    # Disable automatic transition, we want to use transitions from PPTXFileName.csv
    $objPresentation.SlideShowSettings.AdvanceMode = $ppAdvanceOnClick

    # $objPresentation.SlideShowSettings.ShowType = $ppShowTypeKiosk
    $objPresentation.SlideShowSettings.ShowType = $ppSlideShowRehearseNewTimings

    $objPresentation.SlideShowSettings.StartingSlide = 1

    $objPresentation.SlideShowSettings.EndingSlide = $objPresentation.Slides.Count

    $objSlideShow = $objPresentation.SlideShowSettings.Run().View
    Start-Sleep -Seconds 2

    $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden

    #TryCSVBlock
    Try 
    {
        $arrSteps = Import-Csv -Path $SlideCSVfile -header "SlideNumber", "Duration", "Click", "Gender", "Say"
    }
    Catch
    {
        Write-Host "Unable to read slide transitions from CSV file '$($SlideCSVFile)'."
        Break   
    } # End TryCSVBlock

    # Enumerate and run through all slides
    For ($CurSlide=1; $CurSlide -le $objPresentation.Slides.Count; $CurSlide++)
    {
        
        $intSlideDuration = 0
        # Build search pattern; 001 to 999 supported, any more than 999 slides would be #unthinkable
        $strCurSlide = $CurSlide.ToString().PadLeft(3, "0")
        $strPattern = "Slide$StrCurSlide"
        
        $SlideNotes = $arrSteps.Where{ $_.SlideNumber -eq $strPattern }
        If ($SlideNotes.Count -gt 0) # Handle slides that do not have text-to-speech.
        {
            $SlideCurClick = 0
            
            $strSlideMessage = ""

            ForEach ($Transition in $SlideNotes)
            {
                
                $strMessage = $Transition.Say
                $strGender = $Transition.Gender
                Write-Verbose "$strMessage by $strGender"
                # Build a variable holding all text-to-speech for this slide
                # This can be used to create one .wav file per slide
                $strSlideMessage += $strMessage

                $speaktime = Measure-Command { Out-Speech -Message $strMessage -Gender $strGender }
                Write-Verbose $speaktime
                $timetosleep = $Transition.Duration
                $intSlideDuration = $intSlideDuration + [math]::Round($speaktime.TotalSeconds,1) + $timetosleep

                If ($Transition.Duration -gt 0) 
                {
                    Write-Verbose "Sleeping $timetosleep"
                    Start-sleep -Seconds $timetosleep
                }
                
                If ($Transition.Click -like "True" -and $objSlideShow.GetClickCount() -ge $SlideCurClick)
                {
                    $SlideCurClick++
                    # Click for next animation / slides
                    $objSlideShow.GotoClick($SlideCurClick)
                    # Setting PointerType forces screen refresh
                    $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden
                }
            } # All transitions for slide have been processed

            If ($SaveFile) # User wants to save Slidexxx.wav files
            {
                $WAVFileOut = Join-Path -Path $RecordPath -ChildPath "Slide$($strCurSlide).wav"
                fnSaveWAVFile $strSlideMessage -WAVFileName $WAVFileOut
            }
         
        }
        else
        {
            # Nothing to say, let's just show the slide for 5 seconds
            $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden
            $intSlideDuration = 5.0
            Start-Sleep -Seconds $intSlideDuration
        }
        
        # Added errorhandling to first object access - in testing some instances have crashed.
        # Add auto-restart to your own error handling.

        If ($SaveFile) 
        {
          Try   { $objPresentation.Slides($CurSlide).SlideShowTransition.AdvanceTime = $intSlideDuration }
          Catch { Write-Host "An error occurred, did you close PowerPoint - or did it crash?"; Write-Host "Please Retry!"; Break }
        }
         
        If ($CurSlide -lt $objPresentation.Slides.Count)
        { 
            $objPresentation.SlideShowWindow.View.GotoSlide($CurSlide + 1)
        }
        Write-Host "Slide $CurSlide was shown for $intSlideDuration seconds."
        
    } # Reached the end of slide-show

    If ($SaveFile) 
    {
      Try
      {
         $objPresentation.Save() 
         $objPresentation.Close()
         $objPPT.Quit()
      }
      Catch 
      {
         Write-Host "An error occurred on Save."
         Write-Host "- Read-only media?"
         Write-Host "- Did you close PowerPoint?"
         Write-Host "- Did PowerPoint crash?"
         Write-Host "Please Retry!"
      }
    
    }
    
} 
