<#
  .SYNOPSIS
  This script automates the running of a slide show, 
  including text-to-speech and (pending) generation of .SRT subtitle files.
        
  .DESCRIPTION
  Update the Publish-PPTX-Speech.XLSX worksheet to generate .CSV file.
  Alternatively update the .CSV file directly if you do not have Excel available.

  The script will:
  - Run slideshow based on input from Publish-PPTX-Speech.csv
  - Build one .wav file per slide in Record subfolder.
  
  If you are running with the default -SaveFile 1
  - Create and set timing to Advance Slide
  - Insert WAV files (linked) on each slide
    (WAV files are set to autoplay and to hide during show)
  - Save Publish-PPTX-Speech_VoiceOver.pptx 
    * Set to run using timings
    NOTE: If applicable you can "Setup Slidesho" to Run in Kiosk mode manually.
          (When trying to set this via script it corrupted the PPTX file.)

  .LINK
  https://github.com/dotBATmanNO/PSPublish-PPTX-Speech/

  .EXAMPLE
  .\Publish-PPTX-Speech.ps1 Publish-PPTX-Speech.pptx
  
  .EXAMPLE
  .\Publish-PPTX-Speech.ps1 Publish-PPTX-Speech.pptx -SaveFile 0
  e.g. to Run on Read-Only media, will not create WAV files / edit slides.
  (better use is to generate VoiceOver PPTX and run this file)
  
  .EXAMPLE
  .\Publish-PPTX-Speech.ps1 D:\Publish-PPTX-Speech.pptx -SaveVideo 1
  Slide 1 was shown for 8 seconds.
  Slide 2 was shown for 5.3 seconds.
  Slide 3 was shown for 11.5 seconds.
  Slide 4 was shown for 16 seconds.
  Slide 5 was shown for 5 seconds.
  Saving PPTX file with voice-over as 'D:\Publish-PPTX-Speech_VoiceOver.pptx'.
  Save completed.
  Media Task Status: Starting output to 'D:\Publish-PPTX-Speech.mp4'.
  Media Task Status: Done
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
    $SaveFileGender = "Female",
    # If you want the script can directly generate .MP4 version of the file
    $SaveVideo = $false)

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
    # Changed to use SSML to add pause between lines of text.   	
    [CmdletBinding()]
  	param(
	    [String[]]$Message="Test Message.",
	    [String]$WAVFileName)
    
    try   { Add-Type -Assembly System.Speech -ErrorAction Stop }
    catch { Write-Error -Message "Error loading the required assemblies" }

    $voicesave = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop
    $voicestart = "<?xml version=""1.0""?>
    <speak version=""1.0"" xmlns=""http://www.w3.org/2001/10/synthesis""
             xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
             xsi:schemaLocation=""http://www.w3.org/2001/10/synthesis
                       http://www.w3.org/TR/speech-synthesis/synthesis.xsd""
             xml:lang=""en-US"">"
       
    $voicesave.SelectVoiceByHints($SaveFileGender)
    $Voicesave.SetOutputToWaveFile($WAVFileOut)
    ForEach ($strmsg in $Message)
    {
      $VoiceStart += "$($strmsg)<break />"
    }
    
    $VoiceSave.SpeakSsml("$($voicestart)</speak>") 
    $Voicesave.SetOutputToNull()
    $voicesave.Dispose()

} # End function fnSaveWAVFile

# Start of main script. Check input first.
If ( $Path -eq "")
{
  Write-Host "Please run Get-Help $($MyInvocation.MyCommand) for information on how to use this script."
  Break
}

If (($False -eq (Test-Path -Path $Path)) -or ($path.Substring($path.length -5, 5) -ne ".pptx")) 
{
  Write-Host "The script expects the name of a PowerPoint file."
  Write-Host "The file '$($Path)' was not found or is not a .pptx file!"
  Write-Host "Please run Get-Help .\$($MyInvocation.MyCommand) for information on how to use this script."
}
else
{

    $PPTXFileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)                          # Retrieve filename of PPTX file (no extension)
    
    $PPTXPath = Split-Path -Parent $Path                                                          # Retrieve PATH of PPTX file
    If ($PPTXPath -eq ".")
    {
      $Path = [System.IO.Path]::GetFullPath($PSCommandPath)                                       # Use full path to script
      $PPTXPath = Split-Path -Parent $Path                
    }
    
    $PPTXFileOpen = Join-Path -Path $PPTXPath -ChildPath "$($PPTXFileName).pptx"
    $SlideCSVFile = Join-Path -Path $PPTXPath -ChildPath "$($PPTXFileName).csv"                   # Build variable holding name of CSV file
    If ( (Test-Path $SlideCSVFile) -eq $False ) 
    {
      Write-Host "The PowerPoint file needs to be supported by a CSV file named '$($SlideCSVFile)!"
      Write-Host "Generate your .CSV file using the Publish-PPTX-Speech.XLSX file."
      Write-Host ".. or copy and edit Publish-PPTX-Speech.csv if you do not have Excel."
      Break
    }
    
    # All prerequisites seem to be in place, generate variables for output.
    $PPTXVoiceOverFile = Join-Path -Path $PPTXPath -ChildPath "$($PPTXFileName)_VoiceOver.pptx"   # Build variable for saving PPTX with VoiceOver
    $SlideMP4File = Join-Path -Path $PPTXPath -ChildPath "$($PPTXFileName).mp4"                   # Build variable for saving PPTX as MP4 (Video)
    
    If ($SaveFile) # User wants to save Slidexxx.wav files
    {
        $RecordPath = Join-Path -Path $PPTXPath -ChildPath "Record"           # Check for and prepare folder for .WAV files
        If ( (Test-Path $RecordPath) -eq $False ) 
        {
            New-Item -Path $PPTXPath -Name "Record" -ItemType "Directory" | Out-Null
        }
        
    }
    
    # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211582(v=office.11)
    # https://docs.microsoft.com/en-us/office/vba/api/powerpoint(enumerations)
    $ppAdvanceOnClick = 0
    $ppSlideShowUseSlideTimings = 2
    $ppSlideShowPointerAlwaysHidden = 3
    # $ppShowTypeKiosk = 3  # Not using this as it corrupts the PPTX file!
    
    Add-type -AssemblyName office
    Add-Type -AssemblyName microsoft.office.interop.powerpoint
    $msoTrue  = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $objPPT = New-Object -ComObject "PowerPoint.Application"
    $objPPT.Visible = $msoTrue
    $pptFixedFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation 
    
    Try
    {
      $objPresentation = $objPPT.Presentations.Open($PPTXFileOpen, $msoFalse, $msoTrue, $msoTrue)
    }
    Catch 
    {
      write-host "File '$($PPTXFileOpen)' failed to open in PowerPoint."
      Break   
    }

    $objPresentation.SlideShowSettings.StartingSlide = 1
    $objPresentation.SlideShowSettings.EndingSlide = $objPresentation.Slides.Count

    # Ensure the script can control slideshow progress with "clicks"
    $objPresentation.SlideShowSettings.AdvanceMode = $ppAdvanceOnClick
    For ($CurSlide=1; $CurSlide -le $objPresentation.Slides.Count; $CurSlide++)
    {
       # Run through all slides and set the Advance Mode to wait for Click.
       $objPresentation.Slides($CurSlide).SlideShowTransition.AdvanceOnClick = $msoTrue
    }

    $objSlideShow = $objPresentation.SlideShowSettings.Run().View
    Start-Sleep -Seconds 2

    $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden

    Try 
    {
        $arrSteps = Import-Csv -Path $SlideCSVfile -header "SlideNumber", "Duration", "Click", "Gender", "Say"
    }
    Catch
    {
        Write-Host "Unable to read slide transitions from CSV file '$($SlideCSVFile)'."
        Break   
    } 

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
                $strSlideMessage += " $($strMessage) <break time=""1s""/>"

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
                
                # Create WAV file
                $WAVFileOut = Join-Path -Path $RecordPath -ChildPath "Slide$($strCurSlide).wav"
                fnSaveWAVFile $strSlideMessage -WAVFileName $WAVFileOut
                
                # Add WAV file to current slide
                Try
                { 
                  $oWAVFile = $objPresentation.Slides($CurSlide).Shapes.AddMediaObject2($WAVFileOut, $msoFalse, $msoTrue, 10, 10)
                }
                Catch { Write-Host "An error occurred, did you close PowerPoint - or did it crash?"; Write-Host "Please Retry!"; Break }
            
                $oWAVFile.AnimationSettings.PlaySettings.PlayOnEntry = $msoTrue
                $oWAVFile.AnimationSettings.PlaySettings.HideWhileNotPlaying = $msoTrue
                #$oWAVFile.AnimationSettings.AdvanceMode = $ppAdvanceOnTime
                $oWAVFile.AnimationSettings.AnimationOrder = 1
                
                Clear-Variable oWAVFile
            }
         
        }
        else
        {
            # Nothing to say, let's just show the slide for 5 seconds
            $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden
            
            $intSlideDuration = Measure-Command { Start-Sleep -Seconds 5 }
            $intSlideDuration = [math]::Round($intSlideDuration)

        }
        
        If ($SaveFile) 
        {
          Try
          { 
             # Set slide to proceed to next slide automatically after timing established on playback.
             # Round and add 1 to be sure..
             $intSlideDuration = [math]::Round($intSlideDuration+1)
             $objPresentation.Slides($CurSlide).SlideShowTransition.AdvanceTime = $intSlideDuration
          }
          Catch { Write-Host "An error occurred, did you close PowerPoint - or did it crash?"; Write-Host "Please Retry!"; Break }
        }
         
        If ($CurSlide -lt $objPresentation.Slides.Count)
        { 
            $objPresentation.SlideShowWindow.View.GotoSlide($CurSlide + 1)
        }
        Write-Host "Slide $CurSlide was shown for $intSlideDuration seconds."
        
    } # Repeat until the end of slide-show
    
    $objPresentation.SlideShowWindow.View.Exit() 
    $objSlideShow = $null

    If ($SaveFile) 
    {
      For ($CurSlide=1; $CurSlide -le $objPresentation.Slides.Count; $CurSlide++)
      {
         # Run through all slides and set the Advance Mode to use timings from this run-through.
         $objPresentation.Slides($CurSlide).SlideShowTransition.AdvanceOnClick = $msoFalse
         $objPresentation.Slides($CurSlide).SlideShowTransition.AdvanceOnTime = $msoTrue
      }

      # Set the slideshow to use timings and to run in looped Kiosk mode.
      $objPresentation.SlideShowSettings.AdvanceMode = $ppSlideShowUseSlideTimings
      
      # Bug with setting ShowType to Kiosk causes PPTX to be corrupted.
      # $objPresentation.SlideShowSettings.ShowType=3
      
      
      Try
      {
         
         If ($SaveFile)
         {
           Write-Host "Saving PPTX file with voice-over as '$($PPTXVoiceOverFile)'."
           $objPresentation.SaveCopyAs($PPTXVoiceOverFile, $pptFixedFormat)
           Write-Host "Save completed."
         }

         If ($SaveVideo)
         {
           Write-Host "Media Task Status: Starting output to '$($SlideMP4File)'."
           $objPresentation.CreateVideo($SlideMP4File, $True, 5, 1920, 15, 70)
           While ( $objPresentation.CreateVideoStatus -le 2 )
           {
              Start-Sleep -Seconds 3
           }
         
           Switch ($objPresentation.CreateVideoStatus)
           {
             3 { Write-Host "Media Task Status: Done"; Break   }
             4 { Write-Host "Media Task Status: Failed"; Break }
           }
         }
         
         $objPresentation.Close()
         $objPPT.Quit()
         $objPPT = $null
         [gc]::collect()
         [gc]::WaitForPendingFinalizers()
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