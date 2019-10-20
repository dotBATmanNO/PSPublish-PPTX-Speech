<#
   .SYNOPSIS
 		This script automates the running of a slide show, 
        including text-to-speech.
        
	.DESCRIPTION
		Update the Publish-PPTX.Speech.XLSX worksheet, 
        or update the .CSV file if you do not have Excel available.
  #========================================================================
#>

Function Out-Speech 
{
 <#

  #========================================================================
  # Created on:    2/18/2014 2:42 AM
  # Created by:    Guido Oliveira
  # Function name: Out-Speech
  #========================================================================
	.SYNOPSIS
		This is a Text to Speech Function made in powershell.

	.DESCRIPTION
		This is a Text to Speech Function made in powershell.

	.PARAMETER  Message
		Type in the message you want to hear.

	.PARAMETER  Gender
		The description of a the ParameterB parameter.

	.EXAMPLE
		PS C:\> Out-Speech -Message "Testing the function" -Gender 'Female'
		
	.EXAMPLE
		PS C:\> "Testing the function 1","Testing the function 2","Testing the function 3","Testing the function 4","Testing the function 5","Testing the function 6" | Foreach-Object { Out-Speech -Message $_ }
	
	.EXAMPLE
		PS C:\> "Testing the PS-Pipeline" | Out-Speech

	.INPUTS
		System.String

	#>
	[CmdletBinding()]
	param(
	  [Parameter(Mandatory=$true)]
	  [String[]]
	  $Message="Test Message.",
	  [Parameter(Mandatory=$true)]
      [String[]]
	  $Gender="Female")
	begin 
    {
	  try
      {
	   Add-Type -Assembly System.Speech -ErrorAction Stop
	  }
	  catch 
      {
       Write-Error -Message "Error loading the required assemblies"
      }
	}
	process
    {
     $voice = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop
            
     #Write-Verbose "Selecting a $Gender voice"
     $voice.SelectVoiceByHints($Gender)
            			
     #Write-Verbose -Message "Start Speaking"
     $voice.Speak($message) | Out-Null
				
	}
	end
    {
		
	}
}

$ScriptPath = Split-Path -parent $PSCommandPath
$RecordPath = Join-Path -Path $ScriptPath -ChildPath "Record"
If ( (Test-Path $RecordPath) -eq $False ) 
{
 New-Item -Path $ScriptPath -Name "Record" -ItemType "Directory"
}

If ( (Test-Path $PPTXFile) -eq $False ) 
{
 Write-Host "The PPTX file $($PPTXFile) was not found!"
}
else
{

    $ScriptPath = Split-Path -parent $PSCommandPath
    
    $SaveFile = $true
    $SaveFileGender = "Female"

    $ppAdvanceOnClick = 0
    $ppSlideShowRehearseNewTimings = 3
    $ppSlideShowPointerAlwaysHidden = 3

    $ppShowTypeKiosk = 3
    $ppSlideShowDone = 5

    $objPPT = New-Object -ComObject "PowerPoint.Application"
    $objPPT.Visible = 1

    $objPresentation = $objPPT.Presentations.Open($PPTXFile, $true,$false,$false)

    # Disable automatic transition, we want to use transition in NOTES section of slide
    #
    # Slidexxx,TransitionDuration,ClickForNext (Boolean True/False),Gender of voice,Text to say
    # Example:
    # Slide001,7,True,"Male","Risk Management.  A real-world scenario for Car Owners."
    # |- Slide 001   = Must match the slide it is on, this is a
    #                  key to process the line as text-to-speech
    # |- 1           = Number of seconds to pause after text-to-speech completes
    # |- True        = After the transition has completed, 
    #                  automatically click the next animation
    # |- Male        = Use the male text-to-speech engine
    # |- Text To Say = Self-explanatory... Note Language! :)

    $objPresentation.SlideShowSettings.AdvanceMode = $ppAdvanceOnClick

    # $objPresentation.SlideShowSettings.ShowType = $ppShowTypeKiosk
    $objPresentation.SlideShowSettings.ShowType = $ppSlideShowRehearseNewTimings

    $objPresentation.SlideShowSettings.StartingSlide = 1

    $objPresentation.SlideShowSettings.EndingSlide = $objPresentation.Slides.Count


    $objSlideShow = $objPresentation.SlideShowSettings.Run().View
    Start-Sleep -Seconds 2
    $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden

    $arrSteps = Input-Csv -Path $ScriptPath\ -header "SlideNumber", "Duration", "Click", "Gender", "Say"
 
    # Enumerate and run through all slides
    For ($CurSlide=1; $CurSlide -le $objPresentation.Slides.Count; $CurSlide++)
    {
 
     $intSlideDuration = 0
     # Build search pattern; 001 to 999 supported, any more than 999 slides would be #unthinkable
     $strCurSlide = $CurSlide.ToString().PadLeft(3, "0")
     $strPattern = "Slide$StrCurSlide"
 
     # Connect to the Slide Comments text box
     $SlideNotes = $objPresentation.Slides($CurSlide).NotesPage.Shapes.Placeholders(2).TextFrame
 
     # If there are Slide Comments, let's find those that need to be converted to Speech
     If ($SlideNotes.HasText)
     {
      $SlideCurClick = 0
      $NotesLine = 1
      $strSlideMessage = ""
  
      While ($NotesLine -le $SlideNotes.TextRange.Paragraphs().Count)
      {
       $NotesCurLine = $SlideNotes.TextRange.Paragraphs($NotesLine,1).Text
       Write-Verbose $NotesCurLine
       If ($NotesCurLine.Length -gt 8)
       {
        If ($NotesCurLine.SubString(0,8) -eq $strPattern)
        {
    
            $Transition = ConvertFrom-Csv -InputObject $NotesCurLine -header "SlideNumber", "Duration", "Click", "Gender", "Say"
    
            $strMessage = $Transition.Say
            $strGender = $Transition.Gender
            Write-Verbose "$strMessage by $strGender"
    
            # Build a variable holding all text-to-speech for this slide
            $strSlideMessage += $strMessage
        
            $speaktime = measure-command { Out-Speech -Message $strMessage -Gender $strGender }
            Write-Verbose $speaktime
            $timetosleep = ($Transition.Duration - $speaktime.Seconds)
            $intSlideDuration = $intSlideDuration + $speaktime + $timetosleep

            If ($timetosleep -gt 0 -and $Transition.Duration -gt 0) 
            {
             Write-Verbose "Sleeping $timetosleep"
             Start-sleep -Seconds $timetosleep
            }
            Else
            {
             If ($Transition.Duration -gt 0) {Write-Host "Transition $Transition takes too long to say ($speaktime seconds)!" -ForegroundColor Red }
            }
            If ($Transition.Click -like "True" -and $objSlideShow.GetClickCount() -ge $SlideCurClick)
            {
             $SlideCurClick++
             # Click for next animation / slides
             $objSlideShow.GotoClick($SlideCurClick)
             $objSlideShow.PointerType = $ppSlideShowPointerAlwaysHidden
            }
        } 
       }
       $NotesLine++
      }
 
      If ($SaveFile)
      {
       # Build
       try
       {
        Add-Type -Assembly System.Speech -ErrorAction Stop
        $voicesave = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop
       }
       catch 
       {
        Write-Error -Message "Error loading the requered assemblies"
       }

       $WAVFileOut = Join-Path -Path $ScriptPath -ChildPath "Record\Slide$($strCurSlide).wav"
       $voicesave.SelectVoiceByHints($SaveFileGender)
       $Voicesave.SetOutputToWaveFile($WAVFileOut)
       ForEach ($strmsg in $strSlideMessage) { $Voicesave.Speak($strmsg)}
       $voicesave.Dispose()
      }
  
      If ($CurSlide -lt $objPresentation.Slides.Count)
      { 
       $objPresentation.SlideShowWindow.View.GotoSlide($CurSlide + 1)
      }
 
      #Start-Sleep $objSlideShow.Slide.SlideshowTransition.AdvanceTime
      Write-Host "Slide $CurSlide took $intSlideDuration seconds."
     }
 
    }

    $objPresentation.Close()
    $objPPT.Quit()
}
