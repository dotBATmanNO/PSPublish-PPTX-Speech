# PSPublish-PPTX-Speech
Publish your PowerPoint to video using Text-To-Speech and Subtitle generation.

Target Users are: 
- Conferences
- Content Publishers.

Top use case:
- Low-cost narrated video
- Accessibility
- Translations; people in different countries can 
  * translate the presentation to their own language
  * run Localized Windows text-to-speech and/or publish subtitles

Tip: Use Cortana / other professional text-to-speech engines where possible!

Example use:
1. Publish the original presentation with recorded voice as presented at a conference
   or, if no recording is available:  Turn any presentation into a video with audio narration.
2. Transcribe the presentation to be able to publish subtitles that help with indexing and provide accessibility
3. Translate the transcribed text to other language(s), publish alternate subtitles for viewers in other countries
4. Republish the original PowerPoint with Text-To-Speech of the translated text in other language.

Combinations could be:
- English Audio, English Text (accessibility)
- English Audio, Translated Text
- Translated Audio (computer text-to-speech), English Text
- Translated Audio (computer text-to-speech), Translated Text (multiple languages)

This could help overcome reservations against using your own voice / accent.

```
PS C:\scripts\PSPublish-PPTX-Speech> Get-Help .\Publish-PPTX-Speech.ps1 -Full

NAME
    C:\scripts\PSPublish-PPTX-Speech\Publish-PPTX-Speech.ps1

SYNOPSIS
    This script automates the running of a slide show,
    including text-to-speech and (pending) generation of .SRT subtitle files.


SYNTAX
    C:\scripts\PSPublish-PPTX-Speech\Publish-PPTX-Speech.ps1 [[-Path] <String>] [-SaveFile <Object>] [-SaveF
    ileGender <Object>] [-SaveVideo <Object>] [<CommonParameters>]


DESCRIPTION
    Update the Publish-PPTX-Speech.XLSX worksheet to generate .CSV file.
    Alternatively update the .CSV file directly if you do not have Excel available.

    The script will:
    - Run slideshow based on input from Publish-PPTX-Speech.csv
    - Build one .wav file per slide in Record subfolder.

    If you are running with the default -SaveFile 1
    - Create and set timing to Advance Slide
    - Inserts WAV files (linked) on each slide
      (WAV files are set to autoplay and to hide during show)
    - Save Publish-PPTX-Speech_VoiceOver.pptx
      (Set to run full screen and loop)


PARAMETERS
    -Path <String>
        The full path of PPTX source file.
        Note: The corresponding .CSV file must exist as well.

        Required?                    false
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -SaveFile <Object>
        By default the script will generate Slidexxx.wav files in Record Folder
        Use -SaveFile 0 if you want to run the slideshow on read-only media.

        Required?                    false
        Position?                    named
        Default value                True
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -SaveFileGender <Object>
        Only one gender is supported for Slidexxx.wav files.
        Pro-tip: Purchase a professional voice or enable Cortana for Speech-to-Text.

        Required?                    false
        Position?                    named
        Default value                Female
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -SaveVideo <Object>
        If you want the script can directly generate .MP4 version of the file

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>.\Publish-PPTX-Speech.ps1 Publish-PPTX-Speech.pptx

    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>.\Publish-PPTX-Speech.ps1 Publish-PPTX-Speech.pptx -SaveFile 0

    e.g. to Run on Read-Only media, will not create WAV files / edit slides.
    (better use is to generate VoiceOver PPTX and run this file)

    -------------------------- EXAMPLE 3 --------------------------

    PS C:\>.\Publish-PPTX-Speech.ps1 D:\Publish-PPTX-Speech.pptx -SaveVideo 1

    Slide 1 was shown for 8 seconds.
    Slide 2 was shown for 5.3 seconds.
    Slide 3 was shown for 11.5 seconds.
    Slide 4 was shown for 16 seconds.
    Slide 5 was shown for 5 seconds.
    Saving PPTX file with voice-over as 'D:\Publish-PPTX-Speech_VoiceOver.pptx'.
    Save completed.
    Media Task Status: Starting output to 'D:\Publish-PPTX-Speech.mp4'.
    Media Task Status: Done

RELATED LINKS
    https://github.com/dotBATmanNO/PSPublish-PPTX-Speech/

```
