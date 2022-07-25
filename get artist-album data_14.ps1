<#
This script scrapes title and artist data and outputs a csv with "title", "artist" as headers.
Also performs some processing required to work with the buggy but awesome and free NG-Spotify Importer https://nickwanders.com/projects/ng-spotify-importer/
#>


$current_dir = Get-Location
$music_formats = @('.mp3', '.aac', '.3gp', '.m4a', '.m4p', '.ogg', '.oga', '.wav', '.wma', 'webm', 'aiff', 'alac', 'flac', 'opus')

$user_confirm_proceed = $false

While($user_confirm_proceed -eq $false){
$output_file_path = Read-Host -Prompt "`nEnter filename file path to output CSV file, (will overwrite any existing file), or leave blank for current directory'n(eg C:\music\output.csv "
if ($output_file_path -eq ""){
Write-Host "`nBlank dir, using current directory`n"
$output_file_path = [string]$current_dir + "\\output.csv"
}

$target_mp3_folder = Read-Host -Prompt "Enter folder path to scan MP3 files, or leave blank for current directory: "
if ($target_mp3_folder -eq ""){
Write-Host "`nBlank dir, using current directory`n"
$target_mp3_folder = $current_dir
}


#clear content if target file already exists and add headers
Set-Content -Path $output_file_path -Value "title,artist"

#check with user if input parameters are correct
#$full_user_input = "Reading mp3's from: " + $target_mp3_folder + "`n Outputting CSV to $target_mp3_folder"
$user_menu_choice = Read-Host -Prompt "`n`nReading mp3's from:`t $target_mp3_folder `nOutputting CSV to:`t $output_file_path `n`n Is this correct? [y]/n/q"

Switch ($user_menu_choice){
    "y"{$user_confirm_proceed = $true}
    "n"{Write-Host "`nrestarting`n"}
    "q"{exit}
    "" {$user_confirm_proceed = $true}
    DEFAULT {Write-Host "`nrestarting`n"}
}

}

Read-Host -Prompt "Press enter to proceed with scan..."


#get list of file objects from target directory, pipe to loop
Get-ChildItem -path $target_mp3_folder |

    Foreach-Object {

    $FilePath = $_.FullName
    $Folder = Split-Path -Parent -Path $FilePath
    $File = Split-Path -Leaf -Path $FilePath
    $Shell = New-Object -COMObject Shell.Application
    $ShellFolder = $Shell.NameSpace($Folder)
    $ShellFile = $ShellFolder.ParseName($File)
    $current_file_extension = $FilePath.Substring($FilePath.Length-4)

    $song_title = $ShellFile.ExtendedProperty("System.Title")
    $song_artist = $ShellFile.ExtendedProperty("System.Music.Artist")

    #remove anything NOT a letter, NOT a number, and NOT a space. Aka strip all special characters.
    $song_title = $song_title -replace "[^a-zA-Z0-9 ]",""
    $song_artist = $song_artist -replace "[^a-zA-Z0-9 ]",""

    $csv_line = $song_title + "," + $song_artist

    if(-not $music_formats.Contains($current_file_extension)){
    Write-host "not a music file, skipped: " $File
    
    }elseif (-not $song_artist){
     Write-Host "blank artist skipped: " $File
    
    }elseif (-not $song_title) {
    Write-Host "blank title skipped: " $File
    
    }elseif ($csv_line.Length -gt 75){
    
    Write-Host "title/artist too long: skipped" $File

    }else {
    
    Add-Content -Path $output_file_path -Value $csv_line
    Write-Host "pass"
    
    }

    }


Read-Host -Prompt "`nOperation completed, press enter to close"