Powershell script for use with NG-Spotify Importer(not mine) to load your mp3 folders on to spotify. https://nickwanders.com/projects/ng-spotify-importer/


This script scrapes title and artist data and outputs a csv with "title", "artist" as headers.
Also performs some processing required to work with the buggy but awesome and free NG-Spotify Importer.
Defaults to current folder if no folder is entered.

 - Strips special characters
 - skips files with a blank artist or title field
 - skips files with a very long title or artist field
 - skips non-music files


To run you may need to rightclick the file and select "Run with Powershell"
