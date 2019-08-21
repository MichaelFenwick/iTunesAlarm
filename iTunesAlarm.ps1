#Plays a given iTunes playlist starting at a random song. Starts silent and then increases in volume over time. Sets iTunes volume to 100% and increases Windows volume because iTunes is annoyingly non-linear in its volume levels.

# To use this script you must set execution policy to allow it. You can do so with the following command from an Admin Powershell window.
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

# This module is needed to be able to adjust system volume. You can find it at https://github.com/frgnca/AudioDeviceCmdlets
# Make sure you put the module in the same folder as this script.
# You'll get an error on import unless you view the file's properties in Explorer and then click the "Unblock" button.
Import-Module ".\AudioDeviceCmdlets.dll"

### Open iTunes if it's not already
Start-Process "C:\Program Files\iTunes\iTunes.exe"

### Set computer volume to 0%. We'll increase this over time.
Set-AudioDevice -PlaybackVolume 0

### Get variables related to iTunes
$itunes = New-Object -ComObject iTunes.Application
$playlists = $itunes.Sources.ItemByName("Library").Playlists
$playlist = $playlists.ItemByName("3.5+ Rock")

### Make sure iTunes is in the correct state and then start playing the playlist.
$itunes.SoundVolume = 100
$playlist.Shuffle = 1
$playlist.SongRepeat = 0
$playlist.PlayFirstTrack()

### Increase volume over the next three minutes or so.
for ($volume=0; $volume -le 100; $volume++) {
	Start-Sleep -s 4
#	$itunes.SoundVolume = $volume
Set-AudioDevice -PlaybackVolume $volume
}

### Don't play forever in case I'm not home.
Start-Sleep -s 1800
$itunes.Stop()

### TODO:
#Figure out how to make sure this script can be killed when I wake up and replaced with playing music at a comfortable volume.
#Maybe have it launch Airfoil if I can get Airfoil to auto-connect on launch.
#Start-Process "C:\Program Files (x86)\Airfoil\Airfoil.exe"
