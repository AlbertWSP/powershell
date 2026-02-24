# Please note that this is quite old at this point and may or may not work for you. I leave it here for historical purposes
# and just in case it can help someone. I think I may have been using Windows 7 when I wrote this.

function Get-Temperature {
    $t = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"

    $currentTempKelvin = $t.CurrentTemperature / 10
    $currentTempCelsius = $currentTempKelvin - 273.15

    $currentTempFahrenheit = (9/5) * $currentTempCelsius + 32

    return $currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K"  
}

# Save in your c:\users\yourName\Documents\WindowsPowerShell\modules\ directory
# in sub directory get-temperature as get-temperature.psm1
# You **must** run as Administrator.
# It will only work if your system & BIOS support it. If it doesn't work, I can't help you.

# Just type get-temperature in PowerShell and it will spit back the temp in Celsius, Farenheit and Kelvin.

# More info on my blog: http://ammonsonline.com/is-it-hot-in-here-or-is-it-just-my-cpu/