# Default variables
$latitude = 40.7128   # New York City Latitude
$longitude = -74.0060 # New York City Longitude
$temperatureFormat = "C" # Set to "C" for Celsius or "F" for Fahrenheit

# Function to fetch weather data
function Get-WeatherData {
    param (
        [double]$Latitude,
        [double]$Longitude
    )
   
    $apiUrl = "https://api.open-meteo.com/v1/forecast?latitude=$Latitude&longitude=$Longitude&current_weather=true&daily=temperature_2m_max,temperature_2m_min,wind_speed_10m_max,weathercode&hourly=temperature_2m,wind_speed_10m,weathercode&timezone=auto"
    try {
        return Invoke-RestMethod -Uri $apiUrl -Method Get
    } catch {
        return $null
    }
}

# Function to decode weather codes
function Decode-WeatherCode {
    param ([int]$weatherCode)
    switch ($weatherCode) {
        0 { return "Clear Sky" }
        1 { return "Mainly Clear" }
        2 { return "Partly Cloudy" }
        3 { return "Overcast" }
        45 { return "Fog" }
        48 { return "Depositing Rime Fog" }
        51 { return "Light Drizzle" }
        61 { return "Light Rain" }
        71 { return "Light Snowfall" }
        80 { return "Light Showers" }
        95 { return "Thunderstorm" }
        default { return "Unknown Weather" }
    }
}

# Convert Celsius to Fahrenheit
function ConvertTo-Fahrenheit {
    param ([double]$Celsius)
    return [math]::Round(($Celsius * 9 / 5) + 32, 1)
}

# Get day names for the next 7 days
function Get-DayNames {
    $currentDate = Get-Date
    for ($i = 0; $i -lt 7; $i++) {
        (Get-Date ($currentDate.AddDays($i)) -Format dddd)
    }
}

# Get hour names for the next 7 hours
function Get-HourNames {
    $currentDate = Get-Date
    for ($i = 0; $i -lt 7; $i++) {
        (Get-Date ($currentDate.AddHours($i)) -Format HH:mm)
    }
}

# Fetch weather data
$weatherData = Get-WeatherData -Latitude $latitude -Longitude $longitude

if ($weatherData) {
    # Current weather
    $currentTemp = if ($temperatureFormat -eq "F") {
        ConvertTo-Fahrenheit -Celsius $weatherData.current_weather.temperature
    } else {
        $weatherData.current_weather.temperature
    }
    $unit = if ($temperatureFormat -eq "F") { "°F" } else { "°C" }

    Write-Host "Location: Latitude $latitude, Longitude $longitude"
    Write-Host "Current Temperature: $currentTemp$unit"
    Write-Host "Wind Speed: $($weatherData.current_weather.windspeed) km/h"
    Write-Host "Condition: $(Decode-WeatherCode -weatherCode $weatherData.current_weather.weathercode)"

    # Hourly Forecast (Next 7 Hours)
    Write-Host "`nHourly Forecast (Next 7 Hours):"
    $hourNames = Get-HourNames
    for ($i = 0; $i -lt 7; $i++) {
        $hourTemp = if ($temperatureFormat -eq "F") {
            ConvertTo-Fahrenheit -Celsius $weatherData.hourly.temperature_2m[$i]
        } else {
            $weatherData.hourly.temperature_2m[$i]
        }
        Write-Host "$($hourNames[$i]): Temp: $hourTemp$unit, Wind Speed: $($weatherData.hourly.wind_speed_10m[$i]) km/h, Condition: $(Decode-WeatherCode -weatherCode $weatherData.hourly.weathercode[$i])"
    }

    # Daily Forecast (Next 7 Days)
    Write-Host "`nDaily Forecast (Next 7 Days):"
    $dayNames = Get-DayNames
    for ($i = 0; $i -lt 7; $i++) {
        $maxTemp = if ($temperatureFormat -eq "F") {
            ConvertTo-Fahrenheit -Celsius $weatherData.daily.temperature_2m_max[$i]
        } else {
            $weatherData.daily.temperature_2m_max[$i]
        }
        $minTemp = if ($temperatureFormat -eq "F") {
            ConvertTo-Fahrenheit -Celsius $weatherData.daily.temperature_2m_min[$i]
        } else {
            $weatherData.daily.temperature_2m_min[$i]
        }
        Write-Host "$($dayNames[$i]): Max Temp: $maxTemp$unit, Min Temp: $minTemp$unit, Wind: $($weatherData.daily.wind_speed_10m_max[$i]) km/h, Condition: $(Decode-WeatherCode -weatherCode $weatherData.daily.weathercode[$i])"
    }
} else {
    Write-Host "Failed to fetch weather data." -ForegroundColor Red
}
