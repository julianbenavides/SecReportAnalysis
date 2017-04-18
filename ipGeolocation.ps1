param (
	[Parameter( Mandatory=$true)]
	[string]$ip
)

function GetIPGeolocation() {

    param($ipaddress)

    $resource = "http://freegeoip.net/xml/$ipaddress"

    $geoip = Invoke-RestMethod -Method Get -URI $resource

    $hash = @{
        IP = $geoip.Response.IP
        CountryCode = $geoip.Response.CountryCode
        CountryName = $geoip.Response.CountryName
        RegionCode = $geoip.Response.RegionCode
        RegionName = $geoip.Response.RegionName
        City = $geoip.Response.City
        ZipCode = $geoip.Response.ZipCode
        TimeZone = $geoip.Response.TimeZone
        Latitude = $geoip.Response.Latitude
        Longitude = $geoip.Response.Longitude
        MetroCode = $geoip.Response.MetroCode
        }

    $result = New-Object PSObject -Property $hash

    return $result.CountryName

}

GetIPGeolocation $ip
