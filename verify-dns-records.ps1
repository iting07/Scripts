function Import-Xls 
{ 
 
<# 
Ref: https://gallery.technet.microsoft.com/office/17bcabe7-322a-43d3-9a27-f3f96618c74b

.SYNOPSIS 
Import an Excel file. 
 
.DESCRIPTION 
Import an excel file. Since Excel files can have multiple worksheets, you can specify the worksheet you want to import. You can specify it by number (1, 2, 3) or by name (Sheet1, Sheet2, Sheet3). Imports Worksheet 1 by default. 
 
.PARAMETER Path 
Specifies the path to the Excel file to import. You can also pipe a path to Import-Xls. 
 
.PARAMETER Worksheet 
Specifies the worksheet to import in the Excel file. You can specify it by name or by number. The default is 1. 
Note: Charts don't count as worksheets, so they don't affect the Worksheet numbers. 
 
.INPUTS 
System.String 
 
.OUTPUTS 
Object 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet 1 
Import Worksheet 1 from employees.xlsx 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet "Sheet2" 
Import Worksheet "Sheet2" from employees.xlsx 
 
.EXAMPLE 
".\deptA.xslx", ".\deptB.xlsx" | Import-Xls -Worksheet 3 
Import Worksheet 3 from deptA.xlsx and deptB.xlsx. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.EXAMPLE 
Get-ChildItem *.xlsx | Import-Xls -Worksheet "Employees" 
Import Worksheet "Employees" from all .xlsx files in the current directory. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.LINK 
Import-Xls 
http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b 
Export-Xls 
http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2 
Import-Csv 
Export-Csv 
 
.NOTES 
Author: Francis de la Cerna 
Created: 2011-03-27 
Modified: 2011-04-09 
#Requires –Version 2.0 
#> 
 
    [CmdletBinding(SupportsShouldProcess=$true)] 
     
    Param( 
        [parameter( 
            mandatory=$true,  
            position=1,  
            ValueFromPipeline=$true,  
            ValueFromPipelineByPropertyName=$true)] 
        [String[]] 
        $Path, 
     
        [parameter(mandatory=$false)] 
        $Worksheet = 1, 
         
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
 
    Begin 
    { 
        function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
             
        # since an extension like .xls can have multiple formats, this 
        # will need to be changed 
        # 
        $xlFileFormats = @{ 
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        } 
         
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
    } 
 
    Process 
    { 
        $Path | ForEach-Object { 
             
            if ($Force -or $psCmdlet.ShouldProcess($_)) { 
             
                $fileExist = Test-Path $_ 
 
                if (-not $fileExist) { 
                    Write-Error "Error: $_ does not exist" -Category ResourceUnavailable;             
                } else { 
                    # create temporary .csv file from excel file and import .csv 
                    # 
                    $_ = (Resolve-Path $_).toString(); 
                    $wb = $xl.Workbooks.Add($_); 
                    if ($?) { 
                        $csvTemp = GetTempFileName(".csv"); 
                        $ws = $wb.Worksheets.Item($Worksheet); 
                        $ws.SaveAs($csvTemp, $xlFileFormats[".csv"]); 
                        $wb.Close($false); 
                        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
                        Import-Csv $csvTemp; 
                        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
                    } 
                } 
            } 
        } 
    } 
    
    End 
    { 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        [gc]::Collect(); 
    } 
} 
#-------------------------------------------------

# Enter your domain name
$domain = ""
# Enter your DNS Server IP
$dnsServerIp = ""
# Enter your file name to import
$fileName = ""

$hash = @{}
Import-Xls .\$fileName | `
foreach {
	$hash.Add($_.Hosts, $_.IPs)
}

foreach ($hostName in $hash.Keys) {
    $result = ""
    $inputIp = $hash[$hostName]
    $dnsName = $hostName + '.' + $domain
    Write-Host -BackgroundColor Black "Verifying A record and PTR record for $dnsName ($inputIp)"
    try {
        # Check if Host name has an A record
        $actualIp = (nslookup $dnsName $dnsServerIp | Select-String -ErrorAction SilentlyContinue Address | Where-Object -ErrorAction SilentlyContinue LineNumber -eq 5).ToString().Split(' ')[-1]
        if ($actualIp -eq $inputIp) {
            $result = $result + "$dnsName is valid, matches input IP"
            try {
                # Check if input IP has a PTR record
                $ptr = (nslookup $actualIp $dnsServerIp | Select-String -ErrorAction SilentlyContinue Name | Where-Object -ErrorAction SilentlyContinue LineNumber -eq 4).ToString().Split('')[-1].Split('.')[0]
                $result = $result + " and has a valid PTR record"
            } catch {
                $result = $result + ", but no PTR record - Create a PTR record for $dnsName"
            }
        } else {
            $result = $result + "$dnsName is valid, but does not match input IP ($inputIp)."
            try {
                $ptr = (nslookup $actualIp $dnsServerIp | Select-String -ErrorAction SilentlyContinue Name | Where-Object -ErrorAction SilentlyContinue LineNumber -eq 4).ToString().Split('')[-1].Split('.')[0]
                $result = $result + " $actualIp has a valid PTR record - Update input IP to $actualIp"
            } catch {
                $result = $result + " $actualIp does not have a PTR record - Update input IP to $actualIp and create a PTR record for $dnsName"
            }
        }
    } catch {
        $result = $result + "$dnsName is not valid or lookup timed out, please check this one manually"
    }
    
    if ($result.Contains("timed out")) {
        Write-Host -BackgroundColor Red -ForegroundColor Black $result
    } elseif ($result.Contains("Update input IP")) {
        Write-Host -BackgroundColor Yellow -ForegroundColor Black $result
    }elseif ($result.Contains("no PTR")) {
        Write-Host -BackgroundColor Cyan -ForegroundColor Black $result
    } else {
        Write-Host -BackgroundColor Green -ForegroundColor Black $result
    }
    
    Write-Host -BackgroundColor Gray -ForegroundColor Black "***********Next***********"

    <#
        OUTPUT INFO:
            - GREEN     = Valid DNS NAME with A record and PTR record configured.
            - RED       = DNS NAME is not valid or lookup timed out.
            - YELLOW    = Valid DNS NAME with A record and PTR record but incorrect input IP.
            - CYAN      = Valid DNS NAME with A record but no PTR record.
    #>
}
