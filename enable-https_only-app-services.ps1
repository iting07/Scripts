# To enable HTTPS Only for App Services in Azure.

# Enter the Azure Subscriptions you want to look through
$azureSubscriptions = @('Dev/Test', 'Production')

foreach ($subscription in $azureSubscriptions) {
    if ((Get-AzContext).Subscription.Name -ne $subscription) {
        Write-Host "Changing Azure Subscription context to $subscription..."
        Select-AzSubscription $subscription
    } else {
        Write-Host "Currently in $subscription Subscription"
    }

    Write-Host "Getting all App Services without HTTPS Only enabled..."
    $appServicesWithoutHttpsOnly = Get-AzWebApp | Where-Object {$_.HttpsOnly -eq $false} | Select-Object Name -ExpandProperty Name

    if ($appServicesWithoutHttpsOnly.Count -eq 0) {
        Write-Host "All App Services in $subscription have HTTPS Only enabled"
    } else {
        foreach ($appService in $appServicesWithoutHttpsOnly) {
            Write-Host "Enabling HTTPS Only on $appService.."
            Set-AzWebApp -Name $appService -ResourceGroupName (Get-AzWebApp -Name $appService).ResourceGroup -HttpsOnly $true
        }
    }
}
