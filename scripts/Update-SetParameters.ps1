param(
    $parameters, 
    $config
)
Write-Host  ("Update-SetParameters.ps1  " -f $parameters.OuterXml, ( $config| ConvertTo-Json )) -ForegroundColor Yellow
  
# read in the setParameters file
$parameters.parameters.setParameter | ForEach-Object {
    $parameter = $_
    if($parameter.name -eq "AzureWebJobsStorage-Web.config Connection String")
    {
        $parameter.value = $config.StorageAccountConnectionString

    }elseif($config.($parameter.name) -ne $null){
        $parameter.value = $config.($parameter.name)
    } 

}
return $parameters
 
