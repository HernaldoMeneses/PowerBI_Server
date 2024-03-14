#Definir a flag como true ou false
$executarScript = $true
$executarScript = $false
#Verificar se a flag é true antes de importa o módulo do Power BI
if ($executarScript -eq $true) {
    # Instalar o módulo MicrosoftPowerBIMgmt
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force

} else {
    #Hello!
    Write-Host "Hello!"
}

# Importar o módulo do Power BI
Import-Module MicrosoftPowerBIMgmt

# Fazer login no Power BI
Connect-PowerBIServiceAccount

# Listar todos os workspaces
#Get-PowerBIWorkspace

#Explorando as Workspace e seus relatorios
$workspaceName = "Financeiro"
#Obter ID da workspace
$workspaceId = (Get-PowerBIWorkspace -Name $workspaceName).Id
Write-Host $workspaceName

#Listar os relatórios da workspace
#Get-PowerBIReport -WorkspaceId $workspaceId

#Definir a flag como true ou false
$executarScript = $true
$executarScript = $true
#Verificar se a flag é true antes de importa o módulo do Power BI
if ($executarScript -eq $true) {
    #Definir o nome do relatorios
    $reportName = "Fluxo de caixa"
    Write-Host $reportName
    #Obter o ID do Relatório
    $reportId = (Get-PowerBIReport -WorkspaceId $workspaceId -Name $reportName).Id
    
    
    Write-Host $workspaceId
    Write-Host $reportID
    Write-Host $reportName

    # Definir o formato desejado (Pdf ou Excel)
    $formato =  "PDF"

    # Definir o corpo da requisição
    $body = @{
        format = $formato
    }

    # Exportar o relatorio
    $url = "v1.O/myorg/groups/$workspaceId/reports/$reportId/Export"
    $url = "https://api.powerbi.com/v1.O/myorg/groups/$workspaceId/reports/$reportId/Export"
    Invoke-PowerBIRestMethod -Url $url -Method Post -OutFile "C:\Projeto_PowerBI\PowerShell\Reports.$formato" -ContentType "application/$formato" -Body $body 

    # Exportar o relatóro para Excel
    #$excelFilePath = "C:\Projeto_PowerBI\PowerShell\Reports\arq.xlsx"
    #Invoke-PowerBIRestMethod -Url "v1.0/myorg/groups/$workspaceId/reports/$reportId/Export" -Method Post -OutFile $excelFilePath -ContentType "application/xlsx"
    # Exportar o relatóro para Pdf
    #$excelFilePath = "C:\Projeto_PowerBI\PowerShell\Reports\arq.pdf"
    #Invoke-PowerBIRestMethod -Url "v1.0/myorg/groups/$workspaceId/reports/$reportId/Export" -Method Post -OutFile $pdfFilePath -ContentType "application/pdf"


    # Enviar o email com os anexos
    #$smtpServer = "smtp.example.com"
    #$from = " friobombac.pbi@friobombacabal.com.br"
    #$to = "hernaldomeneses@gmail.com"
    #$subject = "Relatrórios do Power BI"
    #$body = "Veja os relatórios do Power Bi em anexo."
    #$attachments = @($excelFilePath, $pdfFilePath)

    #Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject $subject -Body $body -Attachments $attachments

} else {
    #Hello!
    Write-Host "Hello!"
}

#Certifique-se de ter o módulo MicrosoftPowerBIMgmt instalado antes de executar o script. 
#Você pode instalá-lo usando o comando Install-Module -Name MicrosoftPowerBIMgmt se ainda não o tiver instalado.

#Depois de executar o script, você verá uma lista de todos os workspaces do Power BI associados à sua conta.

#Se você tiver alguma dúvida ou precisar de mais alguma coisa, por favor, me avise!


#Definir a flag como true ou false
$executarScript = $true
$executarScript = $false
#Verificar se a flag é true antes de importa o módulo do Power BI
if ($executarScript -eq $true) {
    # Definir a política de execução para RemoteSigned (recomendado para ambientes seguros)
    Set-ExecutionPolicy RemoteSigned

    # Ou, definir a política de execução para Unrestricted (menos seguro)
    Set-ExecutionPolicy Unrestricted
} else {
    #Hello!
    Write-Host "Hello!"
}