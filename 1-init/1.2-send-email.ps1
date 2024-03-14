# Importar o módulo do Power BI
Import-Module MicrosoftPowerBIMgmt

# Fazer login no Power BI
Connect-PowerBIServiceAccount

# Definir o nome do workspace e o nome do relatório
$workspaceName = "Nome do Workspace"
$reportName = "Nome do Relatório"

# Obter o ID do workspace
$workspaceId = (Get-PowerBIWorkspace -Name $workspaceName).Id

# Obter o ID do relatório
$reportId = (Get-PowerBIReport -WorkspaceId $workspaceId -Name $reportName).Id

# Exportar o relatório para Excel
Invoke-PowerBIRestMethod -Url "v1.0/myorg/groups/$workspaceId/reports/$reportId/Export" -Method Post -OutFile "C:\Caminho\para\o\arquivo.xlsx" -ContentType "application/xlsx"

# Exportar o relatório para PDF
Invoke-PowerBIRestMethod -Url "v1.0/myorg/groups/$workspaceId/reports/$reportId/Export" -Method Post -OutFile "C:\Caminho\para\o\arquivo.pdf" -ContentType "application/pdf"

# Enviar o email com os anexos
$smtpServer = "smtp.example.com"
$from = "seuemail@example.com"
$to = "destinatario@example.com"
$subject = "Relatórios do Power BI"
$body = "Veja os relatórios do Power BI em anexo."
$attachments = @("C:\Caminho\para\o\arquivo.xlsx", "C:\Caminho\para\o\arquivo.pdf")

Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject $subject -Body $body -Attachments $attachments
