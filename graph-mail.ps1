$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$request = [Console]::In.ReadToEnd() | ConvertFrom-Json
$context = Get-MgContext
if ($context.AuthType -ne 'Delegated' -or $context.Scopes -notcontains 'Mail.ReadWrite') {
    throw "Sign in through Connect-MgGraph with delegated User.Read and Mail.ReadWrite."
}
$root = 'https://graph.microsoft.com/v1.0'
$me = Invoke-MgGraphRequest -Method GET -Uri "$root/me?`$select=id,mail,userPrincipalName"
if ($request.account -notin @($me.mail,$me.userPrincipalName)) {
    throw "The signed-in Graph account differs from the configured mailbox."
}
if (!$request.message) {
    @{account=$request.account} | ConvertTo-Json -Compress
    exit
}
$mailbox = "$root/users/$($me.id)"
$created = Invoke-MgGraphRequest -Method POST -Uri "$mailbox/mailFolders/drafts/messages" -ContentType 'application/json' -Body ($request.message | ConvertTo-Json -Depth 12 -Compress)
$id = [Uri]::EscapeDataString($created.id)
$saved = Invoke-MgGraphRequest -Method GET -Uri "$mailbox/messages/${id}?`$select=id,subject,body,toRecipients,isDraft,webLink" -Headers @{Prefer='outlook.body-content-type="text"'}
if ($request.message.attachments.Count) {
    $saved.attachments = (Invoke-MgGraphRequest -Method GET -Uri "$mailbox/messages/$id/attachments").value
} else {
    $saved.attachments = @()
}
@{createdId=$created.id;saved=$saved} | ConvertTo-Json -Depth 12 -Compress
