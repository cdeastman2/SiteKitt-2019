Dim CommandLine As String = ("Invoke-Command -ComputerName " & computer & " -ScriptBlock {$Cache = Get-WmiObject -Namespace ‘ROOT\CCM\SoftMgmtAgent’ -Class CacheConfig;$Cache.Size = '" & Cachesize & "';$Cache.Put();Restart-Service -Name CcmExec}")

