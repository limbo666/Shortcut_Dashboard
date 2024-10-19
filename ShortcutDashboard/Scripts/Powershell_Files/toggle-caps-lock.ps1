

try {
	$wsh = New-Object -ComObject WScript.Shell
	$wsh.SendKeys('{CAPSLOCK}')
	exit 0 # success
} catch {
	"⚠️ Error in line $($_.InvocationInfo.ScriptLineNumber): $($Error[0])"
	exit 1
}