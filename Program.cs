using System.Management.Automation.Runspaces;
using System.Threading.Tasks;

namespace PowerShellWSManBug
{
    class Program
    {
        static void Main()
        {
            Execute().Wait();
        }

        static async Task Execute()
        {
            var rs = RunspaceFactory.CreateRunspace();
            rs.Open();
            using (var ps = System.Management.Automation.PowerShell.Create())
            {
                ps.Runspace = rs;
                var initializationScript = $@"
$ErrorActionPreference = 'Stop'
try {{ Set-ExecutionPolicy Unrestricted }} catch {{}} # not supported on non-Windows platforms
$UserCredential = New-Object System.Management.Automation.PSCredential('******', (ConvertTo-SecureString '******' -AsPlainText -Force))
$Option = New-PSSessionOption
$Option.IdleTimeout = [TimeSpan]::FromSeconds(60) # inline setting of this property via New-PSSessionOption is not supported on non-Windows platforms
$Session = New-PSSession -SessionOption $Option -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-Module (Import-PSSession $Session -DisableNameChecking) -Global
";
                await ps.AddScript(initializationScript).InvokeAsync();
            }

        }
    }
}
