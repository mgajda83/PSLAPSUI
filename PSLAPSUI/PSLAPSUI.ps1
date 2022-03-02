Add-Type -AssemblyName PresentationFramework
[xml]$xaml =
@"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PS LAPS UI v1.0.0.1" Height="237" Width="500" MinWidth="500" MinHeight="250"  WindowStyle="ToolWindow" SizeToContent="WidthAndHeight">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="100"/>
        </Grid.ColumnDefinitions>

        <Label Content="Domain Controller" Height="20" Margin="10,10,10,0" VerticalAlignment="Top" FontSize="9" Grid.ColumnSpan="3"/>
        <TextBox x:Name="Domain" Text="<FQDN of Domain Controller>" Height="19" Margin="10,30,10,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
        <Button x:Name="Credential" Content="Credential" Height="19" Width="90" Margin="0,30,10,0" VerticalAlignment="Top" Grid.Column="2"/>
        <Label Content="Computer Name" Height="20" Margin="10,50,10,0" VerticalAlignment="Top" FontSize="9" Grid.ColumnSpan="3"/>
        <TextBox x:Name="Computer" Height="19" Margin="10,70,10,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="" IsEnabled="False"/>
        <Button x:Name="Search" Content="Search" Height="19" Width="90" Margin="0,70,10,0" VerticalAlignment="Top" Grid.Column="2" IsEnabled="False"/>
        <Label Content="LAPS Password" Height="20" Margin="10,90,10,0" VerticalAlignment="Top" FontSize="9" Grid.ColumnSpan="3"/>
        <TextBox x:Name="Password" Height="19" Margin="10,110,10,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="" IsReadOnly="True"/>
        <Label Content="Expiration Date" Height="20" Margin="10,130,10,0" VerticalAlignment="Top" FontSize="9" Grid.ColumnSpan="3"/>
        <TextBox x:Name="Expiration" Height="19" Margin="10,150,10,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="" IsReadOnly="True"/>
    </Grid>
</Window>
"@

Function Get-CredentialClick
{
    $ADModule = Get-Module ActiveDirectory -ListAvailable
    if($null -eq $ADModule)
    {
        $Button = [System.Windows.MessageBoxButton]::OK
        $Icon = [System.Windows.MessageBoxImage]::Error
        $DefaultButton = [System.Windows.MessageBoxResult]::None

        $Result = [System.Windows.MessageBox]::Show("Install Powershell module: ActiveDirectory","Error",$Button,$Icon,$DefaultButton)
    } else {
        $Credential = Get-Credential
        Set-Variable -Name Credential -Value $Credential -Scope Global

        if($Global:Credential)
        {
            $TextBox_Computer.IsEnabled = $true
            $Button_Search.IsEnabled = $true
        }
    }
}

Function Get-ADComputerClick
{
    if($Global:Credential)
    {
        $Credential = $Global:Credential
        $Computer = $TextBox_Computer.Text
        $Server = $TextBox_Domain.Text

        $ADComputer = Get-ADComputer $Computer -Properties ms-Mcs-AdmPwd, ms-Mcs-AdmPwdExpirationTime -Server $Server -Credential $Global:Credential

        if($null -ne $ADComputer)
        {
            $TextBox_Password.Text = $ADComputer.'ms-Mcs-AdmPwd'
            $TextBox_Expiration.Text = $([DateTime]::FromFileTime([Int64]::Parse($ADComputer.'ms-mcs-admpwdexpirationtime')))
        } else {
            $Button = [System.Windows.MessageBoxButton]::OK
            $Icon = [System.Windows.MessageBoxImage]::Error
            $DefaultButton = [System.Windows.MessageBoxResult]::None

            $Result = [System.Windows.MessageBox]::Show("Can't find the computer object: $Computer","Error",$Button,$Icon,$DefaultButton)
        }
    }
}

$XmlNodeReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($XmlNodeReader)

$TextBox_Domain = $Window.FindName("Domain")
$TextBox_Computer = $Window.FindName("Computer")
$TextBox_Password = $Window.FindName("Password")
$TextBox_Expiration = $Window.FindName("Expiration")

$Button_Credential = $Window.FindName("Credential")
$Button_Search = $Window.FindName("Search")

$Method_Credential = $Button_Credential.add_click
$Method_Credential.Invoke({Get-CredentialClick})

$Method_Search = $Button_Search.add_click
$Method_Search.Invoke({Get-ADComputerClick})

$Window.ShowDialog() | Out-Null
