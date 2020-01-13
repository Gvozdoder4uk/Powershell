##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBn/Zrg7CX1UuWf1B0Td
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQC95opyrOCeWY1pgetJkNh0LLFUXKBAGhiq3wnyI30r3HHvW6TacGnqNPeeqEo7E9KeogUnJEDKlMf0Zww72N5olPsBO0pf6bdYVtm91DOV/Ybl+TfezO66pJmnerl0pxmBIOG7KZ9g==
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba

#GLOBAL VARIABLE
$Global:RLS=''

#######################################################################################################################################################
#||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Function Global:New-WPFMessageBox {

    # For examples for use, see my blog:
    # https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
    
    # CHANGES
    # 2017-09-11 - Added some required assemblies in the dynamic parameters to avoid errors when run from the PS console host.
    
    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segoe UI"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
            New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}
#######################################################################################################################################################
#||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

# .Net methods for hiding/showing the console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    # 4 SHOW
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

##########################################################################################################################################################################################################
#AUTOINSTALL FOBO START
#
Function DEPLOYFOBO([string]$Machine, [string]$Server){

####################################################################
# START FUNCTION FOR WMI GET-SERVICE                               #
####################################################################
Function INSTALL_FOBO_NEW([string]$Server,[string]$Machine)
{
#Прочекать сервис
function Get-ServiceNTS(
    [string]$serviceName = $(throw "serviceName is required"), 
    [string]$targetServer = $(throw "targetServer is required"))
{
    $service = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_Service" `
        -ComputerName $targetServer -Filter "Name='$serviceName'" -Impersonation 3    
    return $service
}

#Удалить сервис
function Uninstall-ServiceNTS(
    [string]$serviceName = $(throw "serviceName is required"), 
    [string]$targetServer = $(throw "targetServer is required"))
{
    $service = Get-ServiceNTS $serviceName $targetServer
     
    if (!($service))
    { 
        Write-Warning "Failed to find service $serviceName on $targetServer. Nothing to uninstall."
        return
    }
     
    "Found service $serviceName on $targetServer; checking status"
             
    if ($service.Started)
    {
        "Stopping service $serviceName on $targetServer"
        #could also use Set-Service, net stop, SC, psservice, psexec etc.
        $result = $service.StopService()
        Test-ServiceResult -operation "Stop service $serviceName on $targetServer" -result $result
    }
     
    "Attempting to uninstall service $serviceName on $targetServer"
    $result = $service.Delete()
    Test-ServiceResult -operation "Delete service $serviceName on $targetServer" -result $result   
}

#Проверки на установку и прочие системные проверки
function Test-ServiceResult(
    [string]$operation = $(throw "operation is required"), 
    [object]$result = $(throw "result is required"), 
    [switch]$continueOnError = $false)
{
    $retVal = -1
    if ($result.GetType().Name -eq "UInt32") { $retVal = $result } else {$retVal = $result.ReturnValue}
         
    if ($retVal -eq 0) {return}
     
    $errorcode = 'Success,Not Supported,Access Denied,Dependent Services Running,Invalid Service Control'
    $errorcode += ',Service Cannot Accept Control, Service Not Active, Service Request Timeout'
    $errorcode += ',Unknown Failure, Path Not Found, Service Already Running, Service Database Locked'
    $errorcode += ',Service Dependency Deleted, Service Dependency Failure, Service Disabled'
    $errorcode += ',Service Logon Failure, Service Marked for Deletion, Service No Thread'
    $errorcode += ',Status Circular Dependency, Status Duplicate Name, Status Invalid Name'
    $errorcode += ',Status Invalid Parameter, Status Invalid Service Account, Status Service Exists'
    $errorcode += ',Service Already Paused'
    $desc = $errorcode.Split(',')[$retVal]
     
    $msg = ("{0} failed with code {1}:{2}" -f $operation, $retVal, $desc)
     
    if (!$continueOnError) { Write-Error $msg } else { Write-Warning $msg }        
}
#Конец блоки тестирования

#Установка сервиса
function Install-ServiceNTS(
    [string]$serviceName = $(throw "serviceName is required"), 
    [string]$targetServer = $(throw "targetServer is required"),
    [string]$displayName = "NTSwincash distributor",
    [string]$physicalPath = "C:\NTSwincash\jbin\DistributorService.exe",
    #[string]$userName = $(throw "userName is required"),
    [string]$password = "",
    [string]$startMode = "Automatic",
    [string]$description = " ",
    [bool]$interactWithDesktop = $false
)
{
    # can't use installutil; only for installing services locally
    #[wmiclass]"Win32_Service" | Get-Member -memberType Method | format-list -property:*    
    #[wmiclass]"Win32_Service"::Create( ... )        
          
    # todo: cleanup this section 
    $serviceType = 16          # OwnProcess
    $serviceErrorControl = 1   # UserNotified
    $loadOrderGroup = $null
    $loadOrderGroupDepend = $null
    $dependencies = $null
     
    # description?
    $params = `
        $serviceName, `
        $displayName, `
        $physicalPath, `
        $serviceType, `
        $serviceErrorControl, `
        $startMode, `
        $interactWithDesktop, `
        $userName, `
        $password, `
        $loadOrderGroup, `
        $loadOrderGroupDepend, `
        $dependencies `
          
    $scope = new-object System.Management.ManagementScope("\\$targetServer\root\cimv2", `
        (new-object System.Management.ConnectionOptions))
    "Connecting to $targetServer"
    $scope.Connect()
    $mgt = new-object System.Management.ManagementClass($scope, `
        (new-object System.Management.ManagementPath("Win32_Service")), `
        (new-object System.Management.ObjectGetOptions))
      
    $op = "service $serviceName ($physicalPath) on $targetServer"    
    "Installing $op"
    $result = $mgt.InvokeMethod("Create", $params)    
    Test-ServiceResult -operation "Install $op" -result $result
    "Installed $op"
      
    "Setting $serviceName description to '$description'"
    Set-Service -ComputerName $targetServer -Name $serviceName -Description $description
    "Service install complete"
}
#Конец блока установки сервиса.

# Определение Среды VRX [Магазин - Контур Пакет]
$VRXPACKAGES = @{
##########################
# VRX 1
##########################
    '166'='1 WS-M0';
    '279'='1 WS-M0';
    '105'='1 WS-M1';
    '660'='1 WS-M1';
    '024'='1 WS-M2';
    '050'='1 WS-M2';
    '175'='1 WS-M3';
    '180'='1 WS-M4';
    '061'='1 WS-M4';
##########################
# VRX 2
##########################
    '134'='2 WS-M0';
    '465'='2 WS-M0';
    'A01'='2 WS-M1';
    '217'='2 WS-M1';
    '266'='2 WS-M1';
    '064'='2 WS-M2';
    '299'='2 WS-M3';
    '482'='2 WS-M3';
    '208'='2 WS-M4';
    '469'='2 WS-M5';
    '018'='2 WS-M5';
##########################
# VRX 3
##########################
    '111'='3 WS-M0';
    '123'='3 WS-M0';
    '190'='3 WS-M1';
    '401'='3 WS-M1';
    '191'='3 WS-M1';
    '444'='3 WS-M2';
    '306'='3 WS-M5';
##########################
# VRX 4
##########################
    '025'='4 WS-M0';
    '099'='4 WS-M0';
    '118'='4 WS-M1';
    '119'='4 WS-M1';
    '102'='4 WS-M2';
    '112'='4 WS-M2';
    '230'='4 WS-M3';
##########################
# VRX 5
##########################
    '122'='5 WS-M00';
    '146'='5 WS-M01';
    '494'='5 WS-M02';
    '106'='5 WS-M03';
    '284'='5 WS-M04';
    '127'='5 WS-M05';
    '269'='5 WS-M06';
    '400'='5 WS-M07';
    '224'='5 WS-M08';
    '139'='5 WS-M09';
    '158'='5 WS-M09';
    '110'='5 WS-M10';
    '399'='5 WS-M10';
    '107'='5 WS-M11';
    '278'='5 WS-M11';
    '152'='5 WS-M12';
    '258'='5 WS-M13';
    '434'='5 WS-M14';
    '056'='5 WS-M15';
    '067'='5 WS-M15';
##########################
# VRX 6
##########################
    '014'='6 WS-M0';
    '015'='6 WS-M0';
    '196'='6 WS-M1';
    '461'='6 WS-M1';
    '141'='6 WS-M2';
    '188'='6 WS-M2';
    '130'='6 WS-M3';
    '132'='6 WS-M3';
    '128'='6 WS-M4';
    '131'='6 WS-M4';
    '543'='6 WS-M5';
    '754'='6 WS-M5';

}

# Определение Среды VRQ [Магазин - Контур Пакет]
$VRQPACKAGES = @{
##########################
# VRQ 1
##########################
    '111'='1 WS';
    '123'='1 WS';
    '105'='1 WS-M0';
    '166'='1 WS-M0';
    '279'='1 WS-M0';
    '190'='1 WS-M0';
    'A02'='1 WS-M1';
    '061'='1 WS-M1';
    '660'='1 WS-M2';
##########################
# VRQ 2
##########################
    '064'='2 WS';
    '299'='2 WS';
    '482'='2 WS';
    '266'='2 WS-M0';
    '306'='2 WS-M1';
    '208'='2 WS-M2';
    'A01'='2 WS-M3';
##########################
# VRQ 3
##########################
    '142'='3 WS';
    '102'='3 WS-M0';
    '112'='3 WS-M1';
##########################
# VRQ 4
##########################
    '118'='4 WS-M0';
    '119'='4 WS-M0';
    '120'='4 WS-M1';
##########################
# VRQ 5
##########################
    '217'='5 WS';
    '230'='5 WS-M0';
    '235'='5 WS-M0';
##########################
# VRQ 6
##########################
    '302'='6 WS';
    '401'='6 WS';
    '025'='6 WS-M0';
##########################
# VRQ 7
##########################
    '077'='6 WS';
    '099'='6 WS';
    '191'='6 WS-M0';
    '444'='6 WS-M1';
    '014'='6 WS-M2';
    '015'='6 WS-M3';
}
# Связка пакета с зоной. (Бесполезная операция)
$ZONESX = $VRXPACKAGES
$ZONESQ = $VRQPACKAGES

#Создание и Заполнение DBLINKS
#===============================================================================================================
$Global:FoboStatus.Text = 'Создание DBLINKS. Ожидайте!'
Start-Sleep -Seconds 2
Copy-Item "\\$Server\C$\NTSwincash\config\*" -Filter 'dblink_*' -Destination "\\$Machine\C$\NTSwincash\config\"
$XMLNAME =  Get-ChildItem "\\$Server\C$\NTSwincash\config\*" -Include "dblink_V*" 
$XMLNAME.Name
[xml]$Doc = New-Object System.Xml.XmlDocument                 
$FilePath = "\\$Server\C$\NTSwincash\config\dblinks.xml"
$Path = "\\$Machine\C$\NTSwincash\config\dblinks.xml"
$Path2 = "\\$Machine\C$\NTSwincash\config\dblinks_MobInv.xml"
$doc.Load($filePath)
$doc.linklist.linkref.file = $XMLNAME.name
$doc.Save($Path)
$doc.Save($Path2)
# Конец блока создания и Копирования DBLINKS.
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-==-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=--=
#
$Global:FoboStatus.Text = 'Копирование ПО!'
Start-Sleep -Seconds 2

#
# - Блок проверки Сервера зоны VRX - Выбор пакета.
If($Server -like "*vrx*")
{
    
    foreach($T in $ZONESX.Keys)
    {
        #Write-Host $T
        if($Server -like "*"+$T)
        {
            $PacKet = $ZONESX.$T
            $PacKet = $PacKet.remove(0,2)
            $STR = $ZONESX.$T
            $T = "fobo-vrx-ajb" +  $STR[0]
            $Source = "\\$T\C$\EtalonR3\$PacKet"
            #Invoke-Item $Source
            #Copy-Item -Path $Source\* -Destination \\$Machine\c$\NTSwincash -Recurse -PassThru
            # Копирование с использованием джоба.
            $Copy = [scriptblock]::create("Copy-Item -Path $Source\* -Destination \\$Machine\c$\NTSwincash -Recurse -PassThru")
            Start-Job -Name "COPYFOBO" -scriptblock $Copy  
            Wait-Job -Name "COPYFOBO"
            

        }
    }

}
# - Блок проверки Сервера зоны VRQ - Выбор пакета.
elseif($Server -like "*vrq*")
{
    foreach($T in $ZONESQ.Keys)
    {
        #Write-Host $T
        if($Server -like "*"+$T)
        {
            $PacKet = $ZONESQ.$T
            $PacKet = $PacKet.remove(0,2)
            Write-Host "ВЫБРАН ПАКЕТ: "$PacKet
            $STR = $ZONESQ.$T
            $T = "fobo-vrq-ajb" +  $STR[0]
            Write-Host "ВЫБРАН СЕРВЕР: "$T
            $Source = "\\$T\C$\EtalonR3\$PacKet"
            $Copy = [scriptblock]::create("Copy-Item -Path $Source\* -Destination \\$Machine\c$\NTSwincash -Recurse -PassThru")
            Start-Job -Name "COPYFOBO" -scriptblock $Copy  
            Wait-Job -Name "COPYFOBO"

        }
    }


}


    $Global:FoboStatus.Text = 'Запущена процедура установки службы!'
    Start-Sleep -Seconds 3
    $Service = "NTSwincash distributor"
    $StatusNTS = $null
    $StatusNTS = Get-ServiceNTS $Service $Machine
if($StatusNTS -eq $null)
{
    #Write-Host "Сервиса нет!"
    #[System.Windows.Forms.MessageBox]::Show("Будет выполнена установка сервиса!","NTS","OK")
    Install-ServiceNTS $Service $Machine
    Start-Sleep -Seconds 5
    $StatusNTS = $null
    $StatusNts = Get-ServiceNTS $Service $Machine
    if($StatusNTS -ne $null)
    {
           $Global:FoboStatus.Text = 'Служба установлена! Установка завершена'
           $Global:FoboStatus.ForeColor = 'Green'
    }
    else
    {
           $Global:FoboStatus.Text = 'Служба не установлена!'
           $Global:FoboStatus.ForeColor = 'Red'
    }
}
else
{
    $Global:FoboStatus.Text = 'Служба уже установлена!'
    Start-Sleep -Seconds 6
    $Global:FoboStatus.Text = "Установка выполнена!"
 
}

#End of Function
}
###############################
# END FUNCTION FOR WMI        #
###############################




            $TPath = Test-Path \\$Machine\C$\NTSwincash\config
            $TPath2 = Test-Path \\$Machine\C$\NTSwincash\jbin
            $Global:FoboStatus.Text = 'Выполняются проверки. Ожидайте'
            Start-Sleep -Seconds 2
            if ($TPath -eq $False)
            {
                if($TPath2 -eq $True)
                {
                }
                else
                {
                    New-Item -ItemType Directory -Path \\$Machine\C$\NTSwincash\jbin  
                }
                    $Global:FoboStatus.Text = 'Проверки пройдены! Переход к копированию!'
                    Start-Sleep -Seconds 2
                    New-Item -ItemType Directory  -Path \\$Machine\C$\NTSwincash\config
                    $ACL = ''
                    $FolderAcl = "\\$Machine\C`$\NTSwincash\"
                    $ACL = Get-Acl $FolderAcl
                    $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
                    $ACL.SetAccessRule($AccessRule)
                    $ACL | Set-Acl $FolderAcl
                    INSTALL_FOBO_NEW $SERVER $Machine
            }
            elseif ($TPath -eq $True)
            {
                    $Answer = [System.Windows.Forms.MessageBox]::Show("FOBO уже установлен на ПК
Выполнить переустановку?","Ошибка",'YesNoCancel','Warning')
                    switch($Answer)
                    {
                        "YES"{
                                (Get-WmiObject -Class Win32_Process -ComputerName $Machine -Filter "name='NTSWincash*.exe'").terminate() | Out-Null
                                $Service = "NTSwincash distributor"
                                Uninstall-ServiceNTS $Service $Machine
                                start-sleep -seconds 2
                                $Global:FoboStatus.Text = 'Выполняется удаление старой папки! Ожидайте!'
                                Start-Sleep -Seconds 2 
                                Remove-Item -Path \\$machine\C$\NTSwincash -Recurse -ErrorAction SilentlyContinue | Out-Null
                                Start-Sleep -Seconds 15
                                New-Item -ItemType Directory -Path \\$Machine\C$\NTSwincash\jbin
                                New-Item -ItemType Directory  -Path \\$Machine\C$\NTSwincash\config
                                $ACL = ''
                                $FolderAcl = "\\$Machine\C`$\NTSwincash\"
                                $ACL = Get-Acl $FolderAcl
                                $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
                                $ACL.SetAccessRule($AccessRule)
                                $ACL | Set-Acl $FolderAcl
                                INSTALL_FOBO_NEW $SERVER $Machine


                                
                               }
                    "NO"{return}
                    "CANCEL"{return}
              }

            }      
    }
Function FOBO_INSTALL([string]$Server)
{
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $FontJob = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Fobo_INTERFACE.jpg")

    #Create FOBO FORM
    $Fobo_Form = New-Object System.Windows.Forms.Form
    $Fobo_Form.BackgroundImage = $Image
    $Fobo_Form.BackgroundImageLayout = "None"
    if($Image -ne $null)
    {
    $Fobo_Form.Size = ('350,255')
    }
    else{
    $Fobo_Form.Size = ('350','255')
    #$Fobo_Form.Width = 
    #$Fobo_Form.Height = "500"
    }
    
    $Fobo_Form.StartPosition = "CenterScreen"
    $Fobo_Form.TopMost = $true
    $Fobo_Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $Fobo_Form.Text = "Окно установки FOBO"
    $Fobo_Form.Icon = $Icon

    $Global:FoboStatus = New-Object System.Windows.Forms.Label
    $Global:FoboStatus.Location = New-Object System.Drawing.Point('10','96')
    $Global:FoboStatus.Text = 'ЗДЕСЬ БУДЕТ СТАТУС ВЫПОЛНЕНИЯ РАБОТЫ!'
    $Global:FoboStatus.AutoSize = $True
    #$Global:FoboStatus.BackColor = 'Transparent'
    $Global:FoboStatus.ForeColor = 'Blue'


    $FoboRadio = New-Object System.Windows.Forms.RadioButton
    $FoboRadio.Location = New-Object System.Drawing.Point('10','10')
    $FoboRadio.Text = "Одна станция"
    $FoboRadio.BackColor = 'Transparent'
    $FoboRadio.AutoSize = 'True'
    $FoboRadio.Checked = 'True'

    $FoboRadio1 = New-Object System.Windows.Forms.RadioButton
    $FoboRadio1.Location = New-Object System.Drawing.Point('120','10')
    $FoboRadio1.Text = "Пакетная установка"
    $FoboRadio1.BackColor = 'Transparent'
    $FoboRadio1.AutoSize = 'True'

    $FoboMachine = New-Object System.Windows.Forms.TextBox
    $FoboMachine.Location = New-Object System.Drawing.Point('10','40')
    $FoboMachine.Size = '250,25'

    $FoboMachines = New-Object System.Windows.Forms.ListBox
    $FoboMachines.Location = New-Object System.Drawing.Point('10','40')
    $FoboMachines.Size = '200,150'
    $FoboMachines.Visible = 'False'

#Кнопка запуска процесса одиночного выполнения.
    $FoboProcess = New-Object System.Windows.Forms.Button
    $FoboProcess.Location = New-Object System.Drawing.Point('10','70')
    $FoboProcess.Size = '250,25'
    $FoboProcess.Text = 'Запустить установку!'

#Кнопка запуска процесса пакетного выполнения.
    $FoboProcess1 = New-Object System.Windows.Forms.Button
    $FoboProcess1.Location = New-Object System.Drawing.Point('215','70')
    $FoboProcess1.Text = 'Запустить установку!'

    $FoboFile = New-Object System.Windows.Forms.Button
    $FoboFile.Location = New-Object System.Drawing.Point('215','40')
    $FoboFile.Text = 'FILE'

    $EventMachine = {
        if($FoboRadio1.Checked)
        {
            $FoboStatus.Visible = $False
            $Fobo_Form.Height = $Image.Height
            $Fobo_Form.Width = $Image.Width
            $FoboProcess1.Visible = $True
            $FoboProcess.Visible = $False
            $FoboFile.Visible = $True
            $FoboMachine.Visible = $False
            $FoboMachine.Text = ''
            $FoboMachines.Visible = $True
        }
        elseif ($FoboRadio.Checked)
        {
            $Fobo_Form.Width = '300'
            $FoboStatus.Visible = $True
            $Fobo_Form.Height = '160'
            $FoboProcess.Visible = $True
            $FoboProcess1.Visible = $False
            $FoboFile.Visible = $False
            $FoboMachines.Visible = $False
            $FoboMachine.SelectedText = ''
            $FoboMachine.Text = ''
            $FoboMachine.Visible = $True
        }
                        
    }


    $EventFile = {
                $FoboMachines.Items.Clear()
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = 'Desktop'
                Filter = 'Text (*.txt)|*.txt|Все файлы |*.*'
                Title = 'Выберите список машин'}
                $FileBrowser.ShowDialog()
                Get-Content $FileBrowser.FileName | ForEach-Object {$FoboMachines.Items.Add($_)}

                }

#ФУНКЦИЯ УСТАНОВКИ И ПЕРЕУСТАНОВКИ FOBO


#Заготовка под функцию
$eventStartDeploy = {
    $Machine = $FoboMachine.Text
    
    IF($FoboRadio.Checked)
    {
        if($Machine -eq '' -or $Machine -eq $null)
        {
            $O = [System.Windows.Forms.MessageBox]::Show("Машина для установки FOBO не выбрана!","Ошибка",'OK','ERROR')
            switch($O)
            {
                'OK' {return}
            }
        }
        elseif($Machine -ne '' -or $Machine -ne $null)
        {
            DEPLOYFOBO $Machine $Server
        }
    
    }
    elseif($FoboRadio1.Checked)
    {
        [System.Windows.Forms.MessageBox]::Show("Над Пакетным решением ведется работа!","Ошибка",'OK','WARNING')
         return
    }


                    
}


        
               
    $FoboProcess.add_Click($eventStartDeploy)
    $FoboProcess1.Add_Click($eventStartDeploy)
    $FoboFile.add_Click($EventFile)
    $FoboRadio.add_Click($EventMachine)
    $FoboRadio1.add_Click($EventMachine)
    $Fobo_Form.Controls.AddRange(@($FoboRadio,$FoboRadio1,$FoboMachine,$FoboMachines,$FoboFile,$FoboProcess,$FoboProcess1,$FoboStatus))
    $Fobo_Form.ShowDialog()
}
##AUTOINSTALL FOBO END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#JOB MANIPULATOR START
Function JOB_WORKER([string]$SERVER){
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $FontJob = New-Object System.Drawing.Font("Comic Sans MS",9,[System.Drawing.FontStyle]::Bold)
    $Imagejob =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Worker2.jpg")
    #Create JOB Form
    $JobForm = New-Object System.Windows.Forms.Form
    $JobForm.SizeGripStyle = "Hide"
    $JobForm.BackgroundImage = $Imagejob
    $JobForm.BackgroundImageLayout = "None"
    if($Imagejob -eq $Null)
    {
     $JobForm.Size = ('500,313')
    }
    else{
    $JobForm.Width = $Imagejob.Width
    $JobForm.Height = $Imagejob.Height
    }
    $JobForm.StartPosition = "CenterScreen"
    #$JobForm.Top = $true
    $JobForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $JobForm.Text = "Список заданий сервера - $Server"
    $JobForm.TopMost = $True
    $JobForm.Icon = $Icon

    $JobForm.KeyPreview = $True
    $JobForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {$JobForm.Close()
    }
    })

    
    
    if($Server -like '*vrq-int*')
    {#[System.Windows.Forms.MessageBox]::Show("Для VRQ среды пока беда!","VRQ")
     #$JOBS_OF_SERVER = Get-ScheduledTask - $Server | Select-Object TaskName "FOBO*"
     $Schedule = New-Object -ComObject "Schedule.Service"
     $Schedule.Connect($Server)
     $JOB_OF_SERVER = $Schedule.GetFolder('\')
     $JOB_OF_SERVICE = $JOB_OF_SERVER.GetTasks(1) |  Select @{ Name = 'Name'
     Expression = {if($_.Name -like '*FOBO*')
     {$_.Name}
     else{}}
     },
     @{
      Name = 'State'
      Expression = {switch ($_.State) {
            0 {'Unknown'}
            1 {'Disabled'}
            2 {'Queued'}
            3 {'Ready'}
            4 {'Running'}
          }
       }
     }
    }
    elseif($Server -like "*vrq-ajb*")
    {
     [System.Windows.Forms.MessageBox]::Show("Для серверов центра работа сервиса не предусмотрена!","Контур")
     return
    }
    elseif($Server -like '*vrq-a*')
    {
     $Schedule = New-Object -ComObject "Schedule.Service"
     $Schedule.Connect($Server)
     $JOB_OF_SERVER = $Schedule.GetFolder('\NTSwincash')
     $JOB_OF_SERVICE = $JOB_OF_SERVER.GetTasks(1) |  Select @{ Name = 'Name'
     Expression = {if($_.Name -like '*JOB*')
     {$_.Name}
     else{}}
     },
     @{
      Name = 'State'
      Expression = {switch ($_.State) {
            0 {'Unknown'}
            1 {'Disabled'}
            2 {'Queued'}
            3 {'Ready'}
            4 {'Running'}
          }
       }
     }
    }
    elseif($Server -like '*ajb*')
    {[System.Windows.Forms.MessageBox]::Show("Для серверов центра работа сервиса не предусмотрена!","Контур")
     return
        }
    elseif($Server -like '*int*')
    {#[System.Windows.Forms.MessageBox]::Show("Выбран интерфейсный сервер","Интерфейс")
    $JOBS_OF_SERVER = Get-ScheduledTask -CimSession $Server -TaskName "FOBO*"
        }
    elseif($Server -like '*a*')
    {#[System.Windows.Forms.MessageBox]::Show("Сервер магазина!","Магазин")
     $JOBS_OF_SERVER = Get-ScheduledTask -CimSession $Server -TaskName "JOB*"
        }
    
#Create Listbox
    $JobList = New-Object System.Windows.Forms.ListBox
    $JobList.Location = New-Object System.Drawing.Size(5,10)
    $JobList.Size = '200,260'
    $JobList.ScrollAlwaysVisible = 'False'
    #$JobList.Font = $FontJob
    $JobList.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    if($Server -like '*vrq*')
    {
    $JobList.DataSource = $JOB_OF_SERVICE.Name
    }
    else
    {
    $JobList.DataSource = $JOBS_OF_SERVER.TaskName
    }
#Label Status
    $JobLabel = New-Object System.Windows.Forms.Label
    $JobLabel.Location = New-Object System.Drawing.Size(210,10)
    $JobLabel.Width = '200'
    $JobLabel.Height = '30'
    $JobLabel.ForeColor = 'Green'
    $JobLabel.BackColor = 'Transparent'
    $JobLabel.Text = '    Состояние задачи:'
    

    $JobStatus = New-Object System.Windows.Forms.Textbox
    $JobStatus.Location = New-Object System.Drawing.Size(210,40)
    $JobStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $JobStatus.ReadOnly = 'True'
    $JobStatus.Size = '150,23'
    $JobStatus.Font = $FontJob
    $JobStatus.Text = '  Состояние задачи'

    $JL_SELECT ={
       $JB = $JobList.SelectedItem
       if($Server -like "*vrq*")
       {
       $JBSTAT = $JOB_OF_SERVICE| Where-Object {$_.Name -eq $JB} | Select-Object State
       }
       else
       {
       $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
       }
       $jobStatus.text = "         "+$JBSTAT.State
       $JobLabel.Text = '      Cостояние задачи:     ' + $JB     
    }
    $JobList.add_SelectedIndexChanged($JL_SELECT)


#Label Actions:
    $JobLabel1 = New-Object System.Windows.Forms.Label
    $JobLabel1.Location = New-Object System.Drawing.Size(210,70)
    $JobLabel1.Width = '200'
    $JobLabel1.Height = '30'
    $JobLabel1.ForeColor = 'Black'
    $JobLabel1.BackColor = 'Transparent'
    $JobLabel1.Text = '  Действия:'


#Create START Button
    $JobStart = New-Object System.Windows.Forms.RadioButton
    $JobStart.Location =  New-Object System.Drawing.Point(210,90)
    $JobStart.AutoSize = 'True'
    $JobStart.BackColor = 'Transparent'
    $JobStart.Text = 'Start'
    $JobStart.Font = $FontJob
#Create STOP Button
    $JobStop = New-Object System.Windows.Forms.RadioButton
    $JobStop.Location =  New-Object System.Drawing.Point(210,110)
    $JobStop.AutoSize = 'True'
    $JobStop.BackColor = 'Transparent'
    $JobStop.Text = 'End'
    $JobStop.Font = $FontJob
#Create Enable Button
    $JobEnable = New-Object System.Windows.Forms.RadioButton
    $JobEnable.Location =  New-Object System.Drawing.Point(210,130)
    $JobEnable.AutoSize = 'True'
    $JobEnable.BackColor = 'Transparent'
    $JobEnable.Text = 'Enable'
    $JobEnable.Font = $FontJob
#Create Disable Button
    $JobDisable
    $JobDisable= New-Object System.Windows.Forms.RadioButton
    $JobDisable.Location =  New-Object System.Drawing.Point(210,150)
    $JobDisable.AutoSize = 'True'
    $JobDisable.BackColor = 'Transparent'
    $JobDisable.Text = 'Disable'
    $JobDisable.Font = $FontJob
#Create Processing Button
    $JobProcessButton = New-Object  System.Windows.Forms.Button
    $JobProcessButton.Location = New-Object System.Drawing.Size(215,180)
    $JobProcessButton.Text = "Process"
    $JobProcessButton.add_Click({
              
           if($JobStart.Checked){
                $JB = $JobList.SelectedItem
                if($Server -like "*vrq-int*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\').GetTask($JB)
                    $StartTask.Run($null)
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }
                }
                elseif($Server -like "*vrq-a*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB)
                    $StartTask.Run($null)
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }  
                }
                else
                {
                $JB = $JobList.SelectedItem
                $path = ( Get-ScheduledTask -CimSession $Server -TaskName $JB).TaskPath
                Start-ScheduledTask -CimSession $Server -TaskName $JB -TaskPath $path
                Start-Sleep -Seconds 3
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                }
                
                $jobStatus.text = "         "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB  
              }
          elseif($JobStop.Checked){
                $JB = $JobList.SelectedItem
                $path = ( Get-ScheduledTask -CimSession $Server -TaskName $JB).TaskPath
                Stop-ScheduledTask -CimSession $Server -TaskName $JB -TaskPath $path
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                $jobStatus.text = "          "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }
          elseif($JobEnable.Checked){
                $JB = $JobList.SelectedItem
                if($Server -like "*vrq-int*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\').GetTask($JB)
                    $StartTask.Enabled = $True
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }
                }
                elseif($Server -like "*vrq-a*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB)
                    $StartTask.Enabled = $True
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }  
                }
                else
                {
                Get-ScheduledTask -CimSession $Server -TaskName $JB | Enable-ScheduledTask
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                }
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }

          elseif($JobDisable.Checked){
                $JB = $JobList.SelectedItem
                if($Server -like "*vrq-int*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\').GetTask($JB) 
                    $StartTask.Enabled = $false
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }
                }
                elseif($Server -like "*vrq-a*")
                {
                    ($TaskScheduler = New-Object -ComObject Schedule.Service).Connect($Server)
                    $StartTask = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB)
                    $StartTask.Enabled = $False
                    Start-Sleep 3
                    $JBSTAT = $TaskScheduler.GetFolder('\NTSwincash').GetTask($JB) | Select Name,@{
                                                                                             Name = 'State'
                                                                                             Expression = {switch ($_.State) {
                                                                                                                    0 {'Unknown'}
                                                                                                                    1 {'Disabled'}
                                                                                                                    2 {'Queued'}
                                                                                                                    3 {'Ready'}
                                                                                                                    4 {'Running'}
                                                                                                                    }
                                                                                                                   }
                                                                                                                  }  
                }
                else
                {
                Get-ScheduledTask -CimSession $Server -TaskName $JB | Disable-ScheduledTask
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                }
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }
    
    })
    
    $JobForm.Controls.AddRange(@($JobStart,$JobStop,$JobEnable,$JobDisable))
    $JobForm.Controls.AddRange(@($JobList,$JobStatus,$JobLabel,$JobLabel1,$JobProcessButton))
    $JobForm.Add_Shown({$JobForm.Activate()})
    $JobForm.ShowDialog()

}
#JOB MANIPULATOR END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#CHECK SERVICES FUNCTION START
Function CheckServices([string]$Server)
{
    $FontCheck = New-Object System.Drawing.Font("Colibri",7,[System.Drawing.FontStyle]::Bold)
    $ImageCheck =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Services.jpg")
    $FontLabelCheck = New-Object System.Drawing.Font("Colibri",11,[System.Drawing.FontStyle]::Bold)
    $FontStatus = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

    #CHECK SERVICES MAIN FORM 
    $CheckForm = New-Object System.Windows.Forms.Form
    $CheckForm.SizeGripStyle = "Hide"
    $CheckForm.BackgroundImage = $ImageCheck
    $CheckForm.BackgroundImageLayout = "None"
    #$CheckForm.Size = New-Object System.Drawing.Size(250,110)
    if($ImageCheck -eq $Null)
    {
     $CheckForm.Size = ('598,188')
    }
    else
    {
    $CheckForm.Width = $ImageCheck.Width
    $CheckForm.Height = $ImageCheck.Height
    }
    $CheckForm.StartPosition = "CenterScreen"
    $CheckForm.TopMost = $True
    $CheckForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $CheckForm.Text = "Монитор контроля сервисов $Server"
    $CheckForm.Icon = $Icon

    $CheckForm.KeyPreview = $True
    $CheckForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {$CheckForm.Close()
    }
    })


    Function Checker_Wild([object]$ServiceWildState){
        if($ServiceWildState.Status -eq 'Running')
        {
            $StatusWild.Text = $ServiceWildState.status
            $StatusWild.BackColor = '#90ee90'
        }
        else
        {
            $StatusWild.Text = $ServiceWildState.status
            $StatusWild.BackColor = 'Red'
        }
    }

    Function Checker_NTS([object]$ServiceNTSState){
        if($ServiceNTSState.Status -eq 'Running')
        {
            $StatusNTS.Text = $ServiceNTSState.status
            $StatusNTS.BackColor = '#90ee90'
        }
        else
        {
            $StatusNTS.Text = $ServiceNTSState.status
            $StatusNTS.BackColor = 'Red'
        }
    }



    $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server 
    $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server 

#TEXTBOX WILDFLY
    $CheckLabelWild = New-Object System.Windows.Forms.TextBox
    $CheckLabelWild.Location = New-Object System.Drawing.Size(5,20)
    $CheckLabelWild.Width  = '170'
    $CheckLabelWild.Height = '25'
    $CheckLabelWild.Font = $FontLabelCheck
    $CheckLabelWild.AutoSize = 'True'
    $CheckLabelWild.Text = "  Сервис : Wildfly"
    $CheckLabelWild.ReadOnly = 'True'
    $CheckLabelWild.BackColor = '#90ee90'
    $CheckLabelWild.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    
#TEXTBOX NTS
    $CheckLabelNTS= New-Object System.Windows.Forms.TextBox
    $CheckLabelNTS.Location = New-Object System.Drawing.Size(5,80)
    $CheckLabelNTS.Font = $FontLabelCheck
    $CheckLabelNTS.Width  = '200'
    $CheckLabelNTS.Height = '25'
    $CheckLabelNTS.AutoSize = 'True'
    $CheckLabelNTS.Text = " Сервис : NTSWincash"
    $CheckLabelNTS.ReadOnly = 'True'
    $CheckLabelNTS.BackColor = 'Orange'
    $CheckLabelNTS.SelectionStart = '0'
    $CheckLabelNTS.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

#Подпись Статуса
    $STLabelW = New-Object System.Windows.Forms.Label
    $STLabelW.Location = New-Object System.Drawing.Point(210,5)
    $STLabelW.AutoSize = 'True'
    $STLabelW.Font = $FontCheck
    $STLabelW.Text = 'Состояние задания:'
    $STLabelW.BackColor = 'Transparent'

#Подпись Статуса 2
    $STLabelN = New-Object System.Windows.Forms.Label
    $STLabelN.Location = New-Object System.Drawing.Point(210,67)
    $STLabelN.AutoSize = 'True'
    $STLabelN.Font = $FontCheck
    $STLabelN.Text = 'Состояние задания:'
    $STLabelN.BackColor = 'Transparent'
            
#Окно текущего статуса WILDFLY
    $StatusWild= New-Object System.Windows.Forms.TextBox
    $StatusWild.Location = New-Object System.Drawing.Size(215,20)
    $StatusWild.Font = $FontStatus
    $StatusWild.Width  = '90'
    $StatusWild.Height = '30'
    $StatusWild.ReadOnly  = 'True'
    $StatusWild.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    Checker_Wild($ServiceWildState)
    
#Окно текущего статуса NTS WINCASH
    $StatusNts = New-Object System.Windows.Forms.TextBox
    $StatusNts.Location = New-Object System.Drawing.Size(215,80)
    $StatusNts.Font = $FontStatus
    $StatusNts.Width  = '90'
    $StatusNts.Height = '30'
    $StatusNts.ReadOnly  = 'True'
    $StatusNts.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    Checker_NTS($ServiceNTSState)

#Create ToolTip
    $ToolTipService = New-Object System.Windows.Forms.ToolTip
    $ToolTipService.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    $ToolTipService.SetToolTip($StatusWild,"Click For Update Status")
    $ToolTipService.SetToolTip($StatusNts,"Click For Update Status")
    
#Start Service Wildfly
    $StartWildBtn = New-Object System.Windows.Forms.Button
    $StartWildBtn.Location = New-Object System.Drawing.Size(5,46)
    $StartWildBtn.Size = New-Object System.Drawing.Size(50,20)
    $StartWildBtn.Text = "START"
    $StartWildBtn.Font = $FontCheck
    $StartWildBtn.ForeColor = 'green'

#Restart Service Wildfly
    $RestartWildBtn = New-Object System.Windows.Forms.Button
    $RestartWildBtn.Location = New-Object System.Drawing.Size(57,46)
    $RestartWildBtn.Size = New-Object System.Drawing.Size(65,20)
    $RestartWildBtn.Text = "RESTART"
    $RestartWildBtn.Font = $FontCheck
    #$RestartWildBtn.AutoSize = 'True'
    $RestartWildBtn.ForeColor = 'Blue'

#Stop Service Wildfly
    $StopWildBtn = New-Object System.Windows.Forms.Button
    $StopWildBtn.Location = New-Object System.Drawing.Size(125,46)
    $StopWildBtn.Size = New-Object System.Drawing.Size(50,20)
    $StopWildBtn.Text = "STOP"
    $StopWildBtn.Font = $FontCheck
    $StopWildBtn.ForeColor = 'Red'


#EVENT WILDFLY BTN
    $StartWildBtn.add_Click({
        $ServiceWildState =  Get-Service -Name Wildfly -ComputerName $Server | Start-Service
    })
    $RestartWildBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='java.exe'").terminate() | Out-Null
        $ServiceWildState = Get-Service -Name Wildfly -ComputerName $Server | Restart-Service
    })
    $StopWildBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='java.exe'").terminate() | Out-Null
        $ServiceWildState = Get-Service -Name Wildfly -ComputerName $Server | Stop-Service
    })
            
#Start Service NTS
    $StartNTSBtn = New-Object System.Windows.Forms.Button
    $StartNTSBtn.Location = New-Object System.Drawing.Size(5,106)
    $StartNTSBtn.Size = New-Object System.Drawing.Size(50,20)
    $StartNTSBtn.Text = "START"
    $StartNTSBtn.Font = $FontCheck
    $StartNTSBtn.ForeColor = 'Green'

#Restart Service NTS
    $RestartNTSBtn = New-Object System.Windows.Forms.Button
    $RestartNTSBtn.Location = New-Object System.Drawing.Size(58,106)
    $RestartNTSBtn.Size = New-Object System.Drawing.Size(65,20)
    $RestartNTSBtn.Text = "RESTART"
    $RestartNTSBtn.Font = $FontCheck
    $RestartNTSBtn.ForeColor = 'Blue'

#Stop Service NTS
    $StopNTSBtn = New-Object System.Windows.Forms.Button
    $StopNTSBtn.Location = New-Object System.Drawing.Size(125,106)
    $StopNTSBtn.Size = New-Object System.Drawing.Size(50,20)
    $StopNTSBtn.Text = "STOP"
    $StopNTSBtn.Font = $FontCheck
    $StopNTSBtn.ForeColor = 'Red'
    
#EVENT NTS BTN
    $StartNTSBtn.add_Click({
        $ServiceNTSState =  Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Start-Service
    })
    $RestartNTSBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='javaw.exe'").terminate() | Out-Null
        $ServiceNTSState = Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Restart-Service
    })
    $StopNTSBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='javaw.exe'").terminate() | Out-Null
        $ServiceNTSState = Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Stop-Service

    })

#EVENTS FOR UPDATE STATUS
        $StatusWild.add_Click({
        $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server 
        Checker_Wild($ServiceWildState)
    })

    $StatusNts.add_Click({
        $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server 
        Checker_NTS($ServiceNTSState)
    })

    $CheckForm.Controls.AddRange(@($CheckLabelNTS,$CheckLabelWild,$StatusWild,$StatusNts,$STLabelW,$STLabelN))
    $CheckForm.Controls.AddRange(@($StartWildBtn,$RestartWildBtn,$StopWildBtn))
    $CheckForm.Controls.AddRange(@($StartNTSbtn,$RestartNTSBtn,$StopNTSBtn))
    #$CheckForm.Controls.Add($pbrTest)
    $CheckForm.Add_Shown({$CheckForm.Activate()})
    $CheckForm.ShowDialog()
    $ServiceNTSState = ''
    $ServiceWildState = ''
}
#CHECK SERVICES FUNCTION END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
Function CHECK_SETTINGS(){
    #Проверка Среды
    if($RadioVRX.Checked)
    {
      $SRED = 'vrx'  
    }
    elseif ($RadioVRQ.Checked)
    {
      $SRED = 'vrq'
    }
    else
    {
     [System.Windows.Forms.MessageBox]::Show("НЕ ВЫБРАН КОНТУР!","ВЫБЕРИТЕ КОНТУР",'OK','ERROR')
     return
    }

    #Проверка контура
    if ($RadioContur.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("ajb","Контур",'OK','Info')
     $CONT = "ajb"
     $MACHINE = $Combo_Srez.SelectedItem
    }
    elseif ($RadioMAG.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("a","МАГАЗИН")
     $CONT = "a"
     $MACHINE = $TextBox.Text
    }
    elseif ($RadioINT.Checked = $true)
    {
     #[System.Windows.Forms.MessageBox]::Show("int","ИНТЕРФЕЙС")
     $CONT = "int"
     $MACHINE = $Combo_Srez.SelectedItem
    }



    #Проверка ввода.
    if($MACHINE -eq 'Введите магазин' -or $Machine -eq '' -or ($MACHINE.Length -lt 3 -and $CONT -like "a"))
    {
      [System.Windows.Forms.MessageBox]::Show('Обнаружена ошибка выбора станции. Повторите ввод!',"Ошибка выбора",'RetryCancel','ERROR')
      return
    }
    else
    {
     $Global:SERVER = 'fobo-'+ $SRED + "-" + $CONT + $MACHINE
     $Answer = [System.Windows.Forms.MessageBox]::Show("Выбрана машина: " + $SERVER + ".
Подтверждаем выбор?","Выбор сделан",'YesNo','WARNING')     
    }

 return $Answer
}
##########################################################################################################################################################################################################
Function CHECK_SETTINGS_NO_WINDOW(){
    #Проверка Среды
    if($RadioVRX.Checked)
    {
      $SRED = 'vrx'  
    }
    elseif ($RadioVRQ.Checked)
    {
      $SRED = 'vrq'
    }
    else
    {
     [System.Windows.Forms.MessageBox]::Show("НЕ ВЫБРАН КОНТУР!","ВЫБЕРИТЕ КОНТУР",'OK','ERROR')
     return
    }

    #Проверка контура
    if ($RadioContur.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("ajb","Контур",'OK','Info')
     $CONT = "ajb"
     $MACHINE = $Combo_Srez.SelectedItem
    }
    elseif ($RadioMAG.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("a","МАГАЗИН")
     $CONT = "a"
     $MACHINE = $TextBox.Text
    }
    elseif ($RadioINT.Checked = $true)
    {
     #[System.Windows.Forms.MessageBox]::Show("int","ИНТЕРФЕЙС")
     $CONT = "int"
     $MACHINE = $Combo_Srez.SelectedItem
    }



    #Проверка ввода.
    if($MACHINE -eq 'Введите магазин' -or $Machine -eq '')
    {
      [System.Windows.Forms.MessageBox]::Show('Обнаружена ошибка выбора станции. Повторите ввод!',"Ошибка выбора",'RetryCancel','ERROR')
      return
    }
    else
    {
     $Global:SERVER = 'fobo-'+ $SRED + "-" + $CONT + $MACHINE
    }

 return $SERVER
}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#ReDeploy WARNIKA
Function REDEPLOY([string]$Server){
  
  $WinParamSuccess = @{
        Content = "Переустановка сервиса выполнена успешно!"
        Title = "ReDeploy"
        TitleFontSize = "16"
        TitleTextForeground = "Green"
        ContentBackground = "SpringGreen"

                    }


  $WinParamFail = @{
        Content = "Переустановка сервиса не выполнена!"
        Title = "ReDeploy"
        TitleFontSize = "16"
        TitleTextForeground = "Salmon"
        ContentFontWeight = "Bold"
        ContentBackground = "Salmon"
            }




  $ChoiceD = [System.Windows.Forms.MessageBox]::Show("YES: Выполнить передеплой существующего сервиса.
NO: Выбрать файл и выполнить передеплой.
Cancel: Выход","Выбор действия!","YesNoCancel")
  switch($ChoiceD)
  {
    "YES" {
            if($Server -like '*int*')
            {
                $DestinationPoint = "\\" + $Server + "\C`$\wildfly\wildfly10\standalone\deployments\"

            #Вызов диалога выбора файла с заданными параметрами
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = $DestinationPoint
                Filter = 'Deploy (*.war)|*.war;*.failed|Все файлы |*.*'
                Title = 'Выберите файл сервиса для деплоя'}
                $FileBrowser.ShowDialog()

            #Формирование имени файла.
                $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
                $TST = $FileBrowser.SafeFileName

            #Отработка тестов выбора файла и его наличия.
                if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}

            #Открытие директории и начало выполнения передеплоя файла.
                Invoke-Item $DestinationPoint
                Get-ChildItem -Path  "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.failed" | Remove-Item
                $DeployedServiceName = "$DestinationFileName.deployed"
                $DeployedServiceName = $DeployedServiceName -replace '\s',''
                if(Test-Path "$DestinationFileName.deployed")
                {
                    Rename-Item "$DestinationFileName.deployed" -NewName "$DestinationFileName.undeploy"
                    Start-Sleep 5
                    #Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.undeploy" | Remove-Item
                    Rename-Item "$DestinationFileName.undeploy" -NewName "$DestinationFileName.dodeploy"
                }
                else
                {
                    New-Item -ItemType File "$DestinationFileName.dodeploy"  
                }
                Start-Sleep 5
                $FirstCheck = Test-Path  "$DestinationFileName.deployed"
                $SecondCheck = Test-Path  "$DestinationFileName.isdeploying"
                #$FirstCheck  = Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.isdeploying"
                #$SecondCheck = Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.dodeployed"
                if($FirstCheck -eq $True -or $SecondCheck -eq $True)
                {
                    #$Result = [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса выполнена успешно","REDEPLOY","OK","INFO")
                    $objForm.TopMost = $False
                    New-WPFMessageBox @WinParamSuccess
                    $objForm.TopMost = $True
                    #$Result
                }
                else
                {
                    #$Result = [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса не выполнена!","REDEPLOY","OK","WARNING")
                    #$Result
                    $objForm.TopMost = $False
                    New-WPFMessageBox @WinParamFail
                    $objForm.TopMost = $True
                }
            }

            else
            {
            #Формирование пути к серверу и файлу.
                $DestinationPoint = "\\" + $Server + "\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
            #Вызов диалога выбора файла с заданными параметрами
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = $DestinationPoint
                Filter = 'Deploy (*.war)|*.war;*.failed|Все файлы |*.*'
                Title = 'Выберите файл сервиса для деплоя'}
                $FileBrowser.ShowDialog()

            #Формирование имени файла.
                $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
                $TST = $FileBrowser.SafeFileName
                if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}
            #Выполнение передеплоя  
                Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
                Get-ChildItem -Path  "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.failed" | Remove-Item

            #Проверка наличия установленного сервиса
                if(Test-Path "$DestinationFileName.deployed")
                {
                    Rename-Item "$DestinationFileName.deployed" -NewName "$DestinationFileName.undeploy"
                    Start-Sleep 5
                    #Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.undeploy" | Remove-Item
                    Rename-Item "$DestinationFileName.undeploy" -NewName "$DestinationFileName.dodeploy"
                }
                else
                {
                    New-Item -ItemType File "$DestinationFileName.dodeploy"  
                }

                Start-Sleep 10
                $FirstCheck = Test-Path  "$DestinationFileName.deployed"
                $SecondCheck = Test-Path  "$DestinationFileName.isdeploying"
                #$FirstCheck  = Get-ChildItem -Path "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST.isdeploying"
                #$SecondCheck = Get-ChildItem -Path "\\$Server\C$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST.deployed"
                if($FirstCheck -eq $True -or $SecondCheck -eq $True)
                {
                    #$Result = [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса выполнена успешно","REDEPLOY","OKCancel","INFO")
                    #$Result
                    $objForm.TopMost = $False
                    New-WPFMessageBox @WinParamSuccess
                    $objForm.TopMost = $True
                }
                else
                {
                    #$Result = [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса не выполнена!","REDEPLOY","OKCancel","WARNING")
                    #$Result
                    $objForm.TopMost = $False
                    New-WPFMessageBox @WinParamFail
                    $objForm.TopMost = $True
                }

            }
            
    
    }
    "NO"{
            if($Server -like '*int*')
            { [System.Windows.Forms.MessageBox]::Show("Данная функция для Интерфейсных серверов в Разработке!","В РАЗРАБОТКЕ",'OK','ERROR');return}
            Add-Type -AssemblyName System.Windows.Forms
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Filter = 'Deploy (*.war)|*.war|Все файлы |*.*'
            Title = 'Выберите файл сервиса для деплоя'}
            $FileBrowser.ShowDialog()

            $DestinationPoint = "\\" + $Server + "\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"

            $PathTest = Test-Path $DestinationPoint

            
            $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
            $TST = $FileBrowser.SafeFileName
            if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}
            $FileTestName = $FileBrowser.SafeFileName -split ".war"
            $ProvFile = [System.Windows.Forms.MessageBox]::Show("Будет выполнен деплой файла " + $TST+ ". На сервер :" + $Server,"Путь деплоя","OKCancel",'Info')

            switch($ProvFile)
            {
             "Cancel"{[System.Windows.Forms.MessageBox]::Show("Отмена Деплоя!");return}
            }

            $FileTest = Test-Path $DestinationFileName
  

            if($PathTest -eq $False)
            {
                [System.Windows.MessageBox]::Show("Путь не существует, или недоступен,проверьте выбор сервера")
                return
            }
            elseif($FileTest -eq $False)
            {
                [System.Windows.MessageBox]::Show("Сервиса ранее не было на сервере, проверьте выбор файла")
                return
            }
            elseif($PathTest -eq $True -and $FileTest -eq $True)
            {

                Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
                Get-ChildItem -Path  "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.deployed","$TST*.failed" | Remove-Item
                Rename-Item $DestinationFileName -NewName "$DestinationFileName.backup"
                Copy-Item -Path $FileBrowser.FileName -Destination $DestinationFileName
                Start-Sleep 13
                Get-ChildItem -Path "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.undeployed" | Remove-Item         
                $Hash1 = Get-FileHash $FileBrowser.FileName
                $Hash2 = Get-FileHash $DestinationFileName

                if ($Hash1.Hash -eq $Hash2.Hash -and $Hash1.Hash -ne $NULL -and $Hash2.Hash -ne $NULL)
                    {
                    [System.Windows.MessageBox]::Show("Файл успешно перенесен $DestinationFileName","Перенос файла успешен")
                    }
                    else
                    {
                    $CHECKHASH = [System.Windows.MessageBox]::Show("Файл перенесен в $DestinationFileName с ошибками
                    Будет выполнено восстановление файла!","Перенос файла провалился","OK",'ERROR')
                    Switch($CHECKHASH){
                    "OK"{
                    Rename-Item $DestinationFileName -NewName "$DestinationFileName.backup.FAILED"
                    Rename-Item "$DestinationFileName.backup" -NewName "$DestinationFileName"}
                    }
        
            }
         }
         #Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
    }

    "CANCEL"{
    return}


  }

}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#RESTART WILDFLY DELETE FILES
Function KillWildfly([string]$SRV)
{
    if($Server -like '*int*')
    {
     #[System.Windows.Forms.MessageBox]::Show("ИНТЕРФЕЙС")
     (Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'").terminate() | Out-Null    
     Get-Service -Name Wildfly -ComputerName $SRV -ErrorAction SilentlyContinue | Stop-Service
     Start-Sleep -Seconds 3
     Get-Service -Name Wildfly -ComputerName $server | Start-Service
     Start-Sleep -Seconds 2
     Invoke-Item "\\$Server\C`$\wildfly\wildfly10\standalone\deployments" 
    }
    else{
    #Get-Process -Name java -ComputerName $SRV -ErrorAction SilentlyContinue | Format-List
    (Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'").terminate() | Out-Null    
    Get-Service -Name Wildfly -ComputerName $SRV -ErrorAction SilentlyContinue | Stop-Service
    Start-Sleep -Seconds 3
    if($RLS -eq '19'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.failed" | Remove-Item
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Exclude "vfs" | Remove-Item -Recurse 
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '20'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Exclude "vfs" | Remove-Item 
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '21'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\"  -Exclude "vfs" | Remove-Item 
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '22'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\"  -Exclude "vfs" | Remove-Item 
        #Remove-Item -Path "\\$SRV\C$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 5
    #Progress
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    Start-Sleep -Seconds 2
    Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
    }
}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#PROGRESS BAR FUNCTION START
Function Progress(){

$ProcessForm = New-Object System.Windows.Forms.Form
$ProcessForm.SizeGripStyle = "Hide"
$ProcessForm.BackgroundImage = $ImageRelease
$ProcessForm.BackgroundImageLayout = "None"
$ProcessForm.Size = New-Object System.Drawing.Size(250,110)
$ProcessForm
$ProcessForm.StartPosition = "CenterScreen"
$ProcessForm.TopMost = $true
$ProcessForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$ProcessForm.Text = "Процесс выполнения задания"

$pbrTest = New-Object System.Windows.Forms.ProgressBar
$pbrTest.Maximum = 250
$pbrTest.Minimum = 0
$pbrTest.Location = new-object System.Drawing.Size(10,10)
$pbrTest.size = new-object System.Drawing.Size(200,50)
$pbrTest.Name = 'Выполнение перезапуска службы'

Function StartProgressBar{
   $i = 0
        While ($i -le 250) {
        $pbrTest.Value = $i
        Start-Sleep -m 30
        "VALLUE EQ"
        $i
        $i += 1
        $ProcessForm.Refresh()
    }
    $ProcessForm.Close()     
}
$pbrTest.Add_MouseEnter({StartProgressBar})

$ProcessForm.Controls.Add($pbrTest)
$ProcessForm.Add_Shown({$ProcessForm.Activate()})
$ProcessForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
$ProcessForm.ShowDialog()
$ProcessForm.Focused
$ProcessForm.Refresh()

}
#RELEASE CHOICE FUNCTION END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#RELEASE CHOICE FUNCTION START 
Function RELEASE_WINDOW(){

$FontRelease = New-Object System.Drawing.Font("Eras Bold ITC",10,[System.Drawing.FontStyle]::Bold)
$ImageRelease =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\REL.jpg")
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

#Initialize Release Choice FORM
$ReleaseForm = New-Object System.Windows.Forms.Form
$ReleaseForm.SizeGripStyle = "Hide"
$ReleaseForm.BackgroundImage = $ImageRelease
$ReleaseForm.BackgroundImageLayout = "None"
$ReleaseForm.Size = New-Object System.Drawing.Size(150,202)
$ReleaseForm.StartPosition = "CenterScreen"
$ReleaseForm.TopMost = $True
$ReleaseForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$ReleaseForm.Text = "ВЫБЕРИТЕ РЕЛИЗ"
$ReleaseForm.Icon  = $Icon
$ReleaseForm.ControlBox = $False



  $WinParamFail = @{
        Content = "Отмена операции перезагрузки сервисов!"
        Title = "Decline Process"
        TitleFontSize = "16"
        TitleTextForeground = "Salmon"
        ContentFontWeight = "Bold"
        ContentBackground = "Salmon"
            }



#Release Button 0
$ReleaseButton0 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton0.Location = New-Object System.Drawing.Size(10,15)
$ReleaseButton0.Text = 'Релиз 19.0.0'
$ReleaseButton0.AutoSize = 'True'
$ReleaseButton0.Font = $FontRelease
$ReleaseButton0.Backcolor = 'Transparent'
$ReleaseButton0.Checked = $True
#Release Button 1
$ReleaseButton1 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton1.Location = New-Object System.Drawing.Size(10,35)
$ReleaseButton1.Text = 'Релиз 20.0.0'
$ReleaseButton1.AutoSize = 'True'
$ReleaseButton1.Font = $FontRelease
$ReleaseButton1.Backcolor = 'Transparent'
$ReleaseButton1.Checked = $False
#Release Button 2
$ReleaseButton2 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton2.Location = New-Object System.Drawing.Size(10,55)
$ReleaseButton2.Text = 'Релиз 21.0.0'
$ReleaseButton2.AutoSize = 'True'
$ReleaseButton2.Font = $FontRelease
$ReleaseButton2.Backcolor = 'Transparent'
$ReleaseButton2.Checked = $False
#Release Button 3
$ReleaseButton3 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton3.Location = New-Object System.Drawing.Size(10,75)
$ReleaseButton3.Text = 'Релиз 22.0.0'
$ReleaseButton3.AutoSize = 'True'
$ReleaseButton3.Font = $FontRelease
$ReleaseButton3.Backcolor = 'Transparent'
$ReleaseButton3.Checked = $False
#Burron Accept
$ButtonAccept =  New-Object System.Windows.Forms.Button
$ButtonAccept.Location = New-Object System.Drawing.Size(-3,100)
#$ButtonAccept.Size = New-Object System.Drawing.Size(75,23)
$ButtonAccept.Text = "Подтвердить"
$ButtonAccept.Width = $ReleaseForm.Width
$ButtonAccept.Height = '33'

$ButtonDecline = New-Object System.Windows.Forms.Button 
$ButtonDecline.Location = New-Object System.Drawing.Size(-3,130)
$ButtonDecline.Text = "Отменить"
$ButtonDecline.Width = $ReleaseForm.Width
$ButtonDecline.Height = '33'

if ($ReleaseButton0.Checked -eq $true) {$Global:RLS = 19} 
if ($ReleaseButton1.Checked -eq $true) {$Global:RLS = 20}
if ($ReleaseButton2.Checked -eq $true) {$Global:RLS = 21}
if ($ReleaseButton3.Checked -eq $true) {$Global:RLS = 22}

function BTN_CLICK()
{
  $Global:QA="1"
  if ($ReleaseButton0.Checked -eq $true) {$Global:RLS = 19} 
  if ($ReleaseButton1.Checked -eq $true) {$Global:RLS = 20}
  if ($ReleaseButton2.Checked -eq $true) {$Global:RLS = 21}
  if ($ReleaseButton3.Checked -eq $true) {$Global:RLS = 22}

}
$ButtonAccept.Add_Click(
{
BTN_CLICK;$ReleaseForm.Close()
})

$ButtonDecline.Add_Click({
                    
                    $Global:QA="2"
                    $objForm.TopMost = $False
                    New-WPFMessageBox @WinParamFail
                    $objForm.TopMost = $True
                    $ReleaseForm.Close()
                    return $QA

})
$ReleaseForm.Controls.AddRange(@($ButtonAccept,$ButtonDecline))
$ReleaseForm.Add_Shown({$ReleaseForm.Activate()})
$ReleaseForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
$ReleaseForm.ShowDialog()
}
#RELEASE CHOICE FUNCTION END
##########################################################################################################################################################################################################
#
#START CLEAR TEMPOROS
Workflow CLEAR_TEMP([string]$Server){
    Get-ChildItem "\\$Server\C$\Windows\Temp\*" | Remove-Item -ErrorAction SilentlyContinue -Force -Recurse
}
#END CLEAR TEMPOROS
#
##########################################################################################################################################################################################################
##########################################################################################################################################################################################################
#MAIN FUNCTION FOR ALL PROGRAMM
#CONTAINS MAIN FORM AND BUTTONS FOR START ALL UPPER FUNCTIONS
#
function GENERATOR{
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$VRX = ('1','2','3','4','5','6')
$VRQ = ('1','2','3','4','5','6','7')
# Create base form.


function ENTERCOLOR($ELEMENT){
		$ELEMENT.BackColor = 'LightGreen'
	}
function LEAVECOLOR($ELEMENT){
        $ELEMENT.BackColor = 'Control'
    }



$Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\NTS.jpg")
$Font = New-Object System.Drawing.Font("Comic Sans MS",8,[System.Drawing.FontStyle]::Bold)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Regular)

# Initialize Main Form #
$objForm = New-Object System.Windows.Forms.Form 
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$objForm.SizeGripStyle = "Hide"
$objForm.BackgroundImage = $Image
$objForm.BackgroundImageLayout = "None"
$objForm.Text = "Программа для безумного управления сервисами V1.9"
$objForm.StartPosition = "CenterScreen"
$objForm.Height = '370'
    if($Image -eq $null){
        $objForm.Width = '580'}
    else{
        $objForm.Width = $Image.Width 
        }
$objForm.Icon = $Icon

# Configure keyboard intercepts for ESC & ENTER.

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
        $objForm.Close()
    }
})
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {
        $objForm.Close()
    }
})

#BANKSY LABEL
$FOKINLAB = New-Object System.Windows.Forms.Label
$FOKINLAB.Location = ('455,280')
$FOKINLAB.Text = "Created By Fokin"
$FOKINLAB.Font = $FontBanksy
#$FOKINLAB.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$FOKINLAB.BackColor =  'Transparent'
$FOKINLAB.AutoSize = $True

#GROUP BOX SRED
$MyGroupBox = New-Object System.Windows.Forms.GroupBox
$MyGroupBox.Location = '5,180'
$MyGroupBox.size = '120,80'
$MyGroupBox.Font = $Font
$MyGroupBox.text = "ВЫБОР СРЕДЫ:"
$MyGroupBox.Backcolor = 'Transparent'
#$objForm.Controls.Add($MyGroupBox)

#RADIO VRX
$RadioVRX = New-Object System.Windows.Forms.RadioButton
$RadioVRX.Location = New-Object System.Drawing.Size(10,19)
$RadioVRX.Checked = $True
$RadioVRX.Text = "VRX"

#RADIO VRQ
$RadioVRQ = $RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioVRQ.Location = New-Object System.Drawing.Size(10,39)
$RadioVRQ.Text = "VRQ"
$RadioVRQ.Checked = $False
# ADD GROUP BOX ON FORM
$objForm.Controls.AddRange(@($MyGroupBox))

#TEST LABEL
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,122)
$objLabel.Font = $Font
$objLabel.AutoSize = 'True'
$objLabel.BackColor = 'Transparent'
$objLabel.Text = "!!!!!!!!!!!!!!"
$objLabel.Visible = 'TRUE'
#$objForm.Controls.Add($objLabel) 

#TEXTBOX
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Size(10,30)
$TextBox.Visible = 'False'
$TextBox.Size = '120,80'
$TextBox.MaxLength = '3'


#GROUP BOX TRIGGERS
$MyGroupBox2 = New-Object System.Windows.Forms.GroupBox
$MyGroupBox2.Location = '130,180'
$MyGroupBox2.size = '115,80'
$MyGroupBox2.Font = $Font
$MyGroupBox2.text = "ТРИГГЕРЫ:"
$MyGroupBox2.Backcolor = 'Transparent'

#КОНТУР
$RadioContur = New-Object System.Windows.Forms.RadioButton
$RadioContur.Location = New-Object System.Drawing.Size(10,14)
$RadioContur.Text = "Контур"
$RadioContur.BackColor = 'Transparent'
$RadioContur.Checked = 'True'
#МАГАЗИНЫ
$RadioMAG = New-Object System.Windows.Forms.RadioButton
$RadioMAG.Location = New-Object System.Drawing.Size(10,34)
$RadioMAG.Text = "Магазины"
$RadioMAG.BackColor = 'Transparent'
#ИНТЕРФЕЙСЫ
$RadioINT = New-Object System.Windows.Forms.RadioButton
$RadioINT.Location = New-Object System.Drawing.Size(10,54)
$RadioINT.Text = "Интерфейс"
#ADD SECOND GROUP BOX
$objForm.Controls.Add($MyGroupBox2)
#EVENT FOR CHECK GROUP BOX 2 RADIO BUTTONS
$eventMAG = {
             if($RadioMAG.Checked){
             $TextBox.Text = 'Введите магазин'
             $Combo_Srez.Visible = $False
             $TextBox.Visible = $True
             }
             elseif($RadioContur.Checked){
             $Combo_Srez.Visible = $True
             $TextBox.Text = ''
             $TextBox.Visible = $False
             }
             elseif($RadioINT.Checked){
             $Combo_Srez.Visible = $True
             $TextBox.Text = ''
             $TextBox.Visible = $False
             }
            }
#EVENT FOR TEXT BOX MAGAZINES
$eventBOX = {
             $TextBox.Text = ''
             }
$TextBox.Add_DoubleClick($eventBOX)
$RadioMAG.Add_Click($eventMAG)
$RadioContur.Add_Click($eventMAG)
$RadioINT.Add_Click($eventMAG)



$CheckTop = New-Object System.Windows.Forms.CheckBox
$CheckTop.Location = ('10,5')
$CheckTop.Text = "Window ON Top"
$CheckTop.Checked = $True
$CheckTop.AutoSize = $True
$CheckTop.BackColor = 'Transparent'
$CheckTop.Font = $Font
$CheckTop.add_Click({
    if($CheckTop.Checked -eq $True)
    {
        $objForm.TopMost = $True
    }
    elseif($CheckTop.Checked -eq $False)
    {
        $objForm.TopMost = $False
    }
    })

#GROUP BLOCK CHOICE
$MyGroupBox3 = New-Object System.Windows.Forms.GroupBox
$MyGroupBox3.Location = '250,180'
$MyGroupBox3.size = '140,80'
$MyGroupBox3.Font = $Font
$MyGroupBox3.text = "ВЫБОР СТАНЦИИ:"
$MyGroupBox3.Backcolor = 'Transparent'


#COMBO
$Combo_Srez = New-Object System.Windows.Forms.ComboBox
$Combo_Srez.AutoSize = 'True'
$Combo_Srez.Location = New-Object System.Drawing.Size(10,30)
$Combo_Srez.Text = 'Выберите станцию'
$Combo_Srez.DropDownStyle = 'DropDownList'
            if($RadioVRQ.Checked)
            {
             $Combo_Srez.DataSource = $VRQ}
            elseif ($RadioVRX.Checked){
             $Combo_Srez.DataSource = $VRX}
            else{
             $Combo_Srez.DataSource = $VRX}
$MyGroupBox.Controls.AddRange(@($RadioVRX,$RadioVRQ))
$Combo_Srez.Text
$eventSRED = {
            #$Combo_Srez.Items.Clear()
            if($RadioVRQ.Checked)
            {
             $Combo_Srez.DataSource = $VRQ}
            elseif ($RadioVRX.Checked){
             $Combo_Srez.DataSource = $VRX}
            else{
             $Combo_Srez.DataSource = $VRX}
            }


$objForm.Controls.Add($MyGroupBox3)
$MyGroupBox3.Controls.AddRange(@($Combo_Srez,$TextBox)) 
$RadioVRQ.Add_Click($eventSRED)
$RadioVRX.Add_Click($eventSRED)
$Combo_Srez.add_SelectedIndexChanged($eventSRED)
#$Combo_Srez.Add_Click($eventSRED)
$MyGroupBox2.Controls.AddRange(@($RadioContur,$RadioMAG,$RadioINT)) 


# Create BUTTON FOR START REDEPLOY WILDFLY
$RestartButton = New-Object System.Windows.Forms.Button
$RestartButton.Location = New-Object System.Drawing.Size(5,275)
$RestartButton.Size = New-Object System.Drawing.Size(125,24)
$RestartButton.Text = "RESTART WILDFLY"
$RestartButton.Font = $Font
#$RestartButton.AutoSize = 'True'




#Обработка ВЫБОРА + RESTART SERVICES
$RestartButton.Add_Click(
{
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{ 
                
               if($Server -like '*int*')
               { 
                KillWildfly($SERVER) 
               }
               else{

                    RELEASE_WINDOW
                        if($QA -eq "2")
                        { 
                        return 
                        }
               elseif($QA -eq "1")
               {
               KillWildfly($SERVER)
               }
               }
               $Server = ''
               $RLS = ''
               $Global:QA = ''
             }
        "NO"{ return }
        }    
})


$DeployWAR = New-Object System.Windows.Forms.Button
$DeployWAR.Location = New-Object System.Drawing.Size(130,275)
$DeployWAR.Size = New-Object System.Drawing.Size(115,24)
$DeployWAR.Text = "DEPLOY "".WAR"""
$DeployWAR.Font = $Font
#$DeployWAR.AutoSize = 'True'

$DeployWAR.Add_Click({
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{
               #RELEASE_WINDOW
               REDEPLOY($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    

})

$CheckServicesBTN = New-Object System.Windows.Forms.Button
$CheckServicesBTN.Location = New-Object System.Drawing.Size(5,300)
$CheckServicesBTN.Size = New-Object System.Drawing.Size(145,24)
$CheckServicesBTN.Text = "ПРОВЕРКА СЛУЖБ"
$CheckServicesBTN.Font = $Font
#$CheckServicesBTN.AutoSize = 'True'

$CheckServicesBTN.add_Click({
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{
               CheckServices($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    
            
})


$JobButton = New-Object System.Windows.Forms.Button
$JobButton.Location = New-Object System.Drawing.Size(150,300)
$JobButton.Size = New-Object System.Drawing.Size(75,24)
$JobButton.Font = $Font
$JobButton.Text = "JOB'S"

$JobButton.add_Click({
            
        $Answer = CHECK_SETTINGS
        switch($Answer){
        "YES"{ 
               JOB_WORKER($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    

})


#FOBO INSTALL
$FoboButton = New-Object System.Windows.Forms.Button
$FoboButton.Location = New-Object System.Drawing.Size(225,300)
$FoboButton.Size = New-Object System.Drawing.Size(140,24)
$FoboButton.Font = $Font
$FoboButton.Text = "FOBO INSTALL V1.0"

$FoboButton.Add_Click({

    $Server = CHECK_SETTINGS_NO_WINDOW
               if($Server -like '*int*' -or $Server -like '*ajb*')
               { [System.Windows.Forms.MessageBox]::Show("Нельзя использовать в качестве коннекторов сервера типа INT и AJB")
                 return
               }
               else{
               $DBLINKS = [System.Windows.Forms.MessageBox]::Show("Будут использованы DBLINKS сервера: " + $Server,'DBLINKS','YesNo','INFO')
               switch($DBLINKS)
               {
                "NO"{return}
                "YES"{ FOBO_INSTALL($Server)
               $Server = ''}
               }
               }  
        

})

#Кнопка очистки Темп файлов на Магазина WIP!
$TEMPBTN = New-Object System.Windows.Forms.Button
$TEMPBTN.Location = New-Object System.Drawing.Size(395,185)
$TEMPBTN.Size = New-Object System.Drawing.Size(120,23)
$TEMPBTN.Font = $Font
$TEMPBTN.Text = "CLEAR TEMP"

$TEMPBTN.add_Click({
    $Server = CHECK_SETTINGS_NO_WINDOW
    if($Server -like "fobo*")
    {
      CLEAR_TEMP($Server)
    }
    else
    {
      return
    }

})


#Кнопка открыть шару сервера.
$OpenFLDR = New-Object System.Windows.Forms.Button
$OpenFLDR.Location = New-Object System.Drawing.Size(395,215)
$OpenFLDR.Size = New-Object System.Drawing.Size(120,23)
$OpenFLDR.Font = $Font
$OpenFLDR.Text = "OPEN FOLDER"

$OpenFLDR.add_Click({
    $Server = CHECK_SETTINGS_NO_WINDOW
        if($Server -like "fobo*")
        {
        Invoke-Item "\\$Server\C$\"
        }
        else
        {
            return
        }
    
})


# Cancel EXIT Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(440,300)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Закрыть программу"
$CancelButton.Font = $Font
$CancelButton.AutoSize = 'True'
$CancelButton.Add_Click({$objForm.Close()})


#COLOR BUTTON SELECTION BLOCK {
$CancelButton.add_MouseHover({ENTERCOLOR($CancelButton)})
$CancelButton.add_MouseLeave({LEAVECOLOR($CancelButton)})

$RestartButton.add_MouseHover({ENTERCOLOR($RestartButton)})
$RestartButton.add_MouseLeave({LEAVECOLOR($RestartButton)})

$OpenFLDR.add_MouseHover({ENTERCOLOR($OpenFLDR)})
$OpenFLDR.add_MouseLeave({LEAVECOLOR($OpenFLDR)})

$TEMPBTN.add_MouseHover({ENTERCOLOR($TEMPBTN)})
$TEMPBTN.add_MouseLeave({LEAVECOLOR($TEMPBTN)})

$FoboButton.add_MouseHover({ENTERCOLOR($FoboButton)})
$FoboButton.add_MouseLeave({LEAVECOLOR($FoboButton)})

$JobButton.add_MouseHover({ENTERCOLOR($JobButton)})
$JobButton.add_MouseLeave({LEAVECOLOR($JobButton)})

$CheckServicesBTN.add_MouseHover({ENTERCOLOR($CheckServicesBTN)})
$CheckServicesBTN.add_MouseLeave({LEAVECOLOR($CheckServicesBTN)})

$DeployWAR.add_MouseHover({ENTERCOLOR($DeployWAR)})
$DeployWAR.add_MouseLeave({LEAVECOLOR($DeployWAR)})

# } ###########################


#ADD TO FORM
$objForm.Controls.Add($CheckTop)
$objForm.Controls.Add($FOKINLAB)
$objForm.Controls.Add($OpenFLDR)
$objForm.Controls.Add($TEMPBTN)
$objForm.Controls.Add($FoboButton)
$objForm.Controls.Add($CancelButton)
$objForm.Controls.Add($JobButton)
$objForm.Controls.Add($RestartButton)
$objForm.Controls.Add($CheckServicesBTN)
$objForm.Controls.Add($DeployWAR)
$objForm.TopMost = $true
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()
}
##########################################################################################################################################################################################################
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#START MAIN PROGRAMM
#
#EXECUTION MAIN PROGRAMM
Hide-Console
GENERATOR
$SERVER = ''
##########################################################################################################################################################################################################



##########################################################################################################################################################################################################
#SOME TEST SCURB
<#[void] $objListBox.Items.Add($SREZ + ' 1')
[void] $objListBox.Items.Add($SREZ + ' 2')
[void] $objListBox.Items.Add($SREZ + ' 3')
[void] $objListBox.Items.Add($SREZ + ' 4')
[void] $objListBox.Items.Add($SREZ + ' 5')
[void] $objListBox.Items.Add($SREZ + ' 6')
#$objForm.Controls.Add($objListBox) 

#LABEL SREDA
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,122)
$objLabel.Font = $Font
$objForm
$objLabel.AutoSize = 'True'
$objLabel.BackColor = 'Transparent'
$objLabel.Text = "ВЫБЕРИТЕ СРЕДУ:"
#$objForm.Controls.Add($objLabel) 

#LABEL TRIGGERS
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(130,122)
$objLabel1.BackColor = 'Transparent' 
$objLabel1.Font = $Font
$objLabel1.Size = New-Object System.Drawing.Size(100,20) 
$objLabel1.Text = "ТРИГГЕРЫ:"
#$objForm.Controls.Add($objLabel1) 

#LABEL COMBO
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(250,122) 
$objLabel2.Font = $Font
$objLabel2.BackColor = 'Transparent'
$objLabel2.Text = "МАШИНА:"
$objForm.Controls.Add($objLabel2)



$Server = "C:\1\"
  #
  $DestinationPoint = $Server 
  $DestinationPoint +=$FileBrowser.SafeFileName
  [System.Windows.Forms.MessageBox]::Show($DEST)
  Rename-Item $DestinationPoint -NewName "$Dest.backup" 
  Copy-Item -Path $FileBrowser.FileName -Destination $DestinationPoint
  $Hash1 = Get-FileHash $FileBrowser.FileName
  $Hash2 = Get-FileHash $DestinationPoint
  #
#>

<#
    $StatusNts.add_paint(
    {if($ServiceNTSState.Status -eq 'Running'){
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"orange","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    
    else{
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}

    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring($ServiceNTSState.Status,(new-object System.Drawing.Font("times new roman",11,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(20,3)))
    })
    #>

#Копирование файлов на удаленную машину DBLINK и JBIN
                               <# $Global:FoboStatus.Text = 'Выполняется копирование DBLINKS. Ожидайте!'
                                Start-Sleep -Seconds 2
                                Copy-Item "\\$Server\C$\NTSwincash\config\*" -Filter 'dblink_*' -Destination "\\$Machine\C$\NTSwincash\config\"
                                $XMLNAME =  Get-ChildItem "\\$Server\C$\NTSwincash\config\*" -Include "dblink_V*" 
                                $XMLNAME.Name
                                [xml]$Doc = New-Object System.Xml.XmlDocument                 
                                $FilePath = "\\$Server\C$\NTSwincash\config\dblinks.xml"
                                $Path = "\\$Machine\C$\NTSwincash\config\dblinks.xml"
                                $Path2 = "\\$Machine\C$\NTSwincash\config\dblinks_MobInv.xml"
                                $doc.Load($filePath)
                                $doc.linklist.linkref.file = $XMLNAME.name
                                $doc.Save($Path)
                                $doc.Save($Path2)
                                $Global:FoboStatus.Text = 'Выполняется копирование Jbin. Ожидайте!'
                                Start-Sleep -Seconds 2
                                xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "\\$Machine\C$\NTSWincash\jbin" /S /E /d
                                $Global:FoboStatus.Text = 'Копирование Завершено!'
                                Start-Sleep -Seconds 2       
                                #Invoke-Command -ComputerName $Machine {cmd.exe "/c start C:\NTSWincash\jbin\InstallDistributor-NT.bat"}
                                $Global:FoboStatus.Text = 'Выполняется установка службы!'
                                $PS = Test-Path C:\Windows\System32\PsExec.exe
                                    if($PS -eq $True)
                                    {
                                        psexec -d \\$machine cmd /c "C:\NTSwincash\jbin\NTSWincash Service Installer.exe" DistributorService /install
                                    }
                                    else
                                    {
                                        $PSEXEC = [System.Windows.Forms.MessageBox]::Show("Не установлен PSEXEC!!!","PSexec","YesNoCancel")
                                        switch($PSEXEC)
                                    {
                                        "YES"{[System.Windows.Forms.MessageBox]::Show("Поиск решения корректной установки")}
                                        #xcopy "\\dubovenko\D\SOFT\PSEXEC\" "C:\Windows\System32"}
                                        "NO" {return}
                                        "CANCEL" {return}
                                    }
                                    return  
                                    }
            
                                    if(Get-Service -Name "NTSwincash distributor")
                                    {
                                        $Global:FoboStatus.Text = 'Служба установлена! Установка завершена'
                                        $Global:FoboStatus.ForeColor = 'Green'
                                        #[System.Windows.Forms.MessageBox]::Show("Служба NTSWincash успешно установлена","Успех",'OK','INFO')
                                        #psexec -d \\$machine cmd /c 'C:\NTSwincash\jbin\Configurator.exe'
                                        #start 'C:\NTSwincash\jbin\Configurator.exe'
                                        (Get-WmiObject -Class Win32_Process -ComputerName $Machine -Filter "name='NTSWincash*.exe'").terminate() | Out-Null
                                    }
                                    else
                                    {
                                        $Global:FoboStatus.Text = 'Ошибка установки службы!'
                                        $Global:FoboStatus.ForeColor = 'RED'
                                        [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash не была установлена!","Ошибка",'OK','ERROR')
                                        (Get-WmiObject -Class Win32_Process -ComputerName $Machine -Filter "name='NTSWincash*.exe'").terminate() | Out-Null
                                    }
                                  #>