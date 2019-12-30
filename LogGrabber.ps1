##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBnxRooHXVVlgizuSRi5FNETR+Eau99cYhQkK/0c8f/8FODkaqQM3Op8ZIU=
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
##NEW LOG GRABER
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

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

FUNCTION LOGGRABBER(){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $Font = New-Object System.Drawing.Font("Tempus Sans ITC",9,[System.Drawing.FontStyle]::Bold)
    $Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\grabber.jpg")
    $FontGroup = New-Object System.Drawing.Font("Eras Bold ITC",9,[System.Drawing.FontStyle]::Regular)
    $NTSwin = "C:\NTSWincash"


    $LOG_FORM = New-Object System.Windows.Forms.Form
    $LOG_FORM.SizeGripStyle = "Hide"
    $LOG_FORM.BackgroundImage = $Image
    $LOG_FORM.BackgroundImageLayout = "None"
    $LOG_FORM.Size = ('360,200')
    $LOG_FORM.BackColor = 'White'
    $LOG_FORM.StartPosition = "CenterScreen"
    $LOG_FORM.TopMost = $true
    $LOG_FORM.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $LOG_FORM.Text = "Сборщик Логов [Version 1.1]"
    $LOG_FORM.TopMost = $True
    $LOG_FORM.Icon = $Icon

    #Firm Label
    $FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Bold)
    $FOKINLAB = New-Object System.Windows.Forms.Label
    $FOKINLAB.Location = ('250,115')
    $FOKINLAB.Text = "Created By Fokin"
    $FOKINLAB.Font = $FontBanksy
    $FOKINLAB.BackColor =  'Transparent'
    $FOKINLAB.AutoSize = $True

    $LOGGROUP = New-Object System.Windows.Forms.GroupBox
    $LOGGROUP.Location = ('100,5')
    $LOGGROUP.Text = "Choice Your Module:"
    $LOGGROUP.Font = $FontGroup
    $LOGGROUP.Size = ('220,110')

    $LOG_SALES = New-Object System.Windows.Forms.RadioButton
    $LOG_SALES.Location = ('10,15')
    $LOG_SALES.Text = "SALES"
    $LOG_SALES.AutoSize = $True
    

    $Log_Stages = @('Standart','+MSP')
    $LOG_SALES_EXT = New-Object System.Windows.Forms.ComboBox
    $LOG_SALES_EXT.Location = ('95,15')
    $LOG_SALES_EXT.DataSource = $Log_Stages
    $LOG_SALES_EXT.Font = $FontGroup
    $LOG_SALES_EXT.Size = ('80,15')
    $LOG_SALES_EXT.Visible = $false

    $LOG_LOGISTIC = New-Object System.Windows.Forms.RadioButton
    $LOG_LOGISTIC.Location = ('10,35')
    $LOG_LOGISTIC.Text = "LOGISTIC"
    $LOG_LOGISTIC.AutoSize = $True


    $LOG_MSP = New-Object System.Windows.Forms.RadioButton
    $LOG_MSP.Location =  ('10,55')
    $LOG_MSP.AutoSize = $True
    $LOG_MSP.Text = 'MSP'


    $LOG_TSD = New-Object System.Windows.Forms.RadioButton
    $LOG_TSD.Location =  ('10,75')
    $LOG_TSD.AutoSize = $True
    $LOG_TSD.Text = 'TSD'

    
    $LOG_BUTTON = New-Object System.Windows.Forms.Button
    $LOG_BUTTON.Location = ('110,120')
    $LOG_BUTTON.Text = "Grab LOGS"
    $LOG_BUTTON.Font = $FontGroup
    $LOG_BUTTON.Size = ('120,24')

    $CLEAR_LOGS_BUTTON 
    $CLEAR_LOGS_BUTTON = New-Object System.Windows.Forms.Button
    $CLEAR_LOGS_BUTTON.Location = ('110,120')
    $CLEAR_LOGS_BUTTON.Text = "Grab LOGS"
    $CLEAR_LOGS_BUTTON.Font = $FontGroup
    $CLEAR_LOGS_BUTTON.Size = ('120,24')

    $LOGGROUP.Controls.AddRange(@($LOG_SALES,$LOG_LOGISTIC,$LOG_SALES_EXT,$LOG_MSP))


    $LOG_SALES.Add_Click({
        if($LOG_SALES.Checked)
        {
            $LOG_SALES_EXT.Visible = $True
        }
    })

    $LOG_LOGISTIC.Add_Click({
        if($LOG_LOGISTIC.Checked)
        {
            $LOG_SALES_EXT.Visible = $False
        }
    })




    




    $USERPROFILE = $env:USERPROFILE
    $LOG_BUTTON.Add_Click({
        $Date = Get-Date -Format dd.MM.yyyy
        if($LOG_SALES.Checked -eq $true)
        {
             
            if($LOG_SALES_EXT.Text -like "Standart")
            {
                if((Test-Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\") -eq $False)
                {
                    New-Item -ItemType Directory "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                    Copy-Item "$NTSwin\log\sales\*" -Recurse -Include "general.log","database.log","eft.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                }
                else
                {
                    Copy-Item "$NTSwin\log\sales\*" -Include "general.log","database.log","eft.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                }
                if(Test-Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip")
                {
                    Remove-Item "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip" -Force
                    Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip"
                }
                else
                {
                 Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip"
                }
                    Start-Sleep -Seconds 2
                    Remove-Item "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\" -Recurse -Force
                    $LOG_FORM.TopMost = $False
                    #$LOG_FORM.Bottom = $True
                    New-WPFMessageBox -Content "Заберите Ваши Логи!" -Title "Archive LOGS" -ContentBackground Cornsilk
                    Invoke-Item "$USERPROFILE\Documents\Sales_LOG\"
                    $LOG_FORM.TopMost = $True
                    #$LOG_FORM.Bottom = $False
            }
            elseif($LOG_SALES_EXT.Text -like "+MSP")
            {
                if((Test-Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\") -eq $False)
                {
                    New-Item -ItemType Directory "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                    Copy-Item "$NTSwin\log\sales\*" -Recurse -Include "general.log","database.log","eft.log","msp.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                }
                else
                {
                    Copy-Item "$NTSwin\log\sales\*" -Include "general.log","database.log","eft.log","msp.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\"
                }
                if(Test-Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip")
                {
                    Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip" -Update
                }
                else
                {
                 Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOG_$Date.zip"
                }
                    Start-Sleep -Seconds 2
                    Remove-Item "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\" -Recurse -Force
                    $LOG_FORM.TopMost = $False
                    #$LOG_FORM.Bottom = $True
                    New-WPFMessageBox -Content "Заберите Ваши Логи!" -Title "Archive LOGS" -ContentBackground Cornsilk
                    Invoke-Item "$USERPROFILE\Documents\Sales_LOG\"
                    $LOG_FORM.TopMost = $True
                    #$LOG_FORM.Bottom = $False
            }

        }
        elseif($LOG_LOGISTIC.Checked)
        {

            if((Test-Path "$USERPROFILE\Documents\Sales_LOG\LOG_$Date\") -eq $False)
            {
                New-Item -ItemType Directory "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date\"
                Copy-Item "$NTSwin\log\logistics\*" -Include "general.log","database.log","eft.log","msp.log" -Destination "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date\"
            }
            else
            {
            Copy-Item "$NTSwin\log\logistics\*" -Include "general.log","database.log","eft.log","msp.log" -Destination "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date\"
            }
            if(Test-Path "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date.zip")
            {  
                Compress-Archive -Path "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date.zip" -Update
            }
            else
            {
                Compress-Archive -Path "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date\*" -DestinationPath "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date.zip"
            }
            Start-Sleep -Seconds 2
            Remove-Item "$USERPROFILE\Documents\Logistic_LOG\LOG_$Date" -Recurse -Force
            $LOG_FORM.TopMost = $False
            #$LOG_FORM.Bottom = $True
            New-WPFMessageBox -Content "Заберите Ваши Логи!" -Title "Archive LOGS" -ContentBackground Cornsilk
            Invoke-Item "$USERPROFILE\Documents\Logistic_LOG\"
            $LOG_FORM.TopMost = $True
            #$LOG_FORM.Bottom = $False

        }
        elseif($LOG_MSP.Checked)
        {
                if((Test-Path "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\") -eq $False)
                {
                    New-Item -ItemType Directory "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\"
                    Copy-Item "$NTSwin\log\sales\*" -Recurse -Include "msp.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\"
                }
                else
                {
                    Copy-Item "$NTSwin\log\sales\*" -Include "msp.log" -Destination "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\"
                }
                if(Test-Path "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date.zip")
                {
                    Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date.zip" -Update
                }
                else
                {
                 Compress-Archive -Path "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\*" -DestinationPath "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date.zip"
                }
                    Start-Sleep -Seconds 2
                    Remove-Item "$USERPROFILE\Documents\Sales_LOG\LOGMSP_$Date\" -Recurse -Force
                    $LOG_FORM.TopMost = $False
                    #$LOG_FORM.Bottom = $True
                    New-WPFMessageBox -Content "Заберите Ваши Логи!" -Title "Archive LOGS" -ContentBackground Cornsilk
                    Invoke-Item "$USERPROFILE\Documents\Sales_LOG\"
                    $LOG_FORM.TopMost = $True
                    #$LOG_FORM.Bottom = $False  
        }
        else
        {
            $LOG_FORM.TopMost = $False
            New-WPFMessageBox -Content "Выберите Модуль!" -Title "[Choose your Destiny]" -ContentBackground RosyBrown
            $LOG_FORM.TopMost = $True
            return
        }
    })


    function ENTERCOLOR($ELEMENT){
		$ELEMENT.BackColor = 'LightGreen'
	}
    function LEAVECOLOR($ELEMENT){
        $ELEMENT.BackColor = 'Control'
    }

    $LOG_BUTTON.add_MouseHover({ENTERCOLOR($LOG_BUTTON)})
    $LOG_BUTTON.add_MouseLeave({LEAVECOLOR($LOG_BUTTON)})

    $LOG_FORM.Controls.AddRange(@($LOGGROUP,$LOG_BUTTON))
    $LOG_FORM.Controls.Add($FOKINLAB)
    $LOG_FORM.ShowDialog()
}

Hide-Console
LOGGRABBER