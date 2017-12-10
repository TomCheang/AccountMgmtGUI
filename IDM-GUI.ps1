<# IDM Helper Tool.  Necessary as IDM is turned down.
    v1.0
    1.Update samaccountname
    2.Update OU
    3.Update UserPrincipalName
    4.Update Exchange Alias
    5.Update PrimarySmtpAddress
#>

$GuiRunspaceScript = {
#region XAML for WPF
  $raw = @"
<Window x:Name="IDM_Helper_Tool"  x:Class="WpfApplication4.MainWindow"
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="IDM Helper Tool" Width="950" Height="700" WindowStyle="ThreeDBorderWindow"
  WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize" >
  <Window.Resources>
    <Style TargetType="StackPanel">
      <Setter Property="Margin" Value="20,10,10,10" />
      <Setter Property="Width" Value="900" />
      <Setter Property="HorizontalAlignment" Value="Left" />
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Margin" Value="10,10,10,10" />
      <Setter Property="Padding" Value="4" />
      <Setter Property="BorderThickness" Value="2" />
    </Style>
    <Style TargetType="Button">
      <Setter Property="Height" Value="20" />
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Margin" Value="10,5,10,10" />
    </Style>
    <Style TargetType="Expander">
      <Setter Property="Margin" Value="10,5,10,10" />
    </Style>
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="FontSize" Value="9" />
    </Style>
  </Window.Resources>
  <Grid>
    <StackPanel Margin="0" Width="2000" Background="#FF39579A" Height="28"
      VerticalAlignment="Top" Orientation="Horizontal">
      <Image Name="imgTopPanel" DockPanel.Dock="Top" Stretch="None"
             HorizontalAlignment="Left" />
      <Separator Margin="0,0,600,0" />
      <TextBlock HorizontalAlignment="Right" Margin="0,5,10,5"
        Foreground="White">Admin:</TextBlock>
      <TextBlock Name="txbAdminName" Margin="0,5,20,10" Foreground="White"
        FontWeight="Bold" Text="unknown" />
    </StackPanel>
    <StackPanel Margin="25,40,10,15">
      <GroupBox Header="1. Locate User">
        <StackPanel Orientation="Horizontal">
          <TextBox Name="txbFindUser" Width="250"
            ToolTip="Ex: tomcheang , tomcheang@company.com, or Tom Cheang" />
          <Button Name="btnFindUser">Find User</Button>
          <Separator  Margin="10,0,20,0" />
          <TextBlock Text="Search by sAMAccountName, PrimaryEmailAddress or 'FirstName LastName'"
            TextWrapping="Wrap" Width="150" FontSize="11" FontStyle="Italic"
            Foreground="DarkGray" />
        </StackPanel>
      </GroupBox>
        <WrapPanel Margin="50,0,0,0" Width="750" HorizontalAlignment="Left"
          Orientation="Horizontal">
          <ProgressBar Name ="prgFindUser" Height="15" Width="250"
            Visibility="Hidden"/>
          <Separator Margin="50,0,0,0"/>
          <TextBlock Name="txbWarning" FontStyle="Italic"
            Text="**sAMAccountName searches are wildcard-based and slower."
            Visibility="Hidden"/>
        </WrapPanel>
        <TextBlock Name="txbErrorFindUser" Visibility="Hidden"
          Foreground="Red" Text="N/A" />
      <Expander Name="expResults" Header="2. Select User Account"  >
        <WrapPanel Orientation="Horizontal">
          <ListView Name="lvwResults"  Width="700" Height="125"
            HorizontalAlignment="Left" ScrollViewer.CanContentScroll="True">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="FirstName" Width="100"
                  DisplayMemberBinding="{Binding givenName}" />
                <GridViewColumn Header="LastName" Width="100"
                  DisplayMemberBinding="{Binding surname}" />
                <GridViewColumn Header="Primary Email" Width="150">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBlock Text="{Binding EmailAddress}"
                        TextDecorations="Underline" Foreground="Blue"
                          Cursor="Hand" />
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="UPN" Width="150"
                  DisplayMemberBinding="{Binding UserPrincipalName}" />
                <GridViewColumn Header="sAMAccountName" Width="100"
                  DisplayMemberBinding="{Binding sAMAccountName}" />
                <GridViewColumn Header="Exch Alias" Width="100"
                  DisplayMemberBinding="{Binding mailNickName}" />
                <GridViewColumn Header="DistinguishedName" Width="0"
                  DisplayMemberBinding="{Binding DistinguishedName}" />
              </GridView>
            </ListView.View>
        </ListView>
        <Separator Margin="10,0,10,0" />
          <WrapPanel Orientation="Vertical">
            <TextBlock Text="Property to update:" Margin="0,0,0,5"
              FontWeight="Bold"/>
            <RadioButton Name="rbnUPN" Margin="15,1,0,1"
              Content="UserPrincipalName" IsEnabled="False" />
            <RadioButton Name="rbnSamaccountname" Margin="15,0,0,0"
              Content="sAMAccountName" IsEnabled="False" />
            <RadioButton Name="rbnAlias" Margin="15,0,0,0"
              Content="Exchange Alias" IsEnabled="False" />
            <RadioButton Name="rbnOU" Margin="15,0,0,0"
              Content="Organizational Unit" IsEnabled="False" />
            <RadioButton Name="rbnPrimaryEmail" Margin="15,0,0,0"
              Content="Primary Email" IsEnabled="False" />
          </WrapPanel>
        </WrapPanel>
      </Expander>
    <Expander Name="expUpdate" Header="3. Update User Properties"
      MaxHeight="250" Height="200">
      <StackPanel Orientation="Vertical">
        <StackPanel Orientation="Horizontal" Margin="0,0,0,0" >
            <TextBlock Text="Selected User: " />
            <TextBlock x:Name="txbSelectedUser" FontWeight="Bold" Text="none"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,0" >
            <TextBlock Text="Property to Update: " />
            <TextBlock x:Name="txbSelectedProperty" FontWeight="Bold"
              Text="none" />
        </StackPanel>
        <Separator Margin="0,0,0,20" Visibility="Hidden" />
        <StackPanel Orientation="Horizontal" Margin="150,0,0,15"
          HorizontalAlignment="Left" Width="440">
          <TextBlock Text="Current Setting: " />
          <TextBlock x:Name="txbCurrentSetting" FontWeight="Bold" Text="none"/>
        </StackPanel >
        <StackPanel Orientation="Horizontal" Margin="0,0,0,0"
          HorizontalAlignment="Center" Width="600" >
          <ComboBox Name="cmbOrgUnit" Width="175" Height="25"
            Text="  - Select OU -" Visibility="Hidden" >
            <ComboBox.Items>
              <ComboBoxItem Name="cmbItem1"
                Content="OU=Employees"></ComboBoxItem>
              <ComboBoxItem Name="cmbItem2"
                Content="OU=Contractors"></ComboBoxItem>
              <ComboBoxItem Name="cmbItem3"
                Content="OU=DisabledAccounts"></ComboBoxItem>
            </ComboBox.Items>
          </ComboBox>
            <TextBox Name="txbNewProperty"  Width="250" BorderThickness="5"/>
          <Button Name="btnUpdate" Margin="10,0,0,0" Padding="4,0,4,0"
            Content=" Update " IsEnabled="False"/>
        </StackPanel>
        <WrapPanel>
          <TextBlock Name="txbSuccess" Foreground="Green"
            Margin="50,10,60,10" FontWeight="Bold"
            HorizontalAlignment="Center"/>
          <TextBlock Name="txbUpdateError" Foreground="Red"
            Margin="50,10,60,10" HorizontalAlignment="Center"/>
        </WrapPanel>
      </StackPanel>
      </Expander>
    </StackPanel>
  </Grid>
</Window>
"@
  #endregion XAML

  #Fix up XAML that is copied and pasted from Visual Studio
  #Make sure <Window has no whitespace before it!!
  [XML]$XAML = $raw -replace "x:Name",'Name' -replace '^<Windo.*', '<Window'
  $reader=(New-Object System.Xml.XmlNodeReader $XAML)
  $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )

  #region Declare Runspace functions
  function Get-RunspaceAdUser {
    param($syncHash,$txbFindUserText,$lvwResults)

    $ScriptBlock = {
      $Properties = @('GivenName', 'SurName', 'EmailAddress', 'UserPrincipalName',
        'samaccountname', 'mailNickName', 'DistinguishedName')
      Try {
        if ($txbFindUserText.Contains(' ')) {
          $FirstName = ($txbFindUserText -split ' ')[0]
          $LastName = ($txbFindUserText -split ' ')[-1]
           $Results = Get-AdUser -Filter {givenName -eq $FirstName -and
            sn -eq $LastName} -Properties $Properties | select $Properties
        }
        elseif ($txbFindUserText.contains('@')) {
          $Results = Get-AdUser -Filter {mail -eq $txbFindUserText} -Properties  $Properties |
            select $Properties
        }
        else {
          $syncHash.Window.Dispatcher.Invoke(
            [action]{
            $syncHash.txbWarning.Visibility = 'Visible'
        })
          $Results = Get-Aduser -Filter "samaccountname -like `'`*$txbFindUserText`*`'" -Properties  $Properties |
            select $Properties
        }

        if ($Results) {
          $syncHash.Window.Dispatcher.invoke(
           [action]{foreach ($r in $Results){$lvwResults.AddChild($r)}
          $syncHash.expResults.IsExpanded = $true
          $syncHash.prgFindUser.isIndeterminate = $false
           $syncHash.prgFindUser.Visibility = "Hidden"
           $syncHash.txbWarning.Visibility = "Hidden"
          })
        }
        else {

          $syncHash.Window.Dispatcher.Invoke([action]{
          $msg = "None found with search query"
            $syncHash.txbErrorFindUser.Text = $msg
            $syncHash.txbErrorFindUser.Visibility = "Visible"
            $syncHash.prgFindUser.isIndeterminate = $false
            $syncHash.prgFindUser.Visibility = "Hidden"
            $syncHash.txbWarning.Visibility = "Hidden"
          })
        }
      }
      Catch {

        $syncHash.Window.Dispatcher.Invoke([action]{
          $msg = $Error[0].Exception
          $syncHash.txbErrorFindUser.Text = $msg
          $syncHash.txbErrorFindUser.Visibility = "Visible"
          $syncHash.prgFindUser.isIndeterminate = $false
          $syncHash.prgFindUser.Visibility = "Hidden"
          $syncHash.txbWarning.Visibility = "Hidden"
        })

      }
      Finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
          $syncHash.prgFindUser.IsIndeterminate =$false
          $syncHash.prgFindUser.Visibility = 'Hidden'
          $syncHash.txbWarning.Visibility = 'Hidden'
          $syncHash.btnFindUser.IsEnabled = $true
        })
      }
    }

    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("txbFindUserText",$txbFindUserText)
    $Runspace.SessionStateProxy.SetVariable("lvwResults",$lvwResults)

    $PSinstance = [powershell]::Create().AddScript($ScriptBlock)
    $PSinstance.Runspace = $Runspace
    $ADRunspace = $PSinstance.BeginInvoke()
  }

  function Set-RunspaceADUser {
    param(
      $syncHash,$txbSelectedUser,$txbSelectedProperty,
      $txbNewProperty,$cmbOrgUnit
    )

    $ScriptBlock = {
      Try{
        if ($txbSelectedProperty -eq 'sAMAccountName') {
          Set-Aduser -Identity $txbSelectedUser -SamAccountName $txbNewProperty
          Rename-ADObject -Identity $txbSelectedUser -NewName $txbNewProperty
        }
        elseif ($txbSelectedProperty -eq 'Organizational Unit') {
         $OU = $cmbOrgUnit + ',dc=yourcompany,dc=com'
         Move-AdObject -Identity $txbSelectedUser -TargetPath $OU
        }
        elseif ($txbSelectedProperty -eq 'UserPrincipalName') {
         Set-ADUser -Identity $txbSelectedUser -UserPrincipalName $txbNewProperty
        }
        else {
          $syncHash.Window.Dispatcher.Invoke([action]{
            $msg  = "Something unexpected happened here. " +
              " Restart program, please."
            $syncHash.txbUpdateError.Text = $msg

          })
        }

        #Write success message
        $syncHash.Window.Dispatcher.Invoke([action]{
          $syncHash.txbSuccess.Text = "Success: " + $txbSelectedProperty +
             " updated to " + $txbNewProperty + $cmbOrgUnit
          $syncHash.txbUpdateError.Text = ""
        })
      }
      Catch {
        #write error message
        $syncHash.Window.Dispatcher.Invoke([action]{
          $msg = $Error[0].Exception
          $syncHash.txbUpdateError.Text = $msg
          $syncHash.txbSuccess.Text = ""
        })
      }
      Finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
          $syncHash.btnUpdate.IsEnabled = $true
        })
      }
    }

    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("txbSelectedUser",$txbSelectedUser)
    $Runspace.SessionStateProxy.SetVariable(
      "txbSelectedProperty",$txbSelectedProperty
    )
    $Runspace.SessionStateProxy.SetVariable("txbNewProperty",$txbNewProperty)
    $Runspace.SessionStateProxy.SetVariable("cmbOrgUnit",$cmbOrgUnit)
    $PSinstance = [powershell]::Create().AddScript($ScriptBlock)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
  }

  function Initialize-ExchangeRunspace {
    $ScriptBlock = {
      Try {
        $RandomCHUB =  @("CAS01", "CAS02", "CAS03") | Get-Random
        $ConnectionUri = "http://$RandomCHUB/PowerShell/"
        $splPsSession = @{
          ConfigurationName = 'Microsoft.Exchange'
          ConnectionUri = $ConnectionUri
          Authentication = 'Kerberos'
        }
        $session = New-PSSession @splPsSession
        Import-PSSession $session
      }
      Catch {
        #Nothing to do here, for now..
      }
    }

    $syncHash.Host = $host
    $RunspaceExch = [runspacefactory]::CreateRunspace()
    $RunspaceExch.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $RunspaceExch.Open()
    $PSinstanceExch = [powershell]::Create().AddScript($ScriptBlock)
    $PSinstanceExch.Runspace = $Runspace
    $ExchRunspace = $PSinstanceExch.BeginInvoke()
  }

  function Set-RunspaceMailbox {
    Param (
      $syncHash,$txbSelectedUser, $txbSelectedProperty, $txbNewProperty
    )
    $ScriptBlock = {
      Try {
        $MailboxType = Get-Recipient -Identity $txbSelectedUser |
          Select RecipientTypeDetails

        if ($MailboxType -eq 'UserMailbox') {
          if ($txbSelectedProperty -eq 'Exchange Alias') {
             Set-Mailbox -Identity $txbSelectedUser -Alias $txbNewProperty
          }
          elseif ($txbSelectedProperty -eq 'Primary Email') {
            $splPrimarySmtpAddress = @{
              Identity = $txbSelectedUser
              PrimarySmtpAddress = $txbNewProperty
              EmailAddressPolicyEnabled = $false
            }
            Set-Mailbox @splPrimarySmtpAddress
          }

        }
        elseif ($MailboxType -eq 'RemoteUserMailbox') {
          if ($txbSelectedProperty -eq 'Exchange Alias') {
            Set-RemoteMailbox -Identity $txbSelectedUser -Alias $txbNewProperty
          }
          elseif ($txbSelectedProperty -eq 'Primary Email') {
            $splPrimarySmtpAddress = @{
              Identity = $txbSelectedUser
              PrimarySmtpAddress = $txbNewProperty
              EmailAddressPolicyEnabled = $false
            }
            Set-RemoteMailbox @splPrimarySmtpAddress
          }
        }

        $syncHash.Window.Dispatcher.Invoke([action]{
          $syncHash.txbSuccess.Text = "Success: " + $txbSelectedProperty +
             " updated to " + $txbNewProperty
          $syncHash.txbUpdateError.Text = ""
          })

      }
      Catch {
        $syncHash.Window.Dispatcher.Invoke([action]{
          $msg = $Error[0].Exception
          $syncHash.txbUpdateError.Text = $msg
          $syncHash.txbUpdateError.Text = ""
        })
      }
      Finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
          $syncHash.btnUpdate.IsEnabled = $true
        })
      }
    }

    $syncHash.Host = $host
    $RunspaceExch = [runspacefactory]::CreateRunspace()
    $RunspaceExch.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $RunspaceExch.Open()
    $PSinstanceExch = [powershell]::Create().AddScript($ScriptBlock)
    $PSinstanceExch.Runspace = $Runspace
    $ExchRunspace = $PSinstanceExch.BeginInvoke()
    }
  #endregion GUI functions

  #region Populate syncHash with GUI controls

  $syncHash.txbFindUser = $syncHash.window.FindName("txbFindUser")
  $syncHash.btnFindUser = $syncHash.window.FindName("btnFindUser")
  $syncHash.imgTopPanel = $syncHash.window.FindName("imgTopPanel")
  $syncHash.txbAdminName = $syncHash.window.FindName("txbAdminName")
  $syncHash.prgFindUser = $syncHash.window.FindName("prgFindUser")
  $syncHash.txbWarning = $syncHash.window.FindName("txbWarning")
  $syncHash.txbErrorFindUser = $syncHash.window.FindName("txbErrorFindUser")
  $syncHash.expResults = $syncHash.window.FindName("expResults")
  $syncHash.lvwResults = $syncHash.window.FindName("lvwResults")
  $syncHash.rbnUPN = $syncHash.window.FindName("rbnUPN")
  $syncHash.rbnSamaccountname = $syncHash.window.FindName("rbnSamaccountname")
  $syncHash.rbnAlias = $syncHash.window.FindName("rbnAlias")
  $syncHash.rbnPrimaryEmail = $syncHash.window.FindName("rbnPrimaryEmail")
  $syncHash.rbnOU = $syncHash.window.FindName("rbnOU")
  $syncHash.txbNewProperty = $syncHash.window.FindName("txbNewProperty")
  $syncHash.expUpdate = $syncHash.window.FindName("expUpdate")
  $syncHash.txbSelectedUser = $syncHash.window.FindName("txbSelectedUser")
  $syncHash.txbSelectedProperty = $syncHash.window.FindName(
    "txbSelectedProperty"
  )
  $syncHash.txbCurrentSetting = $syncHash.window.FindName("txbCurrentSetting")
  $syncHash.cmbOrgUnit = $syncHash.window.FindName("cmbOrgUnit")
  $syncHash.btnUpdate = $syncHash.window.FindName("btnUpdate")
  $syncHash.txbUpdateError = $syncHash.window.FindName("txbUpdateError")
  $syncHash.txbSuccess = $syncHash.window.FindName("txbSuccess")
  $syncHash.cmbOrgUnit.cmbItem1 = $syncHash.window.FindName("cmbItem1")
  $syncHash.cmbOrgUnit.cmbItem2 = $syncHash.window.FindName("cmbItem2")
  $syncHash.cmbOrgUnit.cmbItem3 = $syncHash.Window.FindName("cmbItem3")

  #endregion


  #region GUI Events - equivalent to C# code behind

  #FindUser button click
  $syncHash.btnFindUser.Add_Click({
    $syncHash.Window.Dispatcher.Invoke(
    [action]{$global:Query = $syncHash.txbFindUser.Text})

    if ($Global:Query) {
      $syncHash.Window.Dispatcher.invoke([action]{
        $syncHash.btnFindUser.IsEnabled = $false
        $syncHash.lvwResults.Items.clear()
        $syncHash.txbErrorFindUser.Visibility = "Hidden"
        $syncHash.expResults.IsExpanded = $false
        $syncHash.rbnOU.IsEnabled = $false
        $syncHash.rbnSamaccountname.IsEnabled = $false
        $syncHash.rbnAlias.IsEnabled = $false
        $syncHash.rbnUPN.IsEnabled = $false
        $syncHash.rbnPrimaryEmail.isEnabled = $false
        $syncHash.txbSuccess.
        $syncHash.txbUpdateError.
        $syncHash.expUpdate.IsExpanded = $false
        $syncHash.prgFindUser.Visibility = 'Visible'
        $syncHash.prgFindUser.IsIndeterminate =$true
        $syncHash.txbWarning.Visibility = 'Hidden'
        $syncHash.txbNewProperty.Clear()
        $syncHash.btnUpdate.IsEnabled = $false
        $syncHash.rbnSamaccountname.Checked = $false
        $syncHash.rbnUPN.Checked = $false
        $syncHash.rbnPrimaryEmail.Checked = $false
        $syncHash.rbnOU.Checked = $false
        $syncHash.rbnAlias.Checked = $false
      })
      $splFindUser = @{
        syncHash = $syncHash
        txbFindUserText = $syncHash.txbFindUser.Text
        lvwResults = $syncHash.lvwResults
      }
      Get-RunspaceAdUser @splFindUser
    }
  })

  #GridView selection changed
  $syncHash.lvwResults.Add_SelectionChanged({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.rbnOU.IsEnabled = $true
      $syncHash.rbnSamaccountname.IsEnabled = $true
      $syncHash.rbnAlias.IsEnabled = $true
      $syncHash.rbnUPN.isEnabled = $true
      $syncHash.rbnPrimaryEmail.isEnabled = $true
      $syncHash.txbSelectedUser.Text = $syncHash.lvwResults.SelectedItem.DistinguishedName

      if ($syncHash.rbnUPN.IsChecked) {
       $syncHash.txbSelectedProperty.Text = $syncHash.rbnUPN.Content
       $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.UserPrincipalName
        }
      elseif ($syncHash.rbnSamaccountname.IsChecked) {
        $syncHash.txbSelectedProperty.Text = $syncHash.rbnSamaccountname.Content
        $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.Samaccountname
      }
      elseif ($syncHash.rbnAlias.IsChecked) {
        $syncHash.txbSelectedProperty.Text = $syncHash.rbnAlias.Content
        $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.mailNickName
      }
      elseif ($syncHash.rbnOU.IsChecked) {
        $syncHash.txbSelectedProperty.Text = $syncHash.rbnOU.Content
        $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.distinguishedName.Split(',')[1]
      }
      elseif ($syncHash.rbnPrimaryEmail.IsChecked) {
        $syncHash.txbSelectedProperty.Text = $syncHash.rbnPrimaryEmail.Content
        $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.EmailAddress
      }
      $syncHash.txbNewProperty.Clear()
      $syncHash.cmbOrgUnit.SelectedIndex = '-1'
    })
  })

  #rbnSamaccountname checked
  $syncHash.rbnSamaccountname.Add_Checked({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.expUpdate.isExpanded = $true
      $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.Samaccountname
      $syncHash.txbSelectedProperty.Text = $syncHash.rbnSamaccountname.Content
      $syncHash.txbNewProperty.Visibility = "Visible"
      $syncHash.cmbOrgUnit.Visibility = "Hidden"
    })
  })

  #rbnOU checked
  $syncHash.rbnOU.Add_Checked({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.expUpdate.isExpanded = $true
      $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.DistinguishedName.Split(',')[1]
      $syncHash.txbSelectedProperty.Text = $syncHash.rbnOU.Content
      $syncHash.txbNewProperty.Visibility = "Hidden"
      $syncHash.cmbOrgUnit.Visibility = "Visible"
    })
  })

  #rbnAlias checked
  $syncHash.rbnAlias.Add_Checked({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.expUpdate.isExpanded = $true
      $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.mailNickName
      $syncHash.txbSelectedProperty.Text = $syncHash.rbnAlias.Content
      $syncHash.txbNewProperty.Visibility = "Visible"
      $syncHash.cmbOrgUnit.Visibility = "Hidden"
    })
  })

  #rbnUPN checked
  $syncHash.rbnUPN.Add_Checked({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.expUpdate.isExpanded = $true
      $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.UserPrincipalName
      $syncHash.txbSelectedProperty.Text = $syncHash.rbnUPN.Content
      $syncHash.txbNewProperty.Visibility = "Visible"
      $syncHash.cmbOrgUnit.Visibility = "Hidden"
    })
  })

  #rbnPrimaryEmail checked
  $syncHash.rbnPrimaryEmail.Add_Checked({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.expUpdate.isExpanded = $true
      $syncHash.txbCurrentSetting.Text = $syncHash.lvwResults.SelectedItem.EmailAddress
      $syncHash.txbSelectedProperty.Text = $syncHash.rbnPrimaryEmail.Content
      $syncHash.txbNewProperty.Visibility = "Visible"
      $syncHash.cmbOrgUnit.Visibility = "Hidden"
    })
  })

  #txbNewProperty text changed
  $syncHash.txbNewProperty.Add_TextChanged({
    $syncHash.btnUpdate.IsEnabled = $true
  })
    #cmbOrgUnit Selection changed

  $syncHash.cmbOrgUnit.Add_SelectionChanged({
    if ($syncHash.cmbOrgUnit.SelectedItem.Content -eq $syncHash.txbCurrentSetting.Text) {
      $syncHash.btnUpdate.IsEnabled = $false
    }
    else {
      $syncHash.btnUpdate.IsEnabled = $true
    }
  })
  #btnUpdate Click
  $syncHash.btnUpdate.Add_Click({
    $syncHash.Window.Dispatcher.Invoke([action]{
      $syncHash.txbErrorFindUser.Text = ""
      $syncHash.txbSuccess.Text = ""
      $syncHash.txbUpdateError.Text = ""
      $syncHash.btnUpdate.IsEnabled = $false
    })
    if ($syncHash.txbSelectedProperty.Text -in @(
      "sAMAccountName", "UserPrincipalName", "Organizational Unit")) {
      $splSetADUser = @{
        syncHash = $syncHash
        txbSelectedUser = $syncHash.txbSelectedUser.Text
        txbSelectedProperty = $syncHash.txbSelectedProperty.Text
        txbNewProperty = $syncHash.txbNewProperty.Text
        cmbOrgUnit = $syncHash.cmbOrgUnit.SelectedItem.Content
      }
      Set-RunspaceADUser @splSetADUser
    }
    else {
      $splSetMailbox = @{
        syncHash = $syncHash
        txbSelectedUser = $syncHash.txbSelectedUser.Text
        txbSelectedProperty = $syncHash.txbSelectedProperty.Text
        txbNewProperty = $syncHash.txbNewProperty.Text
      }
      Set-RunspaceMailbox @splSetMailbox
    }
  })

  #endregion Events

  #Let's set the starting state
  Initialize-ExchangeRunspace
  $base64string = 'AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKBgRcOjYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP///////////6NiSP+jYkj/o2JI/6NiSP+iYEfEAAAAAAAAAACjYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj///////////+jYkj/o2JI/6NiSP+jYkj/o2JI/wAAAAAAAAAAo2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj/o2JI////////////o2JI/6NiSP+jYkj/o2JI/6NiSP8AAAAAAAAAAKNiSP+jYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP///////////6NiSP+jYkj/o2JI/6NiSP+jYkj/AAAAAAAAAACjYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj///////////+jYkj/o2JI/6NiSP+jYkj/o2JI/wAAAAAAAAAAo2JI/6NiSP+jYkj/o2JI/6NiSP+CTjr/gk46////////////gk46/4JOOv+SWED/o2JI/6NiSP8AAAAAAAAAAKNiSP+jYkj/o2JI/6NiSP+jYkj/////////////////////////////////x66k/6NiSP+jYkj/AAAAAAAAAACjYkj/o2JI/6NiSP+jYkj/o2JI//////////////////////////////////Hn4/+jYkj/o2JI/wAAAAAAAAAAo2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj/o2JI////////////ml1E/6NiSP+jYkj/o2JI/6NiSP8AAAAAAAAAAKNiSP+jYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP/59vX//////6R+b/+CTjr/gk46/6NiSP+jYkj/AAAAAAAAAACjYkj/o2JI/6NiSP+ jYkj/o2JI/6NiSP+jYkj/3si///////////////////////+jYkj/o2JI/wAAAAAAAAAAo2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj/o2JI/6ptVv/izsb//Pv6//fx7//r3dj/o2JI/6NiSP8AAAAAAAAAAJtdRP+jYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP+jYkj/o2JI/6NiSP+bXUT/AAAAAAAAAACATTbDgk46/4JOOv+CTjr/gk46/4JOOv+CTjr/gk46/4JOOv+CTjr/gk46/4JOOv+CTjr/gE02wwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAIABAACAAQAAgAEAAIABAACAAQAAgAEAAIABAACAAQAAgAEAAIABAACAAQAAgAEAAIABAACAAQAA//8AAA=='
  $temp = "$env:Temp\f.ico"
  $ico = [convert]::FromBase64String($base64string)
  [io.file]::WriteAllBytes($temp, $ico)
  $syncHash.Window.Icon = $temp
  $syncHash.txbFindUser.IsFocused = $true
  $syncHash.txbAdminName.Text=$env:USERNAME
  $syncHash.Window.ShowDialog() | Out-Null
  $syncHash.Error = $Error
} #end GuiRunspaceScript

#Fire up the background runspace
Add-Type -AssemblyName PresentationFramework
$syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
$GuiRunspace = [powershell]::Create().AddScript($GuiRunspaceScript)
$GuiRunspace.Runspace = $newRunspace
$GuiTool = $GuiRunspace.BeginInvoke()
