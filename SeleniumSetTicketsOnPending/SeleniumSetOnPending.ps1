[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function Get-Creds ($title,$user, $password, $domain1) {
###################Load Assembly for creating form & button######
[void][System.Reflection.Assembly]::LoadWithPartialName( “System.Windows.Forms”)
[void][System.Reflection.Assembly]::LoadWithPartialName( “Microsoft.VisualBasic”)

$signature=@'

      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]

      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);

'@

$SendMouseClick = Add-Type  -name "Win32MouseEventNew" -memberDefinition $signature -namespace Win32Functions -passThru
 

#####Define the form size & placement
$form = New-Object “System.Windows.Forms.Form”;
$form.Width = 500;
$form.Height = 160;
$form.Text = $title;
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
##############Define text label1
$textLabel1 = New-Object “System.Windows.Forms.Label”;
$textLabel1.Left = 25;
$textLabel1.Top = 15;
$textLabel1.Text = $user;
##############Define text label2
$textLabel2 = New-Object “System.Windows.Forms.Label”;
$textLabel2.Left = 25;
$textLabel2.Top = 50;
$textLabel2.Text = $password;
##############Define text label3
$textLabel3 = New-Object “System.Windows.Forms.Label”;
$textLabel3.Left = 25;
$textLabel3.Top = 85;
$textLabel3.Text = $domain1;
############Define text box1 for input
$textBox1 = New-Object “System.Windows.Forms.TextBox”;
$textBox1.Left = 150;
$textBox1.Top = 10;
$textBox1.width = 200;
############Define text box2 for input
$textBox2 = New-Object “System.Windows.Forms.TextBox”;
$textBox2.Left = 150;
$textBox2.Top = 50;
$textBox2.width = 200;
############Define text box3 for input
$textBox3 = New-Object “System.Windows.Forms.TextBox”;
$textBox3.Left = 150;
$textBox3.Top = 90;
$textBox3.width = 200;
#############Define default values for the input boxes
$defaultValue1 = “”
$defaultValue2 = “”
$defaultValue3 = “”
$textBox1.Text = $defaultValue1;
$textBox2.Text = $defaultValue2;
$textBox3.Text = $defaultValue3;
#############define OK button
$button = New-Object “System.Windows.Forms.Button”;
$button.Left = 360;
$button.Top = 85;
$button.Width = 100;
$button.Text = “Ok”;
############# This is when you have to close the form after getting values
$eventHandler = [System.EventHandler]{
$textBox1.Text;
$textBox2.Text;
$textBox3.Text;
$form.Close();};
$button.Add_Click($eventHandler) ;
#############Add controls to all the above objects defined
$form.Controls.Add($button);
$form.Controls.Add($textLabel1);
$form.Controls.Add($textLabel2);
$form.Controls.Add($textLabel3);
$form.Controls.Add($textBox1);
$form.Controls.Add($textBox2);
$form.Controls.Add($textBox3);
$ret = $form.ShowDialog();
#################return values
return $textBox1.Text, $textBox2.Text, $textBox3.Text
}

if (Get-Module -ListAvailable -Name Selenium) {
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist"
    Install-Module Selenium
    Import-Module Selenium
     }

$return = Get-Creds “Please provide login info:” “User email:” “Password:” “Domain:”
$userS = $return[0]
$passS = $return[1]
$domainS = $return[2]


$seleniumDir = "C:\Program Files\WindowsPowerShell\Modules\Selenium\1.1\assemblies"
Add-Type -Path "$seleniumDir\WebDriver.dll"
Add-Type -Path "$seleniumDir\WebDriver.Support.dll"

$signature=@'

      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]

      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);

'@
$SendMouseClick = Add-Type  -name "Win32MouseEventNew" -memberDefinition $signature -namespace Win32Functions -passThru
#$seleniumOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
#$seleniumOptions.AddArguments(@('--start-maximized', '--allow-running-insecure-content', '--disable-infobars', '--enable-automation'))



#$driver = New-Object OpenQA.Selenium.Chrome.Chromedriver


$chrome = New-Object OpenQA.Selenium.Chrome.ChromeDriver "$seleniumDir"

$driverWait = New-Object -TypeName OpenQA.Selenium.Support.UI.WebDriverWait($chrome, (New-TimeSpan -Seconds 10))

$chrome.manage().Window.Maximize();

$website1 =  "https://smartdesk.hcl.com"

$chrome.Navigate().GoToUrl($website1);

$Browser = $chrome

[OpenQA.Selenium.Support.UI.WebDriverWait]$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait ($Browser,[System.TimeSpan]::FromSeconds(14))

$wait.PollingInterval = 100



<#
try {
  [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Id('i0116')))
} catch [exception]{
  Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
}
#>
$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('i0116')))
#Login
$user = $Browser.FindElements([OpenQA.Selenium.By]::Id('i0116'))
$user.SendKeys($userS)

<#
try {
  [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Id('idSIButton9')))
} catch [exception]{
  Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
}
#>
$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('idSIButton9')))
$button = $Browser.FindElements([OpenQA.Selenium.By]::Id('idSIButton9'))
if ($button -ne $null) {$button.Click();}




<#
try {
  [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Id('idChkBx_PWD_KMSI0Pwd')))
} catch [exception]{
  Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
}
#>
$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('idChkBx_PWD_KMSI0Pwd')))
$checkbox = $Browser.FindElements([OpenQA.Selenium.By]::Id('idChkBx_PWD_KMSI0Pwd'))
if ($checkbox -ne $null) { $checkbox.Click() }



<#
try {
  [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Id('i0118')))
} catch [exception]{
  Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
}
#>
$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('i0118')))
$pass = $Browser.FindElements([OpenQA.Selenium.By]::Id('i0118'))
$pass.SendKeys($passS)



try {
  [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Id('idSIButton9')))
} catch [exception]{
  Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
}

$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('idSIButton9')))
$button_Signin = $Browser.FindElements([OpenQA.Selenium.By]::Id('idSIButton9'))
$button_Signin[0].Click() 


$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Id('idSIButton9')))
$button_StaySigned = $Browser.FindElements([OpenQA.Selenium.By]::Id('idSIButton9'))
if ($button_StaySigned[0] -ne $null) { $button_StaySigned[0].Click() } 

Start-Sleep -seconds 35 ;

WHILE ( 1 )
{

$driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::ClassName('BaseTableInner')))
$table = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('BaseTableInner'))
$rows =  $table.FindElements([OpenQA.Selenium.By]::TagName("tr"))

foreach ($line in $rows)
 {
   if ($line.Text -like "*INC000*")
    {
     if ($line.Text -notlike "*Pending*")
      {
        $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
        [array]$arr = $line.Text[0..12]
        [string]$name = [system.String]::Join("", $arr)
        $name
        $line.Click()
        $selectedline = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('SelPrimary'))
        $selectedline[0].Click()
        #$selectedline[0].Click()
        $point = $selectedline[0].Location
        
        $X = $point.X + 100; 
        $Y = $point.Y + 120 ;  ## deducting chrome favourites bar
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X , $Y)
        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
        
        Start-Sleep -Seconds 16 ;
        
        try 
        {
        [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_3_7"]/div')))
        } 
        catch [exception]
        {
        Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
        }
        $pendingCheckBox = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_3_7"]/div')) 
        $pendingCheckBox.Click()
        $pendingTable = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('MenuTable'))
        $rowsPending =  $pendingTable.FindElements([OpenQA.Selenium.By]::TagName("tr"))
        $pendingVal = $rowsPending[3]
        $pendingVal.Click()       
           
        Start-Sleep -Seconds 2
        try 
        {
        [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_3_1000000881"]')))
        } 
        catch [exception]
        {
        Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
        }
        $statusReasonCheckBox = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_3_1000000881"]')) 
        $statusReasonCheckBox.Click()

        Start-Sleep -Seconds 2

        $statusReasonTable = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('MenuTable'))
        $rowsStatusReason =  $statusReasonTable.FindElements([OpenQA.Selenium.By]::TagName("tr"))
        $statusVal = $rowsStatusReason[14]
        $statusVal.Click()

        Start-Sleep -Seconds 2


        $button_Save = $Browser.FindElements([OpenQA.Selenium.By]::Id('WIN_3_301614800'))
        if ($button_Save[0] -ne $null) { $button_Save[0].Click() } 
        
        Start-Sleep -Seconds 2
                
        $driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_0_304248710"]/fieldset/div/dl/dd[3]/span[2]/a')))       
        $IThome = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_0_304248710"]/fieldset/div/dl/dd[3]/span[2]/a'))
        $TimeToResolve = $stopwatch.Elapsed.TotalSeconds
        Write-Output " $name set on pending in $TimeToResolve Seconds "
        $IThome[0].Click()
        Start-Sleep -Seconds 2
        }
    }

    if ($line.Text -like "*WO000*")
    {
     if ($line.Text -notlike "*Pending*")
      {
        $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
        [array]$arr = $line.Text[0..12]
        [string]$name = [system.String]::Join("", $arr)
        $name
        $line.Click()
        $selectedline = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('SelPrimary'))
        $selectedline[0].Click()
        $selectedline[0].Click()
        $point = $selectedline[0].Location
        
        $X = $point.X + 100; 
        $Y = $point.Y + 120 ;  ## deducting chrome favourites bar
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X , $Y)
        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
        
        Start-Sleep -Seconds 20 ;
        
        try 
        {
        [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_4_7"]/div')))
        } 
        catch [exception]
        {
        Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
        }
        $pendingCheckBox = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_4_7"]/div')) 
        $pendingCheckBox.Click()
        $pendingTable = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('MenuTable'))
        $rowsPending =  $pendingTable.FindElements([OpenQA.Selenium.By]::TagName("tr"))
        $pendingVal = $rowsPending[1]
        $pendingVal.Click()       
           
        Start-Sleep -Seconds 2
        try 
        {
        [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_3_1000000881"]')))
        } 
        catch [exception]
        {
        Write-Output ("Exception with {0}: {1} ...`n(ignored)" -f $id1,(($_.Exception.Message) -split "`n")[0])
        }
        $statusReasonCheckBox = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_4_1000000881"]')) 
        $statusReasonCheckBox.Click()

        Start-Sleep -Seconds 2

        $statusReasonTable = $Browser.FindElements([OpenQA.Selenium.By]::ClassName('MenuTable'))
        $rowsStatusReason =  $statusReasonTable.FindElements([OpenQA.Selenium.By]::TagName("tr"))
        $statusVal = $rowsStatusReason[4]
        $statusVal.Click()

        Start-Sleep -Seconds 2


        $button_Save = $Browser.FindElements([OpenQA.Selenium.By]::Id('WIN_4_300000300'))
        if ($button_Save[0] -ne $null) { $button_Save[0].Click() } 
        
        Start-Sleep -Seconds 5
              
        $driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_0_304248710"]/fieldset/div/dl/dd[3]/span[2]/a')))       
        $IThome = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('//*[@id="WIN_0_304248710"]/fieldset/div/dl/dd[3]/span[2]/a'))
        $TimeToResolve = $stopwatch.Elapsed.TotalSeconds
        Write-Output " $name set on pending in $TimeToResolve Seconds "
        $IThome[0].Click()
        Start-Sleep -Seconds 2
        }
       }
      }
      Start-Sleep -Seconds 20
      $driverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::Xpath('/html/body/div[4]/div[2]/table/tbody/tr[1]/td[2]')))       
      $IThome = $Browser.FindElements([OpenQA.Selenium.By]::Xpath('/html/body/div[4]/div[2]/table/tbody/tr[1]/td[2]'))
      $IThome[0].Click()
      Start-Sleep -Seconds 2
 }

