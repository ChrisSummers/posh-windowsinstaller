$WindowsInstaller = New-Object -com WindowsInstaller.Installer
$WindowsInstaller.psobject.TypeNames[0] = "WindowsInstaller.Installer"
$WindowsInstaller | Add-Member -Name OpenDatabase -MemberType ScriptMethod -Value {
    Param (
        [IO.FileInfo] $FilePath
    )
    # Build the Database object
    
    $db = $this.GetType().InvokeMember("OpenDatabase", [System.Reflection.BindingFlags]::InvokeMethod, $Null, $this, @($FilePath.FullName, 0))
    #$db.psobject.TypeNames[0] = "WindowsInstaller.Installer.Database"
    
    $db | Add-Member -Name GenerateTransform -MemberType ScriptMethod -Value {
        Write-Error "GenerateTransform function has not yet been implemented." -Category NotImplemented
    }
    
    $db | Add-Member -Name OpenView -MemberType ScriptMethod -Value {
        Param (
            [string] $Query
        )
        # Build the View object
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa372518(v=vs.85).aspx

        $view = $this.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $this, ($Query))
        $view.psobject.TypeNames[0] = "WindowsInstaller.Installer.View"
        $view | Add-Member -Name "Execute" -MemberType ScriptMethod -Value {
            $this.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $this, $Null)
        
        }

        $view | Add-Member -Name ColumnInfo -MemberType ScriptMethod -Value {
            Param (
                $arg
            )
            # Returns Record object containing column names or column data types.
            # http://msdn.microsoft.com/en-us/library/windows/desktop/aa372502(v=vs.85).aspx
            
            # Params
            # msiColumnInfoNames = 0
            # msiColumnInfoTypes = 1
            
            $record = $this.GetType().InvokeMember("ColumnInfo", "GetProperty", $null, $this, $arg)
            #$record.psobject.TypeNames[0] = "WindowsInstaller.Installer.Record"
            $record | Add-Member -Name "StringData" -MemberType ScriptMethod -Value {
                Param(
                    [int]$Column
                )
                $this.GetType().InvokeMember("StringData", "GetProperty", $Null, $this, $Column)
            }
            $record | Add-Member -Name "IntegerData" -MemberType ScriptMethod -Value {
                Param(
                    [int]$Column
                )
                $this.GetType().InvokeMember("IntegerData", "GetProperty", $Null, $this, $Column)
            }
            $record | Add-Member -Name FieldCount -MemberType ScriptProperty -Value {
                $this.GetType().InvokeMember("FieldCount", "GetProperty", $Null, $this, $Null)
            }

            return $record

        }

        $view | Add-Member -Name "Fetch" -MemberType ScriptMethod -Value {
            # Build the Record object
            # http://msdn.microsoft.com/en-us/library/windows/desktop/aa371136(v=vs.85).aspx

            $record = $this.GetType().InvokeMember("Fetch", "InvokeMethod", $Null, $this, $Null)
            
            $record | Add-Member -Name "StringData" -MemberType ScriptMethod -Value {
                Param(
                    [int]$Column
                )
                $this.GetType().InvokeMember("StringData", "GetProperty", $Null, $this, $Column)
            }
            $record | Add-Member -Name "IntegerData" -MemberType ScriptMethod -Value {
                Param(
                    [int]$Column
                )
                $this.GetType().InvokeMember("IntegerData", "GetProperty", $Null, $this, $Column)
            }
            $record | Add-Member -Name FieldCount -MemberType ScriptProperty -Value {
                $this.GetType().InvokeMember("FieldCount", "GetProperty", $Null, $this, $Null)
            }
            #$record.psobject.TypeNames[0] = "WindowsInstaller.Installer.Record"
            return $record
        }
        
        return $view
    }

    $db | Add-Member -Name "SummaryInformation" -MemberType ScriptProperty -Value {
        $SummaryInfo = $this.GetType().InvokeMember("SummaryInformation", "GetProperty", $Null, $this, $null)
        $SummaryInfo | Add-Member -Name "Property" -MemberType ScriptMethod -Value {
            Param (
                [int]$PropertyId
            )

            $this.GetType().InvokeMember("Property", "GetProperty", $null, $this, $PropertyId)

        }
        $SummaryInfo


    }
    return $db

}

$WindowsInstaller | Add-Member -Name "Version" -MemberType ScriptProperty -Value {
    $this.GetType().InvokeMember("Version", "GetProperty", $Null, $this, $Null)
}


Function Get-MsiTable {
    [CmdletBinding()]
    Param (
        [string]$table,
        $database,
        [string]$where
    )

    $rows = @()
    $query = "SELECT * FROM $table"
    if ($where) {
        $query = $query + " WHERE $where"
    }
    
    $view = $database.OpenView($query)
    [void]$view.Execute()

    $colNamesRecord = $view.ColumnInfo(0)
    $colCount = $colNamesRecord.FieldCount
    $colNames = @()
    for ($i=1; $i -le $colCount; $i++) {
        $colNames += $colNamesRecord.StringData($i)

    }
    $colTypesRecord = $view.ColumnInfo(1)
    $colTypes = @()
    for ($i=1; $i -le $colCount; $i++) {
        $colTypes += $colTypesRecord.StringData($i)

    }

    Write-Debug $colNames.Count
    
    while ($true) {
        $record = $view.Fetch()
        if ($record) {
            $row = New-Object "psobject"
            for ($i=0; $i -lt $colNames.Count; $i++) {
                $colName = $colNames[$i]

                #Write-Debug ("$colName is {0}" -f $colTypes[$i])

                if ($colTypes[$i].Chars(0) -eq "i") {
                    $colValue = [int]$record.IntegerData($i + 1)
                } else {
                    $colValue = $record.StringData($i + 1)
                }
                Write-Debug "$colname = $colValue"
                
                $row | Add-Member -Name $colName -MemberType NoteProperty -Value $colValue
                #Write-Debug ("Row: {0}" -f $row | gm)
            }
            
            $rows += $row
        } else {
            break
        }
    }
    
    return $rows

}

Function Get-MsiSummaryInfo {
    Param(
        $database
    )


    $info = $database.SummaryInformation
    [pscustomobject]@{
        Codepage = $info.Property(1)
        Title = $info.Property(2)
        Subject = $info.Property(3)
        Author = $info.Property(4)
        Keywords = $info.Property(5)
        Comments = $info.Property(6)
        Template = $info.Property(7)
        LastSavedBy = $info.Property(8)
        RevisionNumber = $info.Property(9)
        LastPrinted = $info.Property(11)
        CreateTimeDate = $info.Property(12)
        LastSaveTimeDate = $info.Property(13)
        PageCount = $info.Property(14)
        WordCount = $info.Property(15)
        CharacterCount = $info.Property(16)
        CreatingApplication = $info.Property(18)
        Security = $info.Property(19)
    }
}

Function Get-Msi {
    param (
        [IO.FileInfo] $FilePath
    )

    If (-not (Test-Path $FilePath)) {
        Write-Error "File not found" -Category ObjectNotFound -TargetObject $FilePath.Name
        return
    }

    if ($FilePath.Extension -ne ".msi") {
        Write-Error "File is not a Windows Installer file." -Category InvalidOperation -TargetObject $FilePath.Name
        return
    }
    
    Try {
        #$db = $WindowsInstaller.OpenDatabase($FilePath)
    } 
    Catch {
        Write-Error "Error opening Windows Installer datbase." -Category InvalidData -TargetObject $FilePath.Name
        return
    }
    
     
    $db = $WindowsInstaller.OpenDatabase($FilePath)

    $Properties = Get-MsiTable -table "Property" -database $db
    $PublicProperties = $Properties | %{if ($_.Property.CompareTo($_.Property.ToUpper()) -eq 0) {$_}}

    $prophash = @{}
    $Properties | %{$prophash.Add($_.Property,$_.Value)}

    $pubhash = @{}
    $PublicProperties | %{$pubhash.Add($_.Property,$_.Value)}

    
    $obj = "" | Select ProductName, ProductVersion, Manufacturer, ProductCode, UpgradeCode, TargetDir, Path, Properties, PublicProperties, Database
    $obj.Properties = $prophash
    $obj.ProductName = $obj.Properties.ProductName
    $obj.ProductVersion = $obj.Properties.ProductVersion
    $obj.Manufacturer = $obj.Properties.Manufacturer
    $obj.ProductCode = $obj.Properties.ProductCode
    $obj.UpgradeCode = $obj.Properties.UpgradeCode
    $obj.TargetDir = $obj.Properties.TargetDir
    $obj.Path = $FilePath
    $obj.PublicProperties = $pubhash
    $obj.Database = $db
    #$obj.UIProperties = Get-MsiTable -table "Control" -database $db | %{if ($_.Property -and $_.Property.CompareTo($_.Property.ToUpper()) -eq 0) {$_}} | Select Dialog_, Control, Type, Property

    $defaultProperties = @("ProductName", "ProductVersion", "ProductCode", "UpgradeCode", "TargetDir", "Path")
    $DefaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet",[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]$DefaultDisplayPropertySet 

    # Attach default display property set
    $obj | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers
    $obj | Add-Member -MemberType ScriptMethod -Name GetTable -Value {
        Param(
            [string]$table
        )
        Get-MsiTable -table $table -database $this.Database
    }
    $obj | Add-Member -MemberType ScriptProperty -Name UIProperties -Value {
        Get-MsiTable -table "Control" -database $this.Database | %{if ($_.Property -and $_.Property.CompareTo($_.Property.ToUpper()) -eq 0) {$_}} | Select Property | Sort Property -Unique
    }

    $obj | Add-Member -MemberType ScriptMethod -Name GetPublicPropertyValues -Value {
        Param (
            [string]$property
        )

        if ($property.CompareTo($property.ToUpper()) -eq 0) {
            $listbox = Get-MsiTable -table "ListBox" -database $this.Database -where "Property = `'$property`'" | Select Value, @{name="Description"; ex={$_.Text}}
            $listview = Get-MsiTable -table "ListView" -database $this.Database -where "Property = `'$property`'" | Select Value, @{name="Description"; ex={$_.Text}}
            $radiobutton = Get-MsiTable -table "RadioButton" -database $this.Database -where "Property = `'$property`'" | Select Value, @{name="Description"; ex={$_.Text}}
            $checkbox = Get-MsiTable -table "CheckBox" -database $this.Database -where "Property = `'$property`'" | Select Value, @{name="Description"; ex={$_.Text}}
            $combobox = Get-MsiTable -table "CheckBox" -database $this.Database -where "Property = `'$property`'" | Select Value, @{name="Description"; ex={$_.Text}}


        } else {
            Write-Error "Must supply a PUBLIC property" -Category InvalidArgument
            return
        }
        $property
        "Default Value: {0}" -f $this.Properties.$property
        $listbox
        $listview
        $radiobutton
    }
    

    $obj | Add-Member -MemberType ScriptProperty -Name "SummaryInformation" -Value {
        Get-MsiSummaryInfo -database $this.Database
    }

    $obj | Add-Member -MemberType ScriptProperty -Name "Platform" -Value {
        $this.SummaryInformation.Template.Split(";")[0]
    }

    $obj | Add-Member -MemberType ScriptProperty -Name "Language" -Value {
        $this.SummaryInformation.Template.Split(";")[1].Split(",") | %{try{[System.Globalization.CultureInfo]::GetCultureInfo([int]$_)}catch{}}
    }

    return $obj
}



Function Get-MsiReport {
    Param (
        $msi
    )
    $html = "<!DOCTYPE html>`n<html>"
    $html += "<head><title>{0} {1}</title></head>`n" -f $msi.ProductName, $msi.ProductVersion
    $html += "<body>`n"
    $html += "<div id='General'>{0}</div>"
    $html += "</body>`n"
    $html += "</html>"

    return $html
}

# Useful stuff
# Public Properties in the UI
# Get-MsiTable -table "Control" -database $db | %{if ($_.Property -and $_.Property.CompareTo($_.Property.ToUpper()) -eq 0) {$_}} | Select Dialog_, Control, Type, Property

# Public Properties in custom actions
# Get-MsiTable -table "CustomAction" -database $db | %{if ((-not $_.Source.Contains(".")) -and $_.Source -and $_.Source.CompareTo($_.Source.ToUpper()) -eq 0) {$_}} | Select Action, Source, Target

# Get-MsiTable -table "Control" -database $db | ?{$_.Property -eq "WORKINGDIR_MODE"} | Select Dialog_, Control, Type, Property
# Get-MsiTable -table "RadioButton" -database $db | ?{$_.Property -eq "WORKINGDIR_MODE"} | Select Text, Value



#Get-ChildItem "*.msi" -Recurse | %{Get-Msi -FilePath $_ | Select Manufacturer, ProductName, ProductVersion, ProductCode, UpgradeCode}
# $env:ProgramFiles



# TODO
# Build InstallDir property
# Read transform
# Create transform
# Get-MsiReport
# Copyright
# Sign
# fix hash for duplicate properties
# Finsih summaryinfo translation