# Set custom properties to Word file
# Kosarev Albert, 2016

param(
	[string]$InputDir = "C:\test\"
	)

$PropertiesArray = @{
    "Property1" = "Value"
    "Property2" = "Value"
}


# $InputFileName = ""
$WordFiles = Get-ChildItem -Path $InputDir "*.docx"

$WordApp = New-Object -ComObject "Word.Application"
$WordApp.Visible = $False
    
Function AddPropertyToWord ($CustomPropertyName, $CustomPropertyValue, $InputDoc) {
    $CustomProps = $InputDoc.CustomDocumentProperties
    $TypeCustomProps = $CustomProps.GetType()
    $binding = "System.Reflection.BindingFlags" -as [type]
    # 4 for strings
    [array]$ArrayArgs = $CustomPropertyName, $False, 4, $CustomPropertyValue
    Try {
        $TypeCustomProps.InvokeMember("add", $binding::InvokeMethod, $null, $CustomProps, $ArrayArgs) | Out-Null
        Write-Host -ForegroundColor Green ("SUCCESS :: Property added to file:  " + $CustomPropertyName + " = " + $CustomPropertyValue )
    } Catch [System.Exception] {
        $PropertyObject = $TypeCustomProps.InvokeMember("Item", $binding::GetProperty, $null, $CustomProps, $CustomPropertyName)
        $TypeCustomProps.InvokeMember("Delete", $binding::InvokeMethod, $null, $PropertyObject, $null)
        $TypeCustomProps.InvokeMember("add", $binding::InvokeMethod, $null, $CustomProps, $ArrayArgs) >null
		if ($?)
		{
			Write-Host -ForegroundColor Green ("SUCCESS :: Property changed in file:  " + $CustomPropertyName + " = " + $CustomPropertyValue )     
		}
    }
    

}

ForEach ($WordFile in $WordFiles) {	
    $WordDoc = $WordApp.Documents.Open($WordFile.FullName)
    Write-Host ("File opened: " + $WordFile.FullName)
    # Close word without saving
    ForEach ($Prop in $PropertiesArray.GetEnumerator()) {
        AddPropertyToWord $($Prop.Name) $($Prop.Value) $WordDoc
    }
    Write-Verbose -Message "Updating..."
    $WordDoc.Fields.Update() | Out-Null
    Write-Verbose -Message "Saving..."
    $WordDoc.Saved = $False
    $WordDoc.Save()
    $WordDoc.Close()
    Write-Verbose ("File has been saved: " + $WordFile.FullName)
    Write-Host
}


$WordApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApp)
$WordApp = $null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()