#region begin block of the main script	

	$scriptFolder = Split-Path $MyInvocation.MyCommand.Path -parent

	$SystemIOPackagingAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
	If ($asm.GetName().Name -eq 'System.IO.Packaging') {
			$SystemIOPackagingAssemblyLoaded = $True
		}
	}
	$DocumentFormatOpenXmlAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
		If ($asm.GetName().Name -eq 'DocumentFormat.OpenXml') {
			$DocumentFormatOpenXmlAssemblyLoaded = $True
		}
	}	
	$SpreadsheetLightAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
		If ($asm.GetName().Name -eq 'SpreadsheetLight') {
			$SpreadsheetLightAssemblyLoaded = $True
		}
	}

	if( $SystemIOPackagingAssemblyLoaded -eq $false ){
		$assemblyPath = Join-Path $scriptFolder "System.IO.Packaging.dll"
		$Null = [Reflection.Assembly]::LoadFile($assemblyPath)
	}
	if( $DocumentFormatOpenXmlAssemblyLoaded -eq $false ){
		$assemblyPath = Join-Path $scriptFolder "DocumentFormat.OpenXml.dll"
		$Null = [Reflection.Assembly]::LoadFile($assemblyPath)
	}
	if( $SpreadsheetLightAssemblyLoaded -eq $false ){
		$assemblyPath = Join-Path $scriptFolder "SpreadsheetLight.dll"
		$Null = [Reflection.Assembly]::LoadFile($assemblyPath)
	}
 
	# [Reflection.Assembly]::LoadWithPartialName throws no exception so we test if assembly is loaded
	$SystemIOPackagingAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
		If ($asm.GetName().Name -eq 'System.IO.Packaging') {
			$SystemIOPackagingAssemblyLoaded = $True
		}
	}
	$DocumentFormatOpenXmlAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
		If ($asm.GetName().Name -eq 'DocumentFormat.OpenXml') {
			$DocumentFormatOpenXmlAssemblyLoaded = $True
		}
	}	
	$SpreadsheetLightAssemblyLoaded = $False
	ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
		If ($asm.GetName().Name -eq 'SpreadsheetLight') {
			$SpreadsheetLightAssemblyLoaded = $True
		}
	}

# if assembly could not be loaded throw Erro
# if assembly could not be loaded throw ErrorRecord
If( -not ( `
		         ($SystemIOPackagingAssemblyLoaded -eq $true)  `
		    -and ($DocumentFormatOpenXmlAssemblyLoaded -eq $true)  `
		    -and ($SpreadsheetLightAssemblyLoaded -eq $true)  `
		) `
	) {
	# create custom ErrorRecord
	$message = "Could not load required assemblies"
	$exception = New-Object System.IO.FileNotFoundException $message
	$errorID = 'AssemblyFileNotFound'
	$errorCategory = [Management.Automation.ErrorCategory]::NotInstalled
	$target = ( "Assembly Status, SystemIOPackagingAssemblyLoaded:{0}, DocumentFormatOpenXmlAssemblyLoaded:{1}, SpreadsheetLightAssemblyLoaded:{2}" -f  $SystemIOPackagingAssemblyLoaded, $DocumentFormatOpenXmlAssemblyLoaded, $SpreadsheetLightAssemblyLoaded )
	$errorRecord = New-Object Management.Automation.ErrorRecord $exception,$errorID,$errorCategory,$target
	# throw terminating error
	$PSCmdlet.ThrowTerminatingError($errorRecord)
	# leave script
	return
}
#endregion

#region	declare (helper) functions
####################### 
function Get-DataTableType 
{ 
	[Parameter(Mandatory=$True,
		Position=0,
		ValueFromPipeline=$True,
		ValueFromPipelinebyPropertyName=$True
	)]
	[System.String]$TypeName,
 
	$TypeNames = @{ 
		'System.Boolean'='yes\no'
	; 'System.Byte[]'='HexBinary'
	; 'System.Byte'='0'
	; 'System.Char'='inlineStr'
	; 'System.Datetime'='inlineStr'
	; 'System.Decimal'='0.00_);\(0.00\)'
	; 'System.Double'='0.00_);\(0.00\)'
	; 'System.Guid'='inlineStr'
	; 'System.Int16'='0'
	; 'System.Int32'='0'
	; 'System.Int64'='0'
	; 'System.Single'='0'
	; 'System.UInt16'='0'
	; 'System.UInt32'='0'
	; 'System.UInt64'='0'}
 
	if ( $types -contains $type ) { 
		Write-Output "$type" 
	} 
	else { 
		Write-Output 'System.String' 
         
	} 
} #Get-Type 

Function Test-FileLocked {
# Some file action needs exclusive access by the calling process, so Windows will lock access to a file.
# This Function can detect if a file is locked or not
# If PassThru is not given then the function returns $True if the File is locked Else it returns $False
# If PassThru is given then the function returns a Management.Automation.ErrorRecord Object if the File is locked Else it returns $Null
# the file could become locked the very next millisecond after this check by any other process!
# Peter Kriegel 22.August.2013 Version 1.0.0

	param(
		[Parameter(Mandatory=$True,
			Position=0,
			ValueFromPipeline=$True,
			ValueFromPipelinebyPropertyName=$True
		)]
		[String]$Path,
		[System.IO.FileAccess]$FileAccessMode = [System.IO.FileAccess]::Read,
		[Switch]$PassThru
			
	)
	   
	If(Test-Path $Path) {
		$FileInfo = Get-Item $Path
	} Else {
		Return $False
	}
		
	try
	{
	    $Stream = $FileInfo.Open([System.IO.FileMode]::Open, $FileAccessMode, [System.IO.FileShare]::None)
			
	}
	catch [System.IO.IOException]
	{
	    #the file is unavailable because it is:
	    #still being written to or being processed by another thread
	    #or does not exist (has already been processed)
		If($PassThru.IsPresent) {
		    #Return $_.Exception
						
			$message = $_.Exception.Message
			$exception = $_.Exception
			$errorID = 'FileIsLocked'
			$errorCategory = [Management.Automation.ErrorCategory]::OpenError
			$target = $Path
			$errorRecord = New-Object Management.Automation.ErrorRecord $exception, $errorID,$errorCategory,$target
			Return $errorRecord

		} Else {
		    Return $True
		}
			
	}
	finally
	{
	    if ($stream){
			$stream.Close()
		}
	}

	#file is not locked
	If($PassThru.IsPresent) {
		Return $Null
	} Else {
		$False
	}
}

#endregion declare (helper) functions

#region declare XLSX functions

	Function Get-XLSXWorkSheets {
	<#
		.SYNOPSIS
			List worksheets for Excel document.
			

		.DESCRIPTION
			Function returns a list of worksheet names for Excel document.

		.PARAMETER  Path
			The path of the Excel document.

		.EXAMPLE
			PS C:\> Get-XLSXWorkSheets -Path 'D:\temp\PSExcel.xlsx'
		    PS C:\> $sheetNames = ( Get-ChildItem -Filter MyDocument.xlsx | Get-XLSXWorkSheets )
			
		.INPUTS
			System.String,System.String

		.OUTPUTS
			None, unless error is reported.

		.NOTES
			Author Miron
			Version: 1.0.0 14.August.2013
	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[String]$Path
		)

		Begin {
		} # end begin block

	Process {
			#check file exist
			if($Path -eq $null){
				$message = "File path must not be null"
				$PSCmdlet.WriteError($message)
				return;
			}
			if( -not ( Test-Path $Path ) ){
				$message = "File $Path is not found";
				$PSCmdlet.WriteError($message);
				return;
			}
			# test if File is locked
			If($ErrorRecord = Test-FileLocked $Path -PassThru){
				$PSCmdlet.WriteError($ErrorRecord);
				return
			}
			Try {
				# test if the file could be accessed
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				$xlFile = Get-Item -Path $Path -ErrorAction stop
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
			Try {
			# open Excel document
			$sl = New-Object SpreadsheetLight.SLDocument( $Path );
		    #find all worksheets of the document
		    $sheetNames = $sl.GetSheetNames($true);
			$sl.CloseWithoutSaving();
			$sl.Dispose();
			$sl = $null;
			# done examening Excel document
			
			$worksheetNames = @();
		    if( ( $sheetNames-eq $null ) -or ($sheetNames.Length -eq 0 ) ){
				$message = "Document $Path contains no worksheets";
				$PSCmdlet.WriteVerbose($message);
				return;
			}
			foreach($Name in $sheetNames)
			{
				$worksheetNames += $Name;
			}

			return $worksheetNames;

		} Catch {
			# we dont want to show the ErrorRecord Object in the $Error list 
			$Error.RemoveAt(0)
			# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
			$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
			# throw the ErrorRecord Object
			$PSCmdlet.WriteError($NewError)
			# leave Function
			Return
		}
		} # end Process block
		
		End { 
			} # end End block
	}#end Function Delete-XLSXWorkSheet

	Function Ensure-XLSXWorkSheets {
	<#
		.SYNOPSIS
			It turns out SpreadsheetLight library
			will create Excel with default worksheet "sheet1".
			

		.DESCRIPTION
			Function removes worksheets, except
			those specified in the comma separated list passed in.

		.PARAMETER  Path
			The path of the Excel document.

		.PARAMETER  WorksheetNames
			Array of worksheet names.

		.EXAMPLE
			PS C:\> Ensure-XLSXWorkSheets -Path 'D:\temp\PSExcel.xlsx' -WorksheetNames "Sheet1[,...]"
			
		.INPUTS
			System.String,System.String

		.OUTPUTS
			None, unless error is reported.

		.NOTES
			Author Miron
			Version: 1.0.0 14.August.2013
	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[String]$Path,
			[String]$WorksheetNames
		)

		Begin {
		} # end begin block

	Process {
			#check file exist
			if($Path -eq $null){
				$message = "File path must not be null"
				$PSCmdlet.WriteError($message)
				return;
			}
			if( -not ( Test-Path $Path ) ){
				$message = "File $Path is not found";
				$PSCmdlet.WriteError($message);
				return;
			}
			# test if File is locked
			If($ErrorRecord = Test-FileLocked $Path -PassThru){
				$PSCmdlet.WriteError($ErrorRecord);
				return
			}
			Try {
				# test if the file could be accessed
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				$xlFile = Get-Item -Path $Path -ErrorAction stop                                
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
		    $sl = $null;
			Try {
			$WorksheetNameArray = @();
			$WorksheetNameArray1 = $WorksheetNames.Split(",");
			if( ( $WorksheetNameArray1 -ne $null ) -and ( $WorksheetNameArray1.Length -gt 0 ) )
			{
				$WorksheetNameArray = $WorksheetNameArray1;
			}
			# open Excel document
			$sl = New-Object SpreadsheetLight.SLDocument( $Path );
		    #find all worksheets of the document
		    $sheetNames = $sl.GetSheetNames($true);
		    if( ( $sheetNames-eq$null ) -or ($sheetNames.Length -eq 0 ) ){
				$message = "Document $Path contains no worksheets";
				$PSCmdlet.WriteVerbose($message);
				return;
			}
            foreach($Name in $sheetNames)
			{
				if( $Name -in $WorksheetNameArray)
				{
					$null = $sl.SelectWorksheet($Name);
				}
			}
			foreach($Name in $sheetNames)
			{
				#is the worksheet of interest really present?
				if( -not( $Name -in $WorksheetNameArray )){
					#remove as per requested
					Write-Output "Worksheet $Name is about to be delieted."
					$sl.DeleteWorksheet($Name);
				}
			}
			$sl.Save();
			$sl.Dispose();
			$sl = $null;
		} Catch {
			# we dont want to show the ErrorRecord Object in the $Error list 
			$Error.RemoveAt(0)
			# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
			$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
			# throw the ErrorRecord Object
			$PSCmdlet.WriteError($NewError)

			if( $sl -ne $null ) ## most likely Excel object is there
			{
			    $sl.Dispose();
			    $sl = $null;
			}

			# leave Function
			Return
		}
		} # end Process block
		
		End { 
			} # end End block
	}#end Function Delete-XLSXWorkSheet

	Function Delete-XLSXWorkSheet {
	<#
		.SYNOPSIS
			Function to append an new empty Excel worksheet to an existing Excel .xlsx workbook

		.DESCRIPTION
			Function to append an new empty Excel worksheet to an existing Excel .xlsx workbook

		.PARAMETER  Path
			Path to the existing Excel .xlsx workbook

		.PARAMETER  Name
			The Name of the new Excel worksheet to create
			Worksheet Names must be unique to an Excel workbook.
			If you provide an allready existing worksheet name an warning is generated and a automatic name is used.
			If you dont provide an worksheet name an unique name is autmaticaly generate with the pattern Table + number (eg: Table21)

		.EXAMPLE
			PS C:\> Add-XLSXWorkSheet -Path 'D:\temp\PSExcel.xlsx' -Name "Willy"
			
			Adds a new worksheet with the Name Willy to the Excel workbook stored in path 'D:\temp\PSExcel.xlsx'

		.EXAMPLE
			PS C:\> Add-XLSXWorkSheet -Path 'D:\temp\PSExcel.xlsx'
			
			Adds a new worksheet with the automatic generate name to the Excel workbook stored in path 'D:\temp\PSExcel.xlsx'

		.INPUTS
			System.String,System.String

		.OUTPUTS
			PSObject with Properties: 
				Uri # The Uri of the new generated worksheet in the XLSX package
				WorkbookRelationID # XLSX package relationship ID to the from Workbook to the worksheet
				Name # Name of the worksheet 
				WorkbookPath # Path to the Excel workbook (XLSX package) which holds the worksheet

		.NOTES
			Author PowerShell:  Peter Kriegel, Germany http://www.admin-source.de
			Version: 1.0.0 14.August.2013
	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[String]$Path,
			[String]$Name
		)

		Begin {
		} # end begin block

	Process {
			#check file exist
			if($Path -eq $null){
				$message = "File path must not be null"
				$PSCmdlet.WriteError($message)
				return;
			}
		    if($Name -eq $null){
				$message = "Worksheet name must not be null"
				$PSCmdlet.WriteError($message)
				return;
			}
			if( -not ( Test-Path $Path ) ){
				$message = "File $Path is not found";
				$PSCmdlet.WriteError($message);
				return;
			}
			# test if File is locked
			If($ErrorRecord = Test-FileLocked $Path -PassThru){
				$PSCmdlet.WriteError($ErrorRecord);
				return
			}
			Try {
				# test if the file could be accessed
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				$xlFile = Get-Item -Path $Path -ErrorAction stop                               
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
			Try {
				# open Excel document
				$sl = New-Object SpreadsheetLight.SLDocument( $Path );
				#find all worksheets of the document
				$sheetNames = $sl.GetSheetNames($true);
				if( ( $sheetNames-eq$null ) -or ($sheetNames.Length -eq 0 ) ){
					$sl.CloseWithoutSaving();
					$sl.Dispose();
					$sl = $null;
					$message = "Document $Path contains no worksheets";
					$PSCmdlet.WriteVerbose($message);
					return;
				}
				#is the worksheet of interest really present?
				if($sheetNames.Contains($Name)){
					#remove as per requested
					$sl.DeleteWorksheet($Name);
					$sl.Save();
					$sl.Dispose();
					$sl = $null;
				} else {
					#inform user, no worksheet requested found
					$sl.CloseWithoutSaving();
					$sl.Dispose();
					$sl = $null;
					$message = "Document $Path contains no worksheet $Name";
					$PSCmdlet.WriteVerbose($message);
					return;
				}
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
		} # end Process block
		
		End { 
			} # end End block
	}#end Function Delete-XLSXWorkSheet

	Function Add-XLSXWorkSheet {
	<#
		.SYNOPSIS
			Function to append an new empty Excel worksheet to an existing Excel .xlsx workbook

		.DESCRIPTION
			Function to append an new empty Excel worksheet to an existing Excel .xlsx workbook

		.PARAMETER  Path
			Path to the existing Excel .xlsx workbook

		.PARAMETER  Name
			The Name of the new Excel worksheet to create
			Worksheet Names must be unique to an Excel workbook.
			If you provide an allready existing worksheet name an warning is generated and a automatic name is used.
			If you dont provide an worksheet name an unique name is autmaticaly generate with the pattern Table + number (eg: Table21)

		.EXAMPLE
			PS C:\> Add-XLSXWorkSheet -Path 'D:\temp\PSExcel.xlsx' -Name "Willy"
			
			Adds a new worksheet with the Name Willy to the Excel workbook stored in path 'D:\temp\PSExcel.xlsx'

		.EXAMPLE
			PS C:\> Add-XLSXWorkSheet -Path 'D:\temp\PSExcel.xlsx'
			
			Adds a new worksheet with the automatic generate name to the Excel workbook stored in path 'D:\temp\PSExcel.xlsx'

		.INPUTS
			System.String,System.String

		.OUTPUTS
			PSObject with Properties: 
				Uri # The Uri of the new generated worksheet in the XLSX package
				WorkbookRelationID # XLSX package relationship ID to the from Workbook to the worksheet
				Name # Name of the worksheet 
				WorkbookPath # Path to the Excel workbook (XLSX package) which holds the worksheet

		.NOTES
			Author PowerShell:  Peter Kriegel, Germany http://www.admin-source.de
			Version: 1.0.0 14.August.2013
	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[String]$Path,
			[String]$Name
		)

		Begin {

			# create worksheet XML document

			# create empty XML Document
			$New_Worksheet_xml = New-Object System.Xml.XmlDocument

	        # Obtain a reference to the root node, and then add the XML declaration.
	        $XmlDeclaration = $New_Worksheet_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
	        $Null = $New_Worksheet_xml.InsertBefore($XmlDeclaration, $New_Worksheet_xml.DocumentElement)

	        # Create and append the worksheet node to the document.
	        $workSheetElement = $New_Worksheet_xml.CreateElement("worksheet")
			# add the Excel related office open xml namespaces to the XML document
	        $Null = $workSheetElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	        $Null = $workSheetElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
	        $Null = $New_Worksheet_xml.AppendChild($workSheetElement)

	        # Create and append the sheetData node to the worksheet node.
	        $Null = $New_Worksheet_xml.DocumentElement.AppendChild($New_Worksheet_xml.CreateElement("sheetData"))

		} # end begin block

	Process {

			# test if File is locked
			If($ErrorRecord = Test-FileLocked $Path -PassThru){
				$PSCmdlet.WriteError($ErrorRecord)
				return
			}
			
			Try {
				# test if the file could be accessed
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				$xlFile = Get-Item -Path $Path -ErrorAction stop
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
			
			# open Excel .XLSX package file
			Try {
				
				$exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Open)
			} catch {
				$_
				Return
			}

			# find /xl/workbook.xml
			ForEach ($Part in $exPkg.GetParts()) {
				# remember workbook.xml 
				IF($Part.ContentType -eq "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" -or $Part.Uri.OriginalString -eq "/xl/workbook.xml") {
					$WorkBookPart = $Part
					# found workbook exit foreach loop
					break
				}
			}

			If(-not $WorkBookPart) {
				Write-Error "Excel Workbook not found in : $Path"
				$exPkg.Close()
				return
			}
			
			# get all relationships of Workbook part
			$WorkBookRels = $WorkBookPart.GetRelationships()
			
			$WorkBookRelIds = [System.Collections.ArrayList]@()
			$WorkSheetPartNames = [System.Collections.ArrayList]@()
			
			ForEach($Rel in $WorkBookRels) {
				
				# collect workbook relationship IDs in a Arraylist
				# to easely find a new unique relationship ID
				$Null = $WorkBookRelIds.Add($Rel.ID)

				# collect workbook related worksheet names in an Arraylist
				# to easely find a new unique sheet name
				If($Rel.RelationshipType -like '*worksheet*' ) {
					$WorkSheetName = Split-Path $Rel.TargetUri.ToString() -Leaf
					$Null = $WorkSheetPartNames.Add($WorkSheetName)
				}
			}
			
			# find a new unused relationship ID
			# relationship ID have the pattern rID + Number (eg: reID1, rID2, rID3 ...)
			$IdCounter = 0 # counter for relationship IDs
			$NewWorkBookRelId = '' # Variable to hold the new found relationship ID
			Do{
				$IdCounter++
				If(-not ($WorkBookRelIds -contains "rId$IdCounter")){
					# $WorkBookRelIds does not contain the rID + Number
					# so we have found an unused rID + Number; create it
					$NewWorkBookRelId = "rId$IdCounter"
				}
			} while($NewWorkBookRelId -eq '')

			# find new unused worksheet part name
			# worksheet in the package have names with the pattern Sheet + number + .xml
			$WorksheetCounter = 0 # counter for worksheet numbers
			$NewWorkSheetPartName = '' # Variable to hold the new found worksheet name
			Do{
				$WorksheetCounter++
				If(-not ($WorkSheetPartNames -contains "sheet$WorksheetCounter.xml")){
					# $WorkSheetPartNames does not contain the worksheet name
					# so we have found an unused sheet + Number + .xml; create it
					$NewWorkSheetPartName = "sheet$WorksheetCounter.xml"
				}
			} while($NewWorkSheetPartName -eq '')
			
			# Excel allows only unique WorkSheet names in a workbook
			# test if worksheet name already exist in workbook
			$WorkbookWorksheetNames = [System.Collections.ArrayList]@()

			# open the workbook.xml
			$WorkBookXmlDoc = New-Object System.Xml.XmlDocument
			# load XML document from package part stream
			$WorkBookXmlDoc.Load($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))

			# read all Sheet elements from workbook
			ForEach ($Element in $WorkBookXmlDoc.documentElement.Item("sheets").get_ChildNodes()) {
				# collect sheet names in Arraylist
				$Null = $WorkbookWorksheetNames.Add($Element.Name)
			}
			
			# test if a given worksheet $Name allready exist in workbook
			$DuplicateName = ''
			If(-not [String]::IsNullOrEmpty($Name)){
				If($WorkbookWorksheetNames -Contains $Name) {
					# save old given name to show in warning message
					$DuplicateName = $Name
					# empty name to create a new one
					$Name = ''
				}
			} 
			
			# If the user has not given a worksheet $Name or the name allready exist 
			# we try to use the automatic created name with the pattern Table + Number
			If([String]::IsNullOrEmpty($Name)){
				$WorkSheetNameCounter = 0
				$Name = "Table$WorkSheetNameCounter"
				# while automatic created Name is used in workbook.xml we create a new name
				While($WorkbookWorksheetNames -Contains $Name) {
					$WorkSheetNameCounter++
					$Name = "Table$WorkSheetNameCounter"
				}
				If(-not [String]::IsNullOrEmpty($DuplicateName)){
					Write-Warning "Worksheetname '$DuplicateName' allready exist!`nUsing automatically generated name: $Name"
				}
			}

	#region Create worksheet part
			
			# create URI for worksheet package part
			$Uri_xl_worksheets_sheet_xml = New-Object System.Uri -ArgumentList ("/xl/worksheets/$NewWorkSheetPartName", [System.UriKind]::Relative)
			# create worksheet part
			$Part_xl_worksheets_sheet_xml = $exPkg.CreatePart($Uri_xl_worksheets_sheet_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
			# get writeable stream from part 
			$dest = $part_xl_worksheets_sheet_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
			# write $New_Worksheet_xml XML document to part stream
			$New_Worksheet_xml.Save($dest)
			
			# create workbook to worksheet relationship
			$Null = $WorkBookPart.CreateRelationship($Uri_xl_worksheets_sheet_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $NewWorkBookRelId)
			
	#endregion 	Create worksheet part

	#region edit xl\workbook.xml
					
			# edit the xl\workbook.xml
			
			# create empty XML Document
			$WorkBookXmlDoc = New-Object System.Xml.XmlDocument
			# load XML document from package part stream
			$WorkBookXmlDoc.Load($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))
					
			# create a new XML Node for the sheet 
			$WorkBookXmlSheetNode = $WorkBookXmlDoc.CreateElement('sheet', $WorkBookXmlDoc.DocumentElement.NamespaceURI)
	        $Null = $WorkBookXmlSheetNode.SetAttribute('name',$Name)
	        $Null = $WorkBookXmlSheetNode.SetAttribute('sheetId',$IdCounter)
			# try to create the ID Attribute with the r: Namespace (xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships") 
			$NamespaceR = $WorkBookXmlDoc.DocumentElement.GetNamespaceOfPrefix("r")
			If($NamespaceR) {
	        	$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NamespaceR,$NewWorkBookRelId)
			} Else {
				$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NewWorkBookRelId)
			}
			
			# add the new sheet node to XML document
			$Null = $WorkBookXmlDoc.DocumentElement.Item("sheets").AppendChild($WorkBookXmlSheetNode)
		
			# Save back the edited XML Document to package part stream
			$WorkBookXmlDoc.Save($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Write))

	#endregion edit xl\workbook.xml		
			
			# close main package (flush all changes to disk)
			$exPkg.Close()
			
			# return datas of new created worksheet
			New-Object -TypeName PsObject -Property @{Uri = $Uri_xl_worksheets_sheet_xml;
													WorkbookRelationID = $NewWorkBookRelId;
													Name = $Name;
													WorkbookPath = $Path
													}
			
		} # end Process block
		
		End { 
			} # end End block
	}

	Function New-XLSXWorkBook {
	<#
		.SYNOPSIS
			Function to create a new empty Excel .xlsx workbook (XLSX package)

		.DESCRIPTION
			Function to create a new empty Excel .xlsx workbook (XLSX package)
			
			This creates an empty Excel workbook without any worksheet!
			Worksheets are mandatory for .xlsx Files! So you have to add at least one worksheet!

		.PARAMETER  Path
			Path to the Excel .xlsx workbook to create
			
		.PARAMETER  NoClobber
			Do not overwrite (replace the contents) of an existing file.
			By default, if a file exists in the specified path, New-XLSXWorkBook overwrites the file without warning.
			
		.PARAMETER Force
			Overwrites the file specified in path without prompting.		

		.EXAMPLE
			PS C:\> New-XLSXWorkBook -Path 'D:\temp\PSExcel.xlsx'
			
			Creates the new empty Excel .xlsx workbook in the Path 'D:\temp\PSExcel.xlsx' 

		.OUTPUTS
			System.IO.FileInfo

		.NOTES
			Author PowerShell:  Peter Kriegel, Germany http://www.admin-source.de
			Version: 1.0.0 14.August.2013
	#>

		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[String]$Path,
			[ValidateNotNull()]
			[Switch]$NoClobber,
			[Switch]$Force
		)

		Begin {
				
			# create the Workbook.xml part XML document
			
			# create empty XML Document
			$xl_Workbook_xml = New-Object System.Xml.XmlDocument

	        # Obtain a reference to the root node, and then add the XML declaration.
	        $XmlDeclaration = $xl_Workbook_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
	        $Null = $xl_Workbook_xml.InsertBefore($XmlDeclaration, $xl_Workbook_xml.DocumentElement)

	        # Create and append the workbook node to the document.
	        $workBookElement = $xl_Workbook_xml.CreateElement("workbook")
			# add the office open xml namespaces to the XML document
	        $Null = $workBookElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	        $Null = $workBookElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
	        $Null = $xl_Workbook_xml.AppendChild($workBookElement)

	        # Create and append the sheets node to the workBook node.
	        $Null = $xl_Workbook_xml.DocumentElement.AppendChild($xl_Workbook_xml.CreateElement("sheets"))

		} # end begin block
		
		Process {	
			
			# set the file extension to xlsx
			$Path = [System.IO.Path]::ChangeExtension($Path,'xlsx')
			
			Try {
				# test if the file could be created
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				Out-File -InputObject "" -FilePath $Path -NoClobber:$NoClobber.IsPresent -Force:$Force.IsPresent -ErrorAction stop
				Remove-Item $Path -Force
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
			
			Try {
				# create the main package on disk with filemode create
				$exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Create)
			} Catch {
				$_
				return
			}
			
			# create URI for workbook.xml package part
			$Uri_xl_workbook_xml = New-Object System.Uri -ArgumentList ("/xl/workbook.xml", [System.UriKind]::Relative)
			# create workbook.xml part
			$Part_xl_workbook_xml = $exPkg.CreatePart($Uri_xl_workbook_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
			# get writeable stream from workbook.xml part 
			$dest = $part_xl_workbook_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
			# write workbook.xml XML document to part stream
			$xl_workbook_xml.Save($dest)

			# create package general main relationships
			$Null = $exPkg.CreateRelationship($Uri_xl_workbook_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1")
			
			# close main package
			$exPkg.Close()

			# return the FileInfo for the created XLSX file
			Return Get-Item $Path

		} # end Process block
		
		End {
		} # end End block
	}

	Function Import-DataTable {
	<#
		.SYNOPSIS
			Function executes SQL query and returns an System.Data.DataTable

		.DESCRIPTION
			Function executes SQL query and returns an System.Data.DataTable
			
			The SQL Query provided in the form of string.

			Ideally it should be able to receive an stream of SQL as an input
			and stream out data. But realistically I am going to start with
			string for SQL Query as input and System.Data.DataTable as an 
			return object.

			Even more ideally it would of being nice to have powershell to support something like
			an variable parameter array such that I could say -@parameter_name=value for query parameters.
			However, for now script will accept string for a query verbatim and any substitution must occur
			prior.

		.PARAMETER  QueryString
			An string containing complete SQL query.

		.PARAMETER  SQLConnectionString
			Connection string.
			
		.EXAMPLE
			PS C:\> Import-DataTable -Query "SELECT 1 AS [ColumnName]" -SQLConnection "..."

		.INPUTS
			System.String,System.Int32

		.OUTPUTS
			System.String

		.NOTES
			Version: 1.0.0 7.January.2015

		.LINK
			about_functions_advanced

		.LINK
			::undefined::

	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[System.String]$Query,
			###########################
			[Parameter(Mandatory=$True,
				Position=1,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[System.String]$SQLConnection
			###########################
		)
		begin{
		}
		process{
			$con                       = New-Object System.Data.SqlClient.SqlConnection
			$com                       = New-Object System.Data.SqlClient.SqlCommand
			$SqlAdapter                = New-Object System.Data.SqlClient.SqlDataAdapter
			#"Server=paramount;Database=pictures;User ID=viewer;Password=everything;Trusted_Connection=False;"
			$con.ConnectionString      = $SQLConnection;
			$com.Connection            = $con
			$com.CommandText           = $Query
			$com.CommandTimeout        = 300
			$DataSet                   = New-Object System.Data.DataSet
			$SqlAdapter.SelectCommand  = $com
			if( $com.Connection.State -ne "Open" )
			{
			  $com.Connection.Open()
			}
			$null = $SqlAdapter.Fill($DataSet)
			$com.Connection.Close()
			$com.Connection.Dispose()
			$com.Dispose()
			$DataSet
		}
		end{
		}
	}

	Function Set-XLSXColumnStyle {
	<#
		.SYNOPSIS
			Function sets formatting style for a columkn of an excel document worksheet.

		.DESCRIPTION
			Apply formatting on Microsoft Excel document column.

		.PARAMETER  QueryString
			An string containing complete SQL query.

		.PARAMETER  SQLConnectionString
			Connection string.
			
		.EXAMPLE
			PS C:\> Set-XLSXColumnStyle -Path "C:\test.xlsx" -Worksheet "test" -ColumnIndex 1 -Format "dd\mm\yyyy"

		.INPUTS
			System.String,System.Int32

		.OUTPUTS
			System.String

		.NOTES
			Version: 1.0.0 7.January.2015

		.LINK
			about_functions_advanced

		.LINK
			::undefined::

	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[System.String]$Path,
			###########################
			[Parameter(Mandatory=$True)]
			[System.String]$Worksheet,
			###########################
			[Parameter(Mandatory=$True)]
			[Int32]$ColumnIndex,
			###########################
			[Parameter(Mandatory=$True)]
			[System.String]$Format
			###########################
		)
		Begin {
		} # end begin block

	Process {
			#check file exist
			if($Path -eq $null){
				$message = "File path must not be null"
				$PSCmdlet.WriteError($message)
				return;
			}
			if( -not ( Test-Path $Path ) ){
				$message = "File $Path is not found";
				$PSCmdlet.WriteError($message);
				return;
			}
			# test if File is locked
			If($ErrorRecord = Test-FileLocked $Path -PassThru){
				$PSCmdlet.WriteError($ErrorRecord);
				return
			}
			Try {
				# test if the file could be accessed
				# generate an ErrorRecord Object which is automaticly translated in other languages by the Microsoft mechanism 
				$xlFile = Get-Item -Path $Path -ErrorAction stop                                
			} Catch {
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
		    $sl = $null;
			Try {
			# open Excel document
			$sl = New-Object SpreadsheetLight.SLDocument( $Path );
		    #find all worksheets of the document
		    $sheetNames = $sl.GetSheetNames($true);
		    if( ( $sheetNames-eq$null ) -or ($sheetNames.Length -eq 0 ) ){
				$message = "Document $Path contains no worksheets";
				$PSCmdlet.WriteVerbose($message);
				return;
			}
            if($Worksheet -in $sheetNames)
			{
				$null = $sl.SelectWorksheet($Worksheet);
			}
			else
			{
				# we dont want to show the ErrorRecord Object in the $Error list 
				$Error.RemoveAt(0)
				# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
				$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList ( "Worksheet {0} could not be found in Microsoft Excel document {1}" -f $Worksheet, $Path )
				# throw the ErrorRecord Object
				$PSCmdlet.WriteError($NewError)
				# leave Function
				Return
			}
			$sl.RemoveColumnStyle($ColumnIndex)
			$style = $sl.CreateStyle(); 
			$style.FormatCode = "mm/dd/yyyy";
			$sl.SetColumnStyle($ColumnIndex,$style)
			$sl.Save();
			$sl.Dispose();
			$sl = $null;
		} Catch {
			# we dont want to show the ErrorRecord Object in the $Error list 
			$Error.RemoveAt(0)
			# recreate an ErrorRecord Object with the invocation info from this function from catched translated ErrorRecord Object
			$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
			# throw the ErrorRecord Object
			$PSCmdlet.WriteError($NewError)

			if( $sl -ne $null ) ## most likely Excel object is there
			{
			    $sl.Dispose();
			    $sl = $null;
			}

			# leave Function
			Return
		}
		} # end Process block
		
		End { 
			} # end End block
	}

	Function Export-WorkSheets {
	<#
		.SYNOPSIS
			Function to fill an empty Excel worksheet with datas

		.DESCRIPTION
			Function to fill an empty Excel worksheet with datas
			
			The Export-WorkSheets function fills each noted in WorksheetNames Excel worksheet with 
			the propertys of the objects that you submit. Each object is represented as a line or row of 
			the worksheet. 
			The row consists of a number of Text typed worksheet cells. Each cell will contain the value of a Property of the object.
			Each property is converted into its string representation and stored as type of text inline into the Excel worksheet XML.
			You can use this Function to create real Excel XLSX spreadsheets without having Microsoft Excel installed!

			By default, the first cell row (line) of the worksheet represents the "column headers".
			The cells in this row contains the names of all the properties of the first object.
			
			Additional cell rows (lines) of the worksheet consist of the property values converted to their string representation of each object.
			
			NOTE: Do not format objects before sending them to the Export-WorkSheets function.
			If you do, the format properties are represented in the worksheet,
			instead of the properties of the original objects.
			To export only selected properties of an object, use the Select-Object cmdlet.

		.PARAMETER  Path
			Path to the existing Excel .xlsx workbook

		.PARAMETER  WorksheetUri
			System.Uri Object which points to a existing worksheet inside the e Excel workbook
			
		.PARAMETER InputDataSet
			Specifies the objects to export into the worksheet cells.
			Enter a variable that contains the objects or type a command or expression that gets the objects.
			You can also pipe objects to Export-XLSX.

		.PARAMETER  NoHeader
			Omits the "column header" row from the worksheet.
			
		.EXAMPLE
			PS C:\> Get-Something -ParameterA 'One value' -ParameterB 32

		.EXAMPLE
			PS C:\> Get-Something 'One value' 32

		.INPUTS
			System.String,System.Int32

		.OUTPUTS
			System.String

		.NOTES
			Author PowerShell:  Peter Kriegel, Germany http://www.admin-source.de
			Version: 1.0.0 14.August.2013

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help

	#>
		
		[CmdletBinding()]
		param(
			[Parameter(Mandatory=$True,
				Position=0,
				ValueFromPipeline=$True,
				ValueFromPipelinebyPropertyName=$True
			)]
			[System.String]$Path,
			###########################
			[Parameter(Mandatory=$true,
				Position=1,
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true
			)]
	    	[System.Data.DataSet]$InputDataSet,
			###########################
			[Parameter(Mandatory=$true,
				Position=1,
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true
			)]
	    	[System.String]$WorksheetNames,
			###########################
			[Switch]$NoHeader
			###########################
		)
		
		Begin {
			$WorksheetNameArray = $WorksheetNames.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries )
		}

		Process {
			$existingWorkbook = $false;
			if( Test-Path $Path )
			{
		       $sl = New-Object SpreadsheetLight.SLDocument( $Path );
				$existingWorkbook = $true;
			}else
			{
                $sl = New-Object SpreadsheetLight.SLDocument
			}

			$tableIndex = 0;

			ForEach( $InputTable in $InputDataSet.Tables )
			{
				$WorkSheetName = $WorksheetNameArray[$tableIndex];
				
				Write-Output "Worksheet $WorkSheetName started"
				Write-Output ( "Worksheet $WorkSheetName will carry {0} rows" -f $InputTable.Rows.Count )

				$time = [System.Datetime]::Now
								
				$sl.AddWorksheet($WorkSheetName);
				$sl.SelectWorksheet($WorkSheetName);
				$sl.ImportDataTable( 1, 1, $InputTable, $True );

				$time2 = [System.Datetime]::Now
				
				$days = $time2.Subtract($time).Days
				$hours = $time2.Subtract($time).Hours
				$minutes = $time2.Subtract($time).Minutes
				$seconds = $time2.Subtract($time).Seconds
				$milliseconds = $time2.Subtract($time).Milliseconds
				$milliseconds += ($seconds*1000 + ($minutes*60000) + ($hours*3600000) + ($days*(3600000*24)))

				Write-Output ( "Worksheet $WorkSheetName completed in {0} milliseconds" -f $milliseconds )
				Write-Output "Worksheet $WorkSheetName completed"
				$tableIndex += 1;
			}
			if($existingWorkbook -eq $false){
				$sl.SaveAs($Path);
			}else{
				$sl.Save();
			}

			$time3 = [System.Datetime]::Now
			$days = $time3.Subtract($time2).Days
			$hours = $time3.Subtract($time2).Hours
			$minutes = $time3.Subtract($time2).Minutes
			$seconds = $time3.Subtract($time2).Seconds
			$milliseconds = $time3.Subtract($time2).Milliseconds
			$milliseconds += ($seconds*1000 + ($minutes*60000) + ($hours*3600000) + ($days*(3600000*24)))

			Write-Output ( "Workbook $Path saved in {0} milliseconds" -f $milliseconds )
		} # end Process block
		End {
		} # end End block
	}
#endregion declare XLSX functions