# The below script can be used to create VS Code JSON theme files based on ps1xml ISE theme files.
# All that's required is to specify the path to where the ps1xml ISE theme files are stored:
$ISEThemesRootPath = "C:\PathToPS1XmlFiles"


# Get the full file path for each ps1xml ISE theme file
$ISEThemeFiles = Get-ChildItem -Path "$ISEThemesRootPath\*.ps1xml" | Select Fullname

# Process each ps1xml ISE theme file
foreach ($ISEThemeFilepath in $ISEThemeFiles) {

    # Get the XML content of the ps1xml ISE theme file
    [xml]$ISEThemeXML = Get-Content -Path $ISEThemeFilepath.FullName

    # Extract the theme name from the ps1xml ISE theme file
    $ThemeName = $ISEThemeXML.StorableColorTheme.Name
    Write-Verbose "Converting $ThemeName..." -Verbose

    # Initialise ISEThemeColourValuePairs custom object
    $ISEThemeColourValuePairs = New-Object -TypeName PSObject

    # Get a list of the XML keys in the ISE theme
    $ISEThemeKeys = $ISEThemeXML.StorableColorTheme.Keys.GetEnumerator().'#Text'

    # Extract the relevant name and colour value pairs from the XML ISE theme
    for ($i = 0; $i -lt $ISEThemeKeys.Count; $i++) {

        # Get current colour name
        [string]$ExtractedColourName = ($ISEThemeXML.StorableColorTheme.Keys.GetEnumerator().'#Text')[$i]

        # Check if it's a script pane colour as those are all we're looking for
        if (($ExtractedColourName -notlike "Console*") -and ($ExtractedColourName -notlike "Xml*")) {

            # If it's a TokenColors colour, trim the start of its name to avoid a slash in property name
            if ($ExtractedColourName -like "TokenColors\*") {
                $ColourName = $ExtractedColourName.Replace("TokenColors\","")
            } else {
                $ColourName = $ExtractedColourName
            }

            # Get the current RGB colour and convert it to a hex value
            ($ISEThemeXML.StorableColorTheme.Values.GetEnumerator() | Select R, G, B)[$i] | foreach {

                [int]$RIntValue = $_.R
                [int]$GintValue = $_.G
                [int]$BIntValue = $_.B

                $RHexValue = "{0:x2}" -f $RIntValue
                $GHexValue = "{0:x2}" -f $GintValue
                $BHexValue = "{0:x2}" -f $BIntValue

                [string]$CurrentColour = ("$($RHexValue)$($GHexValue)$($BHexValue)").ToUpper()


                # Also get a slightly lighter version of the colour if possible
                if ($RIntValue -lt 235) {$RHexValueLighter = "{0:x2}" -f ($RIntValue + 20)} else {$RHexValueLighter = "FF"}
                if ($GIntValue -lt 235) {$GHexValueLighter = "{0:x2}" -f ($GIntValue + 20)} else {$GHexValueLighter = "FF"}
                if ($BIntValue -lt 235) {$BHexValueLighter = "{0:x2}" -f ($BIntValue + 20)} else {$BHexValueLighter = "FF"}

                [string]$CurrentColourLighter = ("$($RHexValueLighter)$($GHexValueLighter)$($BHexValueLighter)").ToUpper()


                # Also get a slightly darker version of the colour if possible
                if ($RIntValue -gt 20) {$RHexValueDarker = "{0:x2}" -f ($RIntValue - 20)} else {$RHexValueDarker = "00"}
                if ($GIntValue -gt 20) {$GHexValueDarker = "{0:x2}" -f ($GIntValue - 20)} else {$GHexValueDarker = "00"}
                if ($BIntValue -gt 20) {$BHexValueDarker = "{0:x2}" -f ($BIntValue - 20)} else {$BHexValueDarker = "00"}

                [string]$CurrentColourDarker = ("$($RHexValueDarker)$($GHexValueDarker)$($BHexValueDarker)").ToUpper()

                # Add colour name and hex value pairs to ISEThemeColourValuePairs custom object
                $ISEThemeColourValuePairs | Add-Member -MemberType NoteProperty -Name $ColourName -Value $CurrentColour
                $ISEThemeColourValuePairs | Add-Member -MemberType NoteProperty -Name "$($ColourName)Lighter" -Value $CurrentColourLighter
                $ISEThemeColourValuePairs | Add-Member -MemberType NoteProperty -Name "$($ColourName)Darker" -Value $CurrentColourDarker

            }# ($ISEThemeXML.StorableColorTheme.Values...

        }# if (($ExtractedColourName...

    }# for ($i = 0;...

    # Construct the VS Code JSON theme
    $VSCodeThemeJSON = @"
{
    "name": "$ThemeName",
    "tokenColors": [
        {
            "name": "Comments",
            "scope": [
                "comment"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Comment)"
            }
        },
        {
            "name": "Comments: Preprocessor",
            "scope": "comment.block.preprocessor",
            "settings": {
                "fontStyle": "",
                "foreground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColor)"
            }
        },
        {
            "name": "Comments: Documentation",
            "scope": [
                "comment.documentation",
                "comment.block.documentation"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColor)"
            }
        },
        {
            "name": "Invalid - Deprecated",
            "scope": "invalid.deprecated",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.WarningForegroundColor)"
            }
        },
        {
            "name": "Invalid - Illegal",
            "scope": "invalid.illegal",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.WarningForegroundColor)"
            }
        },
        {
            "name": "Keywords",
            "scope": [
                "keyword.other",
                "keyword.control",
                "punctuation.section.group",
                "punctuation.section.braces"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Keyword)"
            }
        },
        {
            "name": "Types",
            "scope": [
                "storage.type"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Type)"
            }
        },
        {
            "name": "Language Constants",
            "scope": [
                "constant.language",
                "support.constant",
                "variable.language"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Variable)"
            }
        },
        {
            "name": "Variables",
            "scope": [
                "variable",
                "support.variable",
                "punctuation.definition.variable"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Variable)"
            }
        },
        {
            "name": "Tag attribute",
            "scope": [
                "entity.other.attribute-name",
                "variable.parameter",
                "keyword.operator",
                "meta"
            ],
            "settings": {
                "fontStyle": "",
                "foreground": "#$($ISEThemeColourValuePairs.Operator)"

            }
        },
        {
            "name": "Functions",
            "scope": [
                "entity.name.function"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Number)"
            }
        },
        {
            "name": "Functions2",
            "scope": [
                "support.function"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Command)"
            }
        },
        {
            "name": "Classes",
            "scope": [
                "entity.name.tag",
                "entity.name.type",
                "entity.other.inherited-class",
                "support.class"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Keyword)"
            }
        },
        {
            "name": "Exceptions",
            "scope": "entity.name.exception",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.ErrorForegroundColor)"
            }
        },
        {
            "name": "Sections",
            "scope": "entity.name.section",
            "settings": {
            }
        },
        {
            "name": "Numbers, Characters",
            "scope": [
                "constant.numeric",
                "constant"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Number)"
            }
        },
        {
            "name": "User-defined constant",
            "scope": [
                "constant.character",
                "constant.other"
            ],
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.Keyword)"
            }
        },
        {
            "name": "Strings",
            "scope": "string",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.String)"
            }
        },
        {
            "name": "Strings: Escape Sequences",
            "scope": "constant.character.escape",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.String)"
            }
        },
        {
            "name": "Strings: Regular Expressions",
            "scope": "string.regexp",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.String)"
            }
        },
        {
            "name": "Strings: Symbols",
            "scope": "constant.other.symbol",
            "settings": {
                "foreground": "#$($ISEThemeColourValuePairs.String)"
            }
        }
    ],
    "colors": {
        "titleBar.activeBackground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorDarker)",
        "titleBar.activeForeground": "#$($ISEThemeColourValuePairs.OperatorDarker)",
        "editorGroupHeader.tabsBackground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorDarker)",
        "tab.activeBackground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColor)",
        "tab.activeForeground": "#$($ISEThemeColourValuePairs.OperatorLighter)",
        "tab.inactiveBackground": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorLighter)",
        "tab.inactiveForeground": "#$($ISEThemeColourValuePairs.OperatorDarker)",
        "editor.background": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColor)",
        "editor.foreground": "#$($ISEThemeColourValuePairs.Operator)",
        "editor.selectionBackground": "#45494d82",
        "activityBar.background": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorLighter)",
        "activityBar.foreground": "#$($ISEThemeColourValuePairs.Operator)",
        "sideBar.background": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorDarker)",
        "statusBar.background": "#$($ISEThemeColourValuePairs.ScriptPaneBackgroundColorLighter)"
    }
}
"@

    # Save the resulting VS Code JSON theme file to the same directory as the source ps1xml file
    $ThemeFileName = $ThemeName.Replace(" ","-")
    $VSCodeThemeJSON | Out-File -FilePath "$ISEThemesRootPath\$ThemeFileName-color-theme.json" -Encoding utf8

}# foreach ($ISEThemeFilepath in $ISEThemeFiles)...
