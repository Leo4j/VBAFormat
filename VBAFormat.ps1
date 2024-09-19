function VBAFormat{
    param ([string]$InputString)
    # Split the input string into chunks of maximum 100 characters each
    $chunkSize = 100
    $chunks = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt $inputString.Length; $i += $chunkSize) {
        $chunk = $inputString.Substring($i, [Math]::Min($chunkSize, $inputString.Length - $i))
        $chunks.Add($chunk) | Out-Null
    }
    Write-Output ""
    Write-Output "##################"
    Write-Output "      Chunks      "
    Write-Output "##################"
    Write-Output ""
    # Print the code in the specified format
    for ($i = 0; $i -lt $chunks.Count; $i++) {
        Write-Output "str$i = `"$($chunks[$i])`""
    }

    # Maximum number of strings to concatenate at once
    $maxConcatenationSize = 50

    # Create temporary concatenation blocks
    $blockCounter = 0
    $blockParts = @()
    $concatenationString = ""
    Write-Output ""
    Write-Output "##################"
    Write-Output "      Blocks      "
    Write-Output "##################"
    Write-Output ""
    for ($i = 0; $i -lt $chunks.Count; $i++) {
        if ($i % $maxConcatenationSize -eq 0 -and $i -ne 0) {
            # Output the previous block concatenation
            Write-Output ("block$blockCounter = " + ($blockParts -join ' + '))
            $concatenationString += "block$blockCounter + "
            $blockParts = @() # Reset the block parts for the next batch
            $blockCounter++
        }
        
        # Add current str to block parts
        $blockParts += "str$i"
    }

    # Output the last block if it has remaining parts
    if ($blockParts.Count -gt 0) {
        Write-Output ("block$blockCounter = " + ($blockParts -join ' + '))
        $concatenationString += "block$blockCounter"
    }

    # Remove the trailing ' + ' from the concatenation string if necessary
    $concatenationString = $concatenationString.TrimEnd(' + ')

    # Print the final concatenation string
    Write-Output "str = $concatenationString"

    # Print the 'Dim' statement for VBA, grouping by 5 variables per line
    $dimStatement = "Dim "
    for ($i = 0; $i -lt $chunks.Count; $i++) {
        $dimStatement += "str$i As String"
        if (($i + 1) % 5 -eq 0 -or $i -eq $chunks.Count - 1) {
            $dimStatement += "`nDim "
        } else {
            $dimStatement += ", "
        }
    }
    $dimStatement = $dimStatement.TrimEnd("`nDim ") # Clean up the last 'Dim'

    # Output the 'Dim' statements for the block variables
    $blockDimStatement = "`nDim "
    for ($i = 0; $i -le $blockCounter; $i++) {
        $blockDimStatement += "block$i As String"
        if (($i + 1) % 5 -eq 0 -or $i -eq $blockCounter) {
            $blockDimStatement += "`nDim "
        } else {
            $blockDimStatement += ", "
        }
    }
    $blockDimStatement = $blockDimStatement.TrimEnd("`nDim ") # Clean up the last 'Dim'
    Write-Output ""
    Write-Output "##################"
    Write-Output "  Dim statements  "
    Write-Output "##################"
    Write-Output ""
    # Output all Dim statements
    Write-Output "Dim str As String"
    Write-Output $dimStatement
    Write-Output $blockDimStatement
}
