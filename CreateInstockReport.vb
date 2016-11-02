Sub CreateOnlineInstock()

' ###################################################################################
' ###################################################################################
' ###################################################################################
' ########### Excel VB Macro - Create Online Instock Report               ###########
' ########### Created by: Kanno Kuramoto (kkuramoto@crocs.com)            ###########
' ###########                                                             ###########
' ########### -------                                                     ###########
' ########### SUMMARY                                                     ###########
' ########### -------                                                     ###########
' ########### From a directory of ProductStatusReports calculate          ###########
' ########### the following for each region:                              ###########
' ###########      ~ In-Stock % (Sum of # Sizes Salable / Sum of Sizes)   ###########
' ###########      ~ Total Sizes (Sum of Sizes Salable)                   ###########
' ###########      ~ Top 25 In-Stock (Same as above, but for Top 25 SKUs) ###########
' ###########      ~ Top 25 Offline (Sum of SKUs in Top 25 offline)       ###########
' ###########                                                             ###########
' ########### File directory structure is a follows:                      ###########
' ########### - <directoryLoc>         = location of all files & folders  ###########
' ###########      + <PSRfolder>       = unprocessed PSR csv files        ###########
' ###########      + <archiveFolder>   = processed csv files              ###########
' ###########      + <top25Folder>     = Top 25 csv for each region       ###########
' ###########                                                             ###########
' ########### Metrics are defined and calculated as follows:              ###########
' ###########  [Style-Color Sizes]                                        ###########
' ###########   Definition:  Total number of sizes possible for all       ###########
' ###########                style-color SKUs/colorways.                  ###########
' ###########   Calculation: For each item where:                         ###########
' ###########                - VARIATION MASTER ONLINE = YES              ###########
' ###########                Sum(# SIZES FOR COLOR)                       ###########
' ###########                                                             ###########
' ###########  [Overall Instock %]                                        ###########
' ###########   Definition:  The total number of sizes salable divided by ###########
' ###########                the Style-Color Sizes.                       ###########
' ###########   Calculation: For each item where:                         ###########
' ###########                - VARIATION MASTER ONLINE = YES              ###########
' ###########                (Sum(# SIZES FOR COLOR ORDERABLE)            ###########
' ###########                 Less: Sum(# SIZES FOR COLOR BISN ENABLED))  ###########
' ###########                Divided by: Sum(# SIZES FOR COLOR)           ###########
' ###########                                                             ###########
' ###########  [Top 25 Instock %]                                         ###########
' ###########   Definition:  For a region's Top 25 SKUs, the total number ###########
' ###########                of sizes salable divided by the total number ###########
' ###########                of style-color SKUs possible.                ###########
' ###########   Calculation: For each item where:                         ###########
' ###########                - VARIATION MASTER ONLINE = YES              ###########
' ###########                - COLOR SKU is in region's Top 25 SKUs       ###########
' ###########                (Sum(# SIZES FOR COLOR ORDERABLE)            ###########
' ###########                 Less: Sum(# SIZES FOR COLOR BISN ENABLED))  ###########
' ###########                Divided by: Sum(# SIZES FOR COLOR)           ###########
' ###########                                                             ###########
' ###########  [Top 25 Offline]                                           ###########
' ###########   Definition:  A count of a region's Top 25 style-color     ###########
' ###########                SKUs/colorways that are either offline or    ###########
' ###########                not listed in the PSR.                       ###########
' ###########   Calculation: Count of each item where:                    ###########
' ###########                - VARIATION MASTER ONLINE = YES              ###########
' ###########                - # SIZES FOR COLOR ORDERABLE Less: # SIZES  ###########
' ###########                  FOR COLOR BISN ENABLED > 0                 ###########
' ###########                - COLOR SKU is in region's Top 25 SKUs       ###########
' ###########                                                             ###########
' ###################################################################################
' ###################################################################################
' ###################################################################################

  Dim ws As Worksheet, top25Sheet As Worksheet, dataSheet As Worksheet, PSRsheet As Worksheet, lastDataRow As Integer, lastPSRrow As Integer, lastPSRcolumn As Integer
  Dim sizesForColorCol As Integer, sizesOrderableCol As Integer, sizesBisnCol As Integer, variationMasterCol As Integer
  Dim sizesForColor As Long, sizesOrderable As Long, sizesBisn As Long, topSizesForColor As Long, topSizesOrderable As Long, topSizesBisn As Long, topOfflineCount As Integer
  Dim directoryLoc As String, PSRfolder As String, archiveFolder As String, top25Folder As String, top25Csv As String, filePath As String
  Dim PSRnames() As String, focusRegions() As String, focusRegionCount As Integer, PSRregion As String, PSRdate As Date, inFocusRegions As Variant, topSkus(1 To 25) As String, inTopSkus As Variant
  Dim i As Integer, ii As Integer
  directoryLoc = "C:\Users\kkuramoto\Documents\Ad Hoc Analysis\Online Instock\v2\" 'Directory location with PSR files to be processed
  PSRfolder = "PSRs" 'Directory location with PSR files to be processed
  archiveFolder = "archive" 'Directory location to archive processed PSR files
  top25Folder = "top25"
  top25Csv = "top25.csv"
  Set top25Sheet = Sheets("Top 25")
  Set dataSheet = Sheets("Data")
  Set PSRsheet = Sheets("PSR")

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  ' #### Clear sheets ####
  For Each ws In Worksheets
  '     If ws <> dataSheet Then ' Add exceptions here
          With ws.Cells
              .Clear
              .ClearFormats
          End With
      ' End If
  Next
  Set ws = Nothing

  '#### Specify regions to be included in report (via an array)
  filePath = directoryLoc & top25Folder & Application.PathSeparator & top25csv
  With top25Sheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=top25Sheet.Range("A1"))
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .BackgroundQuery = True
    .TablesOnlyFromHTML = False
    .Refresh BackgroundQuery:=False
    .SaveData = True
  End With
  focusRegionCount = top25Sheet.Cells(1, top25Sheet.Columns.Count).End(xlToLeft).Column
  ReDim focusRegions(focusRegionCount)
  For i = 1 To focusRegionCount
      focusRegions(i) = top25Sheet.Cells(1, i).Value
  Next i

  filePath = directoryLoc & PSRfolder & Application.PathSeparator
  PSRnames = AllFilesinDirectory(filePath) 'Array of PSR filenames to be processed

  For i = LBound(PSRnames) To UBound(PSRnames)
    filePath = directoryLoc & PSRfolder & Application.PathSeparator & PSRnames(i)
    PSRregion = Mid(PSRnames(i), 25, 2) 'region of PSR (from filename)
    PSRdate = DateSerial(Mid(PSRnames(i), 32, 4), Mid(PSRnames(i), 30, 2), Mid(PSRnames(i), 28, 2)) 'date of PSR (from filename)

    inFocusRegions = Filter(focusRegions, PSRregion)
    If UBound(inFocusRegions) < 0 Then
      ' ### Skip processing / do nothing ###
    Else
      '### Import each report ###
      With PSRsheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=PSRsheet.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .BackgroundQuery = True
        .TablesOnlyFromHTML = False
        .Refresh BackgroundQuery:=False
        .SaveData = True
      End With

      '*** Set top 25 for region
      For ii = 1 To 25
        topSkus(ii) = top25Sheet.Cells(1, top25Sheet.Range("1:1").Find(What:=PSRregion, LookIn:=xlValues).Column).Offset(ii, 0).Value
      Next ii

      '*** Identify columns by header ***
      lastPSRrow = PSRsheet.Cells(PSRsheet.Rows.Count, "A").End(xlUp).Row
      lastPSRcolumn = PSRsheet.Cells(1, PSRsheet.Columns.Count).End(xlToLeft).Column
      colorSkuCol = PSRsheet.Range("1:1").Find(What:="COLOR SKU", LookIn:=xlValues).Column
      sizesForColorCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR", LookIn:=xlValues).Column
      sizesOrderableCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR ORDERABLE", LookIn:=xlValues).Column
      sizesBisnCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR BISN ENABLED", LookIn:=xlValues).Column
      variationMasterCol = PSRsheet.Range("1:1").Find(What:="VARIATION MASTER ONLINE", LookIn:=xlValues).Column

      '*** Reset variable values ***
      sizesForColor = 0
      sizesOrderable = 0
      sizesBisn = 0
      topSizesForColor = 0
      topSizesOrderable = 0
      topSizesBisn = 0
      topOfflineCount = 25

      '*** Sum data columns ***
      For ii = 2 to lastPSRrow
        If PSRsheet.Cells(ii, variationMasterCol).Value = "YES" Then  'Do not sum if VARATION MASTER ONLINE <> "YES"
          sizesForColor = sizesForColor + PSRsheet.Cells(ii, sizesForColorCol).Value  'Sum # SIZES FOR COLOR
          sizesOrderable = sizesOrderable + PSRsheet.Cells(ii, sizesOrderableCol).Value  ' Sum # SIZES FOR COLOR ORDERABLE
          sizesBisn = sizesBisn + PSRsheet.Cells(ii, sizesBisnCol).Value  '  Sum # SIZES FOR COLOR BISN ENABLED
          inTopSkus = Filter(topSkus, PSRsheet.Cells(ii, colorSkuCol).Value)  '  Compare SKU to region's Top 25 list
          If UBound(inTopSkus) >= 0 Then  '  If on Top 25 list, then
            topOfflineCount = topOfflineCount - 1  '  Decrease the count of Top 25 Style-Colors offline
            topSizesForColor = topSizesForColor + PSRsheet.Cells(ii, sizesForColorCol).Value  '  Sum # SIZES FOR COLOR
            topSizesOrderable = topSizesOrderable + PSRsheet.Cells(ii, sizesOrderableCol).Value  ' Sum # SIZES FOR COLOR ORDERABLE
            topSizesBisn = topSizesBisn + PSRsheet.Cells(ii, sizesBisnCol).Value  '  Sum # SIZES FOR COLOR BISN ENABLED
          End If
        End If
      Next ii
      '*** Populate Data sheet ***
      With dataSheet.Cells(1, 1)
        .Offset(0, 0).Value = "Date"
        .Offset(0, 1).Value = "Region"
        .Offset(0, 2).Value = "Style-Color Sizes"
        .Offset(0, 3).Value = "Overall Instock %"
        .Offset(0, 4).Value = "Top 25 Instock %"
        .Offset(0, 5).Value = "Top 25 Offline"
      End With
      lastDataRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
      With dataSheet.Cells(lastDataRow + 1, 1)
        .Offset(0, 0).Value = PSRdate
        .Offset(0, 1).Value = PSRregion
        .Offset(0, 2).Value = sizesForColor
        .Offset(0, 3).Value = (sizesOrderable - sizesBisn) / sizesForColor
        .Offset(0, 4).Value = (topSizesOrderable - topSizesBisn) / topSizesForColor
        .Offset(0, 5).Value = topOfflineCount
      End With
    End If
    '*** Move PSR report to archiveFolder ***
    ' Name filePath As directoryLoc & archiveFolder & Application.PathSeparator & PSRnames(i)

    '*** Clear PSR data from PSR sheet ***
    PSRsheet.Cells.Clear
  Next i

  '### Build charts ###

End Sub

Function AllFilesinDirectory(folderLoc As String) As String()
  '### Returns an array of filenames in a directory ###
  Dim fileName As String, fileNames() As String, fileCount As Integer
  fileName = Dir(folderLoc & Application.PathSeparator)
  Do Until fileName = ""
    fileCount = fileCount + 1
    ReDim Preserve fileNames(1 To fileCount)
    fileNames(fileCount) = fileName
    fileName = Dir
  Loop
  AllFilesinDirectory = fileNames
End Function
