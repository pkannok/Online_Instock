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
' ########### the following for each Focus Region:                        ###########
' ###########      ~ Total Color SKUs Sizes (Count of Color-Skus sizes    ###########
' ###########        less count of sizes offline)                         ###########
' ###########      ~ In-Stock % (Sum of # Sizes Salable / Sum of Sizes)   ###########
' ###########      ~ Top 25 In-Stock % (Same as above, but for Top 25     ###########
' ###########        SKUs)                                                ###########
' ###########      ~ Top 25 Offline (Sum of SKUs in Top 25 offline)       ###########
' ###########                                                             ###########
' ########### ---------------------------                                 ###########
' ########### FOCUS REGIONS & TOP 25 SKUS                                 ###########
' ########### ---------------------------                                 ###########
' ########### Focus Regions are defined within the top25Csv file as       ###########
' ########### header items in the first row. The region should match the  ###########
' ########### region as indicated in the PSR filenaming convention        ###########
' ########### (e.g., for Japan, use header "jp" -- without quotes-- to    ###########
' ########### process files for ProductStatusReportFeed_jp_........csv).  ###########
' ########### Every Focus Region should additionally have a Top 25 SKU    ###########
' ########### list. This should be in cells 2 - 26 under the region's     ###########
' ########### heading in the top25Csv file and be the style-color SKU.    ###########
' ########### This list may be updated at any time, submitted by the      ###########
' ########### region team.                                                ###########
' ###########                                                             ###########
' ########### ---------------                                             ###########
' ########### WEEK DEFINITION                                             ###########
' ########### ---------------                                             ###########
' ########### Weeks begin on Monday and end on Sunday.                    ###########
' ###########                                                             ###########
' ########### ------------------------                                    ###########
' ########### FILE DIRECTORY STRUCTURE                                    ###########
' ########### ------------------------                                    ###########
' ########### File directory structure is a follows:                      ###########
' ########### - <directoryLoc>         = location of all files & folders  ###########
' ###########      + <PSRfolder>       = unprocessed PSR csv files        ###########
' ###########      + <archiveFolder>   = processed csv files              ###########
' ###########      + <top25Folder>     = Top 25 csv for each region       ###########
' ###########                                                             ###########
' ########### ---------------------                                       ###########
' ########### DEFINITION OF METRICS                                       ###########
' ########### ---------------------                                       ###########
' ########### Metrics are defined and calculated as follows:              ###########
' ###########  [Style-Color Sizes]                                        ###########
' ###########   Definition:  Total number of sizes possible for all       ###########
' ###########                style-color SKUs/colorways.                  ###########
' ###########   Calculation: For each item where:                         ###########
' ###########                - VARIATION MASTER ONLINE = YES              ###########
' ###########                (Sum(# SIZES FOR COLOR)                      ###########
' ###########                 Less: Sum(# SIZES FOR COLOR OFFLINE))       ###########
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
' ########### -------------------------------------                       ###########
' ########### ACCOUNTING FOR MULTIPLE PSRS PER WEEK                       ###########
' ########### -------------------------------------                       ###########
' ########### The PSR reports are delivered at various times and          ###########
' ########### intervals for each region. For any region that has more     ###########
' ########### than one PSR per week, the smallest value across PSRs is    ###########
' ########### reported for that region. The exception to this is Top 25   ###########
' ########### Offline, which is the maximum value calculated for the      ###########
' ########### week.                                                       ###########
' ###########                                                             ###########
' ###################################################################################
' ###################################################################################
' ###################################################################################

  Dim wb As Workbook, ws As Worksheet, top25Sheet As Worksheet, dataSheet As Worksheet, PSRsheet As Worksheet, pivotSheet As Worksheet, graphSheet As Worksheet, lastDataRow As Integer, lastPSRrow As Integer, lastPSRcolumn As Integer
  Dim sizesForColorCol As Integer, sizesOfflineCol As Integer, sizesOrderableCol As Integer, sizesBisnCol As Integer, variationMasterCol As Integer
  Dim sizesForColor As Long, colorSkuCount As Long, sizesOrderable As Long, sizesBisn As Long, topSizesForColor As Long, topSizesOrderable As Long, topSizesBisn As Long, topOfflineCount As Integer
  Dim directoryLoc As String, PSRfolder As String, archiveFolder As String, top25Folder As String, top25Csv As String, reportName As String, filePath As String
  Dim PSRnames() As String, focusRegions() As String, archiveNames() As String, focusRegionCount As Integer, PSRregion As String, PSRdate As Date, weekStart As Date, inFocusRegions As Variant, topSkus(1 To 25) As String, inTopSkus As Variant
  Dim i As Integer, ii As Integer, dataRange As Range, cht As Chart, pts As Points, dl As DataLabel, seriesCnt As Integer
  Set wb = ActiveWorkbook
  directoryLoc = "C:\Users\kkuramoto\Documents\Ad Hoc Analysis\Online Instock\v2\" 'Directory location with PSR files to be processed
  PSRfolder = "PSRs" 'Directory location with PSR files to be processed
  archiveFolder = "archive" 'Directory location to archive processed PSR files
  top25Folder = "top25"
  top25Csv = "top25.csv"
  reportName = "Instock Report"
  Set top25Sheet = Sheets("Top 25")
  Set dataSheet = Sheets("Data")
  Set PSRsheet = Sheets("PSR")
  Set pivotSheet = Sheets("Tables")
  Set graphSheet = Sheets("Graphs")

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  ' #### Clear sheets ####
  For Each ws In Worksheets
    If ws.Name <> pivotSheet.Name And ws.Name <> graphSheet.Name And ws.Name <> dataSheet.Name Then ' Add exceptions here
      With ws.Cells
        .Clear
        .ClearFormats
      End With
    End If
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

  filePath = directoryLoc & archiveFolder & Application.PathSeparator
  archiveNames = AllFilesinDirectory(filePath) 'Array of previously processed PSR filenames
  filePath = directoryLoc & PSRfolder & Application.PathSeparator
  PSRnames = AllFilesinDirectory(filePath) 'Array of PSR filenames to be processed

  For i = LBound(PSRnames) To UBound(PSRnames)
    ' ### Check if file has already been processed - Delete if True ###
    If UBound(Filter(archiveNames, PSRnames(i))) >= 0 Then
      SetAttr filePath & PSRnames(i), vbNormal
      Kill (filePath & PSRnames(i))
    Else
      filePath = directoryLoc & PSRfolder & Application.PathSeparator & PSRnames(i)
      PSRregion = Mid(PSRnames(i), 25, 2) 'region of PSR (from filename)
      PSRdate = DateSerial(Mid(PSRnames(i), 32, 4), Mid(PSRnames(i), 30, 2), Mid(PSRnames(i), 28, 2)) 'date of PSR (from filename)
      weekStart = PSRdate - (Weekday(PSRdate, vbMonday) -1)

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
        sizesOfflineCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR OFFLINE", LookIn:=xlValues).Column
        sizesOrderableCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR ORDERABLE", LookIn:=xlValues).Column
        sizesBisnCol = PSRsheet.Range("1:1").Find(What:="# SIZES FOR COLOR BISN ENABLED", LookIn:=xlValues).Column
        variationMasterCol = PSRsheet.Range("1:1").Find(What:="VARIATION MASTER ONLINE", LookIn:=xlValues).Column

        '*** Reset variable values ***
        sizesForColor = 0
        sizesOrderable = 0
        sizesBisn = 0
        colorSkuCount = 0
        topSizesForColor = 0
        topSizesOrderable = 0
        topSizesBisn = 0
        topOfflineCount = 25

        '*** Sum data columns ***
        For ii = 2 to lastPSRrow
          If PSRsheet.Cells(ii, variationMasterCol).Value = "YES" Then  'Do not sum if VARATION MASTER ONLINE <> "YES"
            sizesForColor = sizesForColor + (PSRsheet.Cells(ii, sizesForColorCol).Value - PSRsheet.Cells(ii, sizesOfflineCol).Value)  'Sum # SIZES FOR COLOR less # SIZES FOR COLOR OFFLINE
            sizesOrderable = sizesOrderable + PSRsheet.Cells(ii, sizesOrderableCol).Value  ' Sum # SIZES FOR COLOR ORDERABLE
            sizesBisn = sizesBisn + PSRsheet.Cells(ii, sizesBisnCol).Value  '  Sum # SIZES FOR COLOR BISN ENABLED
            inTopSkus = Filter(topSkus, PSRsheet.Cells(ii, colorSkuCol).Value)  '  Compare SKU to region's Top 25 list
            If PSRsheet.Cells(ii, colorSkuCol).Value <> "N/A" Then ' Count of Color-Skus that are not N/A (accessories)
              colorSkuCount = colorSkuCount + 1
            End If
            If UBound(inTopSkus) >= 0 Then  '  If on Top 25 list...
              If PSRsheet.Cells(ii, sizesOrderableCol).Value - PSRsheet.Cells(ii, sizesBisnCol).Value > 0 Then  '  And orderable is not BISN, then...
                topOfflineCount = topOfflineCount - 1  '  Decrease the count of Top 25 Style-Colors offline
              End If
              topSizesForColor = topSizesForColor + (PSRsheet.Cells(ii, sizesForColorCol).Value - PSRsheet.Cells(ii, sizesOfflineCol).Value)  '  Sum # SIZES FOR COLOR less # SIZES FOR COLOR OFFLINE
              topSizesOrderable = topSizesOrderable + PSRsheet.Cells(ii, sizesOrderableCol).Value  ' Sum # SIZES FOR COLOR ORDERABLE
              topSizesBisn = topSizesBisn + PSRsheet.Cells(ii, sizesBisnCol).Value  '  Sum # SIZES FOR COLOR BISN ENABLED
            End If
          End If
        Next ii
        '*** Populate Data sheet ***
        ' With dataSheet.Cells(1, 1)
        '   .Offset(0, 0).Value = "Week Start"
        '   .Offset(0, 1).Value = "Report Date"
        '   .Offset(0, 2).Value = "Region"
        '   .Offset(0, 3).Value = "Style-Color Count"
        '   .Offset(0, 4).Value = "Overall Instock %"
        '   .Offset(0, 5).Value = "Top 25 Instock %"
        '   .Offset(0, 6).Value = "Top 25 Offline"
        ' End With
        lastDataRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
        With dataSheet.Cells(lastDataRow + 1, 1)
          .Offset(0, 0).Value = weekStart
          .Offset(0, 1).Value = PSRdate
          .Offset(0, 2).Value = UCase(PSRregion)
          .Offset(0, 3).Value = Format(colorSkuCount, "#,##0")
          .Offset(0, 4).Value = Format((sizesOrderable - sizesBisn) / sizesForColor, "Percent")
          .Offset(0, 5).Value = Format((topSizesOrderable - topSizesBisn) / topSizesForColor, "Percent")
          .Offset(0, 6).Value = Format(topOfflineCount, "#,##0")
        End With
      End If
      '*** Move PSR report to archiveFolder ***
      Name filePath As directoryLoc & archiveFolder & Application.PathSeparator & PSRnames(i)

      '*** Clear PSR data from PSR sheet ***
      PSRsheet.Cells.Clear
    End If
  Next i

  Call KillConnections

  '### Update Pivot Tables and Pivot Charts ###
  lastDataRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
  Set dataRange = dataSheet.Range("A1", dataSheet.Range("G" & lastDataRow))

  pivotSheet.PivotTables("pvtColors").ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange, Version:=xlPivotTableVersion14)
  pivotSheet.PivotTables("pvtAllInstock").ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange, Version:=xlPivotTableVersion14)
  pivotSheet.PivotTables("pvtTopInstock").ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange, Version:=xlPivotTableVersion14)
  pivotSheet.PivotTables("pvtOffline").ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange, Version:=xlPivotTableVersion14)

  Set cht = graphSheet.ChartObjects("chtColors").Chart
  seriesCnt = cht.SeriesCollection.Count
  For i = 1 to seriesCnt
    cht.SeriesCollection(i).ApplyDataLabels Type:=xlDataLabelsShowNone
    Set pts = cht.SeriesCollection(i).Points
    pts(pts.Count).ApplyDataLabels ShowSeriesName:=True
  Next i
  Set cht = graphSheet.ChartObjects("chtAllInstock").Chart
  For i = 1 to seriesCnt
    cht.SeriesCollection(i).ApplyDataLabels Type:=xlDataLabelsShowNone
    Set pts = cht.SeriesCollection(i).Points
    pts(pts.Count).ApplyDataLabels ShowSeriesName:=True
  Next i
  Set cht = graphSheet.ChartObjects("chtTopInstock").Chart
  For i = 1 to seriesCnt
    cht.SeriesCollection(i).ApplyDataLabels Type:=xlDataLabelsShowNone
    Set pts = cht.SeriesCollection(i).Points
    pts(pts.Count).ApplyDataLabels ShowSeriesName:=True
  Next i
  Set cht = graphSheet.ChartObjects("chtOffline").Chart
  For i = 1 to seriesCnt
    cht.SeriesCollection(i).ApplyDataLabels Type:=xlDataLabelsShowNone
    Set pts = cht.SeriesCollection(i).Points
    pts(pts.Count).ApplyDataLabels ShowSeriesName:=True
    With cht.SeriesCollection(i).DataLabels
      .Orientation = xlUpward
      .Font.Size = 8
    End With
  Next i

  '### Export PDF of Graphs and Save ###
  filePath = directoryLoc & reportName & " - " & Format(Application.WorksheetFunction.Max(dataSheet.Range("A:A")), "mmddyyyy")
  graphSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=filePath, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False

    wb.Save

End Sub

Sub KillConnections()
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Connections.Count
    If ActiveWorkbook.Connections.Count = 0 Then Exit Sub
    ActiveWorkbook.Connections.Item(i).Delete
    i = i - 1
    Next i
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
