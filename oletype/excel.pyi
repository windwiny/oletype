# just stub file, import it , declare app obj, get ide auto type hit

import os
import sys
import datetime

# TODO FIXME unparsed
class A_CalculatedMember_object_that_represents_the_new_calculated_field_or_calculated_item_: pass
class A_Double_value_that_represents_the_subtotal_: pass
class A_Name_object_contained_by_the_collection_: pass
class A_PublishObject_object_that_represents_the_new_item_: pass
class A_Range_object_that_represents_the_first_cell_where_that_information_is_found_: pass
class A_String_value_that_represents_the_new_string__after_replacement_: pass
class AboveAverage_object: pass
class An_AddIn_object_that_represents_the_new_add_in_: pass
class An_AllowEditRange_object_that_represents_the_range_: pass
class An_Object_value_that_represents_an_object_contained_by_the_collection_: pass
class An_Object_value_that_represents_the_new_worksheet__chart__or_macro_sheet_: pass
class ColorScale_object: pass
class Databar_object: pass
class IconSetCondition_object: pass
class Top10_object: pass
class UniqueValues_object: pass

# TODO FIXME nofind define
class AddIn: pass
class CALCULATEDMEMBER: pass
class Comment: pass
class CommentThreaded: pass
class CustomProperty: pass
class CustomView: pass
class FormatCondition: pass
class FreeformBuilder: pass
class HPageBreak: pass
class Hyperlink: pass
class MODELRELATIONSHIP: pass
class MODELTABLE: pass
class ModelMeasure: pass
class ModelRelationship: pass
class Name: pass
class ODBCError: pass
class OLEDBError: pass
class ProtectedViewWindow: pass
class QueryTable: pass
class RecentFile: pass
class Shape: pass
class SlicerCache: pass
class SortField: pass
class SparklineGroup: pass
class VPageBreak: pass
class WORKBOOKCONNECTION: pass
class Watch: pass
class WorkbookQuery: pass
class XmlMap: pass
class object: pass


# num=1
class Application(_Application):
  def __init__(self):
    self.ActiveCell: Range
    '''Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.'''
    self.ActiveEncryptionSession: int
    '''Returns a Long that represents the encryption session associated with the active document. Read-only.'''
    self.ActiveMenuBar: MenuBar
    self.ActivePrinter: str
    '''Returns or sets the name of the active printer. Read/write String.'''
    self.ActiveSheet: _Worksheet
    '''Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.'''
    self.ActiveWindow: Window
    '''Returns a Window object that represents the active Excel window (the window on top). Returns Nothing if there are no windows open. Read-only.'''
    self.ActiveWorkbook: Workbook
    '''Returns a Workbook object that represents the workbook in the active window (the window on top). Returns Nothing if there are no windows open or if either the Info window or the Clipboard window is the active window. Read-only.'''
    self.AddIns: AddIns
    '''Returns an AddIns collection that represents all the add-ins listed in the Add-Ins dialog box (Add-Ins command on the Developer tab). Read-only.'''
    self.AddIns2: AddIns2
    '''Returns an AddIns2 collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed. Read-only.'''
    self.AlertBeforeOverwriting: bool
    '''True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation. Read/write Boolean.'''
    self.AltStartupPath: str
    '''Returns or sets the name of the alternate startup folder. Read/write String.'''
    self.AlwaysUseClearType: bool
    '''Returns or sets a Boolean that represents whether to use ClearType to display fonts in the menu, ribbon, and dialog box text. Read/write Boolean.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.ArbitraryXMLSupportAvailable: bool
    '''Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available. Read-only.'''
    self.AskToUpdateLinks: bool
    '''True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box. Read/write Boolean.'''
    self.Assistance: CDispatch
    '''Returns an IAssistance object for Microsoft Excel that represents the Microsoft Office Help Viewer. Read-only.'''
    self.Assistant: CDispatch
    self.AutoCorrect: AutoCorrect
    '''Returns an AutoCorrect object that represents the Microsoft Excel AutoCorrect attributes. Read-only.'''
    self.AutoFormatAsYouTypeReplaceHyperlinks: bool
    '''True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type. Read/write Boolean.'''
    self.AutoPercentEntry: bool
    '''True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered. Read/write Boolean.'''
    self.AutoRecover: AutoRecover
    '''Returns an AutoRecover object, which backs up all file formats on a timed interval.'''
    self.AutomationSecurity: int
    '''Returns or sets an MsoAutomationSecurity constant that represents the security mode that Microsoft Excel uses when programmatically opening files. Read/write.'''
    self.Build: float
    '''Returns the Microsoft Excel build number. Read-only Long.'''
    self.COMAddIns: CDispatch
    '''Returns the COMAddIns collection for Microsoft Excel, which represents the currently installed COM add-ins. Read-only.'''
    self.CSVDisplayNumberConversionWarning: bool
    self.CSVKeepColumnAsTextIfMultipleEntriesAreText: bool
    self.CalculateBeforeSave: bool
    '''True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property. Read/write Boolean.'''
    self.Calculation: int
    '''Returns or sets an XlCalculation value that represents the calculation mode.'''
    self.CalculationInterruptKey: int
    '''Sets or returns an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations. Read/write.'''
    self.CalculationState: int
    '''Returns an XlCalculationState constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel. Read-only.'''
    self.CalculationVersion: int
    '''Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel. Read-only Long.'''
    self.Caller: int
    '''Returns information about how Visual Basic was called (for more information, see the Remarks section).'''
    self.CanPlaySounds: bool
    self.CanRecordSounds: bool
    self.Caption: str
    '''Returns or sets a String value that represents the name that appears in the title bar of the main Microsoft Excel window.'''
    self.CellDragAndDrop: bool
    '''True if dragging and dropping cells is enabled. Read/write Boolean.'''
    self.Cells: Range
    '''Returns a Range object that represents all the cells on the active worksheet. If the active document is not a worksheet, this property fails.'''
    self.ChartDataPointTrack: bool
    '''True causes all charts in newly created documents to use the cell reference tracking behavior. Boolean.'''
    self.Charts: Sheets
    '''Returns a Sheets collection that represents all the chart sheets in the active workbook.'''
    self.ClipboardFormats: tuple
    '''Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate constant listed in the Remarks section. Read-only Variant.'''
    self.ClusterConnector: str
    '''Returns or sets the name of the High Performance Computing (HPC) Cluster Connector that is used to run user-defined functions in XLL add-ins. Read/write.'''
    self.ColorButtons: bool
    self.Columns: Range
    '''Returns a Range object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the Columns property fails.'''
    self.CommandBars: CDispatch
    '''Returns a CommandBars object that represents the Microsoft Excel command bars. Read-only.'''
    self.CommandUnderlines: int
    '''Returns or sets the state of the command underlines in Microsoft Excel for the Macintosh. Can be one of the constants of XlCommandUnderlines. Read/write Long.'''
    self.ConstrainNumeric: bool
    '''True if handwriting recognition is limited to numbers and punctuation only. Read/write Boolean.'''
    self.ControlCharacters: float
    '''True if Microsoft Excel displays control characters for right-to-left languages. Read/write Boolean.'''
    self.ConvertNumbersWithECharacter: bool
    self.CopyObjectsWithCells: bool
    '''True if objects are cut, copied, extracted, and sorted with cells. Read/write Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Cursor: int
    '''Returns or sets the appearance of the mouse pointer in Microsoft Excel. Read/write XlMousePointer.'''
    self.CursorMovement: float
    '''Returns or sets a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants: xlVisualCursor or xlLogicalCursor. Read/write Long.'''
    self.CustomListCount: float
    '''Returns the number of defined custom lists (including built-in lists). Read-only Long.'''
    self.CutCopyMode: int
    '''Returns or sets the status of Cut or Copy mode. Can be True, False, or an XLCutCopyMode constant, as shown in the following tables. Read/write Long.'''
    self.DDEAppReturnCode: float
    '''Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only Long.'''
    self.DataEntryMode: int
    '''Returns or sets Data Entry mode, as shown in the following table. When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range. Read/write Long.'''
    self.DataPrivacyOptions: CDispatch
    self.DecimalSeparator: str
    '''Sets or returns the character used for the decimal separator as a String. Read/write.'''
    self.DefaultFilePath: str
    '''Returns or sets the default path that Microsoft Excel uses when it opens files. Read/write String.'''
    self.DefaultPivotTableLayoutOptions: DefaultPivotTableLayoutOptions
    self.DefaultSaveFormat: int
    '''Returns or sets the default format for saving files. For a list of valid constants, see the FileFormat property. Read/write Long.'''
    self.DefaultSheetDirection: int
    '''Returns or sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following XlReadingOrder constants: xlRTL (right to left) or xlLTR (left to right). Read/write Long.'''
    self.DefaultWebOptions: DefaultWebOptions
    '''Returns the DefaultWebOptions object that contains global application-level attributes used by Microsoft Excel whenever you save a document as a webpage or open a webpage. Read-only.'''
    self.DeferAsyncQueries: bool
    '''Gets or sets whether asynchronous queries to OLAP data sources are executed when a worksheet is calculated by VBA code. Read/write Boolean.'''
    self.DialogSheets: Sheets
    self.Dialogs: Dialogs
    '''Returns a Dialogs collection that represents all built-in dialog boxes. Read-only.'''
    self.DisplayAlerts: bool
    '''True if Microsoft Excel displays certain alerts and messages while a macro is running. Read/write Boolean.'''
    self.DisplayClipboardWindow: bool
    '''Returns True if the Microsoft Office Clipboard can be displayed. Read/write Boolean.'''
    self.DisplayCommentIndicator: int
    '''Returns or sets the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.'''
    self.DisplayDocumentActionTaskPane: bool
    '''Set to True to display the Document Actions task pane; set to False to hide the Document Actions task pane. Read/write Boolean.'''
    self.DisplayDocumentInformationPanel: bool
    '''Returns or sets a Boolean that represents whether the document properties panel is displayed. Read/write Boolean.'''
    self.DisplayExcel4Menus: bool
    '''True if Microsoft Excel displays version 4.0 menu bars. Read/write Boolean.'''
    self.DisplayFormulaAutoComplete: bool
    '''Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Read/write Boolean.'''
    self.DisplayFormulaBar: bool
    '''True if the formula bar is displayed. Read/write Boolean.'''
    self.DisplayFullScreen: bool
    '''True if Microsoft Excel is in full-screen mode. Read/write Boolean.'''
    self.DisplayFunctionToolTips: bool
    '''True if function ToolTips can be displayed. Read/write Boolean.'''
    self.DisplayInfoWindow: bool
    self.DisplayInsertOptions: bool
    '''True if the Insert Options button should be displayed. Read/write Boolean.'''
    self.DisplayNoteIndicator: bool
    '''True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners). Read/write Boolean.'''
    self.DisplayPasteOptions: bool
    '''True if the Paste Options button can be displayed. Read/write Boolean.'''
    self.DisplayRecentFiles: bool
    '''True if the list of recently used files is displayed in the UI. Read/write Boolean.'''
    self.DisplayScrollBars: bool
    '''True if scroll bars are visible for all workbooks. Read/write Boolean.'''
    self.DisplayStatusBar: bool
    '''True if the status bar is displayed. Read/write Boolean.'''
    self.EditDirectlyInCell: bool
    '''True if Microsoft Excel allows editing in cells. Read/write Boolean.'''
    self.EnableAnimations: bool
    self.EnableAutoComplete: bool
    '''True if the AutoComplete feature is enabled. Read/write Boolean.'''
    self.EnableCancelKey: int
    '''Controls how Microsoft Excel handles Ctrl+Break (or Esc or Command+Period) user interruptions to the running procedure. Read/write XlEnableCancelKey.'''
    self.EnableCheckFileExtensions: bool
    '''True to enable the Tell me if Microsoft Excel isn't the default program for viewing and editing spreadsheets dialog box. Read/write Boolean.'''
    self.EnableEvents: bool
    '''True if events are enabled for the specified object. Read/write Boolean.'''
    self.EnableLargeOperationAlert: bool
    '''Sets or returns a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office Center UI. Read/write Boolean.'''
    self.EnableLivePreview: bool
    '''Sets or returns a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Read/write Boolean.'''
    self.EnableMacroAnimations: bool
    '''Controls whether macro animations are enabled. True if user interface animations or chart animations are enabled. Is set to False (no animation) by default. If it is set to True during the running of a macro, it will enable animation, and then will reset to False after the macro runs. Read/write Boolean.'''
    self.EnableSound: bool
    '''True if sound is enabled for Microsoft Office. Read/write Boolean.'''
    self.EnableTipWizard: bool
    self.ErrorCheckingOptions: ErrorCheckingOptions
    '''Returns an ErrorCheckingOptions object, which represents the error checking options for an application.'''
    self.Excel4IntlMacroSheets: Sheets
    '''Returns a Sheets collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.'''
    self.Excel4MacroSheets: Sheets
    '''Returns a Sheets collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.'''
    self.ExtendList: bool
    '''True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. Read/write Boolean.'''
    self.FeatureInstall: int
    '''Returns or sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table. Read/write MsoFeatureInstall.'''
    self.FileExportConverters: FileExportConverters
    '''Returns a FileExportConverters collection that represents all the file converters for saving files available to Microsoft Excel. Read-only.'''
    self.FileValidation: int
    '''Returns or sets how Excel will validate files before opening them. Read/write.'''
    self.FileValidationPivot: int
    '''Returns or sets how Excel will validate the contents of the data caches for PivotTable reports. Read/write.'''
    self.FindFormat: CellFormat
    '''Sets or returns the search criteria for the type of cell formats to find.'''
    self.FixedDecimal: bool
    '''All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property. Read/write Boolean.'''
    self.FixedDecimalPlaces: int
    '''Returns or sets the number of fixed decimal places used when the FixedDecimal property is set to True. Read/write Long.'''
    self.FlashFill: bool
    '''True indicates that the Excel Flash Fill feature has been enabled and active. Read/write Boolean.'''
    self.FlashFillMode: bool
    '''True if the Flash Fill feature is enabled. Read/write Boolean.'''
    self.FormulaBarHeight: int
    '''Allows the user to specify the height of the formula bar in lines. Read/write Long.'''
    self.GenerateGetPivotData: bool
    '''Returns True when Microsoft Excel can get PivotTable report data. Read/write Boolean.'''
    self.GenerateTableRefs: int
    '''The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Read/write.'''
    self.Height: float
    '''Returns or sets a Double value that represents the height, in points, of the main application window.'''
    self.HighQualityModeForGraphics: bool
    '''Returns or sets whether Excel uses high quality mode to print graphics. Read/write.'''
    self.HinstancePtr: int
    '''Returns a handle to the instance of Excel represented by the specified Application object. Read-only Variant.'''
    self.Hwnd: int
    self.IgnoreRemoteRequests: bool
    '''True if remote DDE requests are ignored. Read/write Boolean.'''
    self.Interactive: bool
    '''True if Microsoft Excel is in interactive mode; this property is usually True. If you set this property to False, Excel blocks all input from the keyboard and mouse (except input to dialog boxes that are displayed by your code). Read/write Boolean.'''
    self.International: tuple
    '''Returns information about the current country/region and international settings. Read-only Variant.'''
    self.IsSandboxed: bool
    '''Returns True if the specified workbook is open in a Protected View window. Read-only.'''
    self.Iteration: bool
    '''True if Microsoft Excel uses iteration to resolve circular references. Read/write Boolean.'''
    self.LanguageSettings: CDispatch
    '''Returns the LanguageSettings object, which contains information about the language settings in Microsoft Excel. Read-only.'''
    self.LargeButtons: bool
    self.LargeOperationCellThousandCount: int
    '''Returns or sets the maximum number of cells needed in an operation beyond which an alert is triggered. Read/write Long.'''
    self.Left: float
    '''Returns or sets a Double value that represents the distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.'''
    self.LibraryPath: str
    '''Returns the path to the Library folder, but without the final separator. Read-only String.'''
    self.MailSystem: float
    '''Returns the mail system that's installed on the host machine. Read-only XlMailSystem.'''
    self.MapPaperSize: bool
    '''True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region. Read/write Boolean.'''
    self.MathCoprocessorAvailable: bool
    '''True if a math coprocessor is available. Read-only Boolean.'''
    self.MaxChange: float
    '''Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. Read/write Double.'''
    self.MaxIterations: float
    '''Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference. Read/write Long.'''
    self.MeasurementUnit: int
    '''Specifies the measurement unit used in the application. Read/write XlMeasurementUnits.'''
    self.MemoryFree: int
    self.MemoryTotal: int
    self.MemoryUsed: int
    self.MenuBars: MenuBars
    self.MergeInstances: bool
    '''True to merge multiple instances of the application into a single instance. Read/write Boolean.'''
    self.Modules: Modules
    self.MouseAvailable: bool
    '''True if a mouse is available. Read-only Boolean.'''
    self.MoveAfterReturn: bool
    '''True if the active cell is moved as soon as the Enter (Return) key is pressed. Read/write Boolean.'''
    self.MoveAfterReturnDirection: int
    '''Returns or sets the direction in which the active cell is moved when the user presses Enter. Read/write XlDirection.'''
    self.MultiThreadedCalculation: MultiThreadedCalculation
    '''Returns a MultiThreadedCalculation object that controls the multi-threaded recalculation settings. Read-only.'''
    self.Name: str
    '''Returns a String value that represents the name of the object.'''
    self.Names: Names
    '''Returns a Names collection that represents all the names in the active workbook. Read-only Names object.'''
    self.NetworkTemplatesPath: str
    '''Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. Read-only String.'''
    self.NewWorkbook: CDispatch
    '''Returns a NewFile object.'''
    self.ODBCErrors: ODBCErrors
    '''Returns an ODBCErrors collection that contains all the ODBC errors generated by the most recent query table or PivotTable report operation. Read-only.'''
    self.ODBCTimeout: int
    '''Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. Read/write Long.'''
    self.OLEDBErrors: OLEDBErrors
    '''Returns the OLEDBErrors collection, which represents the error information returned by the most recent OLE DB query. Read-only.'''
    self.OperatingSystem: str
    '''Returns the name and version number of the current operating system. Read-only String.'''
    self.OrganizationName: str
    '''Returns the registered organization name. Read-only String.'''
    self.Parent: Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.Path: str
    '''Returns a String value that represents the complete path to the application, excluding the final separator and name of the application.'''
    self.PathSeparator: str
    '''Returns the path separator character (\). Read-only String.'''
    self.PivotTableSelection: bool
    '''True if PivotTable reports use structured selection. Read/write Boolean.'''
    self.PrintCommunication: bool
    '''Specifies whether communication with the printer is turned on. Read/write Boolean.'''
    self.ProductCode: str
    '''Returns the globally unique identifier (GUID) for Microsoft Excel. Read-only String.'''
    self.PromptForSummaryInfo: bool
    '''True if Microsoft Excel asks for summary information when files are first saved. Read/write Boolean.'''
    self.ProtectedViewWindows: ProtectedViewWindows
    '''Returns a ProtectedViewWindows collection that represents all the Protected View windows that are open in the application. Read-only.'''
    self.QuickAnalysis: QuickAnalysis
    '''Returns a QuickAnalysis object that represents the Quick Analysis options of the application.'''
    self.Quitting: bool
    self.RTD: RTD
    '''Returns an RTD object.'''
    self.Ready: bool
    '''Returns True when the Microsoft Excel application is ready; False when the Excel application is not ready. Read-only Boolean.'''
    self.RecentFiles: RecentFiles
    '''Returns a RecentFiles collection that represents the list of recently used files.'''
    self.RecordRelative: bool
    '''True if macros are recorded by using relative references; False if recording is absolute. Read-only Boolean.'''
    self.ReferenceStyle: int
    '''Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. Read/write XlReferenceStyle.'''
    self.ReplaceFormat: CellFormat
    '''Sets the replacement criteria to use in replacing cell formats. The replacement criteria is then used in a subsequent call to the Replace method of the Range object.'''
    self.RollZoom: bool
    '''True if the IntelliMouse zooms instead of scrolling. Read/write Boolean.'''
    self.Rows: Range
    '''Returns a Range object that represents all the rows on the active worksheet. If the active document isn't a worksheet, the Rows property fails. Read-only Range object.'''
    self.SaveISO8601Dates: bool
    self.ScreenUpdating: bool
    '''True if screen updating is turned on. Read/write Boolean.'''
    self.Selection: Range
    '''Returns the currently selected object on the active worksheet for an Application object. Returns Nothing if no objects are selected. Use the Select method to set the selection, and use the TypeName function to discover the kind of object that is selected.'''
    self.Sheets: Sheets
    '''Returns a Sheets collection that represents all the sheets in the active workbook. Read-only Sheets object.'''
    self.SheetsInNewWorkbook: float
    '''Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks. Read/write Long.'''
    self.ShowChartTipNames: bool
    '''True if charts show chart tip names. The default value is True. Read/write Boolean.'''
    self.ShowChartTipValues: bool
    '''True if charts show chart tip values. The default value is True. Read/write Boolean.'''
    self.ShowConvertToDataType: bool
    self.ShowDevTools: bool
    '''Returns or sets a Boolean that represents whether the Developer tab is displayed in the ribbon. Read/write Boolean.'''
    self.ShowMenuFloaties: bool
    '''Returns or sets a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. False if Mini toolbars are displayed. Read/write Boolean.'''
    self.ShowQuickAnalysis: bool
    '''Controls whether the Quick Analysis contextual user interface is displayed on selection. True means that the Quick Analysis button will show.'''
    self.ShowSelectionFloaties: bool
    '''Returns or sets a Boolean that represents whether Mini toolbars displays when a user selects text. False if Mini toolbars are displayed. Read/write Boolean.'''
    self.ShowStartupDialog: bool
    '''Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application. Read/write Boolean.'''
    self.ShowToolTips: bool
    '''True if ToolTips are turned on. Read/write Boolean.'''
    self.ShowWindowsInTaskbar: bool
    self.SmartArtColors: CDispatch
    '''Returns the set of SmartArtColors styles that are currently loaded in the application. Read-only.'''
    self.SmartArtLayouts: CDispatch
    '''Returns the set of SmartArtLayouts that are currently loaded in the application. Read-only.'''
    self.SmartArtQuickStyles: CDispatch
    '''Returns the set of SmartArtQuickStyles that are currently loaded in the application. Read-only.'''
    self.SmartTagRecognizers: SmartTagRecognizers
    self.Speech: Speech
    '''Returns a Speech object.'''
    self.SpellingOptions: SpellingOptions
    '''Returns a SpellingOptions object that represents the spelling options of the application.'''
    self.StandardFont: str
    '''Returns or sets the name of the standard font. Read/write String.'''
    self.StandardFontSize: float
    '''Returns or sets the standard font size, in points. Read/write Long.'''
    self.StartupPath: str
    '''Returns the complete path of the startup folder, excluding the final separator. Read-only String.'''
    self.StatusBar: bool
    '''Returns or sets the text in the status bar. Read/write String.'''
    self.TemplatesPath: str
    '''Returns the local path where templates are stored. Read-only String.'''
    self.ThousandsSeparator: str
    '''Sets or returns the character used for the thousands separator as a String. Read/write.'''
    self.Toolbars: Toolbars
    self.Top: float
    '''Returns or sets a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.'''
    self.TransitionMenuKey: str
    '''Returns or sets the Microsoft Excel menu or help key, which is usually /. Read/write String.'''
    self.TransitionMenuKeyAction: float
    '''Returns or sets the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp (see the Excel constants enumeration). Read/write Long.'''
    self.TransitionNavigKeys: bool
    '''True if transition navigation keys are active. Read/write Boolean.'''
    self.TruncateLargeNumbers: bool
    self.TruncateLeadingZeros: bool
    self.UILanguage: int
    self.UsableHeight: float
    '''Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.'''
    self.UsableWidth: float
    '''Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.'''
    self.UseClusterConnector: bool
    '''Returns or sets whether Excel allows user-defined functions in XLL add-ins to be run on a compute cluster. Read/write.'''
    self.UseSystemSeparators: bool
    '''True (default) if the system separators of Microsoft Excel are enabled. Read/write Boolean.'''
    self.UsedObjects: UsedObjects
    '''Returns a UsedObjects object representing objects allocated in a workbook. Read-only.'''
    self.UserControl: bool
    '''True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden. Read/write Boolean.'''
    self.UserLibraryPath: str
    '''Returns the path to the location on the user's computer where the COM add-ins are installed. Read-only String.'''
    self.UserName: str
    '''Returns or sets the name of the current user. Read/write String.'''
    self.Value: str
    '''Returns a String value that represents the name of the application.'''
    self.Version: str
    '''Returns a String value that represents the Microsoft Excel version number.'''
    self.Visible: bool
    '''Returns or sets a Boolean value that determines whether the object is visible. Read/write.'''
    self.WarnOnFunctionNameConflict: bool
    '''The WarnOnFunctionNameConflict property, when set to True, raises an alert if a developer tries to create a new function by using an existing function name. Read/write Boolean.'''
    self.Watches: Watches
    '''Returns a Watches object representing a range that is tracked when the worksheet is recalculated.'''
    self.Width: float
    '''Returns or sets a Double value that represents the distance, in points, from the left edge of the application window to its right edge.'''
    self.WindowState: int
    '''Returns or sets the state of the window. Read/write XlWindowState.'''
    self.Windows: Windows
    '''Returns a Windows collection that represents all the windows in all the workbooks. Read-only Windows object.'''
    self.WindowsForPens: bool
    '''True if the computer is running under Microsoft Windows for Pen Computing. Read-only Boolean.'''
    self.Workbooks: Workbooks
    '''Returns a Workbooks collection that represents all the open workbooks. Read-only.'''
    self.WorksheetFunction: WorksheetFunction
    '''Returns the WorksheetFunction object. Read-only.'''
    self.Worksheets: Sheets
    '''For an Application object, returns a Sheets collection that represents all the worksheets in the active workbook.'''
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def ActivateMicrosoftApp(self, Index):
    '''Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.'''
  def AddChartAutoFormat(self, Chart, Name, Description):  pass
  def AddCustomList(self, ListArray, ByRow):
    '''Adds a custom list for custom autofill and/or custom sort.'''
  def Calculate(self):
    '''Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.'''
  def CalculateFull(self):
    '''Forces a full calculation of the data in all open workbooks.'''
  def CalculateFullRebuild(self):
    '''For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.'''
  def CalculateUntilAsyncQueriesDone(self):
    '''Runs all pending queries to OLEDB and OLAP data sources.'''
  def CentimetersToPoints(self, Centimeters) -> float:
    '''Converts a measurement from centimeters to points (one point equals 0.035 centimeters).'''
  def CheckAbort(self, KeepAbort):
    '''Stops recalculation in a Microsoft Excel application.'''
  def CheckSpelling(self, Word, CustomDictionary, IgnoreUppercase) -> bool:
    '''Checks the spelling of a single word.'''
  def ConvertFormula(self, Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo) -> list:
    '''Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both. Variant.'''
  def DDEExecute(self, Channel, String):
    '''Runs a command or performs some other action or actions in another application by way of the specified DDE channel.'''
  def DDEInitiate(self, App, Topic) -> int:
    '''Opens a DDE channel to an application.'''
  def DDEPoke(self, Channel, Item, Data):
    '''Sends data to an application.'''
  def DDERequest(self, Channel, Item) -> list:
    '''Requests information from the specified application. This method always returns an array.'''
  def DDETerminate(self, Channel):
    '''Closes a channel to another application.'''
  def DeleteChartAutoFormat(self, Name):  pass
  def DeleteCustomList(self, ListNum):
    '''Deletes a custom list.'''
  def DisplayXMLSourcePane(self, XmlMap):
    '''Opens the XML Source task pane and displays the XML map specified by the XmlMap argument.'''
  def DoubleClick(self):
    '''Equivalent to double-clicking the active cell.'''
  def Dummy1(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Dummy10(self, arg):  pass
  def Dummy11(self):  pass
  def Dummy12(self, p1, p2):  pass
  def Dummy13(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Dummy14(self):  pass
  def Dummy2(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8):  pass
  def Dummy20(self, grfCompareFunctions):  pass
  def Dummy3(self):  pass
  def Dummy4(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15):  pass
  def Dummy5(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13):  pass
  def Dummy6(self):  pass
  def Dummy7(self):  pass
  def Dummy8(self, Arg1):  pass
  def Dummy9(self):  pass
  def Evaluate(self, Name) -> list:
    '''Converts a Microsoft Excel name to an object or a value.'''
  def ExecuteExcel4Macro(self, String) -> list:
    '''Runs a Microsoft Excel 4.0 macro function and then returns the result of the function. The return type depends on the function.'''
  def FileDialog(self, fileDialogType):
    '''Returns a FileDialog object representing an instance of the file dialog.'''
  def FindFile(self) -> bool:
    '''Displays the Open dialog box.'''
  def GetCaller(self, Index):  pass
  def GetClipboardFormats(self, Index):  pass
  def GetCustomListContents(self, ListNum) -> list:
    '''Returns a custom list (an array of strings).'''
  def GetCustomListNum(self, ListArray) -> int:
    '''Returns the custom list number for an array of strings. Use this method to match both built-in lists and custom-defined lists.'''
  def GetFileConverters(self, Index1, Index2):  pass
  def GetInternational(self, Index):  pass
  def GetOpenFilename(self, FileFilter, FilterIndex, Title, ButtonText, MultiSelect) -> list:
    '''Displays the standard Open dialog box and gets a file name from the user without actually opening any files.'''
  def GetPhonetic(self, Text) -> str:
    '''Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.'''
  def GetPreviousSelections(self, Index):  pass
  def GetRegisteredFunctions(self, Index1, Index2):  pass
  def GetSaveAsFilename(self, InitialFilename, FileFilter, FilterIndex, Title, ButtonText) -> list:
    '''Displays the standard Save As dialog box and gets a file name from the user without actually saving any files.'''
  def Goto(self, Reference, Scroll):
    '''Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it's not already active.'''
  def Help(self, HelpFile, HelpContextID):
    '''Displays a Help topic.'''
  def InchesToPoints(self, Inches) -> float:
    '''Converts a measurement from inches to points.'''
  def InputBox(self, Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type) -> list:
    '''Displays a dialog box for user input. Returns the information entered in the dialog box.'''
  def Intersect(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> Range:
    '''Returns a Range object that represents the rectangular intersection of two or more ranges. If one or more ranges from a different worksheet are specified, an error is returned.'''
  def MacroOptions(self, Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile, ArgumentDescriptions):
    '''Corresponds to options in the Macro Options dialog box. You can also use this method to display a user-defined function (UDF) in a built-in or new category within the Insert Function dialog box.'''
  def MailLogoff(self):
    '''Closes a MAPI mail session established by Microsoft Excel.'''
  def MailLogon(self, Name, Password, DownloadNewMail):
    '''Logs on to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.'''
  def NextLetter(self) -> Workbook:
    '''You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.'''
  def OnKey(self, Key, Procedure):
    '''Runs a specified procedure when a particular key or key combination is pressed.'''
  def OnRepeat(self, Text, Procedure):
    '''Sets the Repeat item and the name of the procedure that will run if you choose the Repeat command after running the procedure that sets this property.'''
  def OnTime(self, EarliestTime, Procedure, LatestTime, Schedule):
    '''Schedules a procedure to be run at a specified time in the future (either at a specific time of day or after a specific amount of time has passed).'''
  def OnUndo(self, Text, Procedure):
    '''Sets the text of the Undo command and the name of the procedure that's run if you choose the Undo command after running the procedure that sets this property.'''
  def Quit(self):
    '''Quits Microsoft Excel.'''
  def Range(self, Cell1, Cell2):
    '''Returns a Range object that represents a cell or a range of cells.'''
  def RecordMacro(self, BasicCode, XlmCode):
    '''Records code if the macro recorder is on.'''
  def RegisterXLL(self, Filename) -> bool:
    '''Loads an XLL code resource and automatically registers the functions and commands contained in the resource.'''
  def Repeat(self):
    '''Repeats the last user-interface action.'''
  def ResetTipWizard(self):  pass
  def Run(self, Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> list:
    '''Runs a macro or calls a function. This can be used to run a macro written in Visual Basic or the Microsoft Excel macro language, or to run a function in a DLL or XLL.'''
  def Save(self, Filename):  pass
  def SaveWorkspace(self, Filename):  pass
  def SendKeys(self, Keys, Wait):
    '''Sends keystrokes to the active application.'''
  def SetDefaultChart(self, FormatName, Gallery):  pass
  def SharePointVersion(self, bstrUrl) -> int:
    '''Returns the version number of SharePoint Foundation instances running at the site for the specified URL.'''
  def ShortcutMenus(self, Index):  pass
  def Support(self, Object, ID, arg):  pass
  def Undo(self):
    '''Cancels the last user-interface action.'''
  def Union(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> Range:
    '''Returns the union of two or more ranges.'''
  def Volatile(self, Volatile):
    '''Marks a user-defined function as volatile. A volatile function must be recalculated whenever calculation occurs in any cells on the worksheet. A nonvolatile function is recalculated only when the input variables change. This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell.'''
  def Wait(self, Time) -> bool:
    '''Pauses a running macro until a specified time. Returns True if the specified time has arrived.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Evaluate(self, Name):  pass
  def _FindFile(self):  pass
  def _MacroOptions(self, Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile):  pass
  def _Run2(self, Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def _WSFunction(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def _Wait(self, Time):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # ActiveChart:  <class 'NoneType'>
    # ActiveDialog:  <class 'NoneType'>
    # ActiveProtectedViewWindow:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # FileConverters:  <class 'NoneType'>
    # MailSession:  <class 'NoneType'>
    # OnCalculate:  <class 'NoneType'>
    # OnData:  <class 'NoneType'>
    # OnDoubleClick:  <class 'NoneType'>
    # OnEntry:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # OnWindow:  <class 'NoneType'>
    # PreviousSelections:  <class 'NoneType'>
    # RegisteredFunctions:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'PyIID'>

  #getattr AttributeError:

  #getattr Exception:
    # AnswerWizard:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Dummy101:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Dummy22:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # Dummy23:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FileFind:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FileSearch:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FormatStaleValues:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # Hinstance:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147418113), None)
    # SensitivityLabelPolicy:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147220726), None)
    # ThisCell:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ThisWorkbook:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # VBE:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不信任到 Visual Basic Project 的程序连接\n', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Application'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:227, methods:94, whats:18,   ok:339, er:0, er:12


# num=2
class Range:
  def __init__(self):
    self.AddIndent: bool
    '''Returns or sets a Variant value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).'''
    self.Address: str
    '''Returns a String value that represents the range reference in the language of the macro.'''
    self.AddressLocal: str
    '''Returns the range reference for the specified range in the language of the user. Read-only String.'''
    self.AllowEdit: bool
    '''Returns a Boolean value that indicates if the range can be edited on a protected worksheet.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Areas: Areas
    '''Returns an Areas collection that represents all the ranges in a multiple-area selection. Read-only.'''
    self.Borders: Borders
    '''Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).'''
    self.Cells: Range
    '''Returns a Range object that represents the cells in the specified range.'''
    self.Characters: Characters
    '''Returns a Characters object that represents a range of characters within the object text. Use the Characters object to format characters within a text string.'''
    self.Column: int
    '''Returns the number of the first column in the first area in the specified range. Read-only Long.'''
    self.ColumnWidth: float
    '''Returns or sets the width of all columns in the specified range. Read/write Double.'''
    self.Columns: Range
    '''Returns a Range object that represents the columns in the specified range.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.CountLarge: int
    '''Returns a value that represents the number of objects in the collection. Read-only Variant.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.CurrentRegion: Range
    '''Returns a Range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. Read-only.'''
    self.DisplayFormat: DisplayFormat
    '''Returns a DisplayFormat object that represents the display settings for the specified range. Read-only.'''
    self.EntireColumn: Range
    '''Returns a Range object that represents the entire column (or columns) that contains the specified range. Read-only.'''
    self.EntireRow: Range
    '''Returns a Range object that represents the entire row (or rows) that contains the specified range. Read-only.'''
    self.Errors: Errors
    '''Allows the user to access error checking options.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the specified object.'''
    self.FormatConditions: FormatConditions
    '''Returns a FormatConditions collection that represents all the conditional formats for the specified range. Read-only.'''
    self.Formula: str
    '''Returns or sets a Variant value that represents the object's implicitly intersecting formula in A1-style notation.'''
    self.Formula2: str
    self.Formula2Local: str
    self.Formula2R1C1: str
    self.Formula2R1C1Local: str
    self.FormulaArray: str
    '''Returns or sets the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns null. Read/write Variant.'''
    self.FormulaHidden: bool
    '''Returns or sets a Variant value that indicates if the formula will be hidden when the worksheet is protected.'''
    self.FormulaLocal: str
    '''Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write Variant.'''
    self.FormulaR1C1: str
    '''Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write Variant.'''
    self.FormulaR1C1Local: str
    '''Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write Variant.'''
    self.HasArray: bool
    '''True if the specified cell is part of an array formula. Read-only Variant.'''
    self.HasFormula: bool
    '''True if all cells in the range contain formulas; False if none of the cells in the range contains a formula; null otherwise. Read-only Variant.'''
    self.HasRichDataType: bool
    '''True if all cells in the range contain a Rich data type. False if none of the cells in the range contains a Rich data type; otherwise, null. Read-only Variant.'''
    self.HasSpill: bool
    self.Height: float
    '''Returns a Double value that represents the height, in points, of the range. Read-only.'''
    self.HorizontalAlignment: int
    '''Returns or sets a Variant value that represents the horizontal alignment for the specified object. Read/write.'''
    self.Hyperlinks: Hyperlinks
    '''Returns a Hyperlinks collection that represents the hyperlinks for the range.'''
    self.ID: str
    '''Returns or sets a String value that represents the identifying label for the specified cell when the page is saved as a webpage.'''
    self.IndentLevel: int
    '''Returns or sets a Variant value that represents the indent level for the cell or range. Can be an integer from 0 to 15.'''
    self.Interior: Interior
    '''Returns an Interior object that represents the interior of the specified object.'''
    self.Left: float
    '''Returns a Variant value that represents the distance, in points, from the left edge of column A to the left edge of the range.'''
    self.LinkedDataTypeState: int
    '''Returns information about the state of any Linked data types, such as Stocks or Geography, in the range. Possible values are from the XlLinkedDataTypeState enumeration. Read-only.'''
    self.ListHeaderRows: int
    '''Returns the number of header rows for the specified range. Read-only Long.'''
    self.Locked: bool
    '''Returns or sets a Variant value that indicates if the object is locked.'''
    self.MDX: str
    '''Returns the MDX name for the specified Range object. Read-only String.'''
    self.MergeArea: Range
    '''Returns a Range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. Read-only Variant.'''
    self.MergeCells: bool
    '''True if the range contains merged cells. Read/write Variant.'''
    self.Next: Range
    '''Returns a Range object that represents the next cell.'''
    self.NumberFormat: str
    '''Returns or sets a Variant value that represents the format code for the object.'''
    self.NumberFormatLocal: str
    '''Returns or sets a Variant value that represents the format code for the object as a string in the language of the user.'''
    self.Offset: Range
    '''Returns a Range object that represents a range that's offset from the specified range.'''
    self.Orientation: int
    '''Returns or sets a Variant value that represents the text orientation.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.Phonetic: Phonetic
    '''Returns the Phonetic object, which contains information about a specific phonetic text string in a cell.'''
    self.Phonetics: Phonetics
    '''Returns the Phonetics collection of the range. Read-only.'''
    self.PrefixCharacter: str
    '''Returns the prefix character for the cell. Read-only Variant.'''
    self.ReadingOrder: int
    '''Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.'''
    self.Resize: Range
    '''Resizes the specified range. Returns a Range object that represents the resized range.'''
    self.Row: int
    '''Returns the number of the first row of the first area in the range. Read-only Long.'''
    self.RowHeight: float
    '''Returns or sets the height of the first row in the range specified, measured in points. Read/write Double.'''
    self.Rows: Range
    '''Returns a Range object that represents the rows in the specified range.'''
    self.SavedAsArray: bool
    self.ShrinkToFit: bool
    '''Returns or sets a Variant value that indicates if text automatically shrinks to fit in the available column width.'''
    self.SmartTags: SmartTags
    self.SoundNote: SoundNote
    self.SparklineGroups: SparklineGroups
    '''Returns a SparklineGroups object that represents an existing group of sparklines from the specified range. Read-only.'''
    self.Style: Style
    '''Returns or sets a Variant value containing a Style object that represents the style of the specified range.'''
    self.Text: str
    '''Returns the formatted text for the specified object. Read-only String.'''
    self.Top: float
    '''Returns a Variant value that represents the distance, in points, from the top edge of row 1 to the top edge of the range.'''
    self.UseStandardHeight: bool
    '''True if the row height of the Range object equals the standard height of the sheet. Returns Null if the range contains more than one row and the rows aren't all the same height. Read/write Variant.'''
    self.UseStandardWidth: bool
    '''True if the column width of the Range object equals the standard width of the sheet. Returns null if the range contains more than one column and the columns aren't all the same width. Read/write Variant.'''
    self.Validation: Validation
    '''Returns the Validation object that represents data validation for the specified range. Read-only.'''
    self.Value: float
    '''Returns or sets a Variant value that represents the value of the specified range.'''
    self.Value2: float
    '''Returns or sets the cell value. Read/write Variant.'''
    self.VerticalAlignment: int
    '''Returns or sets a Variant value that represents the vertical alignment of the specified object. Read/write.'''
    self.Width: float
    '''Returns a Double value that represents the width of a range in points. Read-only.'''
    self.Worksheet: Worksheet
    '''Returns a Worksheet object that represents the worksheet containing the specified range. Read-only.'''
    self.WrapText: bool
    '''Returns or sets a Variant value that indicates if Microsoft Excel wraps the text in the object.'''
    self.XPath: XPath
    '''Returns an XPath object that represents the XPath of the element mapped to the specified Range object. The context of the range determines whether the action succeeds or returns an empty object. Read-only.'''
    self._Default: float
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self) -> list:
    '''Activates a single cell, which must be inside the current selection. To select a range of cells, use the Select method.'''
  def AddComment(self, Text) -> Comment:
    '''Adds a comment to the range.'''
  def AddCommentThreaded(self, Text) -> CommentThreaded:
    '''Adds a new modern threaded comment to the range if no comment already exists.'''
  def AdvancedFilter(self, Action, CriteriaRange, CopyToRange, Unique) -> list:
    '''Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.'''
  def AllocateChanges(self):
    '''Performs a writeback operation for all edited cells in a range based on an OLAP data source.'''
  def ApplyNames(self, Names, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order, AppendLast) -> list:
    '''Applies names to the cells in the specified range.'''
  def ApplyOutlineStyles(self) -> list:
    '''Applies outlining styles to the specified range.'''
  def AutoComplete(self, String) -> str:
    '''Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.'''
  def AutoFill(self, Destination, Type) -> list:
    '''Performs an autofill on the cells in the specified range.'''
  def AutoFilter(self, Field, Criteria1, Operator, Criteria2, VisibleDropDown, SubField) -> list:
    '''Filters a list by using the AutoFilter.'''
  def AutoFit(self) -> list:
    '''Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.'''
  def AutoFormat(self, Format, Number, Font, Alignment, Border, Pattern, Width):  pass
  def AutoOutline(self) -> list:
    '''Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.'''
  def BorderAround(self, LineStyle, Weight, ColorIndex, Color, ThemeColor) -> list:
    '''Adds a border to a range and sets the Color, LineStyle, and Weight properties of the Border object for the new border. Variant.'''
  def Calculate(self) -> list:
    '''Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the table in the Remarks section.'''
  def CalculateRowMajorOrder(self) -> list:
    '''Calculates a specified range of cells.'''
  def CheckSpelling(self, CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang) -> list:
    '''Checks the spelling of an object.'''
  def Clear(self) -> list:
    '''Clears the entire object.'''
  def ClearComments(self):
    '''Clears all cell comments from the specified range.'''
  def ClearContents(self) -> list:
    '''Clears formulas and values from the range.'''
  def ClearFormats(self) -> list:
    '''Clears the formatting of the object.'''
  def ClearHyperlinks(self) -> None:
    '''Removes all hyperlinks from the specified range.'''
  def ClearNotes(self) -> list:
    '''Clears notes and sound notes from all the cells in the specified range.'''
  def ClearOutline(self) -> list:
    '''Clears the outline for the specified range.'''
  def ColumnDifferences(self, Comparison) -> Range:
    '''Returns a Range object that represents all the cells whose contents are different from the comparison cell in each column.'''
  def Consolidate(self, Sources, Function, TopRow, LeftColumn, CreateLinks) -> list:
    '''Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet. Variant.'''
  def ConvertToLinkedDataType(self, ServiceID, LanguageCulture):
    '''Attempts to convert all the cells in the range to a Linked data type such as Stocks or Geography.'''
  def Copy(self, Destination) -> list:
    '''Copies the range to the specified range or to the Clipboard.'''
  def CopyFromRecordset(self, Data, MaxRows, MaxColumns) -> int:
    '''Copies the contents of an ADO or DAO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range. If the Recordset object contains fields with OLE objects in them, this method fails.'''
  def CopyPicture(self, Appearance, Format) -> list:
    '''Copies the selected object to the Clipboard as a picture. Variant.'''
  def CreateNames(self, Top, Left, Bottom, Right) -> list:
    '''Creates names in the specified range, based on text labels in the sheet.'''
  def CreatePublisher(self, Edition, Appearance, ContainsPICT, ContainsBIFF, ContainsRTF, ContainsVALU):  pass
  def Cut(self, Destination) -> list:
    '''Cuts the object to the Clipboard or pastes it into a specified destination.'''
  def DataSeries(self, Rowcol, Type, Date, Step, Stop, Trend) -> list:
    '''Creates a data series in the specified range. Variant.'''
  def DataTypeToText(self):
    '''If any of the cells in the range are a Linked data type, such as Stocks or Geography, this call will convert their values to text.'''
  def Delete(self, Shift) -> list:
    '''Deletes the object.'''
  def DialogBox(self) -> list:
    '''Displays a dialog box defined by a dialog box definition table on a Microsoft Excel 4.0 macro sheet. Returns the number of the chosen control, or returns False if the user chooses the Cancel button.'''
  def Dirty(self):
    '''Designates a range to be recalculated when the next recalculation occurs.'''
  def DiscardChanges(self):
    '''Discards all changes in the edited cells of the range.'''
  def EditionOptions(self, Type, Option, Name, Reference, Appearance, ChartSize, Format) -> list:
    '''You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.'''
  def End(self, Direction):
    '''Returns a Range object that represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW. Read-only Range object.'''
  def ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity):
    '''Exports to a file of the specified format.'''
  def FillDown(self) -> list:
    '''Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.'''
  def FillLeft(self) -> list:
    '''Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.'''
  def FillRight(self) -> list:
    '''Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.'''
  def FillUp(self) -> list:
    '''Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.'''
  def Find(self, What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat) -> A_Range_object_that_represents_the_first_cell_where_that_information_is_found_:
    '''Finds specific information in a range.'''
  def FindNext(self, After) -> Range:
    '''Continues a search that was begun with the Find method. Finds the next cell that matches those same conditions and returns a Range object that represents that cell. This does not affect the selection or the active cell.'''
  def FindPrevious(self, After) -> Range:
    '''Continues a search that was begun with the Find method. Finds the previous cell that matches those same conditions and returns a Range object that represents that cell. Doesn't affect the selection or the active cell.'''
  def FlashFill(self) -> None:
    '''True indicates that the Excel Flash Fill feature has been enabled and is active.'''
  def FunctionWizard(self) -> list:
    '''Starts the Function Wizard for the upper-left cell of the range.'''
  def GetAddress(self, RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo):  pass
  def GetAddressLocal(self, RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo):  pass
  def GetCharacters(self, Start, Length):  pass
  def GetOffset(self, RowOffset, ColumnOffset):  pass
  def GetResize(self, RowSize, ColumnSize):  pass
  def GetValue(self, RangeValueDataType):  pass
  def Get_Default(self, RowIndex, ColumnIndex):  pass
  def GoalSeek(self, Goal, ChangingCell):  pass
  def Group(self, Start, End, By, Periods) -> list:
    '''When the Range object represents a single cell in a PivotTable field's data range, the Group method performs numeric or date-based grouping in that field.'''
  def Insert(self, Shift, CopyOrigin) -> list:
    '''Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.'''
  def InsertIndent(self, InsertAmount):
    '''Adds an indent to the specified range.'''
  def Item(self, RowIndex, ColumnIndex):
    '''Returns a Range object that represents a range at an offset to the specified range.'''
  def Justify(self) -> list:
    '''Rearranges the text in a range so that it fills the range evenly.'''
  def ListNames(self) -> list:
    '''Pastes a list of all nonhidden names onto the worksheet, beginning with the first cell in the range.'''
  def Merge(self, Across):
    '''Creates a merged cell from the specified Range object.'''
  def NavigateArrow(self, TowardPrecedent, ArrowNumber, LinkNumber) -> list:
    '''Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a Range object that represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.'''
  def NoteText(self, Text, Start, Length) -> str:
    '''Returns or sets the cell note associated with the cell in the upper-left corner of the range. Read/write String. Cell notes have been replaced by range comments. For more information, see the Comment object.'''
  def Parse(self, ParseLine, Destination) -> list:
    '''Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.'''
  def PasteSpecial(self, Paste, Operation, SkipBlanks, Transpose) -> list:
    '''Pastes a Range object that has been copied into the specified range.'''
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName) -> list:
    '''Prints the object.'''
  def PrintPreview(self, EnableChanges) -> list:
    '''Shows a preview of the object as it would look when printed.'''
  def Range(self, Cell1, Cell2):
    '''Returns a Range object that represents a cell or a range of cells.'''
  def RefreshLinkedDataType(self, DomainID):  pass
  def RemoveDuplicates(self, Columns, Header):
    '''Removes duplicate values from a range of values.'''
  def RemoveSubtotal(self) -> list:
    '''Removes subtotals from a list.'''
  def Replace(self, What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat, FormulaVersion):  pass
  def RowDifferences(self, Comparison) -> Range:
    '''Returns a Range object that represents all the cells whose contents are different from those of the comparison cell in each row.'''
  def Run(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> list:
    '''Runs the Microsoft Excel macro at this location. The range must be on a macro sheet.'''
  def Select(self) -> list:
    '''Selects the object.'''
  def SetCellDataTypeFromCell(self, SourceCell):
    '''Creates another instance of a Linked data type, such as Stocks or Geography, that exists in another cell. The new instance will be linked to the data source in the same way as the original, so it will refresh from the service if you call the Workbook.RefreshAll method.'''
  def SetItem(self, RowIndex, ColumnIndex, arg2):  pass
  def SetPhonetic(self):
    '''Creates Phonetic objects for all the cells in the specified range.'''
  def SetValue(self, RangeValueDataType, arg1):  pass
  def Set_Default(self, RowIndex, ColumnIndex, arg2):  pass
  def Show(self) -> list:
    '''Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.'''
  def ShowCard(self):
    '''For a cell containing a Linked data type, such as Stocks or Geography, this method causes a card to appear that shows details about the cell (that is, the same card that the user can view by choosing the cell icon).'''
  def ShowDependents(self, Remove) -> list:
    '''Draws tracer arrows to the direct dependents of the range.'''
  def ShowErrors(self) -> list:
    '''Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.'''
  def ShowPrecedents(self, Remove) -> list:
    '''Draws tracer arrows to the direct precedents of the range.'''
  def Sort(self, Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3, SubField1) -> list:
    '''Sorts a range of values.'''
  def SortSpecial(self, SortMethod, Key1, Order1, Type, Key2, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, DataOption1, DataOption2, DataOption3) -> list:
    '''Uses East Asian sorting methods to sort the range, a PivotTable report, or uses the method for the active region if the range contains only one cell. For example, Japanese sorts in the order of the Kana syllabary.'''
  def Speak(self, SpeakDirection, SpeakFormulas):
    '''Causes the cells of the range to be spoken in row order or column order.'''
  def SpecialCells(self, Type, Value) -> Range:
    '''Returns a Range object that represents all the cells that match the specified type and value.'''
  def SubscribeTo(self, Edition, Format) -> list:
    '''You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.'''
  def Subtotal(self, GroupBy, Function, TotalList, Replace, PageBreaks, SummaryBelowData) -> list:
    '''Creates subtotals for the range (or the current region, if the range is a single cell).'''
  def Table(self, RowInput, ColumnInput) -> list:
    '''Creates a data table based on input values and formulas that you define on a worksheet.'''
  def TextToColumns(self, Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers) -> list:
    '''Parses a column of cells that contain text into several columns.'''
  def UnMerge(self):
    '''Separates a merged area into individual cells.'''
  def Ungroup(self) -> list:
    '''Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _AutoFilter(self, Field, Criteria1, Operator, Criteria2, VisibleDropDown):  pass
  def _BorderAround(self, LineStyle, Weight, ColorIndex, Color):  pass
  def _ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr):  pass
  def _PasteSpecial(self, Paste, Operation, SkipBlanks, Transpose):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def _PrintOut_(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate):  pass
  def _Replace(self, What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat):  pass
  def _Sort(self, Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3):  pass
  def __call__(self, RowIndex, ColumnIndex):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # Comment:  <class 'NoneType'>
    # CommentThreaded:  <class 'NoneType'>
    # ListObject:  <class 'NoneType'>
    # SpillParent:  <class 'NoneType'>
    # SpillingToRange:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # CurrentArray:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '未找到单元格。', 'xlmain11.chm', 0, -2146827284), None)
    # Dependents:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '未找到单元格。', 'xlmain11.chm', 0, -2146827284), None)
    # DirectDependents:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '未找到单元格。', 'xlmain11.chm', 0, -2146827284), None)
    # DirectPrecedents:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '未找到单元格。', 'xlmain11.chm', 0, -2146827284), None)
    # FormulaLabel:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Hidden:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 Hidden 属性', 'xlmain11.chm', 0, -2146827284), None)
    # LocationInTable:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 LocationInTable 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Name:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # OutlineLevel:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 OutlineLevel 属性', 'xlmain11.chm', 0, -2146827284), None)
    # PageBreak:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 PageBreak 属性', 'xlmain11.chm', 0, -2146827284), None)
    # PivotCell:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # PivotField:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 PivotField 属性', 'xlmain11.chm', 0, -2146827284), None)
    # PivotItem:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 PivotItem 属性', 'xlmain11.chm', 0, -2146827284), None)
    # PivotTable:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 PivotTable 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Precedents:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '未找到单元格。', 'xlmain11.chm', 0, -2146827284), None)
    # Previous:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 Previous 属性', 'xlmain11.chm', 0, -2146827284), None)
    # QueryTable:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ServerActions:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ShowDetail:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 ShowDetail 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Summary:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Range 的 Summary 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Range'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:86, methods:118, whats:9,   ok:213, er:0, er:20


# num=3
class MenuBar:
  def __init__(self):
    self.Application: Application
    self.BuiltIn: bool
    self.Caption: str
    self.Creator: int
    self.Index: int
    self.Menus: Menus
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self):  pass
  def Delete(self):  pass
  def Reset(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MenuBar'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:8, whats:4,   ok:23, er:0, er:0


# num=4
class _Worksheet:
  def __init__(self):
    self.Application: Application
    self.AutoFilterMode: bool
    self.Cells: Range
    self.CodeName: str
    self.Columns: Range
    self.Comments: Comments
    self.CommentsThreaded: CommentsThreaded
    self.ConsolidationFunction: int
    self.ConsolidationOptions: tuple
    self.Creator: int
    self.CustomProperties: CustomProperties
    self.DisplayAutomaticPageBreaks: bool
    self.DisplayPageBreaks: bool
    self.DisplayRightToLeft: bool
    self.EnableAutoFilter: bool
    self.EnableCalculation: bool
    self.EnableFormatConditionsCalculation: bool
    self.EnableOutlining: bool
    self.EnablePivotTable: bool
    self.EnableSelection: int
    self.FilterMode: bool
    self.HPageBreaks: HPageBreaks
    self.Hyperlinks: Hyperlinks
    self.Index: int
    self.ListObjects: ListObjects
    self.Name: str
    self.NamedSheetViews: NamedSheetViewCollection
    self.Names: Names
    self.Outline: Outline
    self.PageSetup: PageSetup
    self.Parent: _Workbook
    self.PrintedCommentPages: int
    self.ProtectContents: bool
    self.ProtectDrawingObjects: bool
    self.ProtectScenarios: bool
    self.Protection: Protection
    self.ProtectionMode: bool
    self.QueryTables: QueryTables
    self.Rows: Range
    self.Scripts: CDispatch
    self.ScrollArea: str
    self.Shapes: Shapes
    self.SmartTags: SmartTags
    self.Sort: Sort
    self.StandardHeight: float
    self.StandardWidth: float
    self.Tab: Tab
    self.TransitionExpEval: bool
    self.TransitionFormEntry: bool
    self.Type: int
    self.UsedRange: Range
    self.VPageBreaks: VPageBreaks
    self.Visible: int
    self._CodeName: str
    self._DisplayRightToLeft: bool
    self._Sort: Sort
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self):  pass
  def Arcs(self, Index):  pass
  def Buttons(self, Index):  pass
  def Calculate(self):  pass
  def ChartObjects(self, Index):  pass
  def CheckBoxes(self, Index):  pass
  def CheckSpelling(self, CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang):  pass
  def CircleInvalid(self):  pass
  def ClearArrows(self):  pass
  def ClearCircles(self):  pass
  def Copy(self, Before, After):  pass
  def Delete(self):  pass
  def DrawingObjects(self, Index):  pass
  def Drawings(self, Index):  pass
  def DropDowns(self, Index):  pass
  def Evaluate(self, Name):  pass
  def ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity):  pass
  def GroupBoxes(self, Index):  pass
  def GroupObjects(self, Index):  pass
  def Labels(self, Index):  pass
  def Lines(self, Index):  pass
  def ListBoxes(self, Index):  pass
  def Move(self, Before, After):  pass
  def OLEObjects(self, Index):  pass
  def OptionButtons(self, Index):  pass
  def Ovals(self, Index):  pass
  def Paste(self, Destination, Link):  pass
  def PasteSpecial(self, Format, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, NoHTMLFormatting):  pass
  def Pictures(self, Index):  pass
  def PivotTableWizard(self, SourceType, SourceData, TableDestination, TableName, RowGrand, ColumnGrand, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection):  pass
  def PivotTables(self, Index):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas):  pass
  def PrintPreview(self, EnableChanges):  pass
  def Protect(self, Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables):  pass
  def Range(self, Cell1, Cell2):  pass
  def Rectangles(self, Index):  pass
  def ResetAllPageBreaks(self):  pass
  def SaveAs(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local):  pass
  def Scenarios(self, Index):  pass
  def ScrollBars(self, Index):  pass
  def Select(self, Replace):  pass
  def SetBackgroundPicture(self, Filename):  pass
  def ShowAllData(self):  pass
  def ShowDataForm(self):  pass
  def Spinners(self, Index):  pass
  def TextBoxes(self, Index):  pass
  def Unprotect(self, Password):  pass
  def XmlDataQuery(self, XPath, SelectionNamespaces, Map):  pass
  def XmlMapQuery(self, XPath, SelectionNamespaces, Map):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _CheckSpelling(self, CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang, IgnoreFinalYaa, SpellScript):  pass
  def _Evaluate(self, Name):  pass
  def _ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr):  pass
  def _PasteSpecial(self, Format, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def _PrintOut_(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate):  pass
  def _Protect(self, Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly):  pass
  def _SaveAs(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local):  pass
  def _SaveAs_(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # AutoFilter:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # CircularReference:  <class 'NoneType'>
    # ConsolidationSources:  <class 'NoneType'>
    # Next:  <class 'NoneType'>
    # OnCalculate:  <class 'NoneType'>
    # OnData:  <class 'NoneType'>
    # OnDoubleClick:  <class 'NoneType'>
    # OnEntry:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # Previous:  <class 'NoneType'>
    # _AutoFilter:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'PyIID'>

  #getattr AttributeError:

  #getattr Exception:
    # MailEnvelope:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Worksheet'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:60, methods:63, whats:16,   ok:139, er:0, er:1


# num=5
class Window:
  def __init__(self):
    self.ActiveCell: Range
    '''Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.'''
    self.ActivePane: Pane
    '''Returns a Pane object that represents the active pane in the window. Read-only.'''
    self.ActiveSheet: _Worksheet
    '''Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.'''
    self.ActiveSheetView: WorksheetView
    '''Returns an object that represents the view of the active sheet in the specified window. Read-only.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.AutoFilterDateGrouping: bool
    '''True if the auto filter for date grouping is currently displayed in the specified window. Read/write Boolean.'''
    self.Caption: str
    '''Returns or sets a Variant value that represents the name that appears in the title bar of the document window.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DisplayFormulas: bool
    '''True if the window is displaying formulas; False if the window is displaying values. Read/write Boolean.'''
    self.DisplayGridlines: bool
    '''True if gridlines are displayed. Read/write Boolean.'''
    self.DisplayHeadings: bool
    '''True if both row and column headings are displayed; False if no headings are displayed. Read/write Boolean.'''
    self.DisplayHorizontalScrollBar: bool
    '''True if the horizontal scroll bar is displayed. Read/write Boolean.'''
    self.DisplayOutline: bool
    '''True if outline symbols are displayed. Read/write Boolean.'''
    self.DisplayRightToLeft: bool
    '''True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right. Read-only Boolean.'''
    self.DisplayRuler: bool
    '''True if a ruler is displayed for the specified window. Read/write Boolean.'''
    self.DisplayVerticalScrollBar: bool
    '''True if the vertical scroll bar is displayed. Read/write Boolean.'''
    self.DisplayWhitespace: bool
    '''True if whitespace is displayed. Read/write Boolean.'''
    self.DisplayWorkbookTabs: bool
    '''True if the workbook tabs are displayed. Read/write Boolean.'''
    self.DisplayZeros: bool
    '''True if zero values are displayed. Read/write Boolean.'''
    self.EnableResize: bool
    '''True if the window can be resized. Read/write Boolean.'''
    self.FreezePanes: bool
    '''True if split panes are frozen. Read/write Boolean.'''
    self.GridlineColor: float
    '''Returns or sets the gridline color as an RGB value. Read/write Long.'''
    self.GridlineColorIndex: int
    '''Returns or sets the gridline color as an index into the current color palette or as an XlColorIndex constant.'''
    self.Height: float
    '''Returns or sets a Double value that represents the height, in points, of the window.'''
    self.Hwnd: int
    self.Index: int
    '''Returns a Long value that represents the index number of the object within the collection of similar objects.'''
    self.Left: float
    '''Returns or sets a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.'''
    self.Panes: Panes
    '''Returns a Panes collection that represents all the panes in the specified window. Read-only.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.RangeSelection: Range
    '''Returns a Range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.'''
    self.ScrollColumn: float
    '''Returns or sets the number of the leftmost column in the pane or window. Read/write Long.'''
    self.ScrollRow: float
    '''Returns or sets the number of the row that appears at the top of the pane or window. Read/write Long.'''
    self.SelectedSheets: Sheets
    '''Returns a Sheets collection that represents all the selected sheets in the specified window. Read-only.'''
    self.Selection: Range
    '''Returns the specified window, for a Windows object.'''
    self.SheetViews: SheetViews
    '''Returns the SheetViews object for the specified window. Read-only.'''
    self.Split: bool
    '''True if the window is split. Read/write Boolean.'''
    self.SplitColumn: int
    '''Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/write Long.'''
    self.SplitHorizontal: int
    '''Returns or sets the location of the horizontal window split, in points. Read/write Double.'''
    self.SplitRow: int
    '''Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/write Long.'''
    self.SplitVertical: int
    '''Returns or sets the location of the vertical window split, in points. Read/write Double.'''
    self.TabRatio: float
    '''Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.'''
    self.Top: float
    '''Returns or sets a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).'''
    self.Type: int
    '''Returns or sets an XlWindowType value that represents the window type.'''
    self.UsableHeight: float
    '''Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.'''
    self.UsableWidth: float
    '''Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.'''
    self.View: int
    '''Returns or sets the view showing in the window. Read/write XlWindowView.'''
    self.Visible: bool
    '''Returns or sets a Boolean value that determines whether the object is visible. Read/write.'''
    self.VisibleRange: Range
    '''Returns a Range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.'''
    self.Width: float
    '''Returns or sets a Double value that represents the width, in points, of the window.'''
    self.WindowNumber: float
    '''Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1. Read-only Long.'''
    self.WindowState: int
    '''Returns or sets the state of the window. Read/write XlWindowState.'''
    self.Zoom: float
    '''Returns or sets a Variant value that represents the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).'''
    self._DisplayRightToLeft: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self) -> list:
    '''Brings the window to the front of the z-order.'''
  def ActivateNext(self) -> list:
    '''Activates the specified window and then sends it to the back of the window z-order.'''
  def ActivatePrevious(self) -> list:
    '''Activates the specified window and then activates the window at the back of the window z-order.'''
  def Close(self, SaveChanges, Filename, RouteWorkbook) -> bool:
    '''Closes the object.'''
  def LargeScroll(self, Down, Up, ToRight, ToLeft) -> list:
    '''Scrolls the contents of the window by pages.'''
  def NewWindow(self) -> Window:
    '''Creates a new window or a copy of the specified window.'''
  def PointsToScreenPixelsX(self, Points) -> int:
    '''Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.'''
  def PointsToScreenPixelsY(self, Points) -> int:
    '''Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.'''
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName) -> list:
    '''Prints the object.'''
  def PrintPreview(self, EnableChanges) -> list:
    '''Shows a preview of the object as it would look when printed.'''
  def RangeFromPoint(self, x, y) -> object:
    '''Returns the Shape or Range object that is positioned at the specified pair of screen coordinates. If there isn't a shape located at the specified coordinates, this method returns Nothing.'''
  def ScrollIntoView(self, Left, Top, Width, Height, Start):
    '''Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the Start argument).'''
  def ScrollWorkbookTabs(self, Sheets, Position) -> list:
    '''Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.'''
  def SmallScroll(self, Down, Up, ToRight, ToLeft) -> list:
    '''Scrolls the contents of the window by rows or columns.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # ActiveChart:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # OnWindow:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Window 的 OnWindow 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Window'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:57, methods:20, whats:5,   ok:82, er:0, er:1


# num=6
class Workbook(_Workbook):
  def __init__(self):
    self.__dict__: dict
    self.__module__: str
    self._dispobj_: _Workbook

  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __maybe__call__(self, args, kwargs):  pass
  def __maybe__int__(self, args):  pass
  def __maybe__iter__(self):  pass
  def __maybe__len__(self):  pass
  def __maybe__nonzero__(self):  pass
  def __maybe__str__(self, args):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # coclass_interfaces:  <class 'list'>
    # coclass_sources:  <class 'list'>
    # default_interface:  <class 'type'>
    # default_source:  <class 'type'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Workbook'>, <class 'win32com.client.CoClassBaseClass'>, <class 'object'>)", attrs:3, methods:8, whats:6,   ok:17, er:0, er:0


# num=7
class AddIns:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Filename, CopyFile) -> An_AddIn_object_that_represents_the_new_add_in_:
    '''Adds a new add-in file to the list of add-ins. Returns an AddIn object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AddIns'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=8
class AddIns2:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Filename, CopyFile) -> AddIn:
    '''Adds a new add-in to the list of add-ins.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AddIns2'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=9
class CDispatch:
  def __init__(self):
    self.Application: _Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._builtMethods_: dict
    self._lazydata_: tuple
    self._mapCachedItems_: dict
    self._olerepr_: LazyDispatchItem
    self._username_: str

  def Add(self, Anchor, Location, Language, Id, Extended, ScriptText):  pass
  def AddRef(self):  pass
  def Delete(self):  pass
  def GetIDsOfNames(self, riid, rgszNames, cNames, lcid, rgdispid):  pass
  def GetTypeInfo(self, itinfo, lcid, pptinfo):  pass
  def GetTypeInfoCount(self, pctinfo):  pass
  def Invoke(self, dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr):  pass
  def Item(self, Index):  pass
  def QueryInterface(self, riid, ppvObj):  pass
  def Release(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _FlagAsMethod(self, methodNames):  pass
  def _LazyAddAttr_(self, attr):  pass
  def _NewEnum(self):  pass
  def _Release_(self):  pass
  def _UpdateWithITypeInfo_(self, items_dict, typeInfo):  pass
  def __AttrToID__(self, attr):  pass
  def __LazyMap__(self, attr):  pass
  def __bool__(self):  pass
  def __call__(self, args):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, index):  pass
  def __int__(self):  pass
  def __len__(self):  pass
  def __setitem__(self, index, args):  pass
  def _dir_ole_(self):  pass
  def _find_dispatch_type_(self, methodName):  pass
  def _get_good_object_(self, ob, userName, ReturnCLSID):  pass
  def _get_good_single_object_(self, ob, userName, ReturnCLSID):  pass
  def _make_method_(self, name):  pass
  def _print_details_(self):  pass
  def _proc_(self, name, args):  pass
  def _wrap_dispatch_(self, ob, userName, returnCLSID, UnicodeToString):  pass

  #unknown:
    # __weakref__:  <class 'NoneType'>
    # _enum_:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # _unicode_to_string_:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.client.CDispatch'>, <class 'win32com.client.dynamic.CDispatch'>, <class 'object'>)", attrs:11, methods:33, whats:4,   ok:48, er:0, er:0


# num=10
class AutoCorrect:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.AutoExpandListRange: bool
    '''A Boolean value indicating whether automatic expansion is enabled for lists. When you type in a cell of an empty row or column next to a list, the list will expand to include that row or column if automatic expansion is enabled. Read/write Boolean.'''
    self.AutoFillFormulasInLists: bool
    '''Affects the creation of calculated columns created by automatic fill-down lists. Read/write Boolean.'''
    self.CapitalizeNamesOfDays: bool
    '''True if the first letter of day names is capitalized automatically. Read/write Boolean.'''
    self.CorrectCapsLock: bool
    '''True if Microsoft Excel automatically corrects accidental use of the CapsLock key. Read/write Boolean.'''
    self.CorrectSentenceCap: bool
    '''True if Microsoft Excel automatically corrects sentence (first word) capitalization. Read/write Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DisplayAutoCorrectOptions: bool
    '''Allows the user to display or hide the AutoCorrect Options button. The default value is True. Read/write Boolean.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.ReplaceText: bool
    '''True if text in the list of AutoCorrect replacements is replaced automatically. Read/write Boolean.'''
    self.ReplacementList: tuple
    '''Returns the array of AutoCorrect replacements.'''
    self.TwoInitialCapitals: bool
    '''True if words that begin with two capital letters are corrected automatically. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AddReplacement(self, What, Replacement) -> list:
    '''Adds an entry to the array of AutoCorrect replacements.'''
  def DeleteReplacement(self, What) -> list:
    '''Deletes an entry from the array of AutoCorrect replacements.'''
  def GetReplacementList(self, Index):  pass
  def SetReplacementList(self, Index, arg1):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # ConvertNumbersWithECharacter:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # TruncateLargeNumbers:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # TruncateLeadingZeros:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AutoCorrect'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:9, whats:4,   ok:29, er:0, er:3


# num=11
class AutoRecover:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Enabled: bool
    '''True if the object is enabled. Read/write Boolean.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.Path: str
    '''Returns or sets a String value that represents the complete path to where Microsoft Excel will store the AutoRecover temporary files.'''
    self.Time: int
    '''Sets or returns the time interval for the AutoRecover object. Permissible values are integers from 1 to 120 minutes. The default value is 10 minutes. Read/write Long.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AutoRecover'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0


# num=12
class Sheets:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.HPageBreaks: HPageBreaks
    '''Returns an HPageBreaks collection that represents the horizontal page breaks on the sheet. Read-only.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.VPageBreaks: VPageBreaks
    '''Returns a VPageBreaks collection that represents the vertical page breaks on the sheet. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before, After, Count, Type) -> An_Object_value_that_represents_the_new_worksheet__chart__or_macro_sheet_:
    '''Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet.'''
  def Add2(self, Before, After, Count, NewLayout) -> object:
    '''This method is only implemented for the Charts collection object and will produce a run-time error if used on the Sheets and Worksheets objects.'''
  def Copy(self, Before, After):
    '''Copies the sheet to another location in the workbook.'''
  def Delete(self):
    '''Deletes the object.'''
  def FillAcrossSheets(self, Range, Type):
    '''Copies a range to the same area on all other worksheets in a collection.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def Move(self, Before, After):
    '''Moves the sheet to another location in the workbook.'''
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas) -> list:
    '''Prints the object.'''
  def PrintPreview(self, EnableChanges):
    '''Shows a preview of the object as it would look when printed.'''
  def Select(self, Replace):
    '''Selects the object.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def _PrintOut_(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Visible:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Sheets 的 Visible 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Sheets'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:23, whats:4,   ok:37, er:0, er:1


# num=13
class DefaultPivotTableLayoutOptions:
  def __init__(self):
    self.AllowMultipleFilters: bool
    self.Application: Application
    self.CalculatedMembersInFilters: bool
    self.ColumnGrand: bool
    self.CompactRowIndent: int
    self.Creator: int
    self.DisplayContextTooltips: bool
    self.DisplayEmptyColumn: bool
    self.DisplayEmptyRow: bool
    self.DisplayErrorString: bool
    self.DisplayFieldCaptions: bool
    self.DisplayImmediateItems: bool
    self.DisplayMemberPropertyTooltips: bool
    self.DisplayNullString: bool
    self.EnableDrilldown: bool
    self.EnableWriteback: bool
    self.ErrorString: str
    self.FieldListSortAscending: bool
    self.HasAutoFormat: bool
    self.InGridDropZones: bool
    self.LayoutBlankLine: bool
    self.MergeLabels: bool
    self.NullString: str
    self.PageFieldOrder: bool
    self.PageFieldWrapCount: int
    self.Parent: _Application
    self.PreserveFormatting: bool
    self.PrintDrillIndicators: bool
    self.PrintTitles: bool
    self.RefreshOnFileOpen: bool
    self.RepeatAllLabels: int
    self.RepeatItemsOnEachPrintedPage: bool
    self.RowAxisLayout: int
    self.RowGrand: bool
    self.SaveData: bool
    self.ShowDrillIndicators: bool
    self.ShowValuesRow: bool
    self.SortUsingCustomLists: bool
    self.SubtotalHiddenPageItems: bool
    self.SubtotalLocation: bool
    self.Subtotals: bool
    self.TotalsAnnotation: bool
    self.ViewCalculatedMembers: bool
    self.VisualTotals: bool
    self.VisualTotalsForSets: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict
    self.xlMissingItemsNone: int

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DefaultPivotTableLayoutOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:50, methods:5, whats:4,   ok:59, er:0, er:0


# num=14
class DefaultWebOptions:
  def __init__(self):
    self.AllowPNG: bool
    '''True if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a webpage. False if PNG is not allowed as an output format. The default value is False. Read/write Boolean.'''
    self.AlwaysSaveInDefaultEncoding: bool
    '''True if the default encoding is used when you save a webpage or plain text document, independent of the file's original encoding when opened. False if the original encoding of the file is used. The default value is False. Read/write Boolean.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.CheckIfOfficeIsHTMLEditor: bool
    '''True if Microsoft Excel checks to see whether an Office application is the default HTML editor when you start Excel. False if Excel does not perform this check. The default value is True. Read/write Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DownloadComponents: bool
    '''True if the necessary Microsoft Office Web components are downloaded when you view the saved document in a web browser, but only if the components are not already installed. False if the components are not downloaded. The default value is False. Read/write Boolean.'''
    self.Encoding: int
    '''Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. The default is the system code page. Read/write MsoEncoding.'''
    self.FolderSuffix: str
    '''Returns the folder suffix that Microsoft Excel uses when you save a document as a webpage, use long file names, and choose to save supporting files in a separate folder (that is, if the UseLongFileNames and OrganizeInFolder properties are set to True). Read-only String.'''
    self.Fonts: CDispatch
    '''Returns the WebPageFonts collection representing the set of fonts Microsoft Excel uses when you open a webpage in Excel and there is either no font information specified on the webpage, or the current default font can't display the character set on the webpage. Read-only.'''
    self.LoadPictures: bool
    '''True if images are loaded when you open a document in Microsoft Excel, usually when the images and document were not created in Microsoft Excel. False if the images are not loaded. The default value is True. Read/write Boolean.'''
    self.LocationOfComponents: str
    '''Returns or sets the central URL (on the intranet or web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write String.'''
    self.OrganizeInFolder: bool
    '''True if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. False if supporting files are saved in the same folder as the webpage. The default value is True. Read/write Boolean.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.PixelsPerInch: int
    '''Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write Long.'''
    self.RelyOnCSS: bool
    '''True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your webpage, depending on the value of the OrganizeInFolder property. False if HTML <FONT> tags and cascading style sheets are used. The default value is True. Read/write Boolean.'''
    self.RelyOnVML: bool
    '''True if image files are not generated from drawing objects when you save a document as a webpage. False if images are generated. The default value is False. Read/write Boolean.'''
    self.SaveHiddenData: bool
    '''True if data outside of the specified range is saved when you save the document as a webpage. This data may be necessary for maintaining formulas. False if data outside of the specified range is not saved with the webpage. The default value is True. Read/write Boolean.'''
    self.SaveNewWebPagesAsWebArchives: bool
    '''True if new webpages can be saved as web archives. Read/write Boolean.'''
    self.ScreenSize: int
    '''Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Can be one of the MsoScreenSize constants. The default constant is msoScreenSize800x600. Read/write MsoScreenSize.'''
    self.TargetBrowser: int
    '''Returns or sets an MsoTargetBrowser constant indicating the browser version. Read/write.'''
    self.UpdateLinksOnSave: bool
    '''True if hyperlinks and paths to all supporting files are automatically updated before you save the document as a webpage, ensuring that the links are up to date at the time the document is saved. False if the links are not updated. The default value is True. Read/write Boolean.'''
    self.UseLongFileNames: bool
    '''True if long file names are used when you save the document as a webpage. False if long file names are not used and the DOS file name format (8.3) is used. The default value is True. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DefaultWebOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:26, methods:5, whats:4,   ok:35, er:0, er:0


# num=15
class Dialogs:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Dialogs'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=16
class ErrorCheckingOptions:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.BackgroundChecking: bool
    '''Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors. Read/write Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.EmptyCellReferences: bool
    '''When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking. Read/write Boolean.'''
    self.EvaluateToError: bool
    '''When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value. Read/write Boolean.'''
    self.InconsistentFormula: bool
    '''When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check. Read/write Boolean.'''
    self.InconsistentTableFormula: bool
    '''Returns True if the table formula is inconsistent. Read/write Boolean.'''
    self.IndicatorColorIndex: int
    '''Returns or sets the color of the indicator for error checking options. Read/write XlColorIndex.'''
    self.ListDataValidation: bool
    '''A Boolean value that is True if data validation is enabled in a list. Read/write Boolean.'''
    self.MisleadingNumberFormats: bool
    self.NumberAsText: bool
    '''When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text. Read/write Boolean.'''
    self.OmittedCells: bool
    '''When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells. Read/write Boolean.'''
    self.OutdatedLinkedDataType: bool
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.TextDate: bool
    '''When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year. Read/write Boolean.'''
    self.UnlockedFormulaCells: bool
    '''When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ErrorCheckingOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:20, methods:5, whats:4,   ok:29, er:0, er:0


# num=17
class FileExportConverters:
  def __init__(self):
    self.Application: Application
    '''Returns an Application object that represents the Microsoft Excel application. Read-only.'''
    self.Count: int
    '''Returns a Long that represents the number of file converters in the collection. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns an Object that represents the parent object of the specified FileExportConverters object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns an individual FileExportConverter object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.FileExportConverters'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=18
class CellFormat:
  def __init__(self):
    self.Application: Application
    self.Borders: Borders
    self.Creator: int
    self.Font: Font
    self.Interior: Interior
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Clear(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # AddIndent:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # FormulaHidden:  <class 'NoneType'>
    # HorizontalAlignment:  <class 'NoneType'>
    # IndentLevel:  <class 'NoneType'>
    # Locked:  <class 'NoneType'>
    # NumberFormat:  <class 'NoneType'>
    # NumberFormatLocal:  <class 'NoneType'>
    # Orientation:  <class 'NoneType'>
    # ShrinkToFit:  <class 'NoneType'>
    # VerticalAlignment:  <class 'NoneType'>
    # WrapText:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # MergeCells:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CellFormat'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:6, whats:15,   ok:31, er:0, er:1


# num=19
class MenuBars:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name):  pass
  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MenuBars'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=20
class Modules:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.HPageBreaks: HPageBreaks
    self.Parent: _Workbook
    self.VPageBreaks: VPageBreaks
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before, After, Count):  pass
  def Add2(self, Before, After, Count, NewLayout):  pass
  def Copy(self, Before, After):  pass
  def Delete(self):  pass
  def Item(self, Index):  pass
  def Move(self, Before, After):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas):  pass
  def Select(self, Replace):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def _PrintOut_(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Visible:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Sheets 的 Visible 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Modules'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:21, whats:4,   ok:35, er:0, er:1


# num=21
class MultiThreadedCalculation:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Enabled: bool
    '''The Enabled property allows MultiThreadedCalculation objects to be enabled or disabled at run time. Read/write.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.ThreadCount: int
    '''Gets the total count of the process threads that are a part of the specified MultiThreadedCalculation object.'''
    self.ThreadMode: int
    '''Returns or sets the thread mode for the specified MultiThreadedCalculation object. Read/write XlThreadMode.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MultiThreadedCalculation'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0


# num=22
class Names:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, RefersTo, Visible, MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal, RefersToR1C1, RefersToR1C1Local) -> Name:
    '''Defines a new name for a range of cells.'''
  def Item(self, Index, IndexLocal, RefersTo) -> A_Name_object_contained_by_the_collection_:
    '''Returns a single Name object from a Names collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index, IndexLocal, RefersTo):  pass
  def __call__(self, Index, IndexLocal, RefersTo):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Names'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=23
class ODBCErrors:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index) -> ODBCError:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ODBCErrors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=24
class OLEDBErrors:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index) -> OLEDBError:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.OLEDBErrors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=25
class ProtectedViewWindows:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def Open(self, Filename, Password, AddToMru, RepairMode) -> ProtectedViewWindow:
    '''Opens the specified workbook in a new Protected View window.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ProtectedViewWindows'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=26
class QuickAnalysis:
  def __init__(self):
    self.Application: Application
    '''Returns an Application object that represents the Microsoft Excel application. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns an Object that represents the parent object of the specified QuickAnalysis object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Hide(self, XlQuickAnalysisMode) -> None:
    '''Hides specific members of the Analysis Lens user interface.'''
  def Show(self, XlQuickAnalysisMode) -> None:
    '''Displays specific members of the Analysis Lens user interface.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.QuickAnalysis'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:7, whats:4,   ok:18, er:0, er:0


# num=27
class RTD:
  def __init__(self):
    self.ThrottleInterval: int
    '''Returns or sets a Long indicating the time interval between updates. Read/write.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def RefreshData(self):
    '''Requests an update of real-time data from the real-time data server.'''
  def RestartServers(self):
    '''Reconnects to a real-time data server (RTD).'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.RTD'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:7, whats:4,   ok:16, er:0, er:0


# num=28
class RecentFiles:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Maximum: int
    '''Returns or sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. Read/write Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name) -> RecentFile:
    '''Adds a file to the list of recently used files.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.RecentFiles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0


# num=29
class SmartTagRecognizers:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.Recognize: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SmartTagRecognizers'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:12, whats:4,   ok:25, er:0, er:0


# num=30
class Speech:
  def __init__(self):
    self.Direction: int
    '''Returns or sets the order in which the cells will be spoken. The value of the Direction property is an XlSpeakDirection constant. Read/write.'''
    self.SpeakCellOnEnter: bool
    '''Microsoft Excel supports a mode where the active cell is spoken when the Enter key is pressed or when the active cell is finished being edited. Setting the SpeakCellOnEnter property to True turns this mode on. False turns this mode off. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Speak(self, Text, SpeakAsync, SpeakXML, Purge):
    '''Microsoft Excel plays back the text string that is passed as an argument.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Speech'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:6, methods:6, whats:4,   ok:16, er:0, er:0


# num=31
class SpellingOptions:
  def __init__(self):
    self.ArabicModes: int
    '''Returns or sets the mode for the Arabic spelling checker. Read/write XlArabicModes.'''
    self.ArabicStrictAlefHamza: bool
    '''Returns or sets whether the spelling checker uses rules regarding Arabic words beginning with an alef hamza. Read/write.'''
    self.ArabicStrictFinalYaa: bool
    '''Returns or sets whether the spelling checker uses rules regarding Arabic words ending with the letter yaa. Read/write.'''
    self.ArabicStrictTaaMarboota: bool
    '''Returns or sets whether the spelling checker uses rules to flag Arabic words ending with haa instead of taa marboota. Read/write.'''
    self.BrazilReform: int
    '''Returns or sets the mode for checking the spelling of Brazilian Portuguese. Read/write.'''
    self.DictLang: int
    '''Selects the dictionary language used when Microsoft Excel performs spelling checks. Read/write Long.'''
    self.GermanPostReform: bool
    '''True to check the spelling of words by using the German post-reform rules. False cancels this feature. Read/write Boolean.'''
    self.HebrewModes: int
    '''Returns or sets the mode for the Hebrew spelling checker. Read/write XlHebrewModes.'''
    self.IgnoreCaps: bool
    '''False instructs Microsoft Excel to check for uppercase words; True instructs Excel to ignore words in uppercase when using the spelling checker. Read/write Boolean.'''
    self.IgnoreFileNames: bool
    '''False instructs Microsoft Excel to check for Internet and file addresses; True instructs Excel to ignore Internet and file addresses when using the spell checker. Read/write Boolean.'''
    self.IgnoreMixedDigits: bool
    '''False instructs Microsoft Excel to check for mixed digits; True instructs Excel to ignore mixed digits when checking spelling. Read/write Boolean.'''
    self.KoreanCombineAux: bool
    '''When set to True, Microsoft Excel combines Korean auxiliary verbs and adjectives when spelling is checked. Read/write Boolean.'''
    self.KoreanProcessCompound: bool
    '''When set to True, this enables Microsoft Excel to process Korean compound nouns when using the spelling checker. Read/write Boolean.'''
    self.KoreanUseAutoChangeList: bool
    '''When set to True, this enables Microsoft Excel to use the auto-change list for Korean words when using the spelling checker. Read/write Boolean.'''
    self.PortugalReform: int
    '''Returns or sets the mode for checking the spelling of European Portuguese. Read/write.'''
    self.RussianStrictE: bool
    '''Returns or sets whether the spelling checker uses rules regarding Russian words containing the character ë. Read/write.'''
    self.SpanishModes: int
    '''Returns or sets the mode for checking the spelling of Spanish. Read/write.'''
    self.SuggestMainOnly: bool
    '''When set to True, instructs Microsoft Excel to suggest words from only the main dictionary when using the spelling checker. False removes the limits of suggesting words from only the main dictionary when using the spelling checker. Read/write Boolean.'''
    self.UserDict: str
    '''Instructs Microsoft Excel to create a custom dictionary to which new words can be added when performing spelling checks on a worksheet. Read/write String.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SpellingOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:23, methods:5, whats:4,   ok:32, er:0, er:0


# num=32
class Toolbars:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name):  pass
  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Toolbars'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=33
class UsedObjects:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.UsedObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=34
class Watches:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Source) -> Watch:
    '''Adds a range that is tracked when the worksheet is recalculated.'''
  def Delete(self):
    '''Deletes the object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Watches'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=35
class Windows:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.SyncScrollingSideBySide: bool
    '''True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Arrange(self, ArrangeStyle, ActiveWorkbook, SyncHorizontal, SyncVertical) -> list:
    '''Arranges the windows on the screen.'''
  def BreakSideBySide(self) -> bool:
    '''Ends side-by-side mode if two windows are in side-by-side mode. Returns a Boolean value that represents whether the method was successful.'''
  def CompareSideBySideWith(self, WindowName) -> bool:
    '''Opens two windows in side-by-side mode. Returns a Boolean value.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def ResetPositionsSideBySide(self):
    '''Resets the position of two worksheet windows that are being compared side by side.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Windows'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:16, whats:4,   ok:29, er:0, er:0


# num=36
class Workbooks:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Template) -> Workbook:
    '''Creates a new workbook. The new workbook becomes the active workbook.'''
  def CanCheckOut(self, Filename) -> bool:
    '''True if Microsoft Excel can check out a specified workbook from a server. Read/write Boolean.'''
  def CheckOut(self, Filename):
    '''Returns a String representing a specified workbook from a server to a local computer for editing.'''
  def Close(self):
    '''Closes the object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def Open(self, Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad) -> Workbook:
    '''Opens a workbook.'''
  def OpenDatabase(self, Filename, CommandText, CommandType, BackgroundQuery, ImportDataAs) -> Workbook:
    '''Returns a Workbook object representing a database.'''
  def OpenText(self, Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local):
    '''Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.'''
  def OpenXML(self, Filename, Stylesheets, LoadOption) -> Workbook:
    '''Opens an XML data file. Returns a Workbook object.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def _Open(self, Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru):  pass
  def _OpenText(self, Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator):  pass
  def _OpenText_(self, Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout):  pass
  def _OpenXML(self, Filename, Stylesheets):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Workbooks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:24, whats:4,   ok:36, er:0, er:0


# num=37
class WorksheetFunction:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Application
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AccrInt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns the accrued interest for a security that pays periodic interest.'''
  def AccrIntM(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the accrued interest for a security that pays interest at maturity.'''
  def Acos(self, Arg1) -> float:
    '''Returns the arccosine, or inverse cosine, of a number. The arccosine is the angle whose cosine is Arg1. The returned angle is given in radians in the range 0 (zero) to pi.'''
  def Acosh(self, Arg1) -> float:
    '''Returns the inverse hyperbolic cosine of a number. Number must be greater than or equal to 1. The inverse hyperbolic cosine is the value whose hyperbolic cosine is Arg1, so Acosh(Cosh(number)) equals Arg1.'''
  def Acot(self, Arg1) -> float:
    '''Returns the arccotangent of a number, in radians in the range 0 (zero) to pi.'''
  def Acoth(self, Arg1) -> float:
    '''Returns the inverse hyperbolic cotangent of a number.'''
  def Aggregate(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns an aggregate in a list or database.'''
  def AmorDegrc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns the depreciation for each accounting period. This function is provided for the French accounting system.'''
  def AmorLinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns the depreciation for each accounting period. This function is provided for the French accounting system.'''
  def And(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> bool:
    '''Returns True if all its arguments are True; returns False if one or more arguments is False.'''
  def Arabic(self, Arg1) -> float:
    '''Converts a Roman numeral to an Arabic numeral.'''
  def ArrayToText(self, Arg1, Arg2):  pass
  def Asc(self, Arg1) -> str:
    '''For double-byte character set (DBCS) languages, changes full-width (double-byte) characters to half-width (single-byte) characters.'''
  def Asin(self, Arg1) -> float:
    '''Returns the arcsine, or inverse sine, of a number. The arcsine is the angle whose sine is Arg1. The returned angle is given in radians in the range -pi/2 to pi/2.'''
  def Asinh(self, Arg1) -> float:
    '''Returns the inverse hyperbolic sine of a number. The inverse hyperbolic sine is the value whose hyperbolic sine is Arg1, so Asinh(Sinh(number)) equals Arg1.'''
  def Atan2(self, Arg1, Arg2) -> float:
    '''Returns the arctangent, or inverse tangent, of the specified x- and y-coordinates. The arctangent is the angle from the x-axis to a line containing the origin (0, 0) and a point with coordinates (x_num, y_num). The angle is given in radians between -pi and pi, excluding -pi.'''
  def Atanh(self, Arg1) -> float:
    '''Returns the inverse hyperbolic tangent of a number. Number must be between -1 and 1 (excluding -1 and 1).'''
  def AveDev(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the average of the absolute deviations of data points from their mean. AveDev is a measure of the variability in a data set.'''
  def Average(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the average (arithmetic mean) of the arguments.'''
  def AverageIf(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.'''
  def AverageIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29) -> float:
    '''Returns the average (arithmetic mean) of all cells that meet multiple criteria.'''
  def BahtText(self, Arg1) -> str:
    '''Converts a number to Thai text and adds a suffix of Baht.'''
  def Base(self, Arg1, Arg2, Arg3) -> str:
    '''Converts a number into a text representation with the given radix (base).'''
  def BesselI(self, Arg1, Arg2) -> float:
    '''Returns the modified Bessel function, which is equivalent to the Bessel function evaluated for purely imaginary arguments.'''
  def BesselJ(self, Arg1, Arg2) -> float:
    '''Returns the Bessel function.'''
  def BesselK(self, Arg1, Arg2) -> float:
    '''Returns the modified Bessel function, which is equivalent to the Bessel functions evaluated for purely imaginary arguments.'''
  def BesselY(self, Arg1, Arg2) -> float:
    '''Returns the Bessel function, which is also called the Weber function or the Neumann function.'''
  def BetaDist(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the beta cumulative distribution function.'''
  def BetaInv(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = BetaDist(x,...), then BetaInv(probability,...) = x.'''
  def Beta_Dist(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the beta cumulative distribution function.'''
  def Beta_Inv(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = Beta_Dist(x,...), then Beta_Inv(probability,...) = x.'''
  def Bin2Dec(self, Arg1) -> str:
    '''Converts a binary number to decimal.'''
  def Bin2Hex(self, Arg1, Arg2) -> str:
    '''Converts a binary number to hexadecimal.'''
  def Bin2Oct(self, Arg1, Arg2) -> str:
    '''Converts a binary number to octal.'''
  def BinomDist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the individual term binomial distribution probability.'''
  def Binom_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the individual term binomial distribution probability.'''
  def Binom_Dist_Range(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the probability of a trial result using a binomial distribution.'''
  def Binom_Inv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the individual term binomial distribution probability.'''
  def Bitand(self, Arg1, Arg2) -> float:
    '''Returns a bitwise And of two numbers.'''
  def Bitlshift(self, Arg1, Arg2) -> float:
    '''Returns a value number shifted left by shift_amount bits.'''
  def Bitor(self, Arg1, Arg2) -> float:
    '''Returns a bitwise Or of two numbers.'''
  def Bitrshift(self, Arg1, Arg2) -> float:
    '''Returns a value number shifted right by shift_amount bits.'''
  def Bitxor(self, Arg1, Arg2) -> float:
    '''Returns a bitwise Exclusive Or of two numbers.'''
  def Ceiling(self, Arg1, Arg2) -> float:
    '''Returns number rounded up, away from zero, to the nearest multiple of significance.'''
  def Ceiling_Math(self, Arg1, Arg2, Arg3) -> float:
    '''Rounds a number up to the nearest integer or to the nearest multiple of significance.'''
  def Ceiling_Precise(self, Arg1, Arg2) -> float:
    '''Returns the specified number rounded to the nearest multiple of significance.'''
  def ChiDist(self, Arg1, Arg2) -> float:
    '''Returns the one-tailed probability of the chi-squared distribution.'''
  def ChiInv(self, Arg1, Arg2) -> float:
    '''Returns the inverse of the one-tailed probability of the chi-squared distribution.'''
  def ChiSq_Dist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the chi-squared distribution.'''
  def ChiSq_Dist_RT(self, Arg1, Arg2) -> float:
    '''Returns the right-tailed probability of the chi-squared distribution.'''
  def ChiSq_Inv(self, Arg1, Arg2) -> float:
    '''Returns the inverse of the left-tailed probability of the chi-squared distribution.'''
  def ChiSq_Inv_RT(self, Arg1, Arg2) -> float:
    '''Returns the inverse of the right-tailed probability of the chi-squared distribution.'''
  def ChiSq_Test(self, Arg1, Arg2) -> float:
    '''Returns the test for independence.'''
  def ChiTest(self, Arg1, Arg2) -> float:
    '''Returns the test for independence.'''
  def Choose(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> list:
    '''Uses Arg1 as the index to return a value from the list of value arguments.'''
  def Clean(self, Arg1) -> str:
    '''Removes all nonprintable characters from text.'''
  def Combin(self, Arg1, Arg2) -> float:
    '''Returns the number of combinations for a given number of items. Use Combin to determine the total possible number of groups for a given number of items.'''
  def Combina(self, Arg1, Arg2) -> float:
    '''Returns the number of combinations with repetitions for a given number of items.'''
  def Complex(self, Arg1, Arg2, Arg3) -> str:
    '''Converts real and imaginary coefficients into a complex number of the form x + yi or x + yj.'''
  def Concat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Confidence(self, Arg1, Arg2, Arg3) -> float:
    '''Returns a value that you can use to construct a confidence interval for a population mean.'''
  def Confidence_Norm(self, Arg1, Arg2, Arg3) -> float:
    '''Returns a value that you can use to construct a confidence interval for a population mean.'''
  def Confidence_T(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the confidence interval for a population mean, using a Student's t distribution.'''
  def Convert(self, Arg1, Arg2, Arg3) -> float:
    '''Converts a number from one measurement system to another. For example, Convert can translate a table of distances in miles to a table of distances in kilometers.'''
  def Correl(self, Arg1, Arg2) -> float:
    '''Returns the correlation coefficient of the Arg1 and Arg2 cell ranges.'''
  def Cosh(self, Arg1) -> float:
    '''Returns the hyperbolic cosine of a number.'''
  def Cot(self, Arg1) -> float:
    '''Returns the cotangent of an angle.'''
  def Coth(self, Arg1) -> float:
    '''Returns the hyperbolic cotangent of a number.'''
  def Count(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Counts the number of cells that contain numbers and counts numbers within the list of arguments.'''
  def CountA(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Counts the number of cells that are not empty and the values within the list of arguments.'''
  def CountBlank(self, Arg1) -> float:
    '''Counts empty cells in a specified range of cells.'''
  def CountIf(self, Arg1, Arg2) -> float:
    '''Counts the number of cells within a range that meet the given criteria.'''
  def CountIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Counts the number of cells within a range that meet multiple criteria.'''
  def CoupDayBs(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the number of days from the beginning of the coupon period to the settlement date.'''
  def CoupDays(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the number of days in the coupon period that contain the settlement date.'''
  def CoupDaysNc(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the number of days from the settlement date to the next coupon date.'''
  def CoupNcd(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns a number that represents the next coupon date after the settlement date.'''
  def CoupNum(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the number of coupons payable between the settlement date and maturity date, rounded up to the nearest whole coupon.'''
  def CoupPcd(self, Arg1, Arg2, Arg3, Arg4) -> float:  pass
  def Covar(self, Arg1, Arg2) -> float:
    '''Returns covariance, the average of the products of deviations for each data point pair.'''
  def Covariance_P(self, Arg1, Arg2) -> float:
    '''Returns population covariance, the average of the products of deviations for each data point pair.'''
  def Covariance_S(self, Arg1, Arg2) -> float:
    '''Returns the sample covariance, the average of the products of deviations for each data point pair in two data sets.'''
  def CritBinom(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.'''
  def Csc(self, Arg1) -> float:
    '''Returns the cosecant of an angle.'''
  def Csch(self, Arg1) -> float:
    '''Returns the hyperbolic cosecant of an angle.'''
  def CumIPmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the cumulative interest paid on a loan between start_period and end_period.'''
  def CumPrinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the cumulative principal paid on a loan between start_period and end_period.'''
  def DAverage(self, Arg1, Arg2, Arg3) -> float:
    '''Averages the values in a column of a list or database that match conditions that you specify.'''
  def DCount(self, Arg1, Arg2, Arg3) -> float:
    '''Counts the cells that contain numbers in a column of a list or database that match conditions that you specify.'''
  def DCountA(self, Arg1, Arg2, Arg3) -> float:
    '''Counts the nonblank cells in a column of a list or database that match conditions that you specify.'''
  def DGet(self, Arg1, Arg2, Arg3) -> list:
    '''Extracts a single value from a column of a list or database that matches conditions that you specify.'''
  def DMax(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the largest number in a column of a list or database that matches conditions that you specify.'''
  def DMin(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the smallest number in a column of a list or database that matches conditions that you specify.'''
  def DProduct(self, Arg1, Arg2, Arg3) -> float:
    '''Multiplies the values in a column of a list or database that match conditions that you specify.'''
  def DStDev(self, Arg1, Arg2, Arg3) -> float:
    '''Estimates the standard deviation of a population based on a sample by using the numbers in a column of a list or database that match conditions that you specify.'''
  def DStDevP(self, Arg1, Arg2, Arg3) -> float:
    '''Calculates the standard deviation of a population based on the entire population by using the numbers in a column of a list or database that match conditions that you specify.'''
  def DSum(self, Arg1, Arg2, Arg3) -> float:
    '''Adds the numbers in a column of a list or database that match conditions that you specify.'''
  def DVar(self, Arg1, Arg2, Arg3) -> float:
    '''Estimates the variance of a population based on a sample by using the numbers in a column of a list or database that match conditions that you specify.'''
  def DVarP(self, Arg1, Arg2, Arg3) -> float:
    '''Calculates the variance of a population based on the entire population by using the numbers in a column of a list or database that match conditions that you specify.'''
  def Days(self, Arg1, Arg2) -> float:
    '''Returns the number of days between two dates.'''
  def Days360(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the number of days between two dates based on a 360-day year (twelve 30-day months), which is used in some accounting calculations.'''
  def Db(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the depreciation of an asset for a specified period using the fixed-declining balance method.'''
  def Dbcs(self, Arg1) -> str:
    '''Converts half-width (single-byte) letters within a character string to full-width (double-byte) characters. The name of the function (and the characters that it converts) depends upon the language settings. Read/write String.'''
  def Ddb(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify.'''
  def Dec2Bin(self, Arg1, Arg2) -> str:
    '''Converts a decimal number to binary.'''
  def Dec2Hex(self, Arg1, Arg2) -> str:
    '''Converts a decimal number to hexadecimal.'''
  def Dec2Oct(self, Arg1, Arg2) -> str:
    '''Converts a decimal number to octal.'''
  def Decimal(self, Arg1, Arg2) -> float:
    '''Converts a text representation of a number in a given base into a decimal number.'''
  def Degrees(self, Arg1) -> float:
    '''Converts radians into degrees.'''
  def Delta(self, Arg1, Arg2) -> float:
    '''Tests whether two values are equal. Returns 1 if number1 = number2; otherwise, returns 0.'''
  def DevSq(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the sum of squares of deviations of data points from their sample mean.'''
  def Disc(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the discount rate for a security.'''
  def Dollar(self, Arg1, Arg2) -> str:
    '''The function described in this Help topic converts a number to text format and applies a currency symbol. The name of the function (and the symbol that it applies) depends upon your language settings.'''
  def DollarDe(self, Arg1, Arg2) -> float:
    '''Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal number. Use DollarDe to convert fractional dollar numbers, such as securities prices, to decimal numbers.'''
  def DollarFr(self, Arg1, Arg2) -> float:
    '''Converts a dollar price expressed as a decimal number into a dollar price expressed as a fraction. Use DollarFr to convert decimal numbers to fractional dollar numbers, such as securities prices.'''
  def Dummy19(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Dummy21(self, Arg1, Arg2):  pass
  def Duration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the Macauley duration for an assumed par value of $100. Duration is defined as the weighted average of the present value of the cash flows and is used as a measure of a bond price's response to changes in yield.'''
  def EDate(self, Arg1, Arg2) -> float:
    '''Returns the serial number that represents the date that is the indicated number of months before or after a specified date (the start_date). Use EDate to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.'''
  def Effect(self, Arg1, Arg2) -> float:
    '''Returns the effective annual interest rate, given the nominal annual interest rate and the number of compounding periods per year.'''
  def EncodeURL(self, Arg1):  pass
  def EoMonth(self, Arg1, Arg2) -> float:
    '''Returns the serial number for the last day of the month that is the indicated number of months before or after start_date. Use EoMonth to calculate maturity dates or due dates that fall on the last day of the month.'''
  def Erf(self, Arg1, Arg2) -> float:
    '''Returns the error function integrated between lower_limit and upper_limit.'''
  def ErfC(self, Arg1) -> float:
    '''Returns the complementary Erf function integrated between the specified parameter and infinity.'''
  def ErfC_Precise(self, Arg1) -> float:
    '''Returns the complementary error function integrated between the specified value and infinity.'''
  def Erf_Precise(self, Arg1) -> float:
    '''Returns the error function integrated between zero and lower_limit.'''
  def Even(self, Arg1) -> float:
    '''Returns number rounded up to the nearest even integer. Use this function for processing items that come in twos. For example, a packing crate accepts rows of one or two items. The crate is full when the number of items, rounded up to the nearest two, matches the crate's capacity.'''
  def ExponDist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the exponential distribution. Use ExponDist to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use ExponDist to determine the probability that the process takes at most 1 minute.'''
  def Expon_Dist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the exponential distribution. Use Expon_Dist to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use Expon_Dist to determine the probability that the process takes at most 1 minute.'''
  def FDist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the F probability distribution. Use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.'''
  def FInv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the F probability distribution. If p = FDIST(x,...), then FINV(p,...) = x.'''
  def FTest(self, Arg1, Arg2) -> float:
    '''Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.'''
  def FVSchedule(self, Arg1, Arg2) -> float:
    '''Returns the future value of an initial principal after applying a series of compound interest rates. Use FVSchedule to calculate the future value of an investment with a variable or adjustable rate.'''
  def F_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the F probability distribution.'''
  def F_Dist_RT(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the right-tailed F probability distribution. Use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.'''
  def F_Inv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the F probability distribution.'''
  def F_Inv_RT(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the right-tailed F probability distribution. If p = F_DIST_RT(x,...), then F_INV_RT(p,...) = x.'''
  def F_Test(self, Arg1, Arg2) -> float:
    '''Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.'''
  def Fact(self, Arg1) -> float:
    '''Returns the factorial of a number. The factorial of a number is equal to 1*2*3*...* number.'''
  def FactDouble(self, Arg1) -> float:
    '''Returns the double factorial of a number.'''
  def FieldValue(self, Arg1, Arg2):  pass
  def Filter(self, Arg1, Arg2, Arg3):  pass
  def FilterXML(self, Arg1, Arg2) -> list:
    '''Gets specific data from the returned XML, typically from a WebService function call.'''
  def Find(self, Arg1, Arg2, Arg3) -> float:
    '''Finds specific information on a worksheet.'''
  def FindB(self, Arg1, Arg2, Arg3) -> float:
    '''Find and FindB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.'''
  def Fisher(self, Arg1) -> float:
    '''Returns the Fisher transformation at x. This transformation produces a function that is normally distributed rather than skewed. Use this function to perform hypothesis testing on the correlation coefficient.'''
  def FisherInv(self, Arg1) -> float:
    '''Returns the inverse of the Fisher transformation. Use this transformation when analyzing correlations between ranges or arrays of data. If y = FISHER(x), then FISHERINV(y) = x.'''
  def Fixed(self, Arg1, Arg2, Arg3) -> str:
    '''Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.'''
  def Floor(self, Arg1, Arg2) -> float:
    '''Rounds number down, toward zero, to the nearest multiple of significance.'''
  def Floor_Math(self, Arg1, Arg2, Arg3) -> float:
    '''Rounds a number down, to the nearest integer or to the nearest multiple of significance.'''
  def Floor_Precise(self, Arg1, Arg2) -> float:
    '''Rounds the specified number to the nearest multiple of significance.'''
  def Forecast(self, Arg1, Arg2, Arg3) -> float:
    '''Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. Use this function to predict future sales, inventory requirements, or consumer trends.'''
  def Forecast_ETS(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Calculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm.'''
  def Forecast_ETS_ConfInt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns a confidence interval for the forecast value at the specified target date.'''
  def Forecast_ETS_STAT(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns a statistical value as a result of time series forecasting.'''
  def Forecast_ETS_Seasonality(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the length of the repetitive pattern that Excel detects for the specified time series.'''
  def Forecast_Linear(self, Arg1, Arg2, Arg3) -> float:
    '''Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. Use this function to predict future sales, inventory requirements, or consumer trends.'''
  def Frequency(self, Arg1, Arg2) -> list:
    '''Calculates how often values occur within a range of values, and then returns a vertical array of numbers. For example, use Frequency to count the number of test scores that fall within ranges of scores. Because Frequency returns an array, it must be entered as an array formula.'''
  def Fv(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the future value of an investment based on periodic, constant payments and a constant interest rate.'''
  def Gamma(self, Arg1) -> float:
    '''Returns the gamma function value.'''
  def GammaDist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the gamma distribution. Use this function to study variables that may have a skewed distribution. The gamma distribution is commonly used in queuing analysis.'''
  def GammaInv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the gamma cumulative distribution. If p = GAMMADIST(x,...), then GAMMAINV(p,...) = x.'''
  def GammaLn(self, Arg1) -> float:
    '''Returns the natural logarithm of the gamma function, Γ(x).'''
  def GammaLn_Precise(self, Arg1) -> float:
    '''Returns the natural logarithm of the gamma function, Γ(x).'''
  def Gamma_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the gamma distribution. Use this function to study variables that may have a skewed distribution. The gamma distribution is commonly used in queuing analysis.'''
  def Gamma_Inv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the gamma cumulative distribution. If p = GAMMA_DIST(x,...), then GAMMA_INV(p,...) = x.'''
  def Gauss(self, Arg1) -> float:
    '''Returns 0.5 less than the standard normal cumulative distribution.'''
  def Gcd(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the greatest common divisor of two or more integers. The greatest common divisor is the largest integer that divides both number1 and number2 without a remainder.'''
  def GeStep(self, Arg1, Arg2) -> float:
    '''Returns 1 if number ≥ step; otherwise, returns 0 (zero). Use this function to filter a set of values. For example, by summing several GeStep functions, you calculate the count of values that exceed a threshold.'''
  def GeoMean(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the geometric mean of an array or range of positive data. For example, you can use GeoMean to calculate average growth rate given compound interest with variable rates.'''
  def Growth(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Calculates predicted exponential growth by using existing data. Growth returns the y-values for a series of new x-values that you specify by using existing x-values and y-values. You can also use the Growth worksheet function to fit an exponential curve to existing x-values and y-values.'''
  def HLookup(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Searches for a value in the top row of a table or an array of values, and then returns a value in the same column from a row that you specify in the table or array. Use HLookup when your comparison values are located in a row across the top of a table of data, and you want to look down a specified number of rows. Use VLookup when your comparison values are located in a column to the left of the data that you want to find.'''
  def HarMean(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the harmonic mean of a data set. The harmonic mean is the reciprocal of the arithmetic mean of reciprocals.'''
  def Hex2Bin(self, Arg1, Arg2) -> str:
    '''Converts a hexadecimal number to binary.'''
  def Hex2Dec(self, Arg1) -> str:
    '''Converts a hexadecimal number to decimal.'''
  def Hex2Oct(self, Arg1, Arg2) -> str:
    '''Converts a hexadecimal number to octal.'''
  def HypGeomDist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the hypergeometric distribution. HypGeomDist returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use HypGeomDist for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.'''
  def HypGeom_Dist(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the hypergeometric distribution. HypGeom_Dist returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use HypGeom_Dist for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.'''
  def ISO_Ceiling(self, Arg1, Arg2) -> float:
    '''Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance.'''
  def IfError(self, Arg1, Arg2) -> list:
    '''Returns a value that you specify if a formula evaluates to an error; otherwise, returns the result of the formula. Use the IfError function to trap and handle errors in a formula.'''
  def IfNa(self, Arg1, Arg2) -> list:
    '''Returns the value that you specify if the expression resolves to #N/A; otherwise, returns the result of the expression.'''
  def ImAbs(self, Arg1) -> str:
    '''Returns the absolute value (modulus) of a complex number in x + yi or x + yj text format.'''
  def ImArgument(self, Arg1) -> str:
    '''Returns the argument  (theta), an angle expressed in radians, such that:'''
  def ImConjugate(self, Arg1) -> str:
    '''Returns the complex conjugate of a complex number in x + yi or x + yj text format.'''
  def ImCos(self, Arg1) -> str:
    '''Returns the cosine of a complex number in x + yi or x + yj text format.'''
  def ImCosh(self, Arg1) -> str:
    '''Returns the hyperbolic cosine of a complex number.'''
  def ImCot(self, Arg1) -> str:
    '''Returns the cotangent of a complex number.'''
  def ImCsc(self, Arg1) -> str:
    '''Returns the cosecant of a complex number.'''
  def ImCsch(self, Arg1) -> str:
    '''Returns the hyperbolic cosecant of a complex number.'''
  def ImDiv(self, Arg1, Arg2) -> str:
    '''Returns the quotient of two complex numbers in x + yi or x + yj text format.'''
  def ImExp(self, Arg1) -> str:
    '''Returns the exponential of a complex number in x + yi or x + yj text format.'''
  def ImLn(self, Arg1) -> str:
    '''Returns the natural logarithm of a complex number in x + yi or x + yj text format.'''
  def ImLog10(self, Arg1) -> str:
    '''Returns the common logarithm (base 10) of a complex number in x + yi or x + yj text format.'''
  def ImLog2(self, Arg1) -> str:
    '''Returns the base-2 logarithm of a complex number in x + yi or x + yj text format.'''
  def ImPower(self, Arg1, Arg2) -> str:
    '''Returns a complex number in x + yi or x + yj text format raised to a power.'''
  def ImProduct(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> str:
    '''Returns the product of 2 to 29 complex numbers in x + yi or x + yj text format.'''
  def ImReal(self, Arg1) -> float:
    '''Returns the real coefficient of a complex number in x + yi or x + yj text format.'''
  def ImSec(self, Arg1) -> str:
    '''Returns the secant of a complex number.'''
  def ImSech(self, Arg1) -> str:
    '''Returns the hyperbolic secant of a complex number.'''
  def ImSin(self, Arg1) -> str:
    '''Returns the sine of a complex number in x + yi or x + yj text format.'''
  def ImSinh(self, Arg1) -> str:
    '''Returns the hyperbolic sine of a complex number.'''
  def ImSqrt(self, Arg1) -> str:
    '''Returns the square root of a complex number in x + yi or x + yj text format.'''
  def ImSub(self, Arg1, Arg2) -> str:
    '''Returns the difference of two complex numbers in x + yi or x + yj text format.'''
  def ImSum(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> str:
    '''Returns the sum of two or more complex numbers in x + yi or x + yj text format.'''
  def ImTan(self, Arg1) -> str:
    '''Returns the tangent of a complex number.'''
  def Imaginary(self, Arg1) -> float:
    '''Returns the imaginary coefficient of a complex number in x + yi or x + yj text format.'''
  def Index(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Returns a value or the reference to a value from within a table or range. There are two forms of the Index function: the array form and the reference form.'''
  def IntRate(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the interest rate for a fully invested security.'''
  def Intercept(self, Arg1, Arg2) -> float:
    '''Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values. The intercept point is based on a best-fit regression line plotted through the known x-values and known y-values.'''
  def Ipmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the interest payment for a given period for an investment based on periodic, constant payments and a constant interest rate.'''
  def Irr(self, Arg1, Arg2) -> float:
    '''Returns the internal rate of return for a series of cash flows represented by the numbers in values. These cash flows don't have to be even, as they would be for an annuity. However, the cash flows must occur at regular intervals, such as monthly or annually. The internal rate of return is the interest rate received for an investment consisting of payments (negative values) and income (positive values) that occur at regular periods.'''
  def IsErr(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to any error value except #N/A.'''
  def IsError(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to any error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).'''
  def IsEven(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value is even.'''
  def IsFormula(self, Arg1) -> bool:
    '''Checks whether a reference is to a cell containing a formula, and returns True or False.'''
  def IsLogical(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to a logical value.'''
  def IsNA(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to the #N/A (value not available) error value.'''
  def IsNonText(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to any item that is not text. (Note that this function returns True if value refers to a blank cell.)'''
  def IsNumber(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to a number.'''
  def IsOdd(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value is odd.'''
  def IsText(self, Arg1) -> bool:
    '''Checks the type of value and returns True or False depending on whether the value refers to text.'''
  def IsThaiDigit(self, Arg1):  pass
  def IsoWeekNum(self, Arg1, Arg2) -> float:
    '''Returns the ISO week number of the year for a given date.'''
  def Ispmt(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Calculates the interest paid during a specific period of an investment. This function is provided for compatibility with Lotus 1-2-3.'''
  def Kurt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the kurtosis of a data set. Kurtosis characterizes the relative peakedness or flatness of a distribution compared with the normal distribution. Positive kurtosis indicates a relatively peaked distribution. Negative kurtosis indicates a relatively flat distribution.'''
  def Large(self, Arg1, Arg2) -> float:
    '''Returns the k-th largest value in a data set. Use this function to select a value based on its relative standing. For example, you can use Large to return the highest, runner-up, or third-place score.'''
  def Lcm(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the least common multiple of integers. The least common multiple is the smallest positive integer that is a multiple of all integer arguments number1, number2, and so on. Use Lcm to add fractions with different denominators.'''
  def LinEst(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Calculates the statistics for a line by using the least squares method to calculate a straight line that best fits your data, and returns an array that describes the line. Because this function returns an array of values, it must be entered as an array formula.'''
  def Ln(self, Arg1) -> float:
    '''Returns the natural logarithm of a number. Natural logarithms are based on the constant e (2.71828182845904).'''
  def Log(self, Arg1, Arg2) -> float:
    '''Returns the logarithm of a number to the base that you specify.'''
  def Log10(self, Arg1) -> float:
    '''Returns the base-10 logarithm of a number.'''
  def LogEst(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''In regression analysis, calculates an exponential curve that fits your data, and returns an array of values that describes the curve. Because this function returns an array of values, it must be entered as an array formula.'''
  def LogInv(self, Arg1, Arg2, Arg3) -> float:
    '''Use the lognormal distribution to analyze logarithmically transformed data.'''
  def LogNormDist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.'''
  def LogNorm_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.'''
  def LogNorm_Inv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the lognormal cumulative distribution function. Use the lognormal distribution to analyze logarithmically transformed data.'''
  def Lookup(self, Arg1, Arg2, Arg3) -> list:
    '''Returns a value either from a one-row or one-column range or from an array. The Lookup function has two syntax forms: the vector form and the array form.'''
  def MDeterm(self, Arg1) -> float:
    '''Returns the matrix determinant of an array.'''
  def MDuration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the modified Macauley duration for a security with an assumed par value of $100.'''
  def MInverse(self, Arg1) -> list:
    '''Returns the inverse matrix for the matrix stored in an array.'''
  def MIrr(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the modified internal rate of return for a series of periodic cash flows. MIrr considers both the cost of the investment and the interest received on reinvestment of cash.'''
  def MMult(self, Arg1, Arg2) -> list:
    '''Returns the matrix product of two arrays. The result is an array with the same number of rows as array1 and the same number of columns as array2.'''
  def MRound(self, Arg1, Arg2) -> float:
    '''Returns a number rounded to the desired multiple.'''
  def Match(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the relative position of an item in an array that matches a specified value in a specified order. Use Match instead of one of the Lookup functions when you need the position of an item in a range instead of the item itself.'''
  def Max(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the largest value in a set of values.'''
  def MaxIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Median(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the median of the given numbers. The median is the number in the middle of a set of numbers.'''
  def Min(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the smallest number in a set of values.'''
  def MinIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Mode(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the most frequently occurring, or repetitive, value in an array or range of data.'''
  def Mode_Mult(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> list:
    '''Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data.'''
  def Mode_Sngl(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the most frequently occurring, or repetitive, value in an array or range of data.'''
  def MultiNomial(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the ratio of the factorial of a sum of values to the product of factorials.'''
  def Munit(self, Arg1) -> list:
    '''Returns the unit matrix for the specified dimension.'''
  def NPer(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.'''
  def NegBinomDist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the negative binomial distribution. NegBinomDist returns the probability that there will be number_f failures before the number_s-th success, when the constant probability of a success is probability_s. This function is similar to the binomial distribution, except that the number of successes is fixed, and the number of trials is variable. Like the binomial, trials are assumed to be independent.'''
  def NegBinom_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the negative binomial distribution. NegBinom_Dist returns the probability that there will be number_f failures before the number_s-th success, when the constant probability of a success is probability_s. This function is similar to the binomial distribution, except that the number of successes is fixed, and the number of trials is variable. Like the binomial, trials are assumed to be independent.'''
  def NetworkDays(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the number of whole working days between start_date and end_date. Working days exclude weekends and any dates identified in holidays. Use NetworkDays to calculate employee benefits that accrue based on the number of days worked during a specific term.'''
  def NetworkDays_Intl(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.'''
  def Nominal(self, Arg1, Arg2) -> float:
    '''Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.'''
  def NormDist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the normal distribution for the specified mean and standard deviation. This function has a very wide range of applications in statistics, including hypothesis testing.'''
  def NormInv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.'''
  def NormSDist(self, Arg1) -> float:
    '''Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.'''
  def NormSInv(self, Arg1) -> float:
    '''Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.'''
  def Norm_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the normal distribution for the specified mean and standard deviation. This function has a wide range of applications in statistics, including hypothesis testing.'''
  def Norm_Inv(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.'''
  def Norm_S_Dist(self, Arg1, Arg2) -> float:
    '''Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.'''
  def Norm_S_Inv(self, Arg1) -> float:
    '''Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of 0 (zero) and a standard deviation of one.'''
  def Npv(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Calculates the net present value of an investment by using a discount rate and a series of future payments (negative values) and income (positive values).'''
  def NumberValue(self, Arg1, Arg2, Arg3) -> float:
    '''Converts text to number in a locale-independent manner.'''
  def Oct2Bin(self, Arg1, Arg2) -> str:
    '''Converts an octal number to binary.'''
  def Oct2Dec(self, Arg1) -> str:
    '''Converts an octal number to decimal.'''
  def Oct2Hex(self, Arg1, Arg2) -> str:
    '''Converts an octal number to hexadecimal.'''
  def Odd(self, Arg1) -> float:
    '''Returns number rounded up to the nearest odd integer.'''
  def OddFPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9) -> float:
    '''Returns the price per $100 face value of a security having an odd (short or long) first period.'''
  def OddFYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9) -> float:
    '''Returns the yield of a security that has an odd (short or long) first period.'''
  def OddLPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8) -> float:
    '''Returns the price per $100 face value of a security having an odd (short or long) last coupon period.'''
  def OddLYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8) -> float:
    '''Returns the yield of a security that has an odd (short or long) last period.'''
  def Or(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> bool:
    '''Returns True if any argument is True; returns False if all arguments are False.'''
  def PDuration(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the number of periods required by an investment to reach a specified value.'''
  def Pearson(self, Arg1, Arg2) -> float:
    '''Returns the Pearson product moment correlation coefficient, r, a dimensionless index that ranges from -1.0 to 1.0 inclusive and reflects the extent of a linear relationship between two data sets.'''
  def PercentRank(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a value in a data set as a percentage of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use PercentRank to evaluate the standing of an aptitude test score among all scores for the test.'''
  def PercentRank_Exc(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set.'''
  def PercentRank_Inc(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use PercentRank_Inc to evaluate the standing of an aptitude test score among all scores for the test.'''
  def Percentile(self, Arg1, Arg2) -> float:
    '''Returns the k-th percentile of values in a range. Use this function to establish a threshold of acceptance. For example, you can decide to examine candidates who score above the 90th percentile.'''
  def Percentile_Exc(self, Arg1, Arg2) -> float:
    '''Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.'''
  def Percentile_Inc(self, Arg1, Arg2) -> float:
    '''Returns the k-th percentile of values in a range. Use this function to establish a threshold of acceptance. For example, you can examine candidates who score above the 90th percentile.'''
  def Permut(self, Arg1, Arg2) -> float:
    '''Returns the number of permutations for a given number of objects that can be selected from number objects. A permutation is any set or subset of objects or events where internal order is significant. Permutations are different from combinations, for which the internal order is not significant. Use this function for lottery-style probability calculations.'''
  def Permutationa(self, Arg1, Arg2) -> float:
    '''Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.'''
  def Phi(self, Arg1) -> float:
    '''Returns the value of the density function for a standard normal distribution.'''
  def Phonetic(self, Arg1) -> float:
    '''Extracts the phonetic (furigana) characters from a text string.'''
  def Pi(self) -> float:
    '''Returns the number 3.14159265358979, the mathematical constant pi, accurate to 15 digits.'''
  def Pmt(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Calculates the payment for a loan based on constant payments and a constant interest rate.'''
  def Poisson(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in 1 minute.'''
  def Poisson_Dist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in one minute.'''
  def Power(self, Arg1, Arg2) -> float:
    '''Returns the result of a number raised to a power.'''
  def Ppmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the payment on the principal for a given period for an investment based on periodic, constant payments and a constant interest rate.'''
  def Price(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns the price per $100 face value of a security that pays periodic interest.'''
  def PriceDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the price per $100 face value of a discounted security.'''
  def PriceMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the price per $100 face value of a security that pays interest at maturity.'''
  def Prob(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the probability that values in a range are between two limits. If upper_limit is not supplied, returns the probability that values in x_range are equal to lower_limit.'''
  def Product(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Multiplies all the numbers given as arguments and returns the product.'''
  def Proper(self, Arg1) -> str:
    '''Capitalizes the first letter in a text string and any other letters in text that follow any character other than a letter. Converts all other letters to lowercase letters.'''
  def Pv(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the present value of an investment. The present value is the total amount that a series of future payments is worth now. For example, when you borrow money, the loan amount is the present value to the lender.'''
  def Quartile(self, Arg1, Arg2) -> float:
    '''Returns the quartile of a data set. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use Quartile to find the top 25 percent of incomes in a population.'''
  def Quartile_Exc(self, Arg1, Arg2) -> float:
    '''Returns the quartile of the data set, based on percentile values from 0..1, exclusive.'''
  def Quartile_Inc(self, Arg1, Arg2) -> float:
    '''Returns the quartile of a data set based on percentile values from 0..1, inclusive. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use Quartile_Inc to find the top 25 percent of incomes in a population.'''
  def Quotient(self, Arg1, Arg2) -> float:
    '''Returns the integer portion of a division. Use this function when you want to discard the remainder of a division.'''
  def RSq(self, Arg1, Arg2) -> float:
    '''Returns the square of the Pearson product moment correlation coefficient through data points in known_y's and known_x's. For more information, see Pearson. The r-squared value can be interpreted as the proportion of the variance in y attributable to the variance in x.'''
  def RTD(self, progID, server, topic1, topic2, topic3, topic4, topic5, topic6, topic7, topic8, topic9, topic10, topic11, topic12, topic13, topic14, topic15, topic16, topic17, topic18, topic19, topic20, topic21, topic22, topic23, topic24, topic25, topic26, topic27, topic28) -> list:
    '''This method connects to a source to receive real-time data (RTD).'''
  def Radians(self, Arg1) -> float:
    '''Converts degrees to radians.'''
  def RandArray(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def RandBetween(self, Arg1, Arg2) -> float:
    '''Returns a random integer number between the numbers that you specify. A new random integer number is returned every time the worksheet is calculated.'''
  def Rank(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. If you were to sort the list, the rank of the number would be its position.'''
  def Rank_Avg(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a number in a list of numbers; that is, its size relative to other values in the list. If more than one value has the same rank, the average rank is returned.'''
  def Rank_Eq(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. If you were to sort the list, the rank of the number would be its position.'''
  def Rate(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the interest rate per period of an annuity. Rate is calculated by iteration and can have zero or more solutions. If the successive results of Rate don't converge to within 0.0000001 after 20 iterations, Rate returns the #NUM! error value.'''
  def Received(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the amount received at maturity for a fully invested security.'''
  def Replace(self, Arg1, Arg2, Arg3, Arg4) -> A_String_value_that_represents_the_new_string__after_replacement_:
    '''Replaces part of a text string, based on the number of characters that you specify, with a different text string.'''
  def ReplaceB(self, Arg1, Arg2, Arg3, Arg4) -> str:
    '''Replaces part of a text string, based on the number of bytes that you specify, with a different text string.'''
  def Rept(self, Arg1, Arg2) -> str:
    '''Repeats text a given number of times. Use Rept to fill a cell with a number of instances of a text string.'''
  def Roman(self, Arg1, Arg2) -> str:
    '''Converts an arabic numeral to roman, as text.'''
  def Round(self, Arg1, Arg2) -> float:
    '''Rounds a number to a specified number of digits.'''
  def RoundBahtDown(self, Arg1):  pass
  def RoundBahtUp(self, Arg1):  pass
  def RoundDown(self, Arg1, Arg2) -> float:
    '''Rounds a number down, toward 0 (zero).'''
  def RoundUp(self, Arg1, Arg2) -> float:
    '''Rounds a number up, away from 0 (zero).'''
  def Rri(self, Arg1, Arg2, Arg3) -> float:
    '''Returns an equivalent interest rate for the growth of an investment.'''
  def Search(self, Arg1, Arg2, Arg3) -> float:
    '''Search and SearchB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.'''
  def SearchB(self, Arg1, Arg2, Arg3) -> float:
    '''Search and SearchB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string.'''
  def Sec(self, Arg1) -> float:
    '''Returns the secant of an angle.'''
  def Sech(self, Arg1) -> float:
    '''Returns the hyperbolic secant of an angle.'''
  def Sequence(self, Arg1, Arg2, Arg3, Arg4):  pass
  def SeriesSum(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the sum of a power series based on the following formula:'''
  def Single(self, Arg1):  pass
  def Sinh(self, Arg1) -> float:
    '''Returns the hyperbolic sine of a number.'''
  def Skew(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the skewness of a distribution. Skewness characterizes the degree of asymmetry of a distribution around its mean.'''
  def Skew_p(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.'''
  def Sln(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the straight-line depreciation of an asset for one period.'''
  def Slope(self, Arg1, Arg2) -> float:
    '''Returns the slope of the linear regression line through data points in known_y's and known_x's. The slope is the vertical distance divided by the horizontal distance between any two points on the line, which is the rate of change along the regression line.'''
  def Small(self, Arg1, Arg2) -> float:
    '''Returns the k-th smallest value in a data set. Use this function to return values with a particular relative standing in a data set.'''
  def Sort(self, Arg1, Arg2, Arg3, Arg4):  pass
  def SortBy(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def SqrtPi(self, Arg1) -> float:
    '''Returns the square root of (number * pi).'''
  def StDev(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).'''
  def StDevP(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).'''
  def StDev_P(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).'''
  def StDev_S(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).'''
  def StEyx(self, Arg1, Arg2) -> float:
    '''Returns the standard error of the predicted y-value for each x in the regression. The standard error is a measure of the amount of error in the prediction of y for an individual x.'''
  def Standardize(self, Arg1, Arg2, Arg3) -> float:
    '''Returns a normalized value from a distribution characterized by mean and standard_dev.'''
  def StockHistory(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Substitute(self, Arg1, Arg2, Arg3, Arg4) -> str:
    '''Substitutes new_text for old_text in a text string. Use Substitute when you want to replace specific text in a text string; use Replace when you want to replace any text that occurs in a specific location in a text string.'''
  def Subtotal(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> A_Double_value_that_represents_the_subtotal_:
    '''Creates subtotals.'''
  def Sum(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Adds all the numbers in a range of cells.'''
  def SumIf(self, Arg1, Arg2, Arg3) -> float:
    '''Adds the cells specified by a given criteria.'''
  def SumIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29) -> float:
    '''Adds the cells in a range that meet multiple criteria.'''
  def SumProduct(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Multiplies corresponding components in the given arrays, and returns the sum of those products.'''
  def SumSq(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Returns the sum of the squares of the arguments.'''
  def SumX2MY2(self, Arg1, Arg2) -> float:
    '''Returns the sum of the difference of squares of corresponding values in two arrays.'''
  def SumX2PY2(self, Arg1, Arg2) -> float:
    '''Returns the sum of the sum of squares of corresponding values in two arrays. The sum of the sum of squares is a common term in many statistical calculations.'''
  def SumXMY2(self, Arg1, Arg2) -> float:
    '''Returns the sum of squares of differences of corresponding values in two arrays.'''
  def Syd(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the sum-of-years' digits depreciation of an asset for a specified period.'''
  def TBillEq(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the bond-equivalent yield for a Treasury bill.'''
  def TBillPrice(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the price per $100 face value for a Treasury bill.'''
  def TBillYield(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the yield for a Treasury bill.'''
  def TDist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the Percentage Points (probability) for the Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.'''
  def TInv(self, Arg1, Arg2) -> float:
    '''Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom.'''
  def TTest(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the probability associated with a Student's t-Test. Use TTest to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.'''
  def T_Dist(self, Arg1, Arg2, Arg3) -> float:
    '''Returns a Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are computed.'''
  def T_Dist_2T(self, Arg1, Arg2) -> float:
    '''Returns the two-tailed Student t-distribution.'''
  def T_Dist_RT(self, Arg1, Arg2) -> float:
    '''Returns the right-tailed Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.'''
  def T_Inv(self, Arg1, Arg2) -> float:
    '''Returns the left-tailed inverse of the Student t-distribution.'''
  def T_Inv_2T(self, Arg1, Arg2) -> float:
    '''Returns the t-value of the Student t-distribution as a function of the probability and the degrees of freedom.'''
  def T_Test(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the probability associated with a Student t-Test. Use T_Test to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.'''
  def Tanh(self, Arg1) -> float:
    '''Returns the hyperbolic tangent of a number.'''
  def Text(self, Arg1, Arg2) -> str:
    '''Converts a value to text in a specific number format.'''
  def TextJoin(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def ThaiDayOfWeek(self, Arg1):  pass
  def ThaiDigit(self, Arg1):  pass
  def ThaiMonthOfYear(self, Arg1):  pass
  def ThaiNumSound(self, Arg1):  pass
  def ThaiNumString(self, Arg1):  pass
  def ThaiStringLength(self, Arg1):  pass
  def ThaiYear(self, Arg1):  pass
  def Transpose(self, Arg1) -> list:
    '''Returns a vertical range of cells as a horizontal range, or vice versa. Transpose must be entered as an array formula in a range that has the same number of rows and columns, respectively, as an array has columns and rows. Use Transpose to shift the vertical and horizontal orientation of an array on a worksheet.'''
  def Trend(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Returns values along a linear trend. Fits a straight line (using the method of least squares) to the arrays known_y's and known_x's. Returns the y-values along that line for the array of new_x's that you specify.'''
  def Trim(self, Arg1) -> str:
    '''Removes all spaces from text except for single spaces between words. Use Trim on text that you have received from another application that may have irregular spacing.'''
  def TrimMean(self, Arg1, Arg2) -> float:
    '''Returns the mean of the interior of a data set. TrimMean calculates the mean taken by excluding a percentage of data points from the top and bottom tails of a data set. Use this function when you wish to exclude outlying data from your analysis.'''
  def USDollar(self, Arg1, Arg2) -> str:
    '''Converts a number to text format and applies a currency symbol. The name of the method (and the symbol that it applies) depends upon the language settings.'''
  def Unichar(self, Arg1) -> str:
    '''Returns the Unicode character referenced by the given numeric value.'''
  def Unicode(self, Arg1) -> float:
    '''Returns the number (code point) corresponding to the first character of the text.'''
  def Unique(self, Arg1, Arg2, Arg3):  pass
  def VLookup(self, Arg1, Arg2, Arg3, Arg4) -> list:
    '''Searches for a value in the first column of a table array and returns a value in the same row from another column in the table array.'''
  def ValueToText(self, Arg1, Arg2):  pass
  def Var(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Estimates variance based on a sample.'''
  def VarP(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Calculates variance based on the entire population.'''
  def Var_P(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Calculates variance based on the entire population.'''
  def Var_S(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> float:
    '''Estimates variance based on a sample.'''
  def Vdb(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7) -> float:
    '''Returns the depreciation of an asset for any period that you specify, including partial periods, by using the double-declining balance method or some other method that you specify. Vdb stands for variable declining balance.'''
  def WebService(self, Arg1) -> list:
    '''Underlying function that calls the web service asynchronously, using an HTTP GET request, and returns the response.'''
  def WeekNum(self, Arg1, Arg2) -> float:
    '''Returns a number that indicates where the week falls numerically within a year.'''
  def Weekday(self, Arg1, Arg2) -> float:
    '''Returns the day of the week corresponding to a date. The day is given as an integer, ranging from 1 (Sunday) to 7 (Saturday), by default.'''
  def Weibull(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating a device's mean time to failure.'''
  def Weibull_Dist(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating the mean time to failure for a device.'''
  def WorkDay(self, Arg1, Arg2, Arg3) -> float:
    '''Returns a number that represents a date that is the indicated number of working days before or after a date (the starting date). Working days exclude weekends and any dates identified as holidays. Use WorkDay to exclude weekends or holidays when you calculate invoice due dates, expected delivery times, or the number of days of work performed.'''
  def WorkDay_Intl(self, Arg1, Arg2, Arg3, Arg4) -> float:
    '''Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters. Weekend parameters indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.'''
  def XLookup(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def XMatch(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Xirr(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic. To calculate the internal rate of return for a series of periodic cash flows, use the Irr function.'''
  def Xnpv(self, Arg1, Arg2) -> float:
    '''Returns the net present value for a schedule of cash flows that is not necessarily periodic. Read/write Double.'''
  def Xor(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) -> bool:
    '''Returns a logical exclusive OR of all arguments.'''
  def YearFrac(self, Arg1, Arg2, Arg3) -> float:
    '''Calculates the fraction of the year represented by the number of whole days between two dates (the start_date and the end_date). Use the YearFrac worksheet function to identify the proportion of a whole year's benefits or obligations to assign to a specific term.'''
  def YieldDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5) -> float:
    '''Returns the annual yield for a discounted security.'''
  def YieldMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float:
    '''Returns the annual yield of a security that pays interest at maturity.'''
  def ZTest(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, ZTest returns the probability that the sample mean would be greater than the average of observations in the data set (array); that is, the observed sample mean.'''
  def Z_Test(self, Arg1, Arg2, Arg3) -> float:
    '''Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, Z_Test returns the probability that the sample mean would be greater than the average of observations in the data set (array); that is, the observed sample mean.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _WSFunction(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __len__(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WorksheetFunction'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:424, whats:4,   ok:435, er:0, er:0


# num=38
class Areas:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Areas'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=39
class Borders:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Color: float
    '''Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.'''
    self.ColorIndex: int
    '''Returns or sets a Variant value that represents the color of all four borders.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: CellFormat
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a Border object that represents one of the borders of either a range of cells or a style.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # LineStyle:  <class 'NoneType'>
    # ThemeColor:  <class 'NoneType'>
    # TintAndShade:  <class 'NoneType'>
    # Value:  <class 'NoneType'>
    # Weight:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Borders'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:12, whats:9,   ok:31, er:0, er:0


# num=40
class Characters:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the specified object.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self) -> list:
    '''Deletes the object.'''
  def Insert(self, String) -> list:
    '''Inserts a string preceding the selected characters.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Caption:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Characters 的 Caption 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Count:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Characters 的 Count 属性', 'xlmain11.chm', 0, -2146827284), None)
    # PhoneticCharacters:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Text:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Characters 的 Text 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Characters'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:9, whats:4,   ok:21, er:0, er:4


# num=41
class DisplayFormat:
  def __init__(self):
    self.AddIndent: bool
    '''Returns a value that indicates if Microsoft Excel automatically indents text of the associated Range object when the text alignment in a cell is set to equal distribution (either horizontally or vertically), as it is displayed in the current user interface. Read-only.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Borders: Borders
    '''Returns a Borders object that represents the borders of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Characters: Characters
    '''Returns a Characters object that represents a range of characters within the text of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the associated Range as it is displayed in the current user interface. Read-only.'''
    self.FormulaHidden: bool
    '''Returns a value that indicates if the formula of the associated Range object is hidden when the worksheet is protected as it is displayed in the current user interface. Read-only.'''
    self.HorizontalAlignment: int
    '''Returns a value that represents the horizontal alignment of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.IndentLevel: int
    '''Returns a value that represents the indent level of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Interior: Interior
    '''Returns an Interior object that represents the interior of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Locked: bool
    '''Returns a value that indicates if the associated Range object is locked as it is displayed in the current user interface. Read-only.'''
    self.MergeCells: bool
    '''Returns a value that indicates if the associated Range object contains merged cells as it is displayed in the current user interface. Read-only.'''
    self.NumberFormat: str
    '''Returns a value that represents the format code of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.NumberFormatLocal: str
    '''Returns a value that represents the format code of the associated Range object as a string in the language of the user as it is displayed in the current user interface. Read-only.'''
    self.Orientation: int
    '''Returns a value that represents the text orientation of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.ReadingOrder: int
    '''Returns the reading order of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.ShrinkToFit: bool
    '''Returns a value that indicates if Microsoft Excel automatically shrinks text to fit in the available column width of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.Style: Style
    '''Returns a value, containing a Style object, that represents the style of the associated Range object as it is displayed in the current user interface.'''
    self.VerticalAlignment: int
    '''Returns a value that represents the vertical alignment of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.WrapText: bool
    '''Returns a value that indicates if Microsoft Excel wraps the text of the associated Range object as it is displayed in the current user interface. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def GetCharacters(self, Start, Length):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DisplayFormat'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:25, methods:6, whats:4,   ok:35, er:0, er:0


# num=42
class Errors:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Errors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:10, whats:4,   ok:21, er:0, er:0


# num=43
class Font:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Bold: bool
    '''True if the font is bold. Read/write Variant.'''
    self.Color: float
    '''Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.'''
    self.ColorIndex: int
    '''Returns or sets a Variant value that represents the color of the font.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.FontStyle: str
    '''Returns or sets the font style. Read/write String.'''
    self.Italic: bool
    '''True if the font style is italic. Read/write Boolean.'''
    self.Name: str
    '''Returns or sets a Variant value that represents the name of the object.'''
    self.OutlineFont: bool
    self.Parent: DisplayFormat
    '''Returns the parent object for the specified object. Read-only.'''
    self.Shadow: bool
    self.Size: float
    '''Returns or sets the size of the font. Read/write Variant.'''
    self.Strikethrough: bool
    '''True if the font is struck through with a horizontal line. Read/write Boolean.'''
    self.Subscript: bool
    '''True if the font is formatted as subscript. False by default. Read/write Variant.'''
    self.Superscript: bool
    '''True if the font is formatted as superscript; False by default. Read/write Variant.'''
    self.ThemeColor: int
    '''Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write Variant.'''
    self.ThemeFont: int
    '''Returns or sets the theme font in the applied font scheme that is associated with the specified object. Read/write XlThemeFont.'''
    self.TintAndShade: float
    '''Returns or sets a Single that lightens or darkens a color.'''
    self.Underline: int
    '''Returns or sets the type of underline applied to the font. Read/write Variant.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # Background:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Font'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:23, methods:5, whats:5,   ok:33, er:0, er:0


# num=44
class FormatConditions:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, Operator, Formula1, Formula2, String, TextOperator, DateOperator, ScopeType) -> FormatCondition:
    '''Adds a new conditional format.'''
  def AddAboveAverage(self) -> AboveAverage_object:
    '''Returns a new AboveAverage object representing a conditional formatting rule for the specified range.'''
  def AddColorScale(self, ColorScaleType) -> ColorScale_object:
    '''Returns a new ColorScale object representing a conditional formatting rule that uses gradations in cell colors to indicate relative differences in the values of cells included in a selected range.'''
  def AddDatabar(self) -> Databar_object:
    '''Returns a Databar object representing a data bar conditional formatting rule for the specified range.'''
  def AddIconSetCondition(self) -> IconSetCondition_object:
    '''Returns a new IconSetCondition object that represents an icon set conditional formatting rule for the specified range.'''
  def AddTop10(self) -> Top10_object:
    '''Returns a Top10 object representing a conditional formatting rule for the specified range.'''
  def AddUniqueValues(self) -> UniqueValues_object:
    '''Returns a new UniqueValues object representing a conditional formatting rule for the specified range.'''
  def Delete(self):
    '''Deletes the object.'''
  def Item(self, Index) -> An_Object_value_that_represents_an_object_contained_by_the_collection_:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.FormatConditions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:20, whats:4,   ok:32, er:0, er:0


# num=45
class Hyperlinks:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Anchor, Address, SubAddress, ScreenTip, TextToDisplay) -> Hyperlink:
    '''Adds a hyperlink to the specified range or shape.'''
  def Delete(self):
    '''Deletes the object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Hyperlinks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=46
class Interior:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Color: float
    '''Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.'''
    self.ColorIndex: int
    '''Returns or sets a Variant value that represents the color of the interior.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: DisplayFormat
    '''Returns the parent object for the specified object. Read-only.'''
    self.Pattern: int
    '''Returns or sets a Variant value, containing an XlPattern constant, that represents the interior pattern.'''
    self.PatternColor: int
    '''Returns or sets the color of the interior pattern as an RGB value. Read/write Variant.'''
    self.PatternColorIndex: int
    '''Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.'''
    self.PatternThemeColor: int
    '''Returns or sets a theme color pattern for an Interior object. Read/write Variant.'''
    self.PatternTintAndShade: float
    '''Returns or sets a tint and shade pattern for an Interior object. Read/write Variant.'''
    self.ThemeColor: int
    '''Returns or sets a Variant value, containing an XlThemeColor constant, that represents the color. Read/write Variant.'''
    self.TintAndShade: float
    '''Returns or sets a Single that lightens or darkens a color.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # Gradient:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # InvertIfNegative:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Interior'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:5, whats:5,   ok:26, er:0, er:1


# num=47
class Phonetic:
  def __init__(self):
    self.Alignment: int
    '''Returns or sets a Long value that represents the alignment for the specified phonetic text or tick label.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.CharacterType: int
    '''Returns or sets the type of phonetic text in the specified cell. Read/write XlPhoneticCharacterType.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the specified object.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.Text: str
    '''Returns or sets the text for the specified object. Read/write String.'''
    self.Visible: bool
    '''Returns or sets a Boolean value that determines whether the object is visible. Read/write.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Phonetic'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:12, methods:5, whats:4,   ok:21, er:0, er:0


# num=48
class Phonetics:
  def __init__(self):
    self.Alignment: int
    '''Returns or sets a Long value that represents the alignment for the specified phonetic text or tick label.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.CharacterType: int
    '''Returns or sets the type of phonetic text in the specified cell. Read/write XlPhoneticCharacterType.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the specified object.'''
    self.Length: int
    '''Returns a Long value that represents the number of characters of phonetic text from the position you've specified with the Start property.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.Visible: bool
    '''Returns or sets a Boolean value that determines whether the object is visible. Read/write.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Start, Length, Text):
    '''Adds phonetic text to the specified cell.'''
  def Delete(self):
    '''Deletes the object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Start:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Text:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Phonetics'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:13, methods:14, whats:4,   ok:31, er:0, er:2


# num=49
class SmartTags:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, SmartTagType):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SmartTags'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:11, whats:4,   ok:23, er:0, er:0


# num=50
class SoundNote:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):  pass
  def Import(self, Filename):  pass
  def Play(self):  pass
  def Record(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SoundNote'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:9, whats:4,   ok:20, er:0, er:0


# num=51
class SparklineGroups:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the count of sparkline groups in the associated Range object. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Range
    '''Returns the Range object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, SourceData) -> SparklineGroup:
    '''Creates a new sparkline group and returns a SparklineGroup object.'''
  def Clear(self) -> None:
    '''Clears the selected sparklines.'''
  def ClearGroups(self) -> None:
    '''Clears the selected sparkline groups.'''
  def Group(self, Location) -> None:
    '''Groups the selected sparklines.'''
  def Item(self, Index):
    '''Returns a SparklineGroup object from a collection. Read-only.'''
  def Ungroup(self) -> None:
    '''Ungroups the sparklines in the selected sparkline group.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SparklineGroups'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:17, whats:4,   ok:29, er:0, er:0


# num=52
class Style:
  def __init__(self):
    self.AddIndent: bool
    '''Returns or sets a Boolean value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically).'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Borders: Borders
    '''Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).'''
    self.BuiltIn: bool
    '''True if the style is a built-in style. Read-only Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Font: Font
    '''Returns a Font object that represents the font of the specified object.'''
    self.FormulaHidden: bool
    '''Returns or sets a Boolean value that indicates if the formula will be hidden when the worksheet is protected.'''
    self.HorizontalAlignment: int
    '''Returns or sets an XlHAlign value that represents the horizontal alignment for the specified object.'''
    self.IncludeAlignment: bool
    '''True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties of the Style object. Read/write Boolean.'''
    self.IncludeBorder: bool
    '''True if the style includes the Color, ColorIndex, LineStyle, and Weight properties of the Border object. Read/write Boolean.'''
    self.IncludeFont: bool
    '''True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties. Read/write Boolean.'''
    self.IncludeNumber: bool
    '''True if the style includes the NumberFormat property. Read/write Boolean.'''
    self.IncludePatterns: bool
    '''True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex properties of the Interior object. Read/write Boolean.'''
    self.IncludeProtection: bool
    '''True if the style includes the FormulaHidden and Locked protection properties. Read/write Boolean.'''
    self.Interior: Interior
    '''Returns an Interior object that represents the interior of the specified object.'''
    self.Locked: bool
    '''Returns or sets a Boolean value that indicates if the object is locked.'''
    self.Name: str
    '''Returns a String value that represents the name of the object.'''
    self.NameLocal: str
    '''Returns or sets the name of the object, in the language of the user. Read-only String.'''
    self.NumberFormat: str
    '''Returns or sets a String value that represents the format code for the object.'''
    self.NumberFormatLocal: str
    '''Returns or sets a String value that represents the format code for the object as a string in the language of the user.'''
    self.Orientation: int
    '''Returns or sets an XlOrientation value that represents the text orientation.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.ReadingOrder: int
    '''Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.'''
    self.ShrinkToFit: bool
    '''Returns or sets a Boolean value that indicates if text automatically shrinks to fit in the available column width.'''
    self.Value: str
    '''Returns a String value that represents the name of the specified style.'''
    self.VerticalAlignment: int
    '''Returns or sets an XlVAlign value that represents the vertical alignment of the specified object.'''
    self.WrapText: bool
    '''Returns or sets a Boolean value that indicates if Microsoft Excel wraps the text in the object.'''
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self) -> list:
    '''Deletes the object.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # IndentLevel:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # MergeCells:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Style 的 MergeCells 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Style'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:32, methods:8, whats:5,   ok:45, er:0, er:1


# num=53
class Validation:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.ErrorTitle: str
    '''Returns or sets the title of the data-validation error dialog box. Read/write String.'''
    self.InputTitle: str
    '''Returns or sets the title of the data-validation input dialog box. Read/write String. Limited to 32 characters.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.Value: bool
    '''Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, AlertStyle, Operator, Formula1, Formula2):
    '''Adds data validation to the specified range.'''
  def Delete(self):
    '''Deletes the object.'''
  def Modify(self, Type, AlertStyle, Operator, Formula1, Formula2):
    '''Modifies data validation for a range.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # AlertStyle:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ErrorMessage:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Formula1:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Formula2:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # IMEMode:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # IgnoreBlank:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # InCellDropdown:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # InputMessage:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Operator:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ShowError:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ShowInput:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Type:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Validation'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:10, whats:4,   ok:24, er:0, er:12


# num=54
class Worksheet(_Worksheet):
  def __init__(self):
    self.__dict__: dict
    self.__module__: str
    self._dispobj_: _Worksheet

  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __maybe__call__(self, args, kwargs):  pass
  def __maybe__int__(self, args):  pass
  def __maybe__iter__(self):  pass
  def __maybe__len__(self):  pass
  def __maybe__nonzero__(self):  pass
  def __maybe__str__(self, args):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # coclass_interfaces:  <class 'list'>
    # coclass_sources:  <class 'list'>
    # default_interface:  <class 'type'>
    # default_source:  <class 'type'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Worksheet'>, <class 'win32com.client.CoClassBaseClass'>, <class 'object'>)", attrs:3, methods:8, whats:6,   ok:17, er:0, er:0


# num=55
class XPath:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Range
    '''Returns the parent object for the specified object. Read-only.'''
    self.Value: str
    '''Returns a String that represents the XPath for the specified object.'''
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Clear(self):
    '''Clears all XPath schema information for the mapped range.'''
  def SetValue(self, Map, XPath, SelectionNamespace, Repeating):
    '''Maps the specified XPath object to a ListColumn object or Range collection. If the XPath object has previously been mapped to the ListColumn object or Range collection, the SetValue method sets the properties of the XPath object.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Map:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)
    # Repeating:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XPath'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:9, whats:4,   ok:22, er:0, er:2


# num=56
class Menus:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: MenuBar
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Caption, Before, Restore):  pass
  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Menus'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=57
class _Application:
  def __init__(self):
    self.ActiveCell: Range
    self.ActiveEncryptionSession: int
    self.ActiveMenuBar: MenuBar
    self.ActivePrinter: str
    self.ActiveSheet: _Worksheet
    self.ActiveWindow: Window
    self.ActiveWorkbook: Workbook
    self.AddIns: AddIns
    self.AddIns2: AddIns2
    self.AlertBeforeOverwriting: bool
    self.AltStartupPath: str
    self.AlwaysUseClearType: bool
    self.Application: Application
    self.ArbitraryXMLSupportAvailable: bool
    self.AskToUpdateLinks: bool
    self.Assistance: CDispatch
    self.Assistant: CDispatch
    self.AutoCorrect: AutoCorrect
    self.AutoFormatAsYouTypeReplaceHyperlinks: bool
    self.AutoPercentEntry: bool
    self.AutoRecover: AutoRecover
    self.AutomationSecurity: int
    self.Build: float
    self.COMAddIns: CDispatch
    self.CSVDisplayNumberConversionWarning: bool
    self.CSVKeepColumnAsTextIfMultipleEntriesAreText: bool
    self.CalculateBeforeSave: bool
    self.Calculation: int
    self.CalculationInterruptKey: int
    self.CalculationState: int
    self.CalculationVersion: int
    self.Caller: int
    self.CanPlaySounds: bool
    self.CanRecordSounds: bool
    self.Caption: str
    self.CellDragAndDrop: bool
    self.Cells: Range
    self.ChartDataPointTrack: bool
    self.Charts: Sheets
    self.ClipboardFormats: tuple
    self.ClusterConnector: str
    self.ColorButtons: bool
    self.Columns: Range
    self.CommandBars: CDispatch
    self.CommandUnderlines: int
    self.ConstrainNumeric: bool
    self.ControlCharacters: float
    self.ConvertNumbersWithECharacter: bool
    self.CopyObjectsWithCells: bool
    self.Creator: int
    self.Cursor: int
    self.CursorMovement: float
    self.CustomListCount: float
    self.CutCopyMode: int
    self.DDEAppReturnCode: float
    self.DataEntryMode: int
    self.DataPrivacyOptions: CDispatch
    self.DecimalSeparator: str
    self.DefaultFilePath: str
    self.DefaultPivotTableLayoutOptions: DefaultPivotTableLayoutOptions
    self.DefaultSaveFormat: int
    self.DefaultSheetDirection: int
    self.DefaultWebOptions: DefaultWebOptions
    self.DeferAsyncQueries: bool
    self.DialogSheets: Sheets
    self.Dialogs: Dialogs
    self.DisplayAlerts: bool
    self.DisplayClipboardWindow: bool
    self.DisplayCommentIndicator: int
    self.DisplayDocumentActionTaskPane: bool
    self.DisplayDocumentInformationPanel: bool
    self.DisplayExcel4Menus: bool
    self.DisplayFormulaAutoComplete: bool
    self.DisplayFormulaBar: bool
    self.DisplayFullScreen: bool
    self.DisplayFunctionToolTips: bool
    self.DisplayInfoWindow: bool
    self.DisplayInsertOptions: bool
    self.DisplayNoteIndicator: bool
    self.DisplayPasteOptions: bool
    self.DisplayRecentFiles: bool
    self.DisplayScrollBars: bool
    self.DisplayStatusBar: bool
    self.EditDirectlyInCell: bool
    self.EnableAnimations: bool
    self.EnableAutoComplete: bool
    self.EnableCancelKey: int
    self.EnableCheckFileExtensions: bool
    self.EnableEvents: bool
    self.EnableLargeOperationAlert: bool
    self.EnableLivePreview: bool
    self.EnableMacroAnimations: bool
    self.EnableSound: bool
    self.EnableTipWizard: bool
    self.ErrorCheckingOptions: ErrorCheckingOptions
    self.Excel4IntlMacroSheets: Sheets
    self.Excel4MacroSheets: Sheets
    self.ExtendList: bool
    self.FeatureInstall: int
    self.FileExportConverters: FileExportConverters
    self.FileValidation: int
    self.FileValidationPivot: int
    self.FindFormat: CellFormat
    self.FixedDecimal: bool
    self.FixedDecimalPlaces: int
    self.FlashFill: bool
    self.FlashFillMode: bool
    self.FormulaBarHeight: int
    self.GenerateGetPivotData: bool
    self.GenerateTableRefs: int
    self.Height: float
    self.HighQualityModeForGraphics: bool
    self.HinstancePtr: int
    self.Hwnd: int
    self.IgnoreRemoteRequests: bool
    self.Interactive: bool
    self.International: tuple
    self.IsSandboxed: bool
    self.Iteration: bool
    self.LanguageSettings: CDispatch
    self.LargeButtons: bool
    self.LargeOperationCellThousandCount: int
    self.Left: float
    self.LibraryPath: str
    self.MailSystem: float
    self.MapPaperSize: bool
    self.MathCoprocessorAvailable: bool
    self.MaxChange: float
    self.MaxIterations: float
    self.MeasurementUnit: int
    self.MemoryFree: int
    self.MemoryTotal: int
    self.MemoryUsed: int
    self.MenuBars: MenuBars
    self.MergeInstances: bool
    self.Modules: Modules
    self.MouseAvailable: bool
    self.MoveAfterReturn: bool
    self.MoveAfterReturnDirection: int
    self.MultiThreadedCalculation: MultiThreadedCalculation
    self.Name: str
    self.Names: Names
    self.NetworkTemplatesPath: str
    self.NewWorkbook: CDispatch
    self.ODBCErrors: ODBCErrors
    self.ODBCTimeout: int
    self.OLEDBErrors: OLEDBErrors
    self.OperatingSystem: str
    self.OrganizationName: str
    self.Parent: Application
    self.Path: str
    self.PathSeparator: str
    self.PivotTableSelection: bool
    self.PrintCommunication: bool
    self.ProductCode: str
    self.PromptForSummaryInfo: bool
    self.ProtectedViewWindows: ProtectedViewWindows
    self.QuickAnalysis: QuickAnalysis
    self.Quitting: bool
    self.RTD: RTD
    self.Ready: bool
    self.RecentFiles: RecentFiles
    self.RecordRelative: bool
    self.ReferenceStyle: int
    self.ReplaceFormat: CellFormat
    self.RollZoom: bool
    self.Rows: Range
    self.SaveISO8601Dates: bool
    self.ScreenUpdating: bool
    self.Selection: Range
    self.Sheets: Sheets
    self.SheetsInNewWorkbook: float
    self.ShowChartTipNames: bool
    self.ShowChartTipValues: bool
    self.ShowConvertToDataType: bool
    self.ShowDevTools: bool
    self.ShowMenuFloaties: bool
    self.ShowQuickAnalysis: bool
    self.ShowSelectionFloaties: bool
    self.ShowStartupDialog: bool
    self.ShowToolTips: bool
    self.ShowWindowsInTaskbar: bool
    self.SmartArtColors: CDispatch
    self.SmartArtLayouts: CDispatch
    self.SmartArtQuickStyles: CDispatch
    self.SmartTagRecognizers: SmartTagRecognizers
    self.Speech: Speech
    self.SpellingOptions: SpellingOptions
    self.StandardFont: str
    self.StandardFontSize: float
    self.StartupPath: str
    self.StatusBar: bool
    self.TemplatesPath: str
    self.ThousandsSeparator: str
    self.Toolbars: Toolbars
    self.Top: float
    self.TransitionMenuKey: str
    self.TransitionMenuKeyAction: float
    self.TransitionNavigKeys: bool
    self.TruncateLargeNumbers: bool
    self.TruncateLeadingZeros: bool
    self.UILanguage: int
    self.UsableHeight: float
    self.UsableWidth: float
    self.UseClusterConnector: bool
    self.UseSystemSeparators: bool
    self.UsedObjects: UsedObjects
    self.UserControl: bool
    self.UserLibraryPath: str
    self.UserName: str
    self.Value: str
    self.Version: str
    self.Visible: bool
    self.WarnOnFunctionNameConflict: bool
    self.Watches: Watches
    self.Width: float
    self.WindowState: int
    self.Windows: Windows
    self.WindowsForPens: bool
    self.Workbooks: Workbooks
    self.WorksheetFunction: WorksheetFunction
    self.Worksheets: Sheets
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def ActivateMicrosoftApp(self, Index):  pass
  def AddChartAutoFormat(self, Chart, Name, Description):  pass
  def AddCustomList(self, ListArray, ByRow):  pass
  def Calculate(self):  pass
  def CalculateFull(self):  pass
  def CalculateFullRebuild(self):  pass
  def CalculateUntilAsyncQueriesDone(self):  pass
  def CentimetersToPoints(self, Centimeters):  pass
  def CheckAbort(self, KeepAbort):  pass
  def CheckSpelling(self, Word, CustomDictionary, IgnoreUppercase):  pass
  def ConvertFormula(self, Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo):  pass
  def DDEExecute(self, Channel, String):  pass
  def DDEInitiate(self, App, Topic):  pass
  def DDEPoke(self, Channel, Item, Data):  pass
  def DDERequest(self, Channel, Item):  pass
  def DDETerminate(self, Channel):  pass
  def DeleteChartAutoFormat(self, Name):  pass
  def DeleteCustomList(self, ListNum):  pass
  def DisplayXMLSourcePane(self, XmlMap):  pass
  def DoubleClick(self):  pass
  def Dummy1(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Dummy10(self, arg):  pass
  def Dummy11(self):  pass
  def Dummy12(self, p1, p2):  pass
  def Dummy13(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Dummy14(self):  pass
  def Dummy2(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8):  pass
  def Dummy20(self, grfCompareFunctions):  pass
  def Dummy3(self):  pass
  def Dummy4(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15):  pass
  def Dummy5(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13):  pass
  def Dummy6(self):  pass
  def Dummy7(self):  pass
  def Dummy8(self, Arg1):  pass
  def Dummy9(self):  pass
  def Evaluate(self, Name):  pass
  def ExecuteExcel4Macro(self, String):  pass
  def FileDialog(self, fileDialogType):  pass
  def FindFile(self):  pass
  def GetCaller(self, Index):  pass
  def GetClipboardFormats(self, Index):  pass
  def GetCustomListContents(self, ListNum):  pass
  def GetCustomListNum(self, ListArray):  pass
  def GetFileConverters(self, Index1, Index2):  pass
  def GetInternational(self, Index):  pass
  def GetOpenFilename(self, FileFilter, FilterIndex, Title, ButtonText, MultiSelect):  pass
  def GetPhonetic(self, Text):  pass
  def GetPreviousSelections(self, Index):  pass
  def GetRegisteredFunctions(self, Index1, Index2):  pass
  def GetSaveAsFilename(self, InitialFilename, FileFilter, FilterIndex, Title, ButtonText):  pass
  def Goto(self, Reference, Scroll):  pass
  def Help(self, HelpFile, HelpContextID):  pass
  def InchesToPoints(self, Inches):  pass
  def InputBox(self, Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type):  pass
  def Intersect(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def MacroOptions(self, Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile, ArgumentDescriptions):  pass
  def MailLogoff(self):  pass
  def MailLogon(self, Name, Password, DownloadNewMail):  pass
  def NextLetter(self):  pass
  def OnKey(self, Key, Procedure):  pass
  def OnRepeat(self, Text, Procedure):  pass
  def OnTime(self, EarliestTime, Procedure, LatestTime, Schedule):  pass
  def OnUndo(self, Text, Procedure):  pass
  def Quit(self):  pass
  def Range(self, Cell1, Cell2):  pass
  def RecordMacro(self, BasicCode, XlmCode):  pass
  def RegisterXLL(self, Filename):  pass
  def Repeat(self):  pass
  def ResetTipWizard(self):  pass
  def Run(self, Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Save(self, Filename):  pass
  def SaveWorkspace(self, Filename):  pass
  def SendKeys(self, Keys, Wait):  pass
  def SetDefaultChart(self, FormatName, Gallery):  pass
  def SharePointVersion(self, bstrUrl):  pass
  def ShortcutMenus(self, Index):  pass
  def Support(self, Object, ID, arg):  pass
  def Undo(self):  pass
  def Union(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Volatile(self, Volatile):  pass
  def Wait(self, Time):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Evaluate(self, Name):  pass
  def _FindFile(self):  pass
  def _MacroOptions(self, Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile):  pass
  def _Run2(self, Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def _WSFunction(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def _Wait(self, Time):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # ActiveChart:  <class 'NoneType'>
    # ActiveDialog:  <class 'NoneType'>
    # ActiveProtectedViewWindow:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # FileConverters:  <class 'NoneType'>
    # MailSession:  <class 'NoneType'>
    # OnCalculate:  <class 'NoneType'>
    # OnData:  <class 'NoneType'>
    # OnDoubleClick:  <class 'NoneType'>
    # OnEntry:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # OnWindow:  <class 'NoneType'>
    # PreviousSelections:  <class 'NoneType'>
    # RegisteredFunctions:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'PyIID'>

  #getattr AttributeError:

  #getattr Exception:
    # AnswerWizard:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Dummy101:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Dummy22:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # Dummy23:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FileFind:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FileSearch:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # FormatStaleValues:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467263), None)
    # Hinstance:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147418113), None)
    # SensitivityLabelPolicy:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147220726), None)
    # ThisCell:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ThisWorkbook:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # VBE:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不信任到 Visual Basic Project 的程序连接\n', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Application'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:227, methods:94, whats:18,   ok:339, er:0, er:12


# num=58
class Comments:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index) -> Comment:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Comments'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=59
class CommentsThreaded:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index) -> CommentThreaded:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CommentsThreaded'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=60
class CustomProperties:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Value) -> CustomProperty:
    '''Adds custom property information.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CustomProperties'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=61
class HPageBreaks:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Sheets
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before) -> HPageBreak:
    '''Adds a horizontal page break.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.HPageBreaks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=62
class ListObjects:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName):  pass
  def Item(self, Index):  pass
  def _Add(self, SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ListObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=63
class NamedSheetViewCollection:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name):  pass
  def EnterTemporary(self):  pass
  def Exit(self):  pass
  def GetActive(self):  pass
  def GetItem(self, Name):  pass
  def GetItemAt(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.NamedSheetViewCollection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=64
class Outline:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.AutomaticStyles: bool
    '''True if the outline uses automatic styles. Read/write Boolean.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.SummaryColumn: int
    '''Returns or sets the location of the summary columns in the outline. Read/write XlSummaryColumn.'''
    self.SummaryRow: int
    '''Returns or sets the location of the summary rows in the outline. Read/write XlSummaryRow.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def ShowLevels(self, RowLevels, ColumnLevels) -> list:
    '''Displays the specified number of row and/or column levels of an outline.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Outline'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:6, whats:4,   ok:20, er:0, er:0


# num=65
class PageSetup:
  def __init__(self):
    self.AlignMarginsHeaderFooter: bool
    '''Returns True for Excel to align the header and the footer with the margins set in the page setup options. Read/write Boolean.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.BlackAndWhite: bool
    '''True if elements of the document will be printed in black and white. Read/write Boolean.'''
    self.BottomMargin: float
    '''Returns or sets the size of the bottom margin, in points. Read/write Double.'''
    self.CenterFooter: str
    '''Center aligns the footer information in the PageSetup object. Read/write String.'''
    self.CenterFooterPicture: Graphic
    '''Returns a Graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.'''
    self.CenterHeader: str
    '''Center aligns the header information in the PageSetup object. Read/write String.'''
    self.CenterHeaderPicture: Graphic
    '''Returns a Graphic object that represents the picture for the center section of the header. Used to set attributes about the picture.'''
    self.CenterHorizontally: bool
    '''True if the sheet is centered horizontally on the page when it's printed. Read/write Boolean.'''
    self.CenterVertically: bool
    '''True if the sheet is centered vertically on the page when it's printed. Read/write Boolean.'''
    self.ChartSize: int
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DifferentFirstPageHeaderFooter: bool
    '''True if a different header or footer is used on the first page. Read/write Boolean.'''
    self.Draft: bool
    '''True if the sheet will be printed without graphics. Read/write Boolean.'''
    self.EvenPage: Page
    '''Returns or sets the alignment of text on the even page of a workbook or section.'''
    self.FirstPage: Page
    '''Returns or sets the alignment of text on the first page of a workbook or section.'''
    self.FirstPageNumber: int
    '''Returns or sets the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic (Constants). Read/write Long.'''
    self.FitToPagesTall: int
    '''Returns or sets the number of pages tall that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write Variant.'''
    self.FitToPagesWide: int
    '''Returns or sets the number of pages wide that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write Variant.'''
    self.FooterMargin: float
    '''Returns or sets the distance from the bottom of the page to the footer, in points. Read/write Double.'''
    self.HeaderMargin: float
    '''Returns or sets the distance from the top of the page to the header, in points. Read/write Double.'''
    self.LeftFooter: str
    '''Returns or sets the alignment of text on the left footer of a workbook or section.'''
    self.LeftFooterPicture: Graphic
    '''Returns a Graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture.'''
    self.LeftHeader: str
    '''Returns or sets the alignment of text on the left header of a workbook or section.'''
    self.LeftHeaderPicture: Graphic
    '''Returns a Graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.'''
    self.LeftMargin: float
    '''Returns or sets the size of the left margin, in points. Read/write Double.'''
    self.OddAndEvenPagesHeaderFooter: bool
    '''True if the specified PageSetup object has different headers and footers for odd-numbered and even-numbered pages. Read/write Boolean.'''
    self.Order: float
    '''Returns or sets an XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.'''
    self.Orientation: float
    '''Returns or sets an XlPageOrientation value that represents the portrait or landscape printing mode.'''
    self.Pages: Pages
    '''Returns or sets the count or item number of the pages in the Pages collection.'''
    self.PaperSize: float
    '''Returns or sets the size of the paper. Read/write XlPaperSize.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.PrintArea: str
    '''Returns or sets the range to be printed as a String using A1-style references in the language of the macro. Read/write String.'''
    self.PrintComments: int
    '''Returns or sets the way comments are printed with the sheet. Read/write XlPrintLocation.'''
    self.PrintErrors: int
    '''Sets or returns an XlPrintErrors constant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write.'''
    self.PrintGridlines: bool
    '''True if cell gridlines are printed on the page. Applies only to worksheets. Read/write Boolean.'''
    self.PrintHeadings: bool
    '''True if row and column headings are printed with this page. Applies only to worksheets. Read/write Boolean.'''
    self.PrintNotes: bool
    '''True if cell notes are printed as end notes with the sheet. Applies only to worksheets. Read/write Boolean.'''
    self.PrintQuality: tuple
    '''Returns or sets the print quality. Read/write Variant.'''
    self.PrintTitleColumns: str
    '''Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a String in A1-style notation in the language of the macro. Read/write String.'''
    self.PrintTitleRows: str
    '''Returns or sets the rows that contain the cells to be repeated at the top of each page, as a String in A1-style notation in the language of the macro. Read/write String.'''
    self.RightFooter: str
    '''Returns or sets the distance (in points) between the right edge of the page and the right boundary of the footer. Read/write String.'''
    self.RightFooterPicture: Graphic
    '''Returns a Graphic object that represents the picture for the right section of the footer. Used to set attributes of the picture.'''
    self.RightHeader: str
    '''Returns or sets the right part of the header. Read/write String.'''
    self.RightHeaderPicture: Graphic
    '''Returns a Graphic object that represents the picture for the right section of the header. Used to set attributes about the picture.'''
    self.RightMargin: float
    '''Returns or sets the size of the right margin, in points. Read/write Double.'''
    self.ScaleWithDocHeaderFooter: bool
    '''Returns or sets if the header and footer should be scaled with the document when the size of the document changes. Read/write Boolean.'''
    self.TopMargin: float
    '''Returns or sets the size of the top margin, in points. Read/write Double.'''
    self.Zoom: int
    '''Returns or sets a Variant value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def GetPrintQuality(self, Index):  pass
  def SetPrintQuality(self, Index, arg1):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PageSetup'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:53, methods:7, whats:4,   ok:64, er:0, er:0


# num=66
class _Workbook:
  def __init__(self):
    self.AcceptLabelsInFormulas: bool
    self.AccuracyVersion: int
    self.ActiveSheet: _Worksheet
    self.Application: Application
    self.Author: str
    self.AutoSaveOn: bool
    self.AutoUpdateFrequency: int
    self.BuiltinDocumentProperties: CDispatch
    self.CalculationVersion: int
    self.CaseSensitive: bool
    self.ChangeHistoryDuration: int
    self.ChartDataPointTrack: bool
    self.Charts: Sheets
    self.CheckCompatibility: bool
    self.CodeName: str
    self.Colors: tuple
    self.Comments: str
    self.ConflictResolution: int
    self.Connections: Connections
    self.ConnectionsDisabled: bool
    self.CreateBackup: bool
    self.Creator: int
    self.CustomDocumentProperties: CDispatch
    self.CustomViews: CustomViews
    self.CustomXMLParts: CDispatch
    self.Date1904: bool
    self.DefaultPivotTableStyle: TableStyle
    self.DefaultSlicerStyle: TableStyle
    self.DefaultTableStyle: TableStyle
    self.DefaultTimelineStyle: TableStyle
    self.DialogSheets: Sheets
    self.DisplayDrawingObjects: int
    self.DisplayInkComments: bool
    self.DoNotPromptForConvert: bool
    self.DocumentInspectors: CDispatch
    self.EnableAutoRecover: bool
    self.EncryptionProvider: str
    self.EnvelopeVisible: bool
    self.Excel4IntlMacroSheets: Sheets
    self.Excel4MacroSheets: Sheets
    self.Excel8CompatibilityMode: bool
    self.FileFormat: float
    self.Final: bool
    self.ForceFullCalculation: bool
    self.FullName: str
    self.FullNameURLEncoded: str
    self.HasMailer: bool
    self.HasPassword: bool
    self.HasRoutingSlip: bool
    self.HasVBProject: bool
    self.HighlightChangesOnScreen: bool
    self.IconSets: IconSets
    self.InactiveListBorderVisible: bool
    self.IsAddin: bool
    self.IsInplace: bool
    self.KeepChangeHistory: bool
    self.Keywords: str
    self.ListChangesOnNewSheet: bool
    self.Model: Model
    self.Modules: Sheets
    self.MultiUserEditing: bool
    self.Name: str
    self.Names: Names
    self.Parent: _Application
    self.Password: str
    self.PasswordEncryptionAlgorithm: str
    self.PasswordEncryptionFileProperties: bool
    self.PasswordEncryptionKeyLength: int
    self.PasswordEncryptionProvider: str
    self.Path: str
    self.PersonalViewListSettings: bool
    self.PersonalViewPrintSettings: bool
    self.PivotTables: PivotTables
    self.PrecisionAsDisplayed: bool
    self.ProtectStructure: bool
    self.ProtectWindows: bool
    self.PublishObjects: PublishObjects
    self.Queries: Queries
    self.ReadOnly: bool
    self.ReadOnlyRecommended: bool
    self.RemovePersonalInformation: bool
    self.Research: Research
    self.RevisionNumber: int
    self.Routed: bool
    self.RoutingSlip: RoutingSlip
    self.SaveLinkValues: bool
    self.Saved: bool
    self.Sheets: Sheets
    self.ShowConflictHistory: bool
    self.ShowPivotChartActiveFields: bool
    self.ShowPivotTableFieldList: bool
    self.Signatures: CDispatch
    self.SlicerCaches: SlicerCaches
    self.SmartDocument: CDispatch
    self.Styles: Styles
    self.Subject: str
    self.Sync: CDispatch
    self.TableStyles: TableStyles
    self.TemplateRemoveExtData: bool
    self.Theme: CDispatch
    self.Title: str
    self.UpdateLinks: int
    self.UpdateRemoteReferences: bool
    self.UseWholeCellCriteria: bool
    self.UseWildcards: bool
    self.UserStatus: tuple
    self.VBASigned: bool
    self.WebOptions: WebOptions
    self.Windows: Windows
    self.WorkIdentity: str
    self.Worksheets: Sheets
    self.WritePassword: str
    self.WriteReserved: bool
    self.WriteReservedBy: str
    self.XmlMaps: XmlMaps
    self.XmlNamespaces: XmlNamespaces
    self._CodeName: str
    self._ReadOnlyRecommended: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AcceptAllChanges(self, When, Who, Where):  pass
  def Activate(self):  pass
  def AddToFavorites(self):  pass
  def ApplyTheme(self, Filename):  pass
  def BreakLink(self, Name, Type):  pass
  def CanCheckIn(self):  pass
  def ChangeFileAccess(self, Mode, WritePassword, Notify):  pass
  def ChangeLink(self, Name, NewName, Type):  pass
  def CheckIn(self, SaveChanges, Comments, MakePublic):  pass
  def CheckInWithVersion(self, SaveChanges, Comments, MakePublic, VersionType):  pass
  def Close(self, SaveChanges, Filename, RouteWorkbook):  pass
  def ConvertComments(self):  pass
  def CreateForecastSheet(self, Timeline, Values, ForecastStart, ForecastEnd, ConfInt, Seasonality, DataCompletion, Aggregation, ChartType, ShowStatsTable):  pass
  def DeleteNumberFormat(self, NumberFormat):  pass
  def Dummy16(self):  pass
  def Dummy17(self, calcid):  pass
  def Dummy26(self):  pass
  def Dummy27(self):  pass
  def EnableConnections(self):  pass
  def EndReview(self):  pass
  def ExclusiveAccess(self):  pass
  def ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity):  pass
  def FollowHyperlink(self, Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo):  pass
  def ForwardMailer(self):  pass
  def GetColors(self, Index):  pass
  def GetWorkflowTasks(self):  pass
  def GetWorkflowTemplates(self):  pass
  def HighlightChangesOptions(self, When, Who, Where):  pass
  def LinkInfo(self, Name, LinkInfo, Type, EditionRef):  pass
  def LinkSources(self, Type):  pass
  def LockServerFile(self):  pass
  def LookUpInDocs(self, Filename):  pass
  def MergeWorkbook(self, Filename):  pass
  def NewWindow(self):  pass
  def OpenLinks(self, Name, ReadOnly, Type):  pass
  def PivotCaches(self):  pass
  def PivotTableWizard(self, SourceType, SourceData, TableDestination, TableName, RowGrand, ColumnGrand, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection):  pass
  def Post(self, DestName):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas):  pass
  def PrintPreview(self, EnableChanges):  pass
  def Protect(self, Password, Structure, Windows):  pass
  def ProtectSharing(self, Filename, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword, FileFormat):  pass
  def PublishToDocs(self, Title, DisclosureScope, OverwriteUrl):  pass
  def PublishToPBI(self, PublishType, nameConflict, bstrGroupName):  pass
  def PurgeChangeHistoryNow(self, Days, SharingPassword):  pass
  def RecheckSmartTags(self):  pass
  def RefreshAll(self):  pass
  def RejectAllChanges(self, When, Who, Where):  pass
  def ReloadAs(self, Encoding):  pass
  def RemoveDocumentInformation(self, RemoveDocInfoType):  pass
  def RemoveUser(self, Index):  pass
  def Reply(self):  pass
  def ReplyAll(self):  pass
  def ReplyWithChanges(self, ShowMessage):  pass
  def ResetColors(self):  pass
  def Route(self):  pass
  def RunAutoMacros(self, Which):  pass
  def Save(self):  pass
  def SaveAs(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local, WorkIdentity):  pass
  def SaveAsXMLData(self, Filename, Map):  pass
  def SaveCopyAs(self, Filename):  pass
  def SendFaxOverInternet(self, Recipients, Subject, ShowMessage):  pass
  def SendForReview(self, Recipients, Subject, ShowMessage, IncludeAttachment):  pass
  def SendMail(self, Recipients, Subject, ReturnReceipt):  pass
  def SendMailer(self, FileFormat, Priority):  pass
  def SetColors(self, Index, arg1):  pass
  def SetLinkOnData(self, Name, Procedure):  pass
  def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties):  pass
  def ToggleFormsDesign(self):  pass
  def Unprotect(self, Password):  pass
  def UnprotectSharing(self, SharingPassword):  pass
  def UpdateFromFile(self):  pass
  def UpdateLink(self, Name, Type):  pass
  def WebPagePreview(self):  pass
  def XmlImport(self, Url, ImportMap, Overwrite, Destination):  pass
  def XmlImportXml(self, Data, ImportMap, Overwrite, Destination):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def _PrintOut_(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate):  pass
  def _Protect(self, Password, Structure, Windows):  pass
  def _ProtectSharing(self, Filename, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword):  pass
  def _SaveAs(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local):  pass
  def _SaveAs_(self, Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass
  def sblt(self, s):  pass

  #unknown:
    # ActiveChart:  <class 'NoneType'>
    # ActiveSlicer:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # CommandBars:  <class 'NoneType'>
    # OnSave:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # SharedWorkspace:  <class 'NoneType'>
    # SmartTagOptions:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'PyIID'>

  #getattr AttributeError:

  #getattr Exception:
    # AutoUpdateSaveChanges:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Container:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ContentTypeProperties:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '此文档必须包含内容类型属性。在文档管理系统中，内容类型属性是常见的文件必需属性。', 'xlmain11.chm', 0, -2147216381), None)
    # DocumentLibraryVersions:  (-2147352567, '发生意外。', (0, None, '不存在此文件的任何版本。', None, 0, -2147217328), None)
    # HTMLProject:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '此版本的 Excel 不再支持此方法或属性。', 'xlmain11.chm', 0, -2146827284), None)
    # Mailer:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Permission:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)
    # SensitivityLabel:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147220726), None)
    # ServerPolicy:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147023728), None)
    # ServerViewableItems:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # UserControl:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147352573), None)
    # VBProject:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不信任到 Visual Basic Project 的程序连接\n', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Workbook'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:122, methods:89, whats:12,   ok:223, er:0, er:12


# num=67
class Protection:
  def __init__(self):
    self.AllowDeletingColumns: bool
    '''Returns True if the deletion of columns is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowDeletingRows: bool
    '''Returns True if the deletion of rows is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowEditRanges: AllowEditRanges
    '''Returns an AllowEditRanges object.'''
    self.AllowFiltering: bool
    '''Returns True if the user is allowed to make use of an AutoFilter that was created before the sheet was protected. Read-only Boolean.'''
    self.AllowFormattingCells: bool
    '''Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowFormattingColumns: bool
    '''Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowFormattingRows: bool
    '''Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowInsertingColumns: bool
    '''Returns True if the insertion of columns is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowInsertingHyperlinks: bool
    '''Returns True if the insertion of hyperlinks is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowInsertingRows: bool
    '''Returns True if the insertion of rows is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowSorting: bool
    '''Returns True if the sorting option is allowed on a protected worksheet. Read-only Boolean.'''
    self.AllowUsingPivotTables: bool
    '''Returns True if the user is allowed to manipulate PivotTables on a protected worksheet. Read-only Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Protection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:5, whats:4,   ok:25, er:0, er:0


# num=68
class QueryTables:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Connection, Destination, Sql) -> QueryTable:
    '''Creates a new query table.'''
  def Item(self, Index) -> QueryTable:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.QueryTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=69
class Shapes:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add3DModel(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height) -> Shape:
    '''Creates a 3D model from an existing file. Returns a Shape object that represents the new 3D model.'''
  def AddCallout(self, Type, Left, Top, Width, Height) -> Shape:
    '''Creates a borderless line callout. Returns a Shape object that represents the new callout.'''
  def AddCanvas(self, Left, Top, Width, Height):  pass
  def AddChart(self, XlChartType, Left, Top, Width, Height):  pass
  def AddChart2(self, Style, XlChartType, Left, Top, Width, Height, NewLayout) -> Shape:
    '''Adds a chart to the document. Returns a Shape object that represents a chart and adds it to the specified collection.'''
  def AddConnector(self, Type, BeginX, BeginY, EndX, EndY) -> Shape:
    '''Creates a connector. Returns a Shape object that represents the new connector. When a connector is added, it's not connected to anything. Use the BeginConnect and EndConnect methods to attach the beginning and end of a connector to other shapes in the document.'''
  def AddCurve(self, SafeArrayOfPoints) -> Shape:
    '''Returns a Shape object that represents a Bézier curve on a worksheet.'''
  def AddDiagram(self, Type, Left, Top, Width, Height):  pass
  def AddFormControl(self, Type, Left, Top, Width, Height) -> Shape:
    '''Creates a Microsoft Excel control. Returns a Shape object that represents the new control.'''
  def AddLabel(self, Orientation, Left, Top, Width, Height) -> Shape:
    '''Creates a label. Returns a Shape object that represents the new label.'''
  def AddLine(self, BeginX, BeginY, EndX, EndY) -> Shape:
    '''As it applies to the Shapes object, returns a Shape object that represents the new line on a worksheet.'''
  def AddOLEObject(self, ClassType, Filename, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height) -> Shape:
    '''Creates an OLE object. Returns a Shape object that represents the new OLE object.'''
  def AddPicture(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height) -> Shape:
    '''Creates a picture from an existing file. Returns a Shape object that represents the new picture.'''
  def AddPicture2(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height, Compress) -> Shape:
    '''Creates a picture from an existing file. Returns a Shape object that represents the new picture.'''
  def AddPolyline(self, SafeArrayOfPoints) -> Shape:
    '''Creates an open polyline or a closed polygon drawing. Returns a Shape object that represents the new polyline or polygon.'''
  def AddShape(self, Type, Left, Top, Width, Height) -> Shape:
    '''Returns a Shape object that represents the new AutoShape on a worksheet.'''
  def AddSmartArt(self, Layout, Left, Top, Width, Height) -> Shape:
    '''Creates a new SmartArt graphic with the specified layout.'''
  def AddTextEffect(self, PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top) -> Shape:
    '''Creates a WordArt object. Returns a Shape object that represents the new WordArt object.'''
  def AddTextbox(self, Orientation, Left, Top, Width, Height) -> Shape:
    '''Creates a text box. Returns a Shape object that represents the new text box.'''
  def BuildFreeform(self, EditingType, X1, Y1) -> FreeformBuilder:
    '''Builds a freeform object. Returns a FreeformBuilder object that represents the freeform as it is being built.'''
  def Item(self, Index) -> Shape:
    '''Returns a single object from a collection.'''
  def Range(self, Index):
    '''Returns a ShapeRange object that represents a subset of the shapes in a Shapes collection.'''
  def SelectAll(self):
    '''Selects all the shapes in the specified Shapes collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Shapes'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:34, whats:4,   ok:46, er:0, er:0


# num=70
class Sort:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Header: int
    '''Specifies whether the first row contains header information. Read/write XlYesNoGuess.'''
    self.MatchCase: bool
    '''Set to True to perform a case-sensitive sort, or set to False to perform a non-case-sensitive sort. Read/write.'''
    self.Orientation: int
    '''Specifies the orientation for the sort. Read/write XlSortOrientation.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.SortFields: SortFields
    '''Returns the SortFields object that represents the collection of sort fields associated with the Sort object. Read-only.'''
    self.SortMethod: int
    '''Specifies the sort method for Chinese languages. Read/write XlSortMethod.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Apply(self):
    '''Sorts the range based on the currently applied sort states.'''
  def SetRange(self, Rng):
    '''Sets the range over which the sort occurs.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Rng:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Sort'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:12, methods:7, whats:4,   ok:23, er:0, er:1


# num=71
class Tab:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Color: bool
    '''Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.'''
    self.ColorIndex: int
    '''Returns or sets a Variant value that represents the color of the specified worksheet tab or chart tab.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Worksheet
    '''Returns the parent object for the specified object. Read-only.'''
    self.ThemeColor: int
    '''Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write XlThemeColor.'''
    self.TintAndShade: bool
    '''Returns or sets a Single that lightens or darkens a color.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Tab'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:5, whats:4,   ok:20, er:0, er:0


# num=72
class VPageBreaks:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Sheets
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before) -> VPageBreak:
    '''Adds a vertical page break.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.VPageBreaks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=73
class Pane:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Index: int
    '''Returns a Long value that represents the index number of the object within the collection of similar objects.'''
    self.Parent: Window
    '''Returns the parent object for the specified object. Read-only.'''
    self.ScrollColumn: float
    '''Returns or sets the number of the leftmost column in the pane or window. Read/write Long.'''
    self.ScrollRow: float
    '''Returns or sets the number of the row that appears at the top of the pane or window. Read/write Long.'''
    self.VisibleRange: Range
    '''Returns a Range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self) -> bool:
    '''Activates the pane.'''
  def LargeScroll(self, Down, Up, ToRight, ToLeft) -> list:
    '''Scrolls the contents of the window by pages.'''
  def PointsToScreenPixelsX(self, Points) -> int:
    '''Returns or sets a pixel point on the screen.'''
  def PointsToScreenPixelsY(self, Points) -> int:
    '''Returns or sets the location of the pixel on the screen.'''
  def ScrollIntoView(self, Left, Top, Width, Height, Start):
    '''Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the Start argument).'''
  def SmallScroll(self, Down, Up, ToRight, ToLeft) -> list:
    '''Scrolls the contents of the window by rows or columns.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Pane'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:11, whats:4,   ok:26, er:0, er:0


# num=74
class WorksheetView:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DisplayFormulas: bool
    self.DisplayGridlines: bool
    self.DisplayHeadings: bool
    self.DisplayOutline: bool
    self.DisplayZeros: bool
    self.Parent: Window
    self.Sheet: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WorksheetView'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:13, methods:5, whats:4,   ok:22, er:0, er:0


# num=75
class Panes:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Window
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Panes'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=76
class SheetViews:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the number of objects in the collection. Read-only Long.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Window
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a SheetView object that represents views in a workbook. Read-only.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SheetViews'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=77
class LazyDispatchItem:
  def __init__(self):
    self.__dict__: dict
    self.__module__: str
    self.bIsDispatch: int
    self.bIsSink: int
    self.bWritten: int
    self.defaultDispatchName: str
    self.hidden: int
    self.mapFuncs: dict
    self.propMap: dict
    self.propMapGet: dict
    self.propMapPut: dict
    self.typename: str

  def Build(self, typeinfo, attr, bForUser):  pass
  def CountInOutOptArgs(self, argTuple):  pass
  def MakeDispatchFuncMethod(self, entry, name, bMakeClass):  pass
  def MakeFuncMethod(self, entry, name, bMakeClass):  pass
  def MakeVarArgsFuncMethod(self, entry, name, bMakeClass):  pass
  def _AddFunc_(self, typeinfo, fdesc, bForUser):  pass
  def _AddVar_(self, typeinfo, vardesc, bForUser):  pass
  def _propMapGetCheck_(self, key, item):  pass
  def _propMapPutCheck_(self, key, item):  pass

  #unknown:
    # __weakref__:  <class 'NoneType'>
    # clsid:  <class 'NoneType'>
    # co_class:  <class 'NoneType'>
    # doc:  <class 'NoneType'>
    # python_name:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.client.build.LazyDispatchItem'>, <class 'win32com.client.build.DispatchItem'>, <class 'win32com.client.build.OleItem'>, <class 'object'>)", attrs:12, methods:9, whats:5,   ok:26, er:0, er:0


# num=78
class Graphic:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Filename: str
    '''Returns or sets the URL (on the intranet or the web) or path (local or network) to the location where the specified source object was saved. Read/write String.'''
    self.Parent: PageSetup
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Brightness:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ColorType:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Contrast:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # CropBottom:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # CropLeft:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # CropRight:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # CropTop:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Height:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # LockAspectRatio:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Width:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Graphic'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:10


# num=79
class Page:
  def __init__(self):
    self.CenterFooter: HeaderFooter
    '''Specifies a picture or text to be center-aligned in the page footer.'''
    self.CenterHeader: HeaderFooter
    '''Specifies a picture or text to be center-aligned in the page header.'''
    self.LeftFooter: HeaderFooter
    '''Specifies a picture or text to be left-aligned in the page footer.'''
    self.LeftHeader: HeaderFooter
    '''Specifies a picture or text to be left-aligned in the page header.'''
    self.RightFooter: HeaderFooter
    '''Specifies a picture or text to be right-aligned in the page footer.'''
    self.RightHeader: HeaderFooter
    '''Specifies a picture or text to be right-aligned in the page header.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Page'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0


# num=80
class Pages:
  def __init__(self):
    self.Count: int
    '''Returns the number of objects in the collection. Read-only Long.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a Page object that represents a collection of pages in a workbook. Read-only.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Pages'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:12, whats:4,   ok:21, er:0, er:0


# num=81
class Connections:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the number of objects in the collection. Read-only Long.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Description, ConnectionString, CommandText, lCmdtype) -> WorkbookConnection:
    '''Adds a new connection to the workbook.'''
  def Add2(self, Name, Description, ConnectionString, CommandText, lCmdtype, CreateModelConnection, ImportRelationships):  pass
  def AddFromFile(self, Filename, CreateModelConnection, ImportRelationships) -> WorkbookConnection:
    '''Adds a connection from the specified file.'''
  def Item(self, Index) -> WorkbookConnection:
    '''This method creates a connection item.'''
  def _AddFromFile(self, Filename):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Connections'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:16, whats:4,   ok:28, er:0, er:0


# num=82
class CustomViews:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, ViewName, PrintSettings, RowColSettings) -> CustomView:
    '''Creates a new custom view.'''
  def Item(self, ViewName) -> CustomView:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, ViewName):  pass
  def __call__(self, ViewName):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CustomViews'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=83
class TableStyle:
  def __init__(self):
    self.Application: Application
    self.BuiltIn: bool
    self.Creator: int
    self.Name: str
    self.NameLocal: str
    self.Parent: TableStyles
    self.ShowAsAvailablePivotTableStyle: bool
    self.ShowAsAvailableSlicerStyle: bool
    self.ShowAsAvailableTableStyle: bool
    self.ShowAsAvailableTimelineStyle: bool
    self.TableStyleElements: TableStyleElements
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):  pass
  def Duplicate(self, NewTableStyleName):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyle'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:9, whats:4,   ok:29, er:0, er:0


# num=84
class IconSets:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that specifies the number of icon sets available in the workbook. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):
    '''Returns a single IconSet object from the IconSets collection. Read-only.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.IconSets'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=85
class Model:
  def __init__(self):
    self.Application: Application
    '''Returns an Application object that represents the Microsoft Excel application. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.'''
    self.DataModelConnection: WorkbookConnection
    '''Returns the model WorkbookConnection object from the workbook connections collection that connects to the model.'''
    self.ModelFormatBoolean: ModelFormatBoolean
    '''Returns a ModelFormatBoolean object that represents formatting of type True/False in the data model. Read-only.'''
    self.ModelFormatCurrency: ModelFormatCurrency
    '''Returns a ModelFormatCurrency object that represents formatting of type currency in the data model. Read-only.'''
    self.ModelFormatDate: ModelFormatDate
    '''Returns a ModelFormatDate object that represents formatting of type date in the data model. Read-only.'''
    self.ModelFormatDecimalNumber: ModelFormatDecimalNumber
    '''Returns a ModelFormatDecimalNumber object that represents formatting of type decimal number in the data model. Read-only.'''
    self.ModelFormatGeneral: ModelFormatGeneral
    '''Returns a ModelFormatGeneral object that represents formatting of type general in the data model. Read-only.'''
    self.ModelFormatPercentageNumber: ModelFormatPercentageNumber
    '''Returns a ModelFormatPercentageNumber object that represents formatting of type percentage number in the data model. Read-only.'''
    self.ModelFormatScientificNumber: ModelFormatScientificNumber
    '''Returns a ModelFormatScientificNumber object that represents formatting of type scientific number in the data model. Read-only.'''
    self.ModelFormatWholeNumber: ModelFormatWholeNumber
    '''Returns a ModelFormatWholeNumber object that represents formatting of type whole number in the data model. Read-only.'''
    self.ModelMeasures: ModelMeasures
    '''Returns a ModelMeasures object that represents the collection of model measures in the data model. Read-only.'''
    self.ModelRelationships: ModelRelationships
    '''Returns a ModelRelationships object that represents the collection of relationships between data model tables. Read-only.'''
    self.ModelTables: ModelTables
    '''Returns a ModelTables object that represents a collection of tables inside the data model. Read-only.'''
    self.Name: str
    '''Returns a String value representing the name of the Model object. Read-only.'''
    self.Parent: _Workbook
    '''Returns an Object that represents the parent object of the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AddConnection(self, ConnectionToDataSource) -> WORKBOOKCONNECTION:
    '''Adds a new WorkbookConnection to the model with the same properties as the one supplied as an argument.'''
  def CreateModelWorkbookConnection(self, ModelTable) -> WORKBOOKCONNECTION:
    '''Returns a WorkbookConnection object of type ModelConnection.'''
  def GetModelFormatCurrency(self, Symbol, DecimalPlaces):  pass
  def GetModelFormatDate(self, FormatString):  pass
  def GetModelFormatDecimalNumber(self, UseThousandSeparator, DecimalPlaces):  pass
  def GetModelFormatPercentageNumber(self, UseThousandSeparator, DecimalPlaces):  pass
  def GetModelFormatScientificNumber(self, DecimalPlaces):  pass
  def GetModelFormatWholeNumber(self, UseThousandSeparator):  pass
  def Initialize(self) -> None:
    '''Initializes the Workbook's data model. This is called by default the first time the model is used.'''
  def Refresh(self) -> None:
    '''Refreshes all data sources associated with the model, fully reprocesses the model, and updates all Excel data features associated with the model.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Model'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:20, methods:15, whats:4,   ok:39, er:0, er:0


# num=86
class PivotTables:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, PivotCache, TableDestination, TableName, ReadData, DefaultVersion):  pass
  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PivotTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:10, whats:4,   ok:22, er:0, er:0


# num=87
class PublishObjects:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, SourceType, Filename, Sheet, Source, HtmlType, DivID, Title) -> A_PublishObject_object_that_represents_the_new_item_:
    '''Creates an object that represents an item in a document saved to a webpage. Such objects facilitate subsequent updates to the webpage while automated changes are being made to the document in Microsoft Excel. Returns a PublishObject object.'''
  def Delete(self):
    '''Deletes the object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def Publish(self):
    '''Saves a copy of the item or items in the spreadsheet that have been added to the PublishObjects collection to a webpage.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PublishObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0


# num=88
class Queries:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns an integer that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.FastCombine: bool
    '''True to enable the fast combine feature, as long as the workbook is open. Read/write Boolean.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Formula, Description) -> WorkbookQuery:
    '''Adds a new WorkbookQuery object to the Queries collection.'''
  def Item(self, NameOrIndex) -> WorkbookQuery:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, NameOrIndex):  pass
  def __call__(self, NameOrIndex):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Queries'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0


# num=89
class Research:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def IsResearchService(self, ServiceID):  pass
  def Query(self, ServiceID, QueryString, QueryLanguage, UseSelection, LaunchQuery):  pass
  def SetLanguagePair(self, LanguageFrom, LanguageTo):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Research'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:8, whats:4,   ok:19, er:0, er:0


# num=90
class RoutingSlip:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def GetRecipients(self, Index):  pass
  def Reset(self):  pass
  def SetRecipients(self, Index, arg1):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Delivery:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 Delivery 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Message:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 Message 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Recipients:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 Recipients 属性', 'xlmain11.chm', 0, -2146827284), None)
    # ReturnWhenDone:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 ReturnWhenDone 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Status:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 Status 属性', 'xlmain11.chm', 0, -2146827284), None)
    # Subject:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 Subject 属性', 'xlmain11.chm', 0, -2146827284), None)
    # TrackStatus:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 RoutingSlip 的 TrackStatus 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.RoutingSlip'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:8, whats:4,   ok:19, er:0, er:7


# num=91
class SlicerCaches:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent Workbook object for the collection. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Source, SourceField, Name) -> SlicerCache:
    '''Adds a new SlicerCache object to the collection.'''
  def Add2(self, Source, SourceField, Name, SlicerCacheType):  pass
  def Item(self, Index):
    '''Returns a single SlicerCache object from the collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SlicerCaches'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=92
class Styles:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, BasedOn) -> Style:
    '''Creates a new style and adds it to the list of styles that are available for the current workbook.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def Merge(self, Workbook) -> list:
    '''Merges the styles from another workbook into the Styles collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Styles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=93
class TableStyles:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the number of objects in the collection. Read-only Long.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, TableStyleName) -> TableStyle:
    '''Creates a new TableStyle object and adds it to the collection.'''
  def Item(self, Index) -> TableStyle:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=94
class WebOptions:
  def __init__(self):
    self.AllowPNG: bool
    '''True if Portable Network Graphics (PNG) is allowed as an image format when you save documents as a webpage. False if PNG is not allowed as an output format. The default value is False. Read/write Boolean.'''
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DownloadComponents: bool
    '''True if the necessary Microsoft Office Web components are downloaded when you view the saved document in a web browser, but only if the components are not already installed. False if the components are not downloaded. The default value is False. Read/write Boolean.'''
    self.Encoding: int
    '''Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. The default is the system code page. Read/write MsoEncoding.'''
    self.FolderSuffix: str
    '''Returns the folder suffix that Microsoft Excel uses when you save a document as a webpage, use long file names, and choose to save supporting files in a separate folder (that is, if the UseLongFileNames and OrganizeInFolder properties are set to True). Read-only String.'''
    self.LocationOfComponents: str
    '''Returns or sets the central URL (on the intranet or web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write String.'''
    self.OrganizeInFolder: bool
    '''True if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. False if supporting files are saved in the same folder as the webpage. The default value is True. Read/write Boolean.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.PixelsPerInch: int
    '''Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write Long.'''
    self.RelyOnCSS: bool
    '''True if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your webpage, depending on the value of the OrganizeInFolder property. False if HTML <FONT> tags and cascading style sheets are used. The default value is True. Read/write Boolean.'''
    self.RelyOnVML: bool
    '''True if image files are not generated from drawing objects when you save a document as a webpage. False if images are generated. The default value is False. Read/write Boolean.'''
    self.ScreenSize: int
    '''Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Can be one of the MsoScreenSize constants. The default constant is msoScreenSize800x600. Read/write MsoScreenSize.'''
    self.TargetBrowser: int
    '''Returns or sets an MsoTargetBrowser constant indicating the browser version. Read/write.'''
    self.UseLongFileNames: bool
    '''True if long file names are used when you save the document as a webpage. False if long file names are not used and the DOS file name format (8.3) is used. The default value is True. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def UseDefaultFolderSuffix(self):
    '''Sets the folder suffix for the specified document to the default suffix for the language support that you have selected or installed.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WebOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:19, methods:6, whats:4,   ok:29, er:0, er:0


# num=95
class XmlMaps:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Schema, RootElementName) -> XmlMap:
    '''Adds an XML map to the specified workbook.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XmlMaps'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=96
class XmlNamespaces:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.Value: str
    '''Returns a String value that represents the XML namespaces that have been added to the workbook.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def InstallManifest(self, Path, InstallForAllUsers):
    '''Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XmlNamespaces'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0


# num=97
class AllowEditRanges:
  def __init__(self):
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Title, Range, Password) -> An_AllowEditRange_object_that_represents_the_range_:
    '''Adds a range that can be edited on a protected worksheet. Returns an AllowEditRange object.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AllowEditRanges'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:13, whats:4,   ok:22, er:0, er:0


# num=98
class SortFields:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns the number of objects in the collection. Read-only Long.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Sort
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Key, SortOn, Order, CustomOrder, DataOption) -> SortField:
    '''Creates a new sort field and returns a SortFields object.'''
  def Add2(self, Key, SortOn, Order, CustomOrder, DataOption, SubField) -> SortField:
    '''Creates a new sort field and returns a SortFields object that can optionally sort data types with the SubField defined.'''
  def Clear(self):
    '''Clears all the SortFields objects.'''
  def Item(self, Index):
    '''Returns a SortField object that represents a collection of items that can be sorted in a workbook. Read-only.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SortFields'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0


# num=99
class HeaderFooter:
  def __init__(self):
    self.Picture: Graphic
    self.Text: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.HeaderFooter'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:6, methods:5, whats:4,   ok:15, er:0, er:0


# num=100
class TableStyleElements:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: TableStyle
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyleElements'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=101
class WorkbookConnection:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Description: str
    '''Returns or sets a brief description for a WorkbookConnection object. Read/write String.'''
    self.InModel: bool
    '''Specifies whether the WorkbookConnection has been added to the model. Read-only Boolean.'''
    self.ModelConnection: ModelConnection
    '''Returns an object that contains information for the new model connection type introduced in Excel 2013 to interact with the integrated Data Model. Read-only.'''
    self.Name: str
    '''Returns or sets the name of the WorkbookConnection object. Read/write String.'''
    self.Parent: _Workbook
    '''Returns the parent object for the specified object. Read-only.'''
    self.Ranges: Ranges
    '''Returns the range of objects for the specified WorkbookConnection object. Read-only Ranges.'''
    self.RefreshWithRefreshAll: bool
    '''Determines if the connection should be refreshed when Refresh All is executed. Read/write Boolean.'''
    self.Type: int
    '''Returns the workbook connection type. Read-only XlConnectionType.'''
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):
    '''Deletes a workbook connection.'''
  def Refresh(self):
    '''Refreshes a workbook connection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # DataFeedConnection:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ModelTables:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ODBCConnection:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # OLEDBConnection:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # TextConnection:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # WorksheetDataConnection:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WorkbookConnection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:15, methods:9, whats:4,   ok:28, er:0, er:6


# num=102
class ModelFormatBoolean:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatBoolean'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:5, whats:4,   ok:16, er:0, er:0


# num=103
class ModelFormatCurrency:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DecimalPlaces: int
    '''Specifies the number of decimal places after the dot. Read/write Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.Symbol: str
    '''Specifies the symbol to use to represent the currency. Read/write String.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatCurrency'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0


# num=104
class ModelFormatDate:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.FormatString: str
    '''Specifies the date format, for example, "dd/mm/yy". Read/write String.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatDate'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0


# num=105
class ModelFormatDecimalNumber:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DecimalPlaces: int
    '''Specifies the number of decimal places after the dot. Read/write Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.UseThousandSeparator: bool
    '''Specifies whether to display commas between thousands. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatDecimalNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0


# num=106
class ModelFormatGeneral:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatGeneral'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:5, whats:4,   ok:16, er:0, er:0


# num=107
class ModelFormatPercentageNumber:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DecimalPlaces: int
    '''Specifies the number of decimal places after the dot. Read/write Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.UseThousandSeparator: bool
    '''Specifies whether to display commas between thousands. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatPercentageNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0


# num=108
class ModelFormatScientificNumber:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.DecimalPlaces: int
    '''Specifies the number of decimal places after the dot. Read/write Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatScientificNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0


# num=109
class ModelFormatWholeNumber:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.UseThousandSeparator: bool
    '''Specifies whether to display commas between thousands. Read/write Boolean.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatWholeNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0


# num=110
class ModelMeasures:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns an integer that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, MeasureName, AssociatedTable, Formula, FormatInformation, Description) -> ModelMeasure:
    '''Adds a model measure to the model.'''
  def Item(self, Index) -> ModelMeasure:
    '''Returns a single object from a collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelMeasures'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0


# num=111
class ModelRelationships:
  def __init__(self):
    self.Application: Application
    '''Returns an Application object that represents the Microsoft Excel application. Read-only.'''
    self.Count: int
    '''Returns a Long value that represents the number of ModelRelationship objects in a ModelRelationships object. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns an Object that represents the parent object of the specified ModelRelationships object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, ForeignKeyColumn, PrimaryKeyColumn) -> MODELRELATIONSHIP:
    '''Adds a new relationship to the model.'''
  def DetectRelationships(self, PivotTable) -> None:
    '''Detects model relationships in the specified PivotTable object.'''
  def Item(self, Index) -> ModelRelationship:
    '''Returns a single object from the ModelRelationships object.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelRelationships'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0


# num=112
class ModelTables:
  def __init__(self):
    self.Application: Application
    '''Returns an Application object that represents the Microsoft Excel application. Read-only.'''
    self.Count: int
    '''Returns a Long value that represents the number of ModelTable objects in a ModelTables collection. Read-only.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.'''
    self.Parent: Model
    '''Returns an Object that represents the parent object of the specified ModelTables object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index) -> MODELTABLE:
    '''Returns a single object from the ModelTables collection.'''
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=113
class ModelConnection:
  def __init__(self):
    self.ADOConnection: CDispatch
    self.Application: Application
    self.CalculatedMembers: CalculatedMembers
    self.CommandText: str
    self.CommandType: int
    self.Creator: int
    self.Parent: WorkbookConnection
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelConnection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:5, whats:4,   ok:20, er:0, er:0


# num=114
class Ranges:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: WorkbookConnection
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Ranges'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0


# num=115
class CalculatedMembers:
  def __init__(self):
    self.Application: Application
    '''When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.'''
    self.Count: int
    '''Returns a Long value that represents the number of objects in the collection.'''
    self.Creator: int
    '''Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.'''
    self.Parent: ModelConnection
    '''Returns the parent object for the specified object. Read-only.'''
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Formula, SolveOrder, Type, Dynamic, DisplayFolder, HierarchizeDistinct) -> A_CalculatedMember_object_that_represents_the_new_calculated_field_or_calculated_item_:
    '''Adds a calculated field or calculated item to a PivotTable. Returns a CalculatedMember object.'''
  def AddCalculatedMember(self, Name, Formula, SolveOrder, Type, DisplayFolder, MeasureGroup, ParentHierarchy, ParentMember, NumberFormat) -> CALCULATEDMEMBER:
    '''Adds a calculated field or calculated item to a PivotTable.'''
  def Item(self, Index):
    '''Returns a single object from a collection.'''
  def _Add(self, Name, Formula, SolveOrder, Type):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _Default(self, Index):  pass
  def __call__(self, Index):  pass
  def __getattr__(self, attr):  pass
  def __getitem__(self, key):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #unknown:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CalculatedMembers'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0


