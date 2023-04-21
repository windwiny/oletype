# just stub file, import it , declare app obj, get ide auto type hit

class Application:
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
    self.AutoCorrect: AutoCorrect
    self.AutoFormatAsYouTypeReplaceHyperlinks: bool
    self.AutoPercentEntry: bool
    self.AutoRecover: AutoRecover
    self.AutomationSecurity: int
    self.Build: float
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

  #whats:
    # ActiveChart:  <class 'NoneType'>
    # ActiveDialog:  <class 'NoneType'>
    # ActiveProtectedViewWindow:  <class 'NoneType'>
    # Assistance:  <class 'win32com.client.CDispatch'>
    # Assistant:  <class 'win32com.client.CDispatch'>
    # CLSID:  <class 'PyIID'>
    # COMAddIns:  <class 'win32com.client.CDispatch'>
    # CommandBars:  <class 'win32com.client.CDispatch'>
    # DataPrivacyOptions:  <class 'win32com.client.CDispatch'>
    # FileConverters:  <class 'NoneType'>
    # LanguageSettings:  <class 'win32com.client.CDispatch'>
    # MailSession:  <class 'NoneType'>
    # NewWorkbook:  <class 'win32com.client.CDispatch'>
    # OnCalculate:  <class 'NoneType'>
    # OnData:  <class 'NoneType'>
    # OnDoubleClick:  <class 'NoneType'>
    # OnEntry:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # OnWindow:  <class 'NoneType'>
    # PreviousSelections:  <class 'NoneType'>
    # RegisteredFunctions:  <class 'NoneType'>
    # SmartArtColors:  <class 'win32com.client.CDispatch'>
    # SmartArtLayouts:  <class 'win32com.client.CDispatch'>
    # SmartArtQuickStyles:  <class 'win32com.client.CDispatch'>
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
    # Hinstance:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147418113), None)
    # SensitivityLabelPolicy:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147220726), None)
    # ThisCell:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ThisWorkbook:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # VBE:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不信任到 Visual Basic Project 的程序连接\n', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Application'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:217, methods:94, whats:28,   ok:339, er:0, er:11
# just stub file, import it , declare app obj, get ide auto type hit

class Range:
  def __init__(self):
    self.AddIndent: bool
    self.Address: str
    self.AddressLocal: str
    self.AllowEdit: bool
    self.Application: Application
    self.Areas: Areas
    self.Borders: Borders
    self.Cells: Range
    self.Characters: Characters
    self.Column: int
    self.ColumnWidth: float
    self.Columns: Range
    self.Count: int
    self.CountLarge: int
    self.Creator: int
    self.CurrentRegion: Range
    self.DisplayFormat: DisplayFormat
    self.EntireColumn: Range
    self.EntireRow: Range
    self.Errors: Errors
    self.Font: Font
    self.FormatConditions: FormatConditions
    self.Formula: str
    self.Formula2: str
    self.Formula2Local: str
    self.Formula2R1C1: str
    self.Formula2R1C1Local: str
    self.FormulaArray: str
    self.FormulaHidden: bool
    self.FormulaLocal: str
    self.FormulaR1C1: str
    self.FormulaR1C1Local: str
    self.HasArray: bool
    self.HasFormula: bool
    self.HasRichDataType: bool
    self.HasSpill: bool
    self.Height: float
    self.HorizontalAlignment: int
    self.Hyperlinks: Hyperlinks
    self.ID: str
    self.IndentLevel: int
    self.Interior: Interior
    self.Left: float
    self.LinkedDataTypeState: int
    self.ListHeaderRows: int
    self.Locked: bool
    self.MDX: str
    self.MergeArea: Range
    self.MergeCells: bool
    self.Next: Range
    self.NumberFormat: str
    self.NumberFormatLocal: str
    self.Offset: Range
    self.Orientation: int
    self.Parent: _Worksheet
    self.Phonetic: Phonetic
    self.Phonetics: Phonetics
    self.PrefixCharacter: str
    self.ReadingOrder: int
    self.Resize: Range
    self.Row: int
    self.RowHeight: float
    self.Rows: Range
    self.SavedAsArray: bool
    self.ShrinkToFit: bool
    self.SmartTags: SmartTags
    self.SoundNote: SoundNote
    self.SparklineGroups: SparklineGroups
    self.Style: Style
    self.Text: str
    self.Top: float
    self.UseStandardHeight: bool
    self.UseStandardWidth: bool
    self.Validation: Validation
    self.Value: float
    self.Value2: float
    self.VerticalAlignment: int
    self.Width: float
    self.Worksheet: Worksheet
    self.WrapText: bool
    self.XPath: XPath
    self._Default: float
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self):  pass
  def AddComment(self, Text):  pass
  def AddCommentThreaded(self, Text):  pass
  def AdvancedFilter(self, Action, CriteriaRange, CopyToRange, Unique):  pass
  def AllocateChanges(self):  pass
  def ApplyNames(self, Names, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order, AppendLast):  pass
  def ApplyOutlineStyles(self):  pass
  def AutoComplete(self, String):  pass
  def AutoFill(self, Destination, Type):  pass
  def AutoFilter(self, Field, Criteria1, Operator, Criteria2, VisibleDropDown, SubField):  pass
  def AutoFit(self):  pass
  def AutoFormat(self, Format, Number, Font, Alignment, Border, Pattern, Width):  pass
  def AutoOutline(self):  pass
  def BorderAround(self, LineStyle, Weight, ColorIndex, Color, ThemeColor):  pass
  def Calculate(self):  pass
  def CalculateRowMajorOrder(self):  pass
  def CheckSpelling(self, CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang):  pass
  def Clear(self):  pass
  def ClearComments(self):  pass
  def ClearContents(self):  pass
  def ClearFormats(self):  pass
  def ClearHyperlinks(self):  pass
  def ClearNotes(self):  pass
  def ClearOutline(self):  pass
  def ColumnDifferences(self, Comparison):  pass
  def Consolidate(self, Sources, Function, TopRow, LeftColumn, CreateLinks):  pass
  def ConvertToLinkedDataType(self, ServiceID, LanguageCulture):  pass
  def Copy(self, Destination):  pass
  def CopyFromRecordset(self, Data, MaxRows, MaxColumns):  pass
  def CopyPicture(self, Appearance, Format):  pass
  def CreateNames(self, Top, Left, Bottom, Right):  pass
  def CreatePublisher(self, Edition, Appearance, ContainsPICT, ContainsBIFF, ContainsRTF, ContainsVALU):  pass
  def Cut(self, Destination):  pass
  def DataSeries(self, Rowcol, Type, Date, Step, Stop, Trend):  pass
  def DataTypeToText(self):  pass
  def Delete(self, Shift):  pass
  def DialogBox(self):  pass
  def Dirty(self):  pass
  def DiscardChanges(self):  pass
  def EditionOptions(self, Type, Option, Name, Reference, Appearance, ChartSize, Format):  pass
  def End(self, Direction):  pass
  def ExportAsFixedFormat(self, Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity):  pass
  def FillDown(self):  pass
  def FillLeft(self):  pass
  def FillRight(self):  pass
  def FillUp(self):  pass
  def Find(self, What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat):  pass
  def FindNext(self, After):  pass
  def FindPrevious(self, After):  pass
  def FlashFill(self):  pass
  def FunctionWizard(self):  pass
  def GetAddress(self, RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo):  pass
  def GetAddressLocal(self, RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo):  pass
  def GetCharacters(self, Start, Length):  pass
  def GetOffset(self, RowOffset, ColumnOffset):  pass
  def GetResize(self, RowSize, ColumnSize):  pass
  def GetValue(self, RangeValueDataType):  pass
  def Get_Default(self, RowIndex, ColumnIndex):  pass
  def GoalSeek(self, Goal, ChangingCell):  pass
  def Group(self, Start, End, By, Periods):  pass
  def Insert(self, Shift, CopyOrigin):  pass
  def InsertIndent(self, InsertAmount):  pass
  def Item(self, RowIndex, ColumnIndex):  pass
  def Justify(self):  pass
  def ListNames(self):  pass
  def Merge(self, Across):  pass
  def NavigateArrow(self, TowardPrecedent, ArrowNumber, LinkNumber):  pass
  def NoteText(self, Text, Start, Length):  pass
  def Parse(self, ParseLine, Destination):  pass
  def PasteSpecial(self, Paste, Operation, SkipBlanks, Transpose):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def PrintPreview(self, EnableChanges):  pass
  def Range(self, Cell1, Cell2):  pass
  def RefreshLinkedDataType(self, DomainID):  pass
  def RemoveDuplicates(self, Columns, Header):  pass
  def RemoveSubtotal(self):  pass
  def Replace(self, What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat, FormulaVersion):  pass
  def RowDifferences(self, Comparison):  pass
  def Run(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Select(self):  pass
  def SetCellDataTypeFromCell(self, SourceCell):  pass
  def SetItem(self, RowIndex, ColumnIndex, arg2):  pass
  def SetPhonetic(self):  pass
  def SetValue(self, RangeValueDataType, arg1):  pass
  def Set_Default(self, RowIndex, ColumnIndex, arg2):  pass
  def Show(self):  pass
  def ShowCard(self):  pass
  def ShowDependents(self, Remove):  pass
  def ShowErrors(self):  pass
  def ShowPrecedents(self, Remove):  pass
  def Sort(self, Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3, SubField1):  pass
  def SortSpecial(self, SortMethod, Key1, Order1, Type, Key2, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, DataOption1, DataOption2, DataOption3):  pass
  def Speak(self, SpeakDirection, SpeakFormulas):  pass
  def SpecialCells(self, Type, Value):  pass
  def SubscribeTo(self, Edition, Format):  pass
  def Subtotal(self, GroupBy, Function, TotalList, Replace, PageBreaks, SummaryBelowData):  pass
  def Table(self, RowInput, ColumnInput):  pass
  def TextToColumns(self, Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers):  pass
  def UnMerge(self):  pass
  def Ungroup(self):  pass
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

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MenuBar'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:8, whats:4,   ok:23, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
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
    # Scripts:  <class 'win32com.client.CDispatch'>
    # _AutoFilter:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'PyIID'>

  #getattr AttributeError:

  #getattr Exception:
    # MailEnvelope:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Worksheet'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:59, methods:63, whats:17,   ok:139, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class Window:
  def __init__(self):
    self.ActiveCell: Range
    self.ActivePane: Pane
    self.ActiveSheet: _Worksheet
    self.ActiveSheetView: WorksheetView
    self.Application: Application
    self.AutoFilterDateGrouping: bool
    self.Caption: str
    self.Creator: int
    self.DisplayFormulas: bool
    self.DisplayGridlines: bool
    self.DisplayHeadings: bool
    self.DisplayHorizontalScrollBar: bool
    self.DisplayOutline: bool
    self.DisplayRightToLeft: bool
    self.DisplayRuler: bool
    self.DisplayVerticalScrollBar: bool
    self.DisplayWhitespace: bool
    self.DisplayWorkbookTabs: bool
    self.DisplayZeros: bool
    self.EnableResize: bool
    self.FreezePanes: bool
    self.GridlineColor: float
    self.GridlineColorIndex: int
    self.Height: float
    self.Hwnd: int
    self.Index: int
    self.Left: float
    self.Panes: Panes
    self.Parent: _Workbook
    self.RangeSelection: Range
    self.ScrollColumn: float
    self.ScrollRow: float
    self.SelectedSheets: Sheets
    self.Selection: Range
    self.SheetViews: SheetViews
    self.Split: bool
    self.SplitColumn: int
    self.SplitHorizontal: int
    self.SplitRow: int
    self.SplitVertical: int
    self.TabRatio: float
    self.Top: float
    self.Type: int
    self.UsableHeight: float
    self.UsableWidth: float
    self.View: int
    self.Visible: bool
    self.VisibleRange: Range
    self.Width: float
    self.WindowNumber: float
    self.WindowState: int
    self.Zoom: float
    self._DisplayRightToLeft: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self):  pass
  def ActivateNext(self):  pass
  def ActivatePrevious(self):  pass
  def Close(self, SaveChanges, Filename, RouteWorkbook):  pass
  def LargeScroll(self, Down, Up, ToRight, ToLeft):  pass
  def NewWindow(self):  pass
  def PointsToScreenPixelsX(self, Points):  pass
  def PointsToScreenPixelsY(self, Points):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def PrintPreview(self, EnableChanges):  pass
  def RangeFromPoint(self, x, y):  pass
  def ScrollIntoView(self, Left, Top, Width, Height, Start):  pass
  def ScrollWorkbookTabs(self, Sheets, Position):  pass
  def SmallScroll(self, Down, Up, ToRight, ToLeft):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # ActiveChart:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # OnWindow:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Window 的 OnWindow 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Window'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:57, methods:20, whats:5,   ok:82, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class Workbook:
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # coclass_interfaces:  <class 'list'>
    # coclass_sources:  <class 'list'>
    # default_interface:  <class 'type'>
    # default_source:  <class 'type'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Workbook'>, <class 'win32com.client.CoClassBaseClass'>, <class 'object'>)", attrs:3, methods:8, whats:6,   ok:17, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class AddIns:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Filename, CopyFile):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AddIns'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class AddIns2:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Filename, CopyFile):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AddIns2'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class AutoCorrect:
  def __init__(self):
    self.Application: Application
    self.AutoExpandListRange: bool
    self.AutoFillFormulasInLists: bool
    self.CapitalizeNamesOfDays: bool
    self.CorrectCapsLock: bool
    self.CorrectSentenceCap: bool
    self.Creator: int
    self.DisplayAutoCorrectOptions: bool
    self.Parent: _Application
    self.ReplaceText: bool
    self.ReplacementList: tuple
    self.TwoInitialCapitals: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AddReplacement(self, What, Replacement):  pass
  def DeleteReplacement(self, What):  pass
  def GetReplacementList(self, Index):  pass
  def SetReplacementList(self, Index, arg1):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class AutoRecover:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Enabled: bool
    self.Parent: _Application
    self.Path: str
    self.Time: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AutoRecover'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Sheets:
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

  def Add(self, Before, After, Count, Type):  pass
  def Add2(self, Before, After, Count, NewLayout):  pass
  def Copy(self, Before, After):  pass
  def Delete(self):  pass
  def FillAcrossSheets(self, Range, Type):  pass
  def Item(self, Index):  pass
  def Move(self, Before, After):  pass
  def PrintOut(self, From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas):  pass
  def PrintPreview(self, EnableChanges):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Visible:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Sheets 的 Visible 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Sheets'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:23, whats:4,   ok:37, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DefaultPivotTableLayoutOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:50, methods:5, whats:4,   ok:59, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class DefaultWebOptions:
  def __init__(self):
    self.AllowPNG: bool
    self.AlwaysSaveInDefaultEncoding: bool
    self.Application: Application
    self.CheckIfOfficeIsHTMLEditor: bool
    self.Creator: int
    self.DownloadComponents: bool
    self.Encoding: int
    self.FolderSuffix: str
    self.LoadPictures: bool
    self.LocationOfComponents: str
    self.OrganizeInFolder: bool
    self.Parent: _Application
    self.PixelsPerInch: int
    self.RelyOnCSS: bool
    self.RelyOnVML: bool
    self.SaveHiddenData: bool
    self.SaveNewWebPagesAsWebArchives: bool
    self.ScreenSize: int
    self.TargetBrowser: int
    self.UpdateLinksOnSave: bool
    self.UseLongFileNames: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # Fonts:  <class 'win32com.client.CDispatch'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DefaultWebOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:25, methods:5, whats:5,   ok:35, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Dialogs:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Dialogs'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ErrorCheckingOptions:
  def __init__(self):
    self.Application: Application
    self.BackgroundChecking: bool
    self.Creator: int
    self.EmptyCellReferences: bool
    self.EvaluateToError: bool
    self.InconsistentFormula: bool
    self.InconsistentTableFormula: bool
    self.IndicatorColorIndex: int
    self.ListDataValidation: bool
    self.MisleadingNumberFormats: bool
    self.NumberAsText: bool
    self.OmittedCells: bool
    self.OutdatedLinkedDataType: bool
    self.Parent: _Application
    self.TextDate: bool
    self.UnlockedFormulaCells: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ErrorCheckingOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:20, methods:5, whats:4,   ok:29, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class FileExportConverters:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.FileExportConverters'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MenuBars'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Visible:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Sheets 的 Visible 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Modules'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:21, whats:4,   ok:35, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class MultiThreadedCalculation:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Enabled: bool
    self.Parent: _Application
    self.ThreadCount: int
    self.ThreadMode: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.MultiThreadedCalculation'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Names:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, RefersTo, Visible, MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal, RefersToR1C1, RefersToR1C1Local):  pass
  def Item(self, Index, IndexLocal, RefersTo):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Names'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ODBCErrors:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ODBCErrors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class OLEDBErrors:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.OLEDBErrors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ProtectedViewWindows:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Item(self, Index):  pass
  def Open(self, Filename, Password, AddToMru, RepairMode):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ProtectedViewWindows'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class QuickAnalysis:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Hide(self, XlQuickAnalysisMode):  pass
  def Show(self, XlQuickAnalysisMode):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.QuickAnalysis'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:7, whats:4,   ok:18, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class RTD:
  def __init__(self):
    self.ThrottleInterval: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def RefreshData(self):  pass
  def RestartServers(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.RTD'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:7, whats:4,   ok:16, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class RecentFiles:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Maximum: int
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.RecentFiles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SmartTagRecognizers'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:12, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Speech:
  def __init__(self):
    self.Direction: int
    self.SpeakCellOnEnter: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Speak(self, Text, SpeakAsync, SpeakXML, Purge):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Speech'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:6, methods:6, whats:4,   ok:16, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class SpellingOptions:
  def __init__(self):
    self.ArabicModes: int
    self.ArabicStrictAlefHamza: bool
    self.ArabicStrictFinalYaa: bool
    self.ArabicStrictTaaMarboota: bool
    self.BrazilReform: int
    self.DictLang: int
    self.GermanPostReform: bool
    self.HebrewModes: int
    self.IgnoreCaps: bool
    self.IgnoreFileNames: bool
    self.IgnoreMixedDigits: bool
    self.KoreanCombineAux: bool
    self.KoreanProcessCompound: bool
    self.KoreanUseAutoChangeList: bool
    self.PortugalReform: int
    self.RussianStrictE: bool
    self.SpanishModes: int
    self.SuggestMainOnly: bool
    self.UserDict: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SpellingOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:23, methods:5, whats:4,   ok:32, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Toolbars'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class UsedObjects:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.UsedObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Watches:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Source):  pass
  def Delete(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Watches'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Windows:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.SyncScrollingSideBySide: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Arrange(self, ArrangeStyle, ActiveWorkbook, SyncHorizontal, SyncVertical):  pass
  def BreakSideBySide(self):  pass
  def CompareSideBySideWith(self, WindowName):  pass
  def Item(self, Index):  pass
  def ResetPositionsSideBySide(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Windows'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:16, whats:4,   ok:29, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Workbooks:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Template):  pass
  def CanCheckOut(self, Filename):  pass
  def CheckOut(self, Filename):  pass
  def Close(self):  pass
  def Item(self, Index):  pass
  def Open(self, Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad):  pass
  def OpenDatabase(self, Filename, CommandText, CommandType, BackgroundQuery, ImportDataAs):  pass
  def OpenText(self, Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local):  pass
  def OpenXML(self, Filename, Stylesheets, LoadOption):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Workbooks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:24, whats:4,   ok:36, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class WorksheetFunction:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: _Application
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AccrInt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def AccrIntM(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Acos(self, Arg1):  pass
  def Acosh(self, Arg1):  pass
  def Acot(self, Arg1):  pass
  def Acoth(self, Arg1):  pass
  def Aggregate(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def AmorDegrc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def AmorLinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def And(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Arabic(self, Arg1):  pass
  def ArrayToText(self, Arg1, Arg2):  pass
  def Asc(self, Arg1):  pass
  def Asin(self, Arg1):  pass
  def Asinh(self, Arg1):  pass
  def Atan2(self, Arg1, Arg2):  pass
  def Atanh(self, Arg1):  pass
  def AveDev(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Average(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def AverageIf(self, Arg1, Arg2, Arg3):  pass
  def AverageIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def BahtText(self, Arg1):  pass
  def Base(self, Arg1, Arg2, Arg3):  pass
  def BesselI(self, Arg1, Arg2):  pass
  def BesselJ(self, Arg1, Arg2):  pass
  def BesselK(self, Arg1, Arg2):  pass
  def BesselY(self, Arg1, Arg2):  pass
  def BetaDist(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def BetaInv(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Beta_Dist(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Beta_Inv(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Bin2Dec(self, Arg1):  pass
  def Bin2Hex(self, Arg1, Arg2):  pass
  def Bin2Oct(self, Arg1, Arg2):  pass
  def BinomDist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Binom_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Binom_Dist_Range(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Binom_Inv(self, Arg1, Arg2, Arg3):  pass
  def Bitand(self, Arg1, Arg2):  pass
  def Bitlshift(self, Arg1, Arg2):  pass
  def Bitor(self, Arg1, Arg2):  pass
  def Bitrshift(self, Arg1, Arg2):  pass
  def Bitxor(self, Arg1, Arg2):  pass
  def Ceiling(self, Arg1, Arg2):  pass
  def Ceiling_Math(self, Arg1, Arg2, Arg3):  pass
  def Ceiling_Precise(self, Arg1, Arg2):  pass
  def ChiDist(self, Arg1, Arg2):  pass
  def ChiInv(self, Arg1, Arg2):  pass
  def ChiSq_Dist(self, Arg1, Arg2, Arg3):  pass
  def ChiSq_Dist_RT(self, Arg1, Arg2):  pass
  def ChiSq_Inv(self, Arg1, Arg2):  pass
  def ChiSq_Inv_RT(self, Arg1, Arg2):  pass
  def ChiSq_Test(self, Arg1, Arg2):  pass
  def ChiTest(self, Arg1, Arg2):  pass
  def Choose(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Clean(self, Arg1):  pass
  def Combin(self, Arg1, Arg2):  pass
  def Combina(self, Arg1, Arg2):  pass
  def Complex(self, Arg1, Arg2, Arg3):  pass
  def Concat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Confidence(self, Arg1, Arg2, Arg3):  pass
  def Confidence_Norm(self, Arg1, Arg2, Arg3):  pass
  def Confidence_T(self, Arg1, Arg2, Arg3):  pass
  def Convert(self, Arg1, Arg2, Arg3):  pass
  def Correl(self, Arg1, Arg2):  pass
  def Cosh(self, Arg1):  pass
  def Cot(self, Arg1):  pass
  def Coth(self, Arg1):  pass
  def Count(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def CountA(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def CountBlank(self, Arg1):  pass
  def CountIf(self, Arg1, Arg2):  pass
  def CountIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def CoupDayBs(self, Arg1, Arg2, Arg3, Arg4):  pass
  def CoupDays(self, Arg1, Arg2, Arg3, Arg4):  pass
  def CoupDaysNc(self, Arg1, Arg2, Arg3, Arg4):  pass
  def CoupNcd(self, Arg1, Arg2, Arg3, Arg4):  pass
  def CoupNum(self, Arg1, Arg2, Arg3, Arg4):  pass
  def CoupPcd(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Covar(self, Arg1, Arg2):  pass
  def Covariance_P(self, Arg1, Arg2):  pass
  def Covariance_S(self, Arg1, Arg2):  pass
  def CritBinom(self, Arg1, Arg2, Arg3):  pass
  def Csc(self, Arg1):  pass
  def Csch(self, Arg1):  pass
  def CumIPmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def CumPrinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def DAverage(self, Arg1, Arg2, Arg3):  pass
  def DCount(self, Arg1, Arg2, Arg3):  pass
  def DCountA(self, Arg1, Arg2, Arg3):  pass
  def DGet(self, Arg1, Arg2, Arg3):  pass
  def DMax(self, Arg1, Arg2, Arg3):  pass
  def DMin(self, Arg1, Arg2, Arg3):  pass
  def DProduct(self, Arg1, Arg2, Arg3):  pass
  def DStDev(self, Arg1, Arg2, Arg3):  pass
  def DStDevP(self, Arg1, Arg2, Arg3):  pass
  def DSum(self, Arg1, Arg2, Arg3):  pass
  def DVar(self, Arg1, Arg2, Arg3):  pass
  def DVarP(self, Arg1, Arg2, Arg3):  pass
  def Days(self, Arg1, Arg2):  pass
  def Days360(self, Arg1, Arg2, Arg3):  pass
  def Db(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Dbcs(self, Arg1):  pass
  def Ddb(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Dec2Bin(self, Arg1, Arg2):  pass
  def Dec2Hex(self, Arg1, Arg2):  pass
  def Dec2Oct(self, Arg1, Arg2):  pass
  def Decimal(self, Arg1, Arg2):  pass
  def Degrees(self, Arg1):  pass
  def Delta(self, Arg1, Arg2):  pass
  def DevSq(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Disc(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Dollar(self, Arg1, Arg2):  pass
  def DollarDe(self, Arg1, Arg2):  pass
  def DollarFr(self, Arg1, Arg2):  pass
  def Dummy19(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Dummy21(self, Arg1, Arg2):  pass
  def Duration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def EDate(self, Arg1, Arg2):  pass
  def Effect(self, Arg1, Arg2):  pass
  def EncodeURL(self, Arg1):  pass
  def EoMonth(self, Arg1, Arg2):  pass
  def Erf(self, Arg1, Arg2):  pass
  def ErfC(self, Arg1):  pass
  def ErfC_Precise(self, Arg1):  pass
  def Erf_Precise(self, Arg1):  pass
  def Even(self, Arg1):  pass
  def ExponDist(self, Arg1, Arg2, Arg3):  pass
  def Expon_Dist(self, Arg1, Arg2, Arg3):  pass
  def FDist(self, Arg1, Arg2, Arg3):  pass
  def FInv(self, Arg1, Arg2, Arg3):  pass
  def FTest(self, Arg1, Arg2):  pass
  def FVSchedule(self, Arg1, Arg2):  pass
  def F_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def F_Dist_RT(self, Arg1, Arg2, Arg3):  pass
  def F_Inv(self, Arg1, Arg2, Arg3):  pass
  def F_Inv_RT(self, Arg1, Arg2, Arg3):  pass
  def F_Test(self, Arg1, Arg2):  pass
  def Fact(self, Arg1):  pass
  def FactDouble(self, Arg1):  pass
  def FieldValue(self, Arg1, Arg2):  pass
  def Filter(self, Arg1, Arg2, Arg3):  pass
  def FilterXML(self, Arg1, Arg2):  pass
  def Find(self, Arg1, Arg2, Arg3):  pass
  def FindB(self, Arg1, Arg2, Arg3):  pass
  def Fisher(self, Arg1):  pass
  def FisherInv(self, Arg1):  pass
  def Fixed(self, Arg1, Arg2, Arg3):  pass
  def Floor(self, Arg1, Arg2):  pass
  def Floor_Math(self, Arg1, Arg2, Arg3):  pass
  def Floor_Precise(self, Arg1, Arg2):  pass
  def Forecast(self, Arg1, Arg2, Arg3):  pass
  def Forecast_ETS(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Forecast_ETS_ConfInt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def Forecast_ETS_STAT(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Forecast_ETS_Seasonality(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Forecast_Linear(self, Arg1, Arg2, Arg3):  pass
  def Frequency(self, Arg1, Arg2):  pass
  def Fv(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Gamma(self, Arg1):  pass
  def GammaDist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def GammaInv(self, Arg1, Arg2, Arg3):  pass
  def GammaLn(self, Arg1):  pass
  def GammaLn_Precise(self, Arg1):  pass
  def Gamma_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Gamma_Inv(self, Arg1, Arg2, Arg3):  pass
  def Gauss(self, Arg1):  pass
  def Gcd(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def GeStep(self, Arg1, Arg2):  pass
  def GeoMean(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Growth(self, Arg1, Arg2, Arg3, Arg4):  pass
  def HLookup(self, Arg1, Arg2, Arg3, Arg4):  pass
  def HarMean(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Hex2Bin(self, Arg1, Arg2):  pass
  def Hex2Dec(self, Arg1):  pass
  def Hex2Oct(self, Arg1, Arg2):  pass
  def HypGeomDist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def HypGeom_Dist(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def ISO_Ceiling(self, Arg1, Arg2):  pass
  def IfError(self, Arg1, Arg2):  pass
  def IfNa(self, Arg1, Arg2):  pass
  def ImAbs(self, Arg1):  pass
  def ImArgument(self, Arg1):  pass
  def ImConjugate(self, Arg1):  pass
  def ImCos(self, Arg1):  pass
  def ImCosh(self, Arg1):  pass
  def ImCot(self, Arg1):  pass
  def ImCsc(self, Arg1):  pass
  def ImCsch(self, Arg1):  pass
  def ImDiv(self, Arg1, Arg2):  pass
  def ImExp(self, Arg1):  pass
  def ImLn(self, Arg1):  pass
  def ImLog10(self, Arg1):  pass
  def ImLog2(self, Arg1):  pass
  def ImPower(self, Arg1, Arg2):  pass
  def ImProduct(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def ImReal(self, Arg1):  pass
  def ImSec(self, Arg1):  pass
  def ImSech(self, Arg1):  pass
  def ImSin(self, Arg1):  pass
  def ImSinh(self, Arg1):  pass
  def ImSqrt(self, Arg1):  pass
  def ImSub(self, Arg1, Arg2):  pass
  def ImSum(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def ImTan(self, Arg1):  pass
  def Imaginary(self, Arg1):  pass
  def Index(self, Arg1, Arg2, Arg3, Arg4):  pass
  def IntRate(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Intercept(self, Arg1, Arg2):  pass
  def Ipmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Irr(self, Arg1, Arg2):  pass
  def IsErr(self, Arg1):  pass
  def IsError(self, Arg1):  pass
  def IsEven(self, Arg1):  pass
  def IsFormula(self, Arg1):  pass
  def IsLogical(self, Arg1):  pass
  def IsNA(self, Arg1):  pass
  def IsNonText(self, Arg1):  pass
  def IsNumber(self, Arg1):  pass
  def IsOdd(self, Arg1):  pass
  def IsText(self, Arg1):  pass
  def IsThaiDigit(self, Arg1):  pass
  def IsoWeekNum(self, Arg1, Arg2):  pass
  def Ispmt(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Kurt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Large(self, Arg1, Arg2):  pass
  def Lcm(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def LinEst(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Ln(self, Arg1):  pass
  def Log(self, Arg1, Arg2):  pass
  def Log10(self, Arg1):  pass
  def LogEst(self, Arg1, Arg2, Arg3, Arg4):  pass
  def LogInv(self, Arg1, Arg2, Arg3):  pass
  def LogNormDist(self, Arg1, Arg2, Arg3):  pass
  def LogNorm_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def LogNorm_Inv(self, Arg1, Arg2, Arg3):  pass
  def Lookup(self, Arg1, Arg2, Arg3):  pass
  def MDeterm(self, Arg1):  pass
  def MDuration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def MInverse(self, Arg1):  pass
  def MIrr(self, Arg1, Arg2, Arg3):  pass
  def MMult(self, Arg1, Arg2):  pass
  def MRound(self, Arg1, Arg2):  pass
  def Match(self, Arg1, Arg2, Arg3):  pass
  def Max(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def MaxIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Median(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Min(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def MinIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Mode(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Mode_Mult(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Mode_Sngl(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def MultiNomial(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Munit(self, Arg1):  pass
  def NPer(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def NegBinomDist(self, Arg1, Arg2, Arg3):  pass
  def NegBinom_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def NetworkDays(self, Arg1, Arg2, Arg3):  pass
  def NetworkDays_Intl(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Nominal(self, Arg1, Arg2):  pass
  def NormDist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def NormInv(self, Arg1, Arg2, Arg3):  pass
  def NormSDist(self, Arg1):  pass
  def NormSInv(self, Arg1):  pass
  def Norm_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Norm_Inv(self, Arg1, Arg2, Arg3):  pass
  def Norm_S_Dist(self, Arg1, Arg2):  pass
  def Norm_S_Inv(self, Arg1):  pass
  def Npv(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def NumberValue(self, Arg1, Arg2, Arg3):  pass
  def Oct2Bin(self, Arg1, Arg2):  pass
  def Oct2Dec(self, Arg1):  pass
  def Oct2Hex(self, Arg1, Arg2):  pass
  def Odd(self, Arg1):  pass
  def OddFPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9):  pass
  def OddFYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9):  pass
  def OddLPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8):  pass
  def OddLYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8):  pass
  def Or(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def PDuration(self, Arg1, Arg2, Arg3):  pass
  def Pearson(self, Arg1, Arg2):  pass
  def PercentRank(self, Arg1, Arg2, Arg3):  pass
  def PercentRank_Exc(self, Arg1, Arg2, Arg3):  pass
  def PercentRank_Inc(self, Arg1, Arg2, Arg3):  pass
  def Percentile(self, Arg1, Arg2):  pass
  def Percentile_Exc(self, Arg1, Arg2):  pass
  def Percentile_Inc(self, Arg1, Arg2):  pass
  def Permut(self, Arg1, Arg2):  pass
  def Permutationa(self, Arg1, Arg2):  pass
  def Phi(self, Arg1):  pass
  def Phonetic(self, Arg1):  pass
  def Pi(self):  pass
  def Pmt(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Poisson(self, Arg1, Arg2, Arg3):  pass
  def Poisson_Dist(self, Arg1, Arg2, Arg3):  pass
  def Power(self, Arg1, Arg2):  pass
  def Ppmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Price(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def PriceDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def PriceMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Prob(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Product(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Proper(self, Arg1):  pass
  def Pv(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Quartile(self, Arg1, Arg2):  pass
  def Quartile_Exc(self, Arg1, Arg2):  pass
  def Quartile_Inc(self, Arg1, Arg2):  pass
  def Quotient(self, Arg1, Arg2):  pass
  def RSq(self, Arg1, Arg2):  pass
  def RTD(self, progID, server, topic1, topic2, topic3, topic4, topic5, topic6, topic7, topic8, topic9, topic10, topic11, topic12, topic13, topic14, topic15, topic16, topic17, topic18, topic19, topic20, topic21, topic22, topic23, topic24, topic25, topic26, topic27, topic28):  pass
  def Radians(self, Arg1):  pass
  def RandArray(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def RandBetween(self, Arg1, Arg2):  pass
  def Rank(self, Arg1, Arg2, Arg3):  pass
  def Rank_Avg(self, Arg1, Arg2, Arg3):  pass
  def Rank_Eq(self, Arg1, Arg2, Arg3):  pass
  def Rate(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def Received(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def Replace(self, Arg1, Arg2, Arg3, Arg4):  pass
  def ReplaceB(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Rept(self, Arg1, Arg2):  pass
  def Roman(self, Arg1, Arg2):  pass
  def Round(self, Arg1, Arg2):  pass
  def RoundBahtDown(self, Arg1):  pass
  def RoundBahtUp(self, Arg1):  pass
  def RoundDown(self, Arg1, Arg2):  pass
  def RoundUp(self, Arg1, Arg2):  pass
  def Rri(self, Arg1, Arg2, Arg3):  pass
  def Search(self, Arg1, Arg2, Arg3):  pass
  def SearchB(self, Arg1, Arg2, Arg3):  pass
  def Sec(self, Arg1):  pass
  def Sech(self, Arg1):  pass
  def Sequence(self, Arg1, Arg2, Arg3, Arg4):  pass
  def SeriesSum(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Single(self, Arg1):  pass
  def Sinh(self, Arg1):  pass
  def Skew(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Skew_p(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Sln(self, Arg1, Arg2, Arg3):  pass
  def Slope(self, Arg1, Arg2):  pass
  def Small(self, Arg1, Arg2):  pass
  def Sort(self, Arg1, Arg2, Arg3, Arg4):  pass
  def SortBy(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def SqrtPi(self, Arg1):  pass
  def StDev(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def StDevP(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def StDev_P(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def StDev_S(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def StEyx(self, Arg1, Arg2):  pass
  def Standardize(self, Arg1, Arg2, Arg3):  pass
  def StockHistory(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def Substitute(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Subtotal(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Sum(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def SumIf(self, Arg1, Arg2, Arg3):  pass
  def SumIfs(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def SumProduct(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def SumSq(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def SumX2MY2(self, Arg1, Arg2):  pass
  def SumX2PY2(self, Arg1, Arg2):  pass
  def SumXMY2(self, Arg1, Arg2):  pass
  def Syd(self, Arg1, Arg2, Arg3, Arg4):  pass
  def TBillEq(self, Arg1, Arg2, Arg3):  pass
  def TBillPrice(self, Arg1, Arg2, Arg3):  pass
  def TBillYield(self, Arg1, Arg2, Arg3):  pass
  def TDist(self, Arg1, Arg2, Arg3):  pass
  def TInv(self, Arg1, Arg2):  pass
  def TTest(self, Arg1, Arg2, Arg3, Arg4):  pass
  def T_Dist(self, Arg1, Arg2, Arg3):  pass
  def T_Dist_2T(self, Arg1, Arg2):  pass
  def T_Dist_RT(self, Arg1, Arg2):  pass
  def T_Inv(self, Arg1, Arg2):  pass
  def T_Inv_2T(self, Arg1, Arg2):  pass
  def T_Test(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Tanh(self, Arg1):  pass
  def Text(self, Arg1, Arg2):  pass
  def TextJoin(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29):  pass
  def ThaiDayOfWeek(self, Arg1):  pass
  def ThaiDigit(self, Arg1):  pass
  def ThaiMonthOfYear(self, Arg1):  pass
  def ThaiNumSound(self, Arg1):  pass
  def ThaiNumString(self, Arg1):  pass
  def ThaiStringLength(self, Arg1):  pass
  def ThaiYear(self, Arg1):  pass
  def Transpose(self, Arg1):  pass
  def Trend(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Trim(self, Arg1):  pass
  def TrimMean(self, Arg1, Arg2):  pass
  def USDollar(self, Arg1, Arg2):  pass
  def Unichar(self, Arg1):  pass
  def Unicode(self, Arg1):  pass
  def Unique(self, Arg1, Arg2, Arg3):  pass
  def VLookup(self, Arg1, Arg2, Arg3, Arg4):  pass
  def ValueToText(self, Arg1, Arg2):  pass
  def Var(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def VarP(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Var_P(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Var_S(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def Vdb(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7):  pass
  def WebService(self, Arg1):  pass
  def WeekNum(self, Arg1, Arg2):  pass
  def Weekday(self, Arg1, Arg2):  pass
  def Weibull(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Weibull_Dist(self, Arg1, Arg2, Arg3, Arg4):  pass
  def WorkDay(self, Arg1, Arg2, Arg3):  pass
  def WorkDay_Intl(self, Arg1, Arg2, Arg3, Arg4):  pass
  def XLookup(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def XMatch(self, Arg1, Arg2, Arg3, Arg4):  pass
  def Xirr(self, Arg1, Arg2, Arg3):  pass
  def Xnpv(self, Arg1, Arg2):  pass
  def Xor(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def YearFrac(self, Arg1, Arg2, Arg3):  pass
  def YieldDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5):  pass
  def YieldMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6):  pass
  def ZTest(self, Arg1, Arg2, Arg3):  pass
  def Z_Test(self, Arg1, Arg2, Arg3):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def _WSFunction(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __len__(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WorksheetFunction'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:424, whats:4,   ok:435, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Areas:
  def __init__(self):
    self.Application: Application
    self.Count: int
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
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Areas'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Borders:
  def __init__(self):
    self.Application: Application
    self.Color: float
    self.ColorIndex: int
    self.Count: int
    self.Creator: int
    self.Parent: CellFormat
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

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class Characters:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Font: Font
    self.Parent: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):  pass
  def Insert(self, String):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def __len__(self):  pass
  def __nonzero__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class DisplayFormat:
  def __init__(self):
    self.AddIndent: bool
    self.Application: Application
    self.Borders: Borders
    self.Characters: Characters
    self.Creator: int
    self.Font: Font
    self.FormulaHidden: bool
    self.HorizontalAlignment: int
    self.IndentLevel: int
    self.Interior: Interior
    self.Locked: bool
    self.MergeCells: bool
    self.NumberFormat: str
    self.NumberFormatLocal: str
    self.Orientation: int
    self.Parent: Range
    self.ReadingOrder: int
    self.ShrinkToFit: bool
    self.Style: Style
    self.VerticalAlignment: int
    self.WrapText: bool
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.DisplayFormat'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:25, methods:6, whats:4,   ok:35, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Errors'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:10, whats:4,   ok:21, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Font:
  def __init__(self):
    self.Application: Application
    self.Bold: bool
    self.Color: float
    self.ColorIndex: int
    self.Creator: int
    self.FontStyle: str
    self.Italic: bool
    self.Name: str
    self.OutlineFont: bool
    self.Parent: DisplayFormat
    self.Shadow: bool
    self.Size: float
    self.Strikethrough: bool
    self.Subscript: bool
    self.Superscript: bool
    self.ThemeColor: int
    self.ThemeFont: int
    self.TintAndShade: float
    self.Underline: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # Background:  <class 'NoneType'>
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Font'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:23, methods:5, whats:5,   ok:33, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class FormatConditions:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, Operator, Formula1, Formula2, String, TextOperator, DateOperator, ScopeType):  pass
  def AddAboveAverage(self):  pass
  def AddColorScale(self, ColorScaleType):  pass
  def AddDatabar(self):  pass
  def AddIconSetCondition(self):  pass
  def AddTop10(self):  pass
  def AddUniqueValues(self):  pass
  def Delete(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.FormatConditions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:20, whats:4,   ok:32, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Hyperlinks:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Anchor, Address, SubAddress, ScreenTip, TextToDisplay):  pass
  def Delete(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Hyperlinks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Interior:
  def __init__(self):
    self.Application: Application
    self.Color: float
    self.ColorIndex: int
    self.Creator: int
    self.Parent: DisplayFormat
    self.Pattern: int
    self.PatternColor: int
    self.PatternColorIndex: int
    self.PatternThemeColor: int
    self.PatternTintAndShade: float
    self.ThemeColor: int
    self.TintAndShade: float
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # Gradient:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # InvertIfNegative:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Interior'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:5, whats:5,   ok:26, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class Phonetic:
  def __init__(self):
    self.Alignment: int
    self.Application: Application
    self.CharacterType: int
    self.Creator: int
    self.Font: Font
    self.Parent: Range
    self.Text: str
    self.Visible: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Phonetic'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:12, methods:5, whats:4,   ok:21, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Phonetics:
  def __init__(self):
    self.Alignment: int
    self.Application: Application
    self.CharacterType: int
    self.Count: int
    self.Creator: int
    self.Font: Font
    self.Length: int
    self.Parent: Range
    self.Visible: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Start, Length, Text):  pass
  def Delete(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Start:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # Text:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Phonetics'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:13, methods:14, whats:4,   ok:31, er:0, er:2
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SmartTags'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:11, whats:4,   ok:23, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SoundNote'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:9, whats:4,   ok:20, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class SparklineGroups:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, SourceData):  pass
  def Clear(self):  pass
  def ClearGroups(self):  pass
  def Group(self, Location):  pass
  def Item(self, Index):  pass
  def Ungroup(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SparklineGroups'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:17, whats:4,   ok:29, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Style:
  def __init__(self):
    self.AddIndent: bool
    self.Application: Application
    self.Borders: Borders
    self.BuiltIn: bool
    self.Creator: int
    self.Font: Font
    self.FormulaHidden: bool
    self.HorizontalAlignment: int
    self.IncludeAlignment: bool
    self.IncludeBorder: bool
    self.IncludeFont: bool
    self.IncludeNumber: bool
    self.IncludePatterns: bool
    self.IncludeProtection: bool
    self.Interior: Interior
    self.Locked: bool
    self.Name: str
    self.NameLocal: str
    self.NumberFormat: str
    self.NumberFormatLocal: str
    self.Orientation: int
    self.Parent: _Workbook
    self.ReadingOrder: int
    self.ShrinkToFit: bool
    self.Value: str
    self.VerticalAlignment: int
    self.WrapText: bool
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # IndentLevel:  <class 'NoneType'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # MergeCells:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不能取得类 Style 的 MergeCells 属性', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Style'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:32, methods:8, whats:5,   ok:45, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class Validation:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.ErrorTitle: str
    self.InputTitle: str
    self.Parent: Range
    self.Value: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Type, AlertStyle, Operator, Formula1, Formula2):  pass
  def Delete(self):  pass
  def Modify(self, Type, AlertStyle, Operator, Formula1, Formula2):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class Worksheet:
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # coclass_interfaces:  <class 'list'>
    # coclass_sources:  <class 'list'>
    # default_interface:  <class 'type'>
    # default_source:  <class 'type'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Worksheet'>, <class 'win32com.client.CoClassBaseClass'>, <class 'object'>)", attrs:3, methods:8, whats:6,   ok:17, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class XPath:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Range
    self.Value: str
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Clear(self):  pass
  def SetValue(self, Map, XPath, SelectionNamespace, Repeating):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Map:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)
    # Repeating:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XPath'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:9, whats:4,   ok:22, er:0, er:2
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Menus'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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
    self.AutoCorrect: AutoCorrect
    self.AutoFormatAsYouTypeReplaceHyperlinks: bool
    self.AutoPercentEntry: bool
    self.AutoRecover: AutoRecover
    self.AutomationSecurity: int
    self.Build: float
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

  #whats:
    # ActiveChart:  <class 'NoneType'>
    # ActiveDialog:  <class 'NoneType'>
    # ActiveProtectedViewWindow:  <class 'NoneType'>
    # Assistance:  <class 'win32com.client.CDispatch'>
    # Assistant:  <class 'win32com.client.CDispatch'>
    # CLSID:  <class 'PyIID'>
    # COMAddIns:  <class 'win32com.client.CDispatch'>
    # CommandBars:  <class 'win32com.client.CDispatch'>
    # DataPrivacyOptions:  <class 'win32com.client.CDispatch'>
    # FileConverters:  <class 'NoneType'>
    # LanguageSettings:  <class 'win32com.client.CDispatch'>
    # MailSession:  <class 'NoneType'>
    # NewWorkbook:  <class 'win32com.client.CDispatch'>
    # OnCalculate:  <class 'NoneType'>
    # OnData:  <class 'NoneType'>
    # OnDoubleClick:  <class 'NoneType'>
    # OnEntry:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # OnWindow:  <class 'NoneType'>
    # PreviousSelections:  <class 'NoneType'>
    # RegisteredFunctions:  <class 'NoneType'>
    # SmartArtColors:  <class 'win32com.client.CDispatch'>
    # SmartArtLayouts:  <class 'win32com.client.CDispatch'>
    # SmartArtQuickStyles:  <class 'win32com.client.CDispatch'>
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
    # Hinstance:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147418113), None)
    # SensitivityLabelPolicy:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2147220726), None)
    # ThisCell:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # ThisWorkbook:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)
    # VBE:  (-2147352567, '发生意外。', (0, 'Microsoft Excel', '不信任到 Visual Basic Project 的程序连接\n', 'xlmain11.chm', 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Application'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:217, methods:94, whats:28,   ok:339, er:0, er:11
# just stub file, import it , declare app obj, get ide auto type hit

class Comments:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Comments'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class CommentsThreaded:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CommentsThreaded'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class CustomProperties:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Value):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CustomProperties'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class HPageBreaks:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Sheets
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.HPageBreaks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ListObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.NamedSheetViewCollection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Outline:
  def __init__(self):
    self.Application: Application
    self.AutomaticStyles: bool
    self.Creator: int
    self.Parent: _Worksheet
    self.SummaryColumn: int
    self.SummaryRow: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def ShowLevels(self, RowLevels, ColumnLevels):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Outline'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:6, whats:4,   ok:20, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class PageSetup:
  def __init__(self):
    self.AlignMarginsHeaderFooter: bool
    self.Application: Application
    self.BlackAndWhite: bool
    self.BottomMargin: float
    self.CenterFooter: str
    self.CenterFooterPicture: Graphic
    self.CenterHeader: str
    self.CenterHeaderPicture: Graphic
    self.CenterHorizontally: bool
    self.CenterVertically: bool
    self.ChartSize: int
    self.Creator: int
    self.DifferentFirstPageHeaderFooter: bool
    self.Draft: bool
    self.EvenPage: Page
    self.FirstPage: Page
    self.FirstPageNumber: int
    self.FitToPagesTall: int
    self.FitToPagesWide: int
    self.FooterMargin: float
    self.HeaderMargin: float
    self.LeftFooter: str
    self.LeftFooterPicture: Graphic
    self.LeftHeader: str
    self.LeftHeaderPicture: Graphic
    self.LeftMargin: float
    self.OddAndEvenPagesHeaderFooter: bool
    self.Order: float
    self.Orientation: float
    self.Pages: Pages
    self.PaperSize: float
    self.Parent: _Worksheet
    self.PrintArea: str
    self.PrintComments: int
    self.PrintErrors: int
    self.PrintGridlines: bool
    self.PrintHeadings: bool
    self.PrintNotes: bool
    self.PrintQuality: tuple
    self.PrintTitleColumns: str
    self.PrintTitleRows: str
    self.RightFooter: str
    self.RightFooterPicture: Graphic
    self.RightHeader: str
    self.RightHeaderPicture: Graphic
    self.RightMargin: float
    self.ScaleWithDocHeaderFooter: bool
    self.TopMargin: float
    self.Zoom: int
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PageSetup'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:53, methods:7, whats:4,   ok:64, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class _Workbook:
  def __init__(self):
    self.AcceptLabelsInFormulas: bool
    self.AccuracyVersion: int
    self.ActiveSheet: _Worksheet
    self.Application: Application
    self.Author: str
    self.AutoSaveOn: bool
    self.AutoUpdateFrequency: int
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
    self.CustomViews: CustomViews
    self.Date1904: bool
    self.DefaultPivotTableStyle: TableStyle
    self.DefaultSlicerStyle: TableStyle
    self.DefaultTableStyle: TableStyle
    self.DefaultTimelineStyle: TableStyle
    self.DialogSheets: Sheets
    self.DisplayDrawingObjects: int
    self.DisplayInkComments: bool
    self.DoNotPromptForConvert: bool
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
    self.SlicerCaches: SlicerCaches
    self.Styles: Styles
    self.Subject: str
    self.TableStyles: TableStyles
    self.TemplateRemoveExtData: bool
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

  #whats:
    # ActiveChart:  <class 'NoneType'>
    # ActiveSlicer:  <class 'NoneType'>
    # BuiltinDocumentProperties:  <class 'win32com.client.CDispatch'>
    # CLSID:  <class 'PyIID'>
    # CommandBars:  <class 'NoneType'>
    # CustomDocumentProperties:  <class 'win32com.client.CDispatch'>
    # CustomXMLParts:  <class 'win32com.client.CDispatch'>
    # DocumentInspectors:  <class 'win32com.client.CDispatch'>
    # OnSave:  <class 'NoneType'>
    # OnSheetActivate:  <class 'NoneType'>
    # OnSheetDeactivate:  <class 'NoneType'>
    # SharedWorkspace:  <class 'NoneType'>
    # Signatures:  <class 'win32com.client.CDispatch'>
    # SmartDocument:  <class 'win32com.client.CDispatch'>
    # SmartTagOptions:  <class 'NoneType'>
    # Sync:  <class 'win32com.client.CDispatch'>
    # Theme:  <class 'win32com.client.CDispatch'>
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

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9._Workbook'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:114, methods:89, whats:20,   ok:223, er:0, er:12
# just stub file, import it , declare app obj, get ide auto type hit

class Protection:
  def __init__(self):
    self.AllowDeletingColumns: bool
    self.AllowDeletingRows: bool
    self.AllowEditRanges: AllowEditRanges
    self.AllowFiltering: bool
    self.AllowFormattingCells: bool
    self.AllowFormattingColumns: bool
    self.AllowFormattingRows: bool
    self.AllowInsertingColumns: bool
    self.AllowInsertingHyperlinks: bool
    self.AllowInsertingRows: bool
    self.AllowSorting: bool
    self.AllowUsingPivotTables: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Protection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:5, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class QueryTables:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Connection, Destination, Sql):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.QueryTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Shapes:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Worksheet
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add3DModel(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height):  pass
  def AddCallout(self, Type, Left, Top, Width, Height):  pass
  def AddCanvas(self, Left, Top, Width, Height):  pass
  def AddChart(self, XlChartType, Left, Top, Width, Height):  pass
  def AddChart2(self, Style, XlChartType, Left, Top, Width, Height, NewLayout):  pass
  def AddConnector(self, Type, BeginX, BeginY, EndX, EndY):  pass
  def AddCurve(self, SafeArrayOfPoints):  pass
  def AddDiagram(self, Type, Left, Top, Width, Height):  pass
  def AddFormControl(self, Type, Left, Top, Width, Height):  pass
  def AddLabel(self, Orientation, Left, Top, Width, Height):  pass
  def AddLine(self, BeginX, BeginY, EndX, EndY):  pass
  def AddOLEObject(self, ClassType, Filename, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height):  pass
  def AddPicture(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height):  pass
  def AddPicture2(self, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height, Compress):  pass
  def AddPolyline(self, SafeArrayOfPoints):  pass
  def AddShape(self, Type, Left, Top, Width, Height):  pass
  def AddSmartArt(self, Layout, Left, Top, Width, Height):  pass
  def AddTextEffect(self, PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top):  pass
  def AddTextbox(self, Orientation, Left, Top, Width, Height):  pass
  def BuildFreeform(self, EditingType, X1, Y1):  pass
  def Item(self, Index):  pass
  def Range(self, Index):  pass
  def SelectAll(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Shapes'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:34, whats:4,   ok:46, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Sort:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Header: int
    self.MatchCase: bool
    self.Orientation: int
    self.Parent: _Worksheet
    self.SortFields: SortFields
    self.SortMethod: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Apply(self):  pass
  def SetRange(self, Rng):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:
    # Rng:  (-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284), None)

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Sort'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:12, methods:7, whats:4,   ok:23, er:0, er:1
# just stub file, import it , declare app obj, get ide auto type hit

class Tab:
  def __init__(self):
    self.Application: Application
    self.Color: bool
    self.ColorIndex: int
    self.Creator: int
    self.Parent: _Worksheet
    self.ThemeColor: int
    self.TintAndShade: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Tab'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:5, whats:4,   ok:20, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class VPageBreaks:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Sheets
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Before):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.VPageBreaks'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Pane:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Index: int
    self.Parent: Window
    self.ScrollColumn: float
    self.ScrollRow: float
    self.VisibleRange: Range
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Activate(self):  pass
  def LargeScroll(self, Down, Up, ToRight, ToLeft):  pass
  def PointsToScreenPixelsX(self, Points):  pass
  def PointsToScreenPixelsY(self, Points):  pass
  def ScrollIntoView(self, Left, Top, Width, Height, Start):  pass
  def SmallScroll(self, Down, Up, ToRight, ToLeft):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Pane'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:11, methods:11, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WorksheetView'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:13, methods:5, whats:4,   ok:22, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Panes:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Window
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Panes'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class SheetViews:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Window
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SheetViews'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Graphic:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Filename: str
    self.Parent: PageSetup
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class Page:
  def __init__(self):
    self.CenterFooter: HeaderFooter
    self.CenterHeader: HeaderFooter
    self.LeftFooter: HeaderFooter
    self.LeftHeader: HeaderFooter
    self.RightFooter: HeaderFooter
    self.RightHeader: HeaderFooter
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Page'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:4,   ok:19, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Pages:
  def __init__(self):
    self.Count: int
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Pages'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:12, whats:4,   ok:21, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Connections:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Description, ConnectionString, CommandText, lCmdtype):  pass
  def Add2(self, Name, Description, ConnectionString, CommandText, lCmdtype, CreateModelConnection, ImportRelationships):  pass
  def AddFromFile(self, Filename, CreateModelConnection, ImportRelationships):  pass
  def Item(self, Index):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Connections'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:16, whats:4,   ok:28, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class CustomViews:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, ViewName, PrintSettings, RowColSettings):  pass
  def Item(self, ViewName):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CustomViews'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyle'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:16, methods:9, whats:4,   ok:29, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class IconSets:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.IconSets'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Model:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DataModelConnection: WorkbookConnection
    self.ModelFormatBoolean: ModelFormatBoolean
    self.ModelFormatCurrency: ModelFormatCurrency
    self.ModelFormatDate: ModelFormatDate
    self.ModelFormatDecimalNumber: ModelFormatDecimalNumber
    self.ModelFormatGeneral: ModelFormatGeneral
    self.ModelFormatPercentageNumber: ModelFormatPercentageNumber
    self.ModelFormatScientificNumber: ModelFormatScientificNumber
    self.ModelFormatWholeNumber: ModelFormatWholeNumber
    self.ModelMeasures: ModelMeasures
    self.ModelRelationships: ModelRelationships
    self.ModelTables: ModelTables
    self.Name: str
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def AddConnection(self, ConnectionToDataSource):  pass
  def CreateModelWorkbookConnection(self, ModelTable):  pass
  def GetModelFormatCurrency(self, Symbol, DecimalPlaces):  pass
  def GetModelFormatDate(self, FormatString):  pass
  def GetModelFormatDecimalNumber(self, UseThousandSeparator, DecimalPlaces):  pass
  def GetModelFormatPercentageNumber(self, UseThousandSeparator, DecimalPlaces):  pass
  def GetModelFormatScientificNumber(self, DecimalPlaces):  pass
  def GetModelFormatWholeNumber(self, UseThousandSeparator):  pass
  def Initialize(self):  pass
  def Refresh(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Model'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:20, methods:15, whats:4,   ok:39, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PivotTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:10, whats:4,   ok:22, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class PublishObjects:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, SourceType, Filename, Sheet, Source, HtmlType, DivID, Title):  pass
  def Delete(self):  pass
  def Item(self, Index):  pass
  def Publish(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.PublishObjects'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Queries:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.FastCombine: bool
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Formula, Description):  pass
  def Item(self, NameOrIndex):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Queries'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Research'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:8, whats:4,   ok:19, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class SlicerCaches:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Source, SourceField, Name):  pass
  def Add2(self, Source, SourceField, Name, SlicerCacheType):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SlicerCaches'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class Styles:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, BasedOn):  pass
  def Item(self, Index):  pass
  def Merge(self, Workbook):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Styles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class TableStyles:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, TableStyleName):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyles'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class WebOptions:
  def __init__(self):
    self.AllowPNG: bool
    self.Application: Application
    self.Creator: int
    self.DownloadComponents: bool
    self.Encoding: int
    self.FolderSuffix: str
    self.LocationOfComponents: str
    self.OrganizeInFolder: bool
    self.Parent: _Workbook
    self.PixelsPerInch: int
    self.RelyOnCSS: bool
    self.RelyOnVML: bool
    self.ScreenSize: int
    self.TargetBrowser: int
    self.UseLongFileNames: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def UseDefaultFolderSuffix(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.WebOptions'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:19, methods:6, whats:4,   ok:29, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class XmlMaps:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Schema, RootElementName):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XmlMaps'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class XmlNamespaces:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: _Workbook
    self.Value: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def InstallManifest(self, Path, InstallForAllUsers):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.XmlNamespaces'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:13, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class AllowEditRanges:
  def __init__(self):
    self.Count: int
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Title, Range, Password):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.AllowEditRanges'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:5, methods:13, whats:4,   ok:22, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class SortFields:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Sort
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Key, SortOn, Order, CustomOrder, DataOption):  pass
  def Add2(self, Key, SortOn, Order, CustomOrder, DataOption, SubField):  pass
  def Clear(self):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.SortFields'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.HeaderFooter'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:6, methods:5, whats:4,   ok:15, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.TableStyleElements'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class WorkbookConnection:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Description: str
    self.InModel: bool
    self.ModelConnection: ModelConnection
    self.Name: str
    self.Parent: _Workbook
    self.Ranges: Ranges
    self.RefreshWithRefreshAll: bool
    self.Type: int
    self._Default: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Delete(self):  pass
  def Refresh(self):  pass
  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __call__(self):  pass
  def __getattr__(self, attr):  pass
  def __int__(self, args):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
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
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatBoolean:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatBoolean'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:5, whats:4,   ok:16, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatCurrency:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DecimalPlaces: int
    self.Parent: Model
    self.Symbol: str
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatCurrency'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatDate:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.FormatString: str
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatDate'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatDecimalNumber:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DecimalPlaces: int
    self.Parent: Model
    self.UseThousandSeparator: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatDecimalNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatGeneral:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatGeneral'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:7, methods:5, whats:4,   ok:16, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatPercentageNumber:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DecimalPlaces: int
    self.Parent: Model
    self.UseThousandSeparator: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatPercentageNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:9, methods:5, whats:4,   ok:18, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatScientificNumber:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.DecimalPlaces: int
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatScientificNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelFormatWholeNumber:
  def __init__(self):
    self.Application: Application
    self.Creator: int
    self.Parent: Model
    self.UseThousandSeparator: bool
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def _ApplyTypes_(self, dispid, wFlags, retType, argTypes, user, resultCLSID, args):  pass
  def __getattr__(self, attr):  pass
  def __iter__(self):  pass
  def _get_good_object_(self, obj, obUserName, resultCLSID):  pass
  def _get_good_single_object_(self, obj, obUserName, resultCLSID):  pass

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelFormatWholeNumber'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:5, whats:4,   ok:17, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelMeasures:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, MeasureName, AssociatedTable, Formula, FormatInformation, Description):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelMeasures'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:13, whats:4,   ok:25, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelRelationships:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Model
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, ForeignKeyColumn, PrimaryKeyColumn):  pass
  def DetectRelationships(self, PivotTable):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelRelationships'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:14, whats:4,   ok:26, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelTables:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: Model
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelTables'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class ModelConnection:
  def __init__(self):
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

  #whats:
    # ADOConnection:  <class 'win32com.client.CDispatch'>
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.ModelConnection'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:10, methods:5, whats:5,   ok:20, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.Ranges'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:12, whats:4,   ok:24, er:0, er:0
# just stub file, import it , declare app obj, get ide auto type hit

class CalculatedMembers:
  def __init__(self):
    self.Application: Application
    self.Count: int
    self.Creator: int
    self.Parent: ModelConnection
    self.__dict__: dict
    self.__module__: str
    self._prop_map_get_: dict
    self._prop_map_put_: dict

  def Add(self, Name, Formula, SolveOrder, Type, Dynamic, DisplayFolder, HierarchizeDistinct):  pass
  def AddCalculatedMember(self, Name, Formula, SolveOrder, Type, DisplayFolder, MeasureGroup, ParentHierarchy, ParentMember, NumberFormat):  pass
  def Item(self, Index):  pass
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

  #whats:
    # CLSID:  <class 'PyIID'>
    # __weakref__:  <class 'NoneType'>
    # _oleobj_:  <class 'PyIDispatch'>
    # coclass_clsid:  <class 'NoneType'>

  #getattr AttributeError:

  #getattr Exception:

# Summary "(<class 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9.CalculatedMembers'>, <class 'win32com.client.DispatchBaseClass'>, <class 'object'>)", attrs:8, methods:15, whats:4,   ok:27, er:0, er:0
