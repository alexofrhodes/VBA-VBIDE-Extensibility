**WIP**

>Disclaimer
>
>Provided as is, read and test.
There be dragons, there be bugs. Feedback is welcome.


**Tools to extend and facilitate work with**
- Modules
- Procedures
- Codemodule
- Designer
- Userforms
- other objects

**Notable userforms in this project**
- ChangeLog Manager
- Code on the fly
- Code Finder
- Project explorer
- Radial Menu
- Multipage to modern UI

**Other projects of mine**
- standalone : Dynamic Ribbon for Excel, Word, PowerPoint, Access
- standalone : Code Printer
- addin : Table Manager
- game : Go aka Baduk aka WeiQi

..and more..

<br>

**TOC of public procedures in my classes**

# aCodeModule

| Type | Procedure                 | Returns             |
| ---- | ------------------------- | ------------------  |
| Fun  | Active                    | As aCodeModule      |
| Fun  | ActiveEnum                | As aModuleEnumItem  |
| Fun  | ActiveType                | As aModuleTypeItem  |
| Fun  | CaseTo                    | As aCodeModule      |
| Fun  | Component                 | As VBComponent      |
| Fun  | EncapsulateParenthesis    | As aCodeModule      |
| Fun  | EncapsulateQuotes         | As aCodeModule      |
| Fun  | Initialize                | As aCodeModule      |
| Fun  | Procedure                 | As aProcedure       |
| Fun  | ProceduresByDeclaration   | As Collection       |
| Fun  | Region                    | As aCodeModule      |
| Fun  | Regions                   | As Variant          |
| Fun  | SortComma                 | As aCodeModule      |
| Fun  | SortLines                 | As aCodeModule      |
| Fun  | SortLinesByLength         | As aCodeModule      |
| Fun  | Substitute                | As aCodeModule      |
| Fun  | lines                     | As aCodeModule      |
| Get  | ProceduresByName          | As Collection       |
| Sub  | Align                     |                     |
| Sub  | AlignAs                   |                     |
| Sub  | AlignColumn               |                     |
| Sub  | AlignComments             |                     |
| Sub  | BeautifyFunction          |                     |
| Sub  | BringProcedureHere        |                     |
| Sub  | Comment                   |                     |
| Sub  | CommentTargetLine         |                     |
| Sub  | CommentsAddRem            |                     |
| Sub  | CommentsRemoveRem         |                     |
| Sub  | CommentsToggle            |                     |
| Sub  | Copy                      |                     |
| Sub  | Cut                       |                     |
| Sub  | DeleteSelectedLines       |                     |
| Sub  | DimMerge                  |                     |
| Sub  | DimSeparate               |                     |
| Sub  | Duplicate                 |                     |
| Sub  | Encapsulate               |                     |
| Sub  | EncapsulateMultiple       |                     |
| Sub  | FoldLine                  |                     |
| Sub  | Increment                 |                     |
| Sub  | Inject                    |                     |
| Sub  | InjectArgumentStyleFolded |                     |
| Sub  | Insert                    |                     |
| Sub  | Move                      |                     |
| Sub  | ProcedureActivate         |                     |
| Sub  | ProcedureImport           |                     |
| Sub  | Rotate                    |                     |
| Sub  | RotateCommas              |                     |
| Sub  | RotateEqualInLines        |                     |
| Sub  | RotateLines               |                     |
| Sub  | RotateMultiple            |                     |
| Sub  | SetSelection              |                     |
| Sub  | Sort                      |                     |
| Sub  | Todo                      |                     |
| Sub  | UnComment                 |                     |
| Sub  | UnFoldLine                |                     |
| Sub  | UncommentTargetLine       |                     |
| Sub  | UnremTargetLine           |                     |
| Sub  | injectDivider             |                     |

# aColorScheme

| Type | Procedure                | Returns          |
| ---- | ------------------------ | ---------------  |
| Fun  | Init                     | As aColorScheme  |
| Sub  | AssignColors             |                  |
| Sub  | ThemeBlackAndBlueDark    |                  |
| Sub  | ThemeBlackAndBrownDark   |                  |
| Sub  | ThemeBlackAndGrayDark    |                  |
| Sub  | ThemeBlackAndGreenDark   |                  |
| Sub  | ThemeBlackAndOrangeDark  |                  |
| Sub  | ThemeBlackAndPinkDark    |                  |
| Sub  | ThemeBlackAndPurpleDark  |                  |
| Sub  | ThemeBlackAndRedDark     |                  |
| Sub  | ThemeBlackAndYellowDark  |                  |
| Sub  | ThemeBlueAndGreenLight   |                  |
| Sub  | ThemeWhiteAndBlueLight   |                  |
| Sub  | ThemeWhiteAndBrownLight  |                  |
| Sub  | ThemeWhiteAndGrayLight   |                  |
| Sub  | ThemeWhiteAndGreenLight  |                  |
| Sub  | ThemeWhiteAndOrangeLight |                  |
| Sub  | ThemeWhiteAndPinkLight   |                  |
| Sub  | ThemeWhiteAndPurpleLight |                  |
| Sub  | ThemeWhiteAndRedLight    |                  |
| Sub  | ThemeWhiteAndYellowLight |                  |
| Sub  | color                    |                  |

# aComboBox

| Type | Procedure        | Returns       |
| ---- | ---------------- | ------------  |
| Fun  | AutoSizeDropDown | As Long       |
| Fun  | Init             | As aComboBox  |
| Sub  | LoadVBProjects   |               |

# aDesigner

| Type | Procedure                        | Returns             |
| ---- | -------------------------------- | ------------------  |
| Fun  | Active                           | As aDesigner        |
| Fun  | SelectedControl                  | As MSForms.control  |
| Fun  | SelectedControls                 | As Collection       |
| Fun  | SelectedFrameOrMultipageControl  | As MSForms.control  |
| Fun  | SelectedFrameOrMultipageControls | As Collection       |
| Sub  | CenterLabelCaption               |                     |
| Sub  | CopyControlProperties            |                     |
| Sub  | CopySubControlProperties         |                     |
| Sub  | CreateFrameMenu                  |                     |
| Sub  | EditObjectProperties             |                     |
| Sub  | EditObjectsProperty              |                     |
| Sub  | IconDesign                       |                     |
| Sub  | PasteControlProperties           |                     |
| Sub  | PasteSubControlProperties        |                     |
| Sub  | RemoveCaption                    |                     |
| Sub  | RenameControlAndCode             |                     |
| Sub  | ReplaceCommandButtonWithLabel    |                     |
| Sub  | SetHandCursor                    |                     |
| Sub  | SetHandCursorToSubControls       |                     |
| Sub  | SortControlsHorizontally         |                     |
| Sub  | SortControlsVertically           |                     |
| Sub  | SwitchNames                      |                     |
| Sub  | SwitchPositions                  |                     |
| Sub  | addFrameFormCode                 |                     |

# aFrame

| Type | Procedure           | Returns    |
| ---- | ------------------- | ---------  |
| Fun  | Init                | As aFrame  |
| Sub  | AddThemeControls    |            |
| Sub  | ResizeToFitControls |            |

# aListBox

| Type | Procedure                 | Returns                      |
| ---- | ------------------------- | ---------------------------  |
| Fun  | Contains                  | As Boolean                   |
| Fun  | Init                      | As aListBox                  |
| Fun  | LoadCSV                   | As Variant                   |
| Fun  | Parent                    | As Variant                   |
| Fun  | SelectedCount             | As Long                      |
| Fun  | SelectedRowsArray         | As Variant                   |
| Fun  | SelectedRowsText          | As String                    |
| Fun  | SelectedValues            | As Collection single column  |
| Fun  | TotalColumnsWidth         | As Variant                   |
| Fun  | selectedIndexes           | As Collection                |
| Fun  | targetColumn              | As Variant                   |
| Sub  | AcceptFiles               |                              |
| Sub  | AddFilter                 |                              |
| Sub  | AddHeader                 |                              |
| Sub  | AutofitColumns            |                              |
| Sub  | ClearSelection            |                              |
| Sub  | DeselectAll               |                              |
| Sub  | DeselectLike              |                              |
| Sub  | FilterByColumn            |                              |
| Sub  | HeightToEntries           |                              |
| Sub  | ListenToDoubleClick       |                              |
| Sub  | ListenToDragDrop          |                              |
| Sub  | ListenToExtendedSelection |                              |
| Sub  | LoadVBProjects            |                              |
| Sub  | RememberList              |                              |
| Sub  | RemoveSelected            |                              |
| Sub  | SelectAll                 |                              |
| Sub  | SelectItems               |                              |
| Sub  | SelectLike                |                              |
| Sub  | SelectedToRange           |                              |
| Sub  | ShowTheseColumns          |                              |
| Sub  | SortAZ                    |                              |
| Sub  | SortOnColumn              |                              |
| Sub  | SortZA                    |                              |
| Sub  | ToRange                   |                              |
| Sub  | removeHeaders             |                              |

# aListView

| Type | Procedure              | Returns       |
| ---- | ---------------------- | ------------  |
| Fun  | ClickedColumn          | As Variant    |
| Fun  | Init                   | As aListView  |
| Fun  | RowArray               | As Variant    |
| Fun  | SelectionArray         | As Variant    |
| Fun  | value                  | As Variant    |
| Sub  | AppendArray            |               |
| Sub  | AutofitColumns         |               |
| Sub  | DeselectAll            |               |
| Sub  | EnableDragSort         |               |
| Sub  | EnableDropFilesFolders |               |
| Sub  | EventListener          |               |
| Sub  | InitializeFromArray    |               |
| Sub  | RowsFormatOddEven      |               |
| Sub  | clear                  |               |

# aModule

| Type | Procedure                   | Returns                 |
| ---- | --------------------------- | ----------------------  |
| Fun  | Body                        | As String               |
| Fun  | ClassCalls                  | As Variant              |
| Fun  | Code                        | As String               |
| Fun  | Component                   | As VBComponent          |
| Fun  | Contains                    | As Variant              |
| Fun  | Copy                        | As Boolean              |
| Fun  | Duplicate                   | As Boolean              |
| Fun  | Enums                       | As aModuleEnums         |
| Fun  | Extension                   | As String               |
| Fun  | Folders                     | As aModuleFolders       |
| Fun  | Header                      | As String               |
| Fun  | HeaderContains              | As Boolean              |
| Fun  | Ignore                      | As Boolean              |
| Fun  | Initialize                  | As aModule              |
| Fun  | LineLike                    | As Long                 |
| Fun  | LinesLike                   | As Collection           |
| Fun  | ListOfInclude               | As Collection           |
| Fun  | Name                        | As String               |
| Fun  | Procedures                  | As aModuleProcedures    |
| Fun  | TodoList                    | As Variant              |
| Fun  | TypeToLong                  | As vbext_ComponentType  |
| Fun  | TypeToString                | As String               |
| Fun  | Types                       | As aModuleTypes         |
| Get  | Active                      | As aModule              |
| Get  | Project                     | As VBProject            |
| Get  | WorkbookObject              | As Workbook             |
| Sub  | Activate                    |                         |
| Sub  | CodeMove                    |                         |
| Sub  | CodeRemove                  |                         |
| Sub  | CommentsRemove              |                         |
| Sub  | CommentsToOwnLine           |                         |
| Sub  | Delete                      |                         |
| Sub  | DeleteIfEmpty               |                         |
| Sub  | DisableDebugPrint           |                         |
| Sub  | DisableStop                 |                         |
| Sub  | EnableDebugPrint            |                         |
| Sub  | EnableStop                  |                         |
| Sub  | Export                      |                         |
| Sub  | HeaderAdd                   |                         |
| Sub  | Indent                      |                         |
| Sub  | PredeclaredId               |                         |
| Sub  | PrintListOfInclude          |                         |
| Sub  | PrintTodoList               |                         |
| Sub  | ProcedureFoldDeclarations   |                         |
| Sub  | RemoveEmptyLines            |                         |
| Sub  | RemoveEmptyLinesButLeaveOne |                         |
| Sub  | RemoveLinesLike             |                         |
| Sub  | Rename                      |                         |

# aModuleEnumItem

| Type | Procedure    | Returns             |
| ---- | ------------ | ------------------  |
| Fun  | Initialize   | As aModuleEnumItem  |
| Get  | Body         | As String           |
| Get  | Name         | As String           |
| Get  | countoflines | As Long             |
| Get  | firstline    | As Long             |
| Get  | index        | As String           |
| Get  | lastline     | As Long             |
| Let  | Name         |                     |
| Let  | index        |                     |
| Sub  | AssignValues |                     |
| Sub  | ToCase       |                     |

# aModuleEnums

| Type | Procedure  | Returns             |
| ---- | ---------- | ------------------  |
| Fun  | Active     | As aModuleEnumItem  |
| Fun  | Initialize | As aModuleEnums     |
| Get  | Items      | As Variant          |

# aModuleFolders

| Type | Procedure  | Returns            |
| ---- | ---------- | -----------------  |
| Fun  | Exists     | As Variant         |
| Fun  | Initialize | As aModuleFolders  |
| Fun  | ToString   | As String          |
| Sub  | Append     |                    |
| Sub  | Create     |                    |
| Sub  | Delete     |                    |
| Sub  | Overwrite  |                    |

# aModuleProcedures

| Type | Procedure         | Returns               |
| ---- | ----------------- | --------------------  |
| Fun  | Initialize        | As aModuleProcedures  |
| Fun  | Items             | As Variant            |
| Fun  | Names             | As Collection         |
| Fun  | PrivateProcedures | As Collection         |
| Fun  | PublicProcedures  | As Collection         |
| Sub  | Export            |                       |
| Sub  | List              |                       |
| Sub  | ListPublic        |                       |
| Sub  | SortAZ            |                       |
| Sub  | SortByKind        |                       |
| Sub  | SortByScope       |                       |
| Sub  | Update            |                       |

# aModules

| Type | Procedure                   | Returns        |
| ---- | --------------------------- | -------------  |
| Fun  | AddOrSet                    | As aModule     |
| Fun  | ClassNames                  | As Variant     |
| Fun  | Classes                     | As Collection  |
| Fun  | Documents                   | As Collection  |
| Fun  | Exists                      | As Boolean     |
| Fun  | Initialize                  | As aModules    |
| Fun  | Item                        | As aModule     |
| Fun  | ModuleNames                 | As Variant     |
| Fun  | NamesOf                     | As Variant     |
| Fun  | NormalModules               | As Collection  |
| Fun  | UserformNames               | As Variant     |
| Fun  | Userforms                   | As Collection  |
| Get  | Items                       | As Variant     |
| Sub  | CaseProperModulesOfWorkbook |                |
| Sub  | CommentsRemove              |                |
| Sub  | Export                      |                |
| Sub  | ExportProcedures            |                |
| Sub  | ImportPaths                 |                |
| Sub  | ImportPicker                |                |
| Sub  | Indent                      |                |
| Sub  | InjectOptionExplicit        |                |
| Sub  | ListProcedures              |                |
| Sub  | PrintTodoList               |                |
| Sub  | Refresh                     |                |
| Sub  | RemoveEmptyLinesButLeaveOne |                |
| Sub  | RemoveLinesLike             |                |
| Sub  | RemoveProcedureList         |                |
| Sub  | SideBySide                  |                |
| Sub  | UpdateProcedures            |                |

# aModuleTypeItem

| Type | Procedure    | Returns             |
| ---- | ------------ | ------------------  |
| Fun  | Initialize   | As aModuleTypeItem  |
| Get  | Body         | As String           |
| Get  | Name         | As String           |
| Get  | countoflines | As Long             |
| Get  | firstline    | As Long             |
| Get  | index        | As String           |
| Get  | lastline     | As Long             |
| Let  | Name         |                     |
| Let  | index        |                     |
| Sub  | AssignValues |                     |

# aModuleTypes

| Type | Procedure  | Returns             |
| ---- | ---------- | ------------------  |
| Fun  | Active     | As aModuleTypeItem  |
| Fun  | Initialize | As aModuleTypes     |
| Get  | Items      | As Variant          |

# aMultiPage

| Type | Procedure                   | Returns          |
| ---- | --------------------------- | ---------------  |
| Fun  | ActivePage                  | As MSForms.page  |
| Fun  | Init                        | As aMultiPage    |
| Fun  | OutlookCheck                | As Boolean       |
| Sub  | AddContactsToSidebarBottom  |                  |
| Sub  | AddThemeControlsSidbarRight |                  |
| Sub  | BuildMenu                   |                  |
| Sub  | MailDev                     |                  |
| Sub  | SetBackColor                |                  |

# aProcedure

| Type | Procedure        | Returns                        |
| ---- | ---------------- | -----------------------------  |
| Fun  | Active           | As aProcedure                  |
| Fun  | Code             | As aProcedureCode              |
| Fun  | CustomProperties | As aProcedureCustomProperties  |
| Fun  | Dependencies     | As aProcedureDependencies      |
| Fun  | Folder           | As aProcedureFolder            |
| Fun  | Format           | As aProcedureFormat            |
| Fun  | Initialize       | As aProcedure                  |
| Fun  | Inject           | As aProcedureInject            |
| Fun  | Move             | As aProcedureMove              |
| Fun  | Scope            | As aProcedureScope             |
| Fun  | Variables        | As aProcedureVariables         |
| Fun  | arguments        | As aProcedureArguments         |
| Fun  | lines            | As aProcedureLines             |
| Get  | KindAsLong       | As Long                        |
| Get  | KindAsString     | As String                      |
| Get  | Name             | As String                      |
| Get  | Parent           | As VBComponent                 |
| Get  | returnType       | As String                      |
| Sub  | Activate         |                                |
| Sub  | CreateCaller     |                                |
| Sub  | CreateTest       |                                |
| Sub  | Delete           |                                |
| Sub  | Replace          |                                |

# aProcedureArguments

| Type | Procedure  | Returns                 |
| ---- | ---------- | ----------------------  |
| Fun  | Initialize | As aProcedureArguments  |
| Fun  | MultiLine  | As String               |
| Fun  | SingleLine | As String               |
| Get  | AsSeen     | As String               |
| Get  | Items      | As Variant              |
| Get  | count      | As Long                 |

# aProcedureArgumentsItem

| Type | Procedure       | Returns                     |
| ---- | --------------- | --------------------------  |
| Fun  | Initialize      | As aProcedureArgumentsItem  |
| Get  | DefaultValue    | As Variant                  |
| Get  | IsByRef         | As Boolean                  |
| Get  | IsByVal         | As Boolean                  |
| Get  | IsOptional      | As Boolean                  |
| Get  | IsParamArray    | As Boolean                  |
| Get  | IsType          | As String                   |
| Get  | Name            | As String                   |
| Get  | hasDefaultValue | As Boolean                  |
| Let  | DefaultValue    |                             |
| Let  | IsByRef         |                             |
| Let  | IsByVal         |                             |
| Let  | IsOptional      |                             |
| Let  | IsParamArray    |                             |
| Let  | IsType          |                             |
| Let  | Name            |                             |
| Let  | hasDefaultValue |                             |

# aProcedureCode

| Type | Procedure         | Returns              |
| ---- | ----------------- | -------------------  |
| Fun  | Initialize        | As aProcedureCode    |
| Fun  | Inject            | As aProcedureInject  |
| Fun  | lines             | As aProcedureLines   |
| Get  | All               | As Variant           |
| Get  | Body              | As Variant           |
| Get  | BodyAfterComments | As Variant           |
| Get  | Contains          | As Boolean           |
| Get  | ContainsInBody    | As String            |
| Get  | ContainsInHeader  | As String            |
| Get  | Declaration       | As Variant           |
| Get  | DeclarationClean  | As Variant           |
| Get  | Header            | As Variant           |

# aProcedureCustomProperties

| Type | Procedure      | Returns                        |
| ---- | -------------- | -----------------------------  |
| Fun  | Initialize     | As aProcedureCustomProperties  |
| Get  | Ignore         | As Boolean                     |
| Get  | LastModified   | As String                      |
| Get  | ParentAssigned | As String                      |
| Let  | Ignore         |                                |
| Let  | LastModified   |                                |
| Let  | ParentAssigned |                                |

# aProcedureDependencies

| Type | Procedure                 | Returns                    |
| ---- | ------------------------- | -------------------------  |
| Fun  | CallerModules             | As Collection              |
| Fun  | CallerModulesToString     | As String                  |
| Fun  | Callers                   | As Collection              |
| Fun  | CallersToString           | As String                  |
| Fun  | DeclarationsTableKeywords | As Collection              |
| Fun  | Initialize                | As aProcedureDependencies  |
| Fun  | LinkedClasses             | As Collection              |
| Fun  | LinkedDeclarations        | As Collection              |
| Fun  | LinkedProcedures          | As Collection              |
| Fun  | LinkedProceduresDeep      | As Collection              |
| Fun  | LinkedSheets              | As Collection              |
| Fun  | LinkedUserforms           | As Variant                 |
| Fun  | collLinkedProcedures      | As Collection              |
| Fun  | collLinkedProceduresDeep  | As Collection              |
| Sub  | AddToLinkedTable          |                            |
| Sub  | BringLinkedProceduresHere |                            |
| Sub  | BringProcedureHere        |                            |
| Sub  | Export                    |                            |
| Sub  | ExportDeclaration         |                            |
| Sub  | ExportLinkedCode          |                            |
| Sub  | ImportClass               |                            |
| Sub  | ImportDeclaration         |                            |
| Sub  | ImportDependencies        |                            |
| Sub  | ImportProcedure           |                            |
| Sub  | ImportUserform            |                            |
| Sub  | InjectLinkedClasses       |                            |
| Sub  | InjectLinkedDeclarations  |                            |
| Sub  | InjectLinkedLists         |                            |
| Sub  | InjectLinkedProcedures    |                            |
| Sub  | InjectLinkedUserforms     |                            |
| Sub  | RemoveIncludeLines        |                            |
| Sub  | Update                    |                            |

# aProcedureFolder

| Type | Procedure  | Returns              |
| ---- | ---------- | -------------------  |
| Fun  | Exists     | As Boolean           |
| Fun  | Initialize | As aProcedureFolder  |
| Sub  | Append     |                      |
| Sub  | Create     |                      |
| Sub  | Delete     |                      |
| Sub  | Overwrite  |                      |

# aProcedureFormat

| Type | Procedure            | Returns              |
| ---- | -------------------- | -------------------  |
| Fun  | Initialize           | As aProcedureFormat  |
| Sub  | BlankLinesToDividers |                      |
| Sub  | CommentsRemove       |                      |
| Sub  | CommentsToOwnLine    |                      |
| Sub  | CommentsToRem        |                      |
| Sub  | DisableDebugPrint    |                      |
| Sub  | DisableStop          |                      |
| Sub  | EnableDebugPrint     |                      |
| Sub  | EnableStop           |                      |
| Sub  | FoldDeclaration      |                      |
| Sub  | Indent               |                      |
| Sub  | NumbersAdd           |                      |
| Sub  | NumbersRemove        |                      |
| Sub  | RemoveEmptyLines     |                      |
| Sub  | RemoveLinesLike      |                      |
| Sub  | Replace              |                      |
| Sub  | UnfoldDeclaration    |                      |

# aProcedureInject

| Type | Procedure           | Returns              |
| ---- | ------------------- | -------------------  |
| Fun  | Initialize          | As aProcedureInject  |
| Get  | ObjectsReleaseText  | As String            |
| Sub  | BodyAfterComments   |                      |
| Sub  | BodyBottom          |                      |
| Sub  | BodyTop             |                      |
| Sub  | Description         |                      |
| Sub  | HeaderBottom        |                      |
| Sub  | HeaderTop           |                      |
| Sub  | Modification        |                      |
| Sub  | ObjectsReleaseAtEnd |                      |
| Sub  | ObjectsReleaseHere  |                      |
| Sub  | Template            |                      |
| Sub  | TemplateObject      |                      |
| Sub  | Timer               |                      |
| Sub  | test                |                      |

# aProcedureLines

| Type | Procedure                      | Returns             |
| ---- | ------------------------------ | ------------------  |
| Fun  | Initialize                     | As aProcedureLines  |
| Get  | CountOfBody                    | As Long             |
| Get  | CountOfDeclarationLines        | As Long             |
| Get  | CountOfHeaderLines             | As Variant          |
| Get  | FirstOfBody                    | As Long             |
| Get  | FirstOfBodyAfterComments       | As Long             |
| Get  | FirstOfDeclaration             | As Long             |
| Get  | FirstOfHeader                  | As Long             |
| Get  | LastOfBody                     | As Long             |
| Get  | LastOfDeclaration              | As Long             |
| Get  | LastOfHeader                   | As Long             |
| Get  | LikeThis                       | As Long             |
| Get  | Longest                        | As Long             |
| Get  | count                          | As Long             |
| Get  | first                          | As Long             |
| Get  | last                           | As Long             |
| Sub  | EnsureBlankLineBeforeProcedure |                     |

# aProcedureMove

| Type | Procedure        | Returns            |
| ---- | ---------------- | -----------------  |
| Fun  | Initialize       | As aProcedureMove  |
| Get  | IndexInModule    | As Long            |
| Sub  | Bottom           |                    |
| Sub  | Copy             |                    |
| Sub  | Down             |                    |
| Sub  | ToModule         |                    |
| Sub  | ToModuleAssigned |                    |
| Sub  | Top              |                    |
| Sub  | Up               |                    |

# aProcedureScope

| Type | Procedure     | Returns             |
| ---- | ------------- | ------------------  |
| Fun  | Initialize    | As aProcedureScope  |
| Fun  | ToString      | As String           |
| Get  | Suggested     | As String           |
| Sub  | MakePrivate   |                     |
| Sub  | MakePublic    |                     |
| Sub  | MakeSuggested |                     |

# aProcedureVariables

| Type | Procedure               | Returns                 |
| ---- | ----------------------- | ----------------------  |
| Fun  | Initialize              | As aProcedureVariables  |
| Get  | Items                   | As Variant              |
| Get  | count                   | As Long                 |
| Sub  | ToImmediate             |                         |
| Sub  | UpdatableVariableAdd    |                         |
| Sub  | UpdatableVariableRemove |                         |

# aProcedureVariablesItem

| Type | Procedure       | Returns                     |
| ---- | --------------- | --------------------------  |
| Fun  | Initialize      | As aProcedureVariablesItem  |
| Get  | IsType          | As String                   |
| Get  | Line            | As Long                     |
| Get  | Name            | As String                   |
| Get  | isAssignedValue | As Boolean                  |
| Let  | IsType          |                             |
| Let  | Line            |                             |
| Let  | Name            |                             |

# aProject

| Type | Procedure              | Returns                  |
| ---- | ---------------------- | -----------------------  |
| Fun  | Code                   | As String                |
| Fun  | Declarations           | As aProjectDeclarations  |
| Fun  | Extension              | As Variant               |
| Fun  | Name                   | As Variant               |
| Fun  | NameClean              | As Variant               |
| Fun  | ProceduresArray        | As Variant               |
| Fun  | ProceduresLike         | As Collection            |
| Fun  | REFERENCES             | As aProjectReferences    |
| Fun  | TodoList               | As String                |
| Fun  | WorkbookObject         | As Workbook              |
| Get  | Active                 | As aProject              |
| Get  | Initialize             | As aProject              |
| Get  | Items                  | As Variant               |
| Get  | Procedures             | As Collection            |
| Get  | ProceduresNames        | As Collection            |
| Get  | Project                | As VBProject             |
| Get  | this                   | As aProject              |
| Sub  | Backup                 |                          |
| Sub  | CreateLinkedTable      |                          |
| Sub  | CreateLinkedTableSheet |                          |
| Sub  | ExportCodeUnified      |                          |
| Sub  | ExportModules          |                          |
| Sub  | ExportProcedures       |                          |
| Sub  | ExportXML              |                          |
| Sub  | Indent                 |                          |
| Sub  | ModulesMerge           |                          |

# aProjectDeclarations

| Type | Procedure          | Returns                  |
| ---- | ------------------ | -----------------------  |
| Fun  | Initialize         | As aProjectDeclarations  |
| Fun  | Items              | As Collection            |
| Fun  | declaredEnums      | As String                |
| Fun  | declaredFunctions  | As String                |
| Fun  | declaredKeywords   | As Variant               |
| Fun  | declaredSubs       | As String                |
| Fun  | declaredTypes      | As String                |
| Fun  | tableKeywords      | As Collection            |
| Sub  | ExportDeclarations |                          |
| Sub  | ExportTable        |                          |
| Sub  | createTable        |                          |

# aProjectReferences

| Type | Procedure           | Returns                |
| ---- | ------------------- | ---------------------  |
| Fun  | Initialize          | As aProjectReferences  |
| Sub  | AddFromFile         |                        |
| Sub  | AddFromGUID         |                        |
| Sub  | AddScriptControl    |                        |
| Sub  | AddVBIDE            |                        |
| Sub  | Export              |                        |
| Sub  | ImportReferences    |                        |
| Sub  | RemoveByDescription |                        |
| Sub  | RemoveByGUID        |                        |
| Sub  | RemoveByName        |                        |
| Sub  | ToSheet             |                        |

# aTreeView

| Type | Procedure                   | Returns        |
| ---- | --------------------------- | -------------  |
| Fun  | GetLevel                    | As Integer     |
| Fun  | Init                        | As aTreeView   |
| Fun  | ToArray                     | As Variant     |
| Fun  | TreeviewArrayPaths          | As Variant 1d  |
| Fun  | columnCount                 | As Variant     |
| Fun  | rowCount                    | As Variant     |
| Sub  | ActivateProjectElement      |                |
| Sub  | ApplyStandardStyle          |                |
| Sub  | ChildrenCheck               |                |
| Sub  | CollapseAll                 |                |
| Sub  | ExpandAll                   |                |
| Sub  | FilterTV                    |                |
| Sub  | FindCodeEverywhere          |                |
| Sub  | ImageListLoadProjectIcons   |                |
| Sub  | LoadRange                   |                |
| Sub  | LoadTreeArray               |                |
| Sub  | LoadVBProjects              |                |
| Sub  | RemoveEmpty                 |                |
| Sub  | SelectNextNode              |                |
| Sub  | SelectNodes                 |                |
| Sub  | SelectPreviousNode          |                |
| Sub  | TreeviewArrayAppendPaths    |                |
| Sub  | TreeviewAssignProjectImages |                |
| Sub  | clear                       |                |

# aUserform

| Type | Procedure           | Returns                  |
| ---- | ------------------- | -----------------------  |
| Fun  | ClearBit            | As Long                  |
| Fun  | Effect              | As Scripting.Dictionary  |
| Fun  | EnableCloseButton   | As Boolean               |
| Fun  | HWndOfUserForm      | As Long                  |
| Fun  | Initialize          | As aUserform             |
| Fun  | IsFormResizable     | As Boolean               |
| Fun  | IsTitleBarVisible   | As Boolean               |
| Fun  | MakeFormResizable   | As Boolean               |
| Fun  | SetFormOpacity      | As Boolean               |
| Fun  | SetFormParent       | As Boolean               |
| Fun  | ShowCloseButton     | As Boolean               |
| Fun  | ShowMaximizeButton  | As Boolean               |
| Fun  | ShowMinimizeButton  | As Boolean               |
| Fun  | ShowTitleBar        | As Boolean               |
| Fun  | hwnd                | As LongPtr               |
| Sub  | Borderless          |                          |
| Sub  | DockControls        |                          |
| Sub  | IconDesign          |                          |
| Sub  | LoadOptions         |                          |
| Sub  | LoadPosition        |                          |
| Sub  | MouseOnControl      |                          |
| Sub  | OnTop               |                          |
| Sub  | ParentIsVBE         |                          |
| Sub  | Resizable           |                          |
| Sub  | Resizable2          |                          |
| Sub  | ResizeToFitControls |                          |
| Sub  | SaveOptions         |                          |
| Sub  | SavePosition        |                          |
| Sub  | ShowAtCursor        |                          |
| Sub  | TRANSPARENT         |                          |
| Sub  | Transition          |                          |
