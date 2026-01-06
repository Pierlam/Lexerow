namespace Lexerow.Core.System;

/// <summary>
/// Error/warning code.
/// </summary>
public enum ErrorCode
{
    Ok,

    Undefined,

    InternalError,

    UnexpectedException,

    /// <summary>
    /// Try to init the repository with a null value.
    /// </summary>
    RepositIsNull,

    RepositAlreadySet,

    RepositTypeNotImplemented,

    RepositNotSet,

    /// <summary>
    /// An unexcepted error (exception) occurs during the initialization of the repository in the core.
    /// </summary>
    RepositInitExceptionError,

    RepositCantCreateItAlreadyExists,

    RepositStringConnectionIsNullOrEmpty,

    RepositStringConnectionFileNameIsNullOrEmpty,

    RepositStringConnectionPathNameIsNullOrEmpty,

    RepositStringConnectionShouldMatchDatabase,

    RepositCantCreateDataBase,

    RepositCantOpenDataBase,

    RepositDatabaseIsAlreadyOpened,

    RepositNoOpenedDataBase,

    RepositCantCloseDataBase,

    RepositCollectionNameIsEmpty,

    RepositInsertObjectIsNull,

    RepositInsertObjectAlreadyInserted,

    RepositInsertObjectFailed,

    RepositInsertObjectIdExpected,

    RepositInsertObjectLinkedIdExpected,

    RepositUpdateObjectIsNull,

    RepositUpdateObjectNotYetInserted,

    RepositUpdateObjectFailed,

    RepositRemoveObjectIsNull,

    RepositRemoveObjectNotYetInserted,

    RepositRemoveObjectFailed,

    //-----DbDescriptor error

    DbDescriptorIsNull,
    DbDescriptorVersionIsNegative,
    DbDescriptorGuidIsNull,
    DbDescriptorGuidIsWrong,

    //-----others error code

    VarNameNotFound,
    FileNameNotFound,

    NameNullOrEmpty,
    FileNameNullOrEmpty,
    FileNameWrongSyntax,

    FileObjectNameNullOrEmpty,
    ObjectWithSameNameAlreadyExists,

    ExcelFileObjectNameIsNull,
    ExcelFileObjectNameDoesNotExists,
    ExcelFileNameAlreadyOpen,
    ExcelFileObjectNameAlreadyOpen,

    ExcelFileObjectNameSyntaxWrong,

    UnableFindSheetByNum,
    SheetNumValueWrong,
    FirstDatarowNumLineValueWrong,

    FileNotFound,

    //-----Excel errors
    ExcelError,

    ExcelUnableOpenFile,
    ExcelUnableCloseFile,
    ExcelUnableGetSheet,

    ExcelUnableSetCellValue,

    ExcelCellIsEmpty,

    ExcelCellTypeNotManaged,

    ExcelCellNotUnique,
    ExcelFuncNameIsEmpty,
    ExcelColNameIsWrong,
    ExcelFuncNameIsWrong,

    ExcelWrongListCols,
    ExcelColNameIsEmpty,

    /// <summary>
    /// All cells of a row are empty.
    /// </summary>
    ExcelRowIsEmpty,

    AtLeastOneInstrThenExpected,
    AtLeastOneInstrIfColThenExpected,

    InstrNotAllowed,
    IfConditionInstrNotAllowed,
    ThenConditionInstrNotAllowed,
    ExecIfCondTypeMismatch,

    CompOperatorIsEmpty,
    CompOperatorIsWrong,

    ItemsShouldBeNotNullAndUnique,

    UnableCreateInstrNotInStageBuild,

    CellTypeStringExpected,
    EntryListIsEmpty,

    ProgramWrongName,
    ProgramNameAlreadyUsed,
    ProgramNotFound,
    NoCurrentProgramExist,

    LoadScriptLinesNull,
    LoadScriptLinesEmpty,
    LoadScriptFileException,
    LoadScriptFileEmpty,

    //--lexical analyzer error codes
    LexerFoundDoubleWrong,

    LexerFoundSgtringBadFormatted,
    LexerFoundCharUndefined,
    LexerFoundSystNameWrong,

    //--syntax analyzer error codes
    ParserCaseNotManaged,

    ParserTokenNotExpected,
    ParserTokenExpected,

    ParserFctNotManaged,
    ParserFctNameExpected,
    ParserFctNameNotExpected,
    ParserFctResultNotSet,
    ParserFctParamWrong,
    ParserFctParamCountWrong,
    ParserFctParamTypeWrong,
    ParserFctParamVarNotDefined,

    ParserValueStringExpected,
    ParserValueIntExpected,
    ParserValueIntWrong,
    ParserVarNotDefined,
    ParserVarOrFctNameNotDefined,
    ParserVarWrongRightPart,

    ParserOnExcelExpected,
    ParserOnSheetExpected,
    ParserTokenDotExpected,
    ParserColAddressExpected,
    ParserSepComparatorExpected,
    ParserTokenIfExpected,
    ParserTokenThenExpected,

    ParserColNumWrong,
    ParserSepComparatorWrong,
    ParserCompExprWrong,

    ParserReturnTypeWrong,
    ParserThenPartIsEmpty,

    ParserBoolExprMixAndOrNotAllowed,
    ParserBoolExprWrong,

    ParserExpressionWrong,

    //--run program, instructions
    ExecUnableOpenFile,

    ExecUnableCloseFile,

    ExecInstrNotManaged,

    ExecInstrValueExpected,

    ExecInstrVarNotFound,
    ExecInstrVarTypeNotExpected,

    ExecInstrTypeStringExpected,
    ExecValueStringExpected,
    ExecValueIntExpected,

    ExecValueIntWrong,
    ExecValueDateWrong,

    ExecInstrMissing,
    ExecInstrAccessFileWrong,
    ExecInstrFilenameWrong,
    ExecInstrFilePathWrong,

    ExecFuncNotExists,

    ExecFuncOneParamExpected,

    ExecNoFileSelected,
    ExecUnableCreateUpdateVar,

    InstrNotExpected,
    IntMustBePositive,
}