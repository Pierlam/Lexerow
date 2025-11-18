namespace Lexerow.Core.System;

public enum ErrorType
{
    Repository,
    Design,
    Execution
    //License
}

/// <summary>
/// TODO: garder??
/// </summary>
public enum ErrorParamKey
{
    TypeExecResult,
    WrongPos
}

/// <summary>
/// An error parameter.
/// </summary>
public class ErrorParam
{
    public ErrorParamKey Key { get; set; }
    public string Value { get; set; }
}

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
    ExcelCellIsEmpty,

    ExcelCellTypeNotManaged,

    ExcelUnableOpenFile,
    ExcelUnableCloseFile,
    ExcelUnableGetSheet,

    ExcelCellNotUnique,
    ExcelFuncNameIsEmpty,
    ExcelColNameIsWrong,
    ExcelFuncNameIsWrong,

    CompOperatorIsEmpty,
    CompOperatorIsWrong,

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
    IfCondTypeMismatch,

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

    //--syntax analyzer error codes
    ParserCaseNotManaged,

    ParserTokenNotExpected,
    ParserTokenExpected,

    ParserFctNameExpected,
    ParserFctParamCountWrong,
    ParserFctParamTypeWrong,
    ParserFctParamVarNotDefined,
    ParserSetVarWrongRightPart,
    ParserFctResultNotSet,
    ParserConstStringValueExpected,
    ParserConstIntValueExpected,

    ParserConstIntValueWrong,

    ParserOnSheetExpected,
    ParserTokenDotExpected,
    ParserColAddressExpected,
    ParserColNumWrong,
    ParserSepComparatorExpected,
    ParserSepComparatorWrong,
    ParserReturnTypeWrong,
    ParserTokenIfExpected,
    ParserTokenThenExpected,
    ParserThenPartIsEmpty,

    //--run program, instructions
    ExecInstrNotManaged,

    ExecInstrTypeStringExpected,
    ExecInstrVarNotFound,
    ExecInstrVarTypeNotExpected,
    ExecInstrMissing,
    ExecInstrAccessFileWrong,
    ExecInstrFilenameWrong,
    ExecInstrFilePathWrong,

    ExecFuncNotExists,

    ExecFuncOneParamExpected,

    InstrNotExpected,
    IntMustBePositive
}