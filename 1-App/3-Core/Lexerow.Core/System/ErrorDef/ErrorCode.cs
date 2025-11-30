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
    ExecIfCondTypeMismatch,

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
    ParserReturnTypeWrong,
    ParserThenPartIsEmpty,

    //--run program, instructions
    ExecInstrNotManaged,

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

    InstrNotExpected,
    IntMustBePositive
}