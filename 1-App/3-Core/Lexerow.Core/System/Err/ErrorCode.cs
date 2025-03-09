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

    NameNullOrEmpty,
    FileNameNullOrEmpty,
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
    ExcelUnableGetSheet,

    ExcelCellNotUnique,

    ExcelWrongListCols,

    /// <summary>
    /// All cells of a row are empty.
    /// </summary>
    ExcelRowIsEmpty,

    AtLeastOneInstrThenExpected,
    AtLeastOneInstrIfColThenExpected,

    IfConditionInstrNotAllowed,
    ThenConditionInstrNotAllowed,


    UnableCreateInstrNotInStageBuild,
}
