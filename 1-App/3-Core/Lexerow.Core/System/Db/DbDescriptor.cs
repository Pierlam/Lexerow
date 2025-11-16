namespace Lexerow.Core.System;

/// <summary>
/// Application database descriptor.
/// !!The structure can't be changed!!
/// To be able to read all old versions of repository data.
/// </summary>
public class DbDescriptor
{
    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="version"></param>
    //public DbDescriptor(string guid, int version, DateOnly dtCreation, string comment)
    //{
    //    Guid = guid;
    //    Version = version;
    //    DateCreation = dtCreation;

    //    if (comment == null)
    //        comment = string.Empty;
    //    Comment = comment;
    //}

    /// <summary>
    /// Internal/technical id.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// This guid should never change, identify the db of the application.
    /// </summary>
    public string Guid { get; set; }

    /// <summary>
    /// Current/last Version of the db.
    /// Very important information.
    /// </summary>
    public int Version { get; set; }

    /// <summary>
    /// dateTime creation of the db.
    /// </summary>
    public DateTime DateCreation { get; set; }

    public string Comment { get; set; } = string.Empty;
}