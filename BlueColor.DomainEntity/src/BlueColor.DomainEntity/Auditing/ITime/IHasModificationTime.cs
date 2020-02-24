using System;

namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    public interface IHasModificationTime
    {
        /// <summary>
        /// The last modified time for this entity.
        /// </summary>
        DateTime? LastModificationTime { get; set; }
    }
}