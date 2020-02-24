using System;

namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    public interface IHasCreationTime
    {
        /// <summary>
        /// Creation time of this entity.
        /// </summary>
        DateTime CreationTime { get; set; }
    }
}