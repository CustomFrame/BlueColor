using System;

namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TPrimaryKey"></typeparam>
    /// <typeparam name="TUserId"></typeparam>
    [Serializable]
    public abstract class FullAuditedEntityOfPkUk<TPrimaryKey, TUserId> : AuditedEntityOfPkUk<TPrimaryKey, TUserId>, IFullAuditedOfTUserId<TUserId>
    {
        /// <summary>
        /// Is this entity Deleted?
        /// </summary>
        public virtual bool IsDeleted { get; set; }

        /// <summary>
        /// Which user deleted this entity?
        /// </summary>
        public virtual TUserId DeleterUserId { get; set; }

        /// <summary>
        /// Deletion time of this entity.
        /// </summary>
        public virtual DateTime? DeletionTime { get; set; }
    }
}