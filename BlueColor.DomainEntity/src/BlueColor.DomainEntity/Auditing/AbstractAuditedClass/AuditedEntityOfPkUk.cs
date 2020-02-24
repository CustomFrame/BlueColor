using System;

namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TPrimaryKey"></typeparam>
    /// <typeparam name="TUserId"></typeparam>
    [Serializable]
    public abstract class AuditedEntityOfPkUk<TPrimaryKey, TUserId> : CreationAuditedEntityOfPkUk<TPrimaryKey, TUserId>, IAuditedOfTUserId<TUserId>
    {
        /// <summary>
        /// Last modification date of this entity.
        /// </summary>
        public virtual DateTime? LastModificationTime { get; set; }

        /// <summary>
        /// Last modifier user of this entity.
        /// </summary>
        public virtual TUserId LastModifierUserId { get; set; }
    }
}