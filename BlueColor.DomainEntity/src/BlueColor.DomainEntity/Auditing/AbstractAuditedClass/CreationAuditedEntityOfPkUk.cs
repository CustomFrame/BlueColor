using System;

namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TPrimaryKey"></typeparam>
    /// <typeparam name="TUserId"></typeparam>
    [Serializable]
    public abstract class CreationAuditedEntityOfPkUk<TPrimaryKey, TUserId> : Entity<TPrimaryKey>, ICreationAuditedOfTUserId<TUserId>
    {
        /// <summary>
        /// Creation time of this entity.
        /// </summary>
        public virtual DateTime CreationTime { get; set; }

        /// <summary>
        /// Creator of this entity.
        /// </summary>
        public virtual TUserId CreatorUserId { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected CreationAuditedEntityOfPkUk()
        {
            //CreationTime = Clock.Now;
        }
    }
}