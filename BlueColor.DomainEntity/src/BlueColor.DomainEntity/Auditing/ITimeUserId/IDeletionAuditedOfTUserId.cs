namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TUserId"></typeparam>
    public interface IDeletionAuditedOfTUserId<TUserId> : IHasDeletionTime
    {
        /// <summary>
        /// Which user deleted this entity?
        /// </summary>
        TUserId DeleterUserId { get; set; }
    }
}