namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TUserId"></typeparam>
    public interface IModificationAuditedOfTUserId<TUserId> : IHasModificationTime
    {
        /// <summary>
        /// Last modifier user for this entity.
        /// </summary>
        TUserId LastModifierUserId { get; set; }
    }
}