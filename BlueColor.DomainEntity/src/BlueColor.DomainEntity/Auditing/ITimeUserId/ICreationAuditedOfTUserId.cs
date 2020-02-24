namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TUserId"></typeparam>
    public interface ICreationAuditedOfTUserId<TUserId> : IHasCreationTime
    {
        /// <summary>
        /// Id of the creator user of this entity.
        /// </summary>
        TUserId CreatorUserId { get; set; }
    }
}