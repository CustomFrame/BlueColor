namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TUserId"></typeparam>
    public interface IAuditedOfTUserId<TUserId> : ICreationAuditedOfTUserId<TUserId>, IModificationAuditedOfTUserId<TUserId>
    {
    }
}