namespace BlueColor.DomainEntity.Auditing
{
    /// <summary>
    ///
    /// </summary>
    /// <typeparam name="TUserId"></typeparam>
    public interface IFullAuditedOfTUserId<TUserId> : IAuditedOfTUserId<TUserId>, IDeletionAuditedOfTUserId<TUserId>
    {
    }
}