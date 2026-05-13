using System.Collections.Concurrent;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Services;

public interface IGraphSubscriptionStore
{
    IReadOnlyCollection<GraphSubscriptionRecord> List();
    bool TryGet(string id, out GraphSubscriptionRecord? record);
    void Upsert(GraphSubscriptionRecord record);
    bool Remove(string id, out GraphSubscriptionRecord? record);
}

public sealed class GraphSubscriptionStore : IGraphSubscriptionStore
{
    private readonly ConcurrentDictionary<string, GraphSubscriptionRecord> _subscriptions = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyCollection<GraphSubscriptionRecord> List() => _subscriptions.Values.ToArray();

    public bool TryGet(string id, out GraphSubscriptionRecord? record) => _subscriptions.TryGetValue(id, out record);

    public void Upsert(GraphSubscriptionRecord record) => _subscriptions.AddOrUpdate(record.Id, record, (_, _) => record);

    public bool Remove(string id, out GraphSubscriptionRecord? record) => _subscriptions.TryRemove(id, out record);
}
