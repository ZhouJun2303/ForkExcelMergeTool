namespace ExcelMergeFork.Api;

public sealed class WorkSessionStore : IDisposable
{
    private readonly object _gate = new();
    private readonly Dictionary<string, MergeWorkSession> _merges = new(StringComparer.Ordinal);
    private readonly Dictionary<string, CompareWorkSession> _compares = new(StringComparer.Ordinal);
    private static readonly TimeSpan Lifetime = TimeSpan.FromHours(6);

    public MergeWorkSession OpenMerge(string local, string basePath, string remote, string? merged, MergeOptionsBody? options, string? mode)
    {
        Sweep();
        var initial = MergeWorkSession.ResolveOptions(MergeWorkSession.FromSettings(), options, mode);
        var session = MergeWorkSession.Open(local, basePath, remote, string.IsNullOrWhiteSpace(merged) ? null : RepoLayout.ResolvePath(merged), initial);
        lock (_gate)
        {
            _merges[session.Id] = session;
        }

        return session;
    }

    public CompareWorkSession OpenCompare(string left, string right, bool includeSame)
    {
        Sweep();
        var session = CompareWorkSession.Open(left, right, includeSame);
        lock (_gate)
        {
            _compares[session.Id] = session;
        }

        return session;
    }

    public T WithMerge<T>(string id, Func<MergeWorkSession, T> action)
    {
        var session = FindMerge(id);
        lock (session.Gate)
        {
            return action(session);
        }
    }

    public T WithCompare<T>(string id, Func<CompareWorkSession, T> action)
    {
        var session = FindCompare(id);
        lock (session.Gate)
        {
            return action(session);
        }
    }

    public void CloseMerge(string id)
    {
        MergeWorkSession? session;
        lock (_gate)
        {
            _merges.Remove(id, out session);
        }

        session?.Dispose();
    }

    public void CloseCompare(string id)
    {
        CompareWorkSession? session;
        lock (_gate)
        {
            _compares.Remove(id, out session);
        }

        session?.Dispose();
    }

    public void Dispose()
    {
        List<IDisposable> sessions;
        lock (_gate)
        {
            sessions = _merges.Values.Cast<IDisposable>().Concat(_compares.Values).ToList();
            _merges.Clear();
            _compares.Clear();
        }

        foreach (var session in sessions)
        {
            session.Dispose();
        }
    }

    private MergeWorkSession FindMerge(string id)
    {
        lock (_gate)
        {
            if (_merges.TryGetValue(id, out var session))
            {
                return session;
            }
        }

        throw new InvalidOperationException("合并会话不存在或已过期");
    }

    private CompareWorkSession FindCompare(string id)
    {
        lock (_gate)
        {
            if (_compares.TryGetValue(id, out var session))
            {
                return session;
            }
        }

        throw new InvalidOperationException("对比会话不存在或已过期");
    }

    private void Sweep()
    {
        var cutoff = DateTime.UtcNow - Lifetime;
        List<IDisposable> expired = [];
        lock (_gate)
        {
            foreach (var pair in _merges.Where(pair => pair.Value.TouchedUtc < cutoff).ToList())
            {
                _merges.Remove(pair.Key);
                expired.Add(pair.Value);
            }

            foreach (var pair in _compares.Where(pair => pair.Value.TouchedUtc < cutoff).ToList())
            {
                _compares.Remove(pair.Key);
                expired.Add(pair.Value);
            }
        }

        foreach (var session in expired)
        {
            session.Dispose();
        }
    }
}
