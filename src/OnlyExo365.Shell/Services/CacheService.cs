using System.Collections.Concurrent;

namespace OnlyExo365.Shell.Services;

/// <summary>In-memory cache with TTL support for dashboard stats and policies.</summary>
public sealed class CacheService
{
    private readonly ConcurrentDictionary<string, CacheEntry> _cache = new();

    private static readonly TimeSpan DefaultTtl = TimeSpan.FromMinutes(5);

    public static class Keys
    {
        public const string DashboardStats = "dashboard_stats";
        public const string RetentionPolicies = "retention_policies";
    }

    /// <summary>Gets cached value if present and unexpired.</summary>
    public T? Get<T>(string key) where T : class
    {
        if (_cache.TryGetValue(key, out var entry) && !entry.IsExpired)
        {
            return entry.Value as T;
        }

        if (entry?.IsExpired == true)
        {
            _cache.TryRemove(key, out _);
        }

        return null;
    }

    /// <summary>Sets cached value with optional TTL.</summary>
    public void Set<T>(string key, T value, TimeSpan? ttl = null) where T : class
    {
        var expiration = DateTime.UtcNow + (ttl ?? DefaultTtl);
        _cache[key] = new CacheEntry(value, expiration);
    }

    /// <summary>Invalidates specified cache key.</summary>
    public void Invalidate(string key)
    {
        _cache.TryRemove(key, out _);
    }

    /// <summary>Invalidates all cached entries.</summary>
    public void InvalidateAll()
    {
        _cache.Clear();
    }

    /// <summary>Gets cached value or invokes fetch factory.</summary>
    public async Task<T?> GetOrFetchAsync<T>(
        string key,
        Func<Task<T?>> fetchFunc,
        TimeSpan? ttl = null,
        bool forceRefresh = false) where T : class
    {
        if (!forceRefresh)
        {
            var cached = Get<T>(key);
            if (cached != null)
            {
                return cached;
            }
        }

        var value = await fetchFunc();
        if (value != null)
        {
            Set(key, value, ttl);
        }

        return value;
    }

    private sealed class CacheEntry
    {
        public object Value { get; }
        public DateTime Expiration { get; }
        public bool IsExpired => DateTime.UtcNow >= Expiration;

        public CacheEntry(object value, DateTime expiration)
        {
            Value = value;
            Expiration = expiration;
        }
    }
}
