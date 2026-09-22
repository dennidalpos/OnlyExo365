using System.Collections.ObjectModel;

namespace OnlyExo365.Shell.Helpers;

/// <summary>ObservableCollection extension methods optimized for UI stability.</summary>
public static class CollectionExtensions
{
    /// <summary>Synchronizes collection with new items by key, minimizing UI flicker.</summary>
    public static void SyncWith<T, TKey>(
        this ObservableCollection<T> collection,
        IEnumerable<T> newItems,
        Func<T, TKey> keySelector) where TKey : notnull
    {
        var newItemsList = newItems.ToList();
        var newKeys = newItemsList.Select(keySelector).ToHashSet();
        var existingKeys = collection.Select(keySelector).ToHashSet();

        for (int i = collection.Count - 1; i >= 0; i--)
        {
            var key = keySelector(collection[i]);
            if (!newKeys.Contains(key))
            {
                collection.RemoveAt(i);
            }
        }

        var existingKeysAfterRemoval = collection.Select(keySelector).ToHashSet();
        foreach (var item in newItemsList)
        {
            var key = keySelector(item);
            if (!existingKeysAfterRemoval.Contains(key))
            {
                collection.Add(item);
            }
        }
    }

    /// <summary>Replaces items using count heuristic or in-place update.</summary>
    public static void ReplaceAll<T>(this ObservableCollection<T> collection, IEnumerable<T> newItems)
    {
        var newItemsList = newItems.ToList();

        if (collection.Count == 0)
        {
            foreach (var item in newItemsList)
            {
                collection.Add(item);
            }
            return;
        }

        if (newItemsList.Count == 0)
        {
            collection.Clear();
            return;
        }

        // Clear and add if counts differ significantly (faster than per-element update)
        if (Math.Abs(collection.Count - newItemsList.Count) > collection.Count / 2)
        {
            collection.Clear();
            foreach (var item in newItemsList)
            {
                collection.Add(item);
            }
            return;
        }

        // In-place update and size alignment
        while (collection.Count > newItemsList.Count)
        {
            collection.RemoveAt(collection.Count - 1);
        }

        for (int i = 0; i < newItemsList.Count; i++)
        {
            if (i < collection.Count)
            {
                collection[i] = newItemsList[i];
            }
            else
            {
                collection.Add(newItemsList[i]);
            }
        }
    }

    /// <summary>Adds a sequence of items to the collection.</summary>
    public static void AddRange<T>(this ObservableCollection<T> collection, IEnumerable<T> items)
    {
        foreach (var item in items)
        {
            collection.Add(item);
        }
    }
}
