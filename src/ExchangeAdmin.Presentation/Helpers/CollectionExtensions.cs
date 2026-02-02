using System.Collections.ObjectModel;

namespace ExchangeAdmin.Presentation.Helpers;

/// <summary>
/// Extension methods for ObservableCollection to optimize updates and reduce UI flickering.
/// </summary>
public static class CollectionExtensions
{
    /// <summary>
    /// Synchronizes the ObservableCollection with a new set of items using minimal changes.
    /// This avoids clearing and re-adding all items, which causes UI flickering.
    /// </summary>
    /// <typeparam name="T">The type of items in the collection.</typeparam>
    /// <param name="collection">The observable collection to update.</param>
    /// <param name="newItems">The new items to synchronize with.</param>
    /// <param name="keySelector">Function to extract a unique key for comparison.</param>
    public static void SyncWith<T, TKey>(
        this ObservableCollection<T> collection,
        IEnumerable<T> newItems,
        Func<T, TKey> keySelector) where TKey : notnull
    {
        var newItemsList = newItems.ToList();
        var newKeys = newItemsList.Select(keySelector).ToHashSet();
        var existingKeys = collection.Select(keySelector).ToHashSet();

        // Remove items that are no longer present
        for (int i = collection.Count - 1; i >= 0; i--)
        {
            var key = keySelector(collection[i]);
            if (!newKeys.Contains(key))
            {
                collection.RemoveAt(i);
            }
        }

        // Add new items that don't exist
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

    /// <summary>
    /// Replaces all items in the collection efficiently.
    /// If the new items are the same count and can be updated in place, it does so.
    /// Otherwise performs a smart diff.
    /// </summary>
    /// <typeparam name="T">The type of items in the collection.</typeparam>
    /// <param name="collection">The observable collection to update.</param>
    /// <param name="newItems">The new items to replace with.</param>
    public static void ReplaceAll<T>(this ObservableCollection<T> collection, IEnumerable<T> newItems)
    {
        var newItemsList = newItems.ToList();

        // If collection is empty, just add all
        if (collection.Count == 0)
        {
            foreach (var item in newItemsList)
            {
                collection.Add(item);
            }
            return;
        }

        // If new items are empty, clear
        if (newItemsList.Count == 0)
        {
            collection.Clear();
            return;
        }

        // If counts are very different, just clear and add (faster)
        if (Math.Abs(collection.Count - newItemsList.Count) > collection.Count / 2)
        {
            collection.Clear();
            foreach (var item in newItemsList)
            {
                collection.Add(item);
            }
            return;
        }

        // Smart update: remove extras, add missing
        while (collection.Count > newItemsList.Count)
        {
            collection.RemoveAt(collection.Count - 1);
        }

        for (int i = 0; i < newItemsList.Count; i++)
        {
            if (i < collection.Count)
            {
                // Update existing position
                collection[i] = newItemsList[i];
            }
            else
            {
                // Add new item
                collection.Add(newItemsList[i]);
            }
        }
    }

    /// <summary>
    /// Adds a range of items to the collection.
    /// </summary>
    public static void AddRange<T>(this ObservableCollection<T> collection, IEnumerable<T> items)
    {
        foreach (var item in items)
        {
            collection.Add(item);
        }
    }
}
