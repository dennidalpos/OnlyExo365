namespace ExchangeAdmin.Domain.DeltaPlan;

/// <summary>
/// Tipo di azione nel piano delta.
/// </summary>
public enum DeltaActionType
{
    Add,
    Remove,
    Modify,
    NoChange
}

/// <summary>
/// Singola azione nel piano delta.
/// </summary>
public class DeltaAction<T>
{
    public DeltaActionType ActionType { get; init; }
    public T? Item { get; init; }
    public string? Key { get; init; }
    public string? Description { get; init; }

    /// <summary>
    /// Per Modify: valore corrente.
    /// </summary>
    public object? CurrentValue { get; init; }

    /// <summary>
    /// Per Modify: valore desiderato.
    /// </summary>
    public object? DesiredValue { get; init; }
}

/// <summary>
/// Piano delta risultante dal confronto desired vs current state.
/// </summary>
public class DeltaPlan<T>
{
    public List<DeltaAction<T>> Actions { get; init; } = new();

    public int TotalActions => Actions.Count;
    public int AddCount => Actions.Count(a => a.ActionType == DeltaActionType.Add);
    public int RemoveCount => Actions.Count(a => a.ActionType == DeltaActionType.Remove);
    public int ModifyCount => Actions.Count(a => a.ActionType == DeltaActionType.Modify);

    public bool HasChanges => Actions.Any(a => a.ActionType != DeltaActionType.NoChange);
}

/// <summary>
/// Calcolatore di piano delta generico.
/// </summary>
public static class DeltaPlanCalculator
{
    /// <summary>
    /// Calcola il piano delta tra stato desiderato e corrente.
    /// </summary>
    /// <typeparam name="T">Tipo degli elementi.</typeparam>
    /// <param name="desired">Stato desiderato.</param>
    /// <param name="current">Stato corrente.</param>
    /// <param name="keySelector">Selettore chiave univoca.</param>
    /// <param name="equalityComparer">Comparatore per determinare se un elemento è modificato.</param>
    public static DeltaPlan<T> Calculate<T>(
        IEnumerable<T> desired,
        IEnumerable<T> current,
        Func<T, string> keySelector,
        Func<T, T, bool>? equalityComparer = null)
    {
        var actions = new List<DeltaAction<T>>();

        var desiredDict = desired.ToDictionary(keySelector);
        var currentDict = current.ToDictionary(keySelector);

        var comparer = equalityComparer ?? ((a, b) => EqualityComparer<T>.Default.Equals(a, b));

        // Trova elementi da aggiungere o modificare
        foreach (var (key, desiredItem) in desiredDict)
        {
            if (currentDict.TryGetValue(key, out var currentItem))
            {
                // Esiste in entrambi - verifica se modificato
                if (!comparer(desiredItem, currentItem))
                {
                    actions.Add(new DeltaAction<T>
                    {
                        ActionType = DeltaActionType.Modify,
                        Item = desiredItem,
                        Key = key,
                        CurrentValue = currentItem,
                        DesiredValue = desiredItem,
                        Description = $"Modify {key}"
                    });
                }
                else
                {
                    actions.Add(new DeltaAction<T>
                    {
                        ActionType = DeltaActionType.NoChange,
                        Item = desiredItem,
                        Key = key
                    });
                }
            }
            else
            {
                // Non esiste nel corrente - da aggiungere
                actions.Add(new DeltaAction<T>
                {
                    ActionType = DeltaActionType.Add,
                    Item = desiredItem,
                    Key = key,
                    Description = $"Add {key}"
                });
            }
        }

        // Trova elementi da rimuovere
        foreach (var (key, currentItem) in currentDict)
        {
            if (!desiredDict.ContainsKey(key))
            {
                actions.Add(new DeltaAction<T>
                {
                    ActionType = DeltaActionType.Remove,
                    Item = currentItem,
                    Key = key,
                    Description = $"Remove {key}"
                });
            }
        }

        return new DeltaPlan<T> { Actions = actions };
    }
}
