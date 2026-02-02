namespace ExchangeAdmin.Domain.DeltaPlan;

             
                                   
              
public enum DeltaActionType
{
    Add,
    Remove,
    Modify,
    NoChange
}

             
                                   
              
public class DeltaAction<T>
{
    public DeltaActionType ActionType { get; init; }
    public T? Item { get; init; }
    public string? Key { get; init; }
    public string? Description { get; init; }

                 
                                    
                  
    public object? CurrentValue { get; init; }

                 
                                      
                  
    public object? DesiredValue { get; init; }
}

             
                                                                  
              
public class DeltaPlan<T>
{
    public List<DeltaAction<T>> Actions { get; init; } = new();

    public int TotalActions => Actions.Count;
    public int AddCount => Actions.Count(a => a.ActionType == DeltaActionType.Add);
    public int RemoveCount => Actions.Count(a => a.ActionType == DeltaActionType.Remove);
    public int ModifyCount => Actions.Count(a => a.ActionType == DeltaActionType.Modify);

    public bool HasChanges => Actions.Any(a => a.ActionType != DeltaActionType.NoChange);
}

             
                                        
              
public static class DeltaPlanCalculator
{
                 
                                                               
                  
                                                            
                                                       
                                                     
                                                                   
                                                                                                        
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

                                                    
        foreach (var (key, desiredItem) in desiredDict)
        {
            if (currentDict.TryGetValue(key, out var currentItem))
            {
                                                              
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
                                                          
                actions.Add(new DeltaAction<T>
                {
                    ActionType = DeltaActionType.Add,
                    Item = desiredItem,
                    Key = key,
                    Description = $"Add {key}"
                });
            }
        }

                                      
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
