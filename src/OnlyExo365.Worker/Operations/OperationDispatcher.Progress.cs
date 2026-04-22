namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private static string FormatListProgressStatus(int current, int total, string analyzedLabel)
    {
        var safeTotal = Math.Max(total, current);
        var remaining = Math.Max(safeTotal - current, 0);
        return $"{analyzedLabel} {current}/{safeTotal} (rimanenti {remaining})";
    }
}

