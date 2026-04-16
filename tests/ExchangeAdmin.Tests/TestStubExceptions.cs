using System.Runtime.CompilerServices;

namespace ExchangeAdmin.Tests;

internal static class TestStubExceptions
{
    public static NotSupportedException CreateUnsupported([CallerMemberName] string? memberName = null)
        => new($"{memberName ?? "Operation"} is not implemented for this test stub.");
}
