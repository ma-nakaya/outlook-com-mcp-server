using System.Runtime.InteropServices;

namespace OutlookComMcp.Outlook;

internal static class ComObject
{
    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }
    }
}

