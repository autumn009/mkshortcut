using System.Runtime.InteropServices;

if( args.Length != 2 )
{
    Console.WriteLine("Usage: mkshortcut <shortcut_path> <target_path>");
    return;
}
create(args[0], args[1]);

void create(string lnkPath,string fullPath)
{
    dynamic? shell = null;   // IWshRuntimeLibrary.WshShell
    dynamic? lnk = null;     // IWshRuntimeLibrary.IWshShortcut
    try
    {
#pragma warning disable CA1416 // プラットフォームの互換性を検証
        // available in Windows only
        var type = Type.GetTypeFromProgID("WScript.Shell");
#pragma warning restore CA1416 // プラットフォームの互換性を検証
        if (type != null)
        {
            shell = Activator.CreateInstance(type);
            if (shell != null)
            {
                lnk = shell.CreateShortcut(lnkPath);
                if (lnk != null)
                {
                    lnk.TargetPath = fullPath;
                    lnk.Save();
                    Console.WriteLine($"created {lnkPath} to {fullPath}");
                }
            }
        }
    }
    finally
    {
        if (lnk != null) Marshal.ReleaseComObject(lnk);
        if (shell != null) Marshal.ReleaseComObject(shell);
    }
}
