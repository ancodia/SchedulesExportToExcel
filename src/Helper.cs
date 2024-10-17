using System.IO;

public static class Helper
{
    public static bool IsFileOpen(string filePath)
    {
        try
        {
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                fileStream.Close();
            }
        }
        catch (IOException)
        {
            return true;
        }

        return false;
    }
}
