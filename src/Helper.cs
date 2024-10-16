using System.IO;

public static class Helper
{
    /// <summary>
    /// Checks if a file is currently open by another process.
    /// </summary>
    /// <param name="filePath">The path of the file to check.</param>
    /// <returns>True if the file is open, false otherwise.</returns>
    public static bool IsFileOpen(string filePath)
    {
        try
        {
            // Try to open the file with read and write permissions
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                fileStream.Close();
            }
        }
        catch (IOException)
        {
            // If an IOException is caught, the file is currently open
            return true;
        }

        return false;
    }
}
