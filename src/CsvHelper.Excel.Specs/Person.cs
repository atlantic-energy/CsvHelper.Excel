namespace CsvHelper.Excel.Specs;

public class Person
{
    public string Name { get; set; } = "";

    public int Age { get; set; }
}

/// <summary>A folder for one test's files, deleted when the test is done.</summary>
public sealed class ScratchFolder : IDisposable
{
    public ScratchFolder()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "CsvHelper.Excel.Specs", Guid.NewGuid().ToString("n"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        if (Directory.Exists(Path))
        {
            Directory.Delete(Path, recursive: true);
        }
    }
}
