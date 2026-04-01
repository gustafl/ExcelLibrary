namespace ExcelLibrary;

public class WorkbookOptions
{
    private bool includeHidden = false;
    private bool loadSheets = true;

    public bool IncludeHidden
    {
        get { return this.includeHidden; }
        set { this.includeHidden = value; }
    }

    public bool LoadSheets
    {
        get { return this.loadSheets; }
        set { this.loadSheets = value; }
    }
}
