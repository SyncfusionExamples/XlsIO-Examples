using System.Data;

interface IExcelEngine
{
    string Name { get; }
    void Run(string scenario, string sourcePath, DataTable table, string outputStem);
}

