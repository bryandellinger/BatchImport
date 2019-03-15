
using System.Diagnostics;


namespace ADBatchImportForInterProd
{
    public interface IBatchService
    {
        string[] GetColumnNames();
        void WriteToEventLog(string message);
        void ExecuteNonQuery(string v);
        string GetDBValue(string source);
        string getCWOPAInsertStatement();
        string getElapsedTime(Stopwatch stopWatch);
    }
}
