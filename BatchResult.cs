namespace OfficeLegacyConverter;

/// <summary>
/// 整批轉換的結果摘要：總數與失敗清單（單檔失敗時記錄、不中斷整批）。
/// </summary>
internal sealed class BatchResult
{
    public int Total { get; init; }

    public List<string> Failures { get; } = new();

    public int FailureCount => Failures.Count;

    public int SuccessCount => Total - FailureCount;
}
