using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Tests;

public class KeyNormalizerTests
{
    [Theory]
    [InlineData(1, "1")]
    [InlineData(1.0, "1")]
    [InlineData("1", "1")]
    [InlineData("1.0", "1")]
    [InlineData("1.00", "1")]
    [InlineData("01", "01")]
    [InlineData("  k1  ", "k1")]
    [InlineData("", "")]
    public void Normalize_MatchesPythonRules(object value, string expected)
    {
        Assert.Equal(expected, KeyNormalizer.Normalize(value));
    }

    [Fact]
    public void HeaderCompare_IgnoresCaseAndSpaces()
    {
        Assert.Equal(KeyNormalizer.HeaderForCompare("齿轮ID"), KeyNormalizer.HeaderForCompare("齿轮Id"));
    }

    [Fact]
    public void InsertionIndex_GroupsByPrefix()
    {
        var merged = new List<string> { "A-1", "B-1" };
        var idx = KeyNormalizer.InsertionIndex(merged, "A-2");
        Assert.Equal(1, idx);
        var result = KeyNormalizer.MergeOrdered(merged, ["A-2", "B-2"]);
        Assert.Equal(["A-1", "A-2", "B-1", "B-2"], result);
    }
}
