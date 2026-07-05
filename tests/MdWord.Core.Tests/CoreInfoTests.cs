using Xunit;

namespace MdWord.Core.Tests;

public class CoreInfoTests
{
    [Fact]
    public void Name_IsMdWordCore()
    {
        Assert.Equal("MdWord.Core", CoreInfo.Name);
    }
}
