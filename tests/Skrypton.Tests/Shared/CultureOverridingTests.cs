using System;
using System.Globalization;
using System.Threading;

namespace Skrypton.Tests.Shared
{
    public abstract class CultureOverridingTests : TestBase
    {
        protected CultureOverridingTests(string cultureName)
        {
            TestCulture = new CultureInfo(cultureName, useUserOverride: false);
        }
    }
}
