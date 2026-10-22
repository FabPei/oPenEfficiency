using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.UI;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public class ExcelLinkManagerSelectionTests
    {
        [TestMethod]
        public void CheckedSelectionNotifiesTheGridBinding()
        {
            var item = new ExcelLinkItem();
            var changed = new List<string>();
            item.PropertyChanged += (sender, args) => changed.Add(args.PropertyName);

            item.IsSelected = true;

            Assert.IsTrue(item.IsSelected);
            CollectionAssert.AreEqual(new[] { nameof(ExcelLinkItem.IsSelected) }, changed);
        }

        [TestMethod]
        public void ReapplyingCheckedSelectionDoesNotRaiseDuplicateNotification()
        {
            var item = new ExcelLinkItem { IsSelected = true };
            int notifications = 0;
            item.PropertyChanged += (sender, args) => notifications++;

            item.IsSelected = true;

            Assert.AreEqual(0, notifications);
        }
    }
}
