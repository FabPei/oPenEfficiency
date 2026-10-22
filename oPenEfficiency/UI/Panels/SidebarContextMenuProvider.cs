using System;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Dialogs;

namespace oPenEfficiency.UI.Panels
{
    internal sealed class SidebarContextMenuProvider
    {
        private readonly Func<ButtonBase, bool> _attachSpecializedMenu;
        private readonly Action<ButtonBase> _attachCatalogMenu;

        public SidebarContextMenuProvider(
            Func<ButtonBase, bool> attachSpecializedMenu,
            Action<ButtonBase> attachCatalogMenu)
        {
            _attachSpecializedMenu = attachSpecializedMenu ??
                throw new ArgumentNullException(nameof(attachSpecializedMenu));
            _attachCatalogMenu = attachCatalogMenu ??
                throw new ArgumentNullException(nameof(attachCatalogMenu));
        }

        public void Attach(ButtonBase button)
        {
            if (button == null) throw new ArgumentNullException(nameof(button));
            button.ContextMenu = new ContextMenu();
            AddShortcutEntry(button);
            if (!_attachSpecializedMenu(button)) _attachCatalogMenu(button);
        }

        private static void AddShortcutEntry(ButtonBase button)
        {
            bool supportsShortcut = FeatureLibrary.Catalog.Any(record =>
                record.SupportsUserShortcut &&
                string.Equals(record.Id, button.Name, StringComparison.Ordinal));
            if (!supportsShortcut && ShortcutManager.GetAssignment(button.Name) == null) return;

            var item = new MenuItem { Header = "Configure Shortcut" };
            item.Click += (sender, args) =>
            {
                new ShortcutConfigWindow(button.Name).ShowDialog();
            };
            button.ContextMenu.Items.Add(item);
        }
    }
}
