using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Asset Manager guards (QA9-BtnAssetManager-002): the library
    /// containment rule that gates writes and deletes, and the folder-name sanitization, used by
    /// LocalAssetService.
    /// </summary>
    [TestClass]
    public sealed class AssetLibraryPolicyTests
    {
        private static string Root => Path.Combine(Path.GetTempPath(), "oPE-AssetRoot");
        private static string Sep => Path.DirectorySeparatorChar.ToString();

        [TestMethod]
        public void ContainmentAcceptsTheRootItselfAndItsChildren()
        {
            Assert.IsTrue(AssetLibraryPolicy.IsWithinRoot(Root, Root), "The root itself is inside the library.");
            Assert.IsTrue(AssetLibraryPolicy.IsWithinRoot(Root, Root + Sep + "Slides"),
                "A direct child is inside the library.");
            Assert.IsTrue(AssetLibraryPolicy.IsWithinRoot(Root + Sep, Root + Sep + "Slides" + Sep + "Deck"),
                "Trailing separators on the root do not matter.");
        }

        [TestMethod]
        public void ContainmentRejectsSiblingsSharingTheRootPrefix()
        {
            string sibling = Root + "2";
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, sibling),
                "A sibling folder that merely shares the root's prefix (C:\\Root2 vs C:\\Root) is outside.");
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, sibling + Sep + "file.png"),
                "Children of such siblings are outside as well.");
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, Path.GetTempPath()),
                "An unrelated ancestor directory is outside.");
        }

        [TestMethod]
        public void ContainmentRejectsParentTraversalOutsideTheRoot()
        {
            string escaping = Root + Sep + "Slides" + Sep + ".." + Sep + ".." + Sep + "escape.png";
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, escaping),
                "Parent traversal that normalizes outside the root is rejected.");

            string internalPath = Root + Sep + "Slides" + Sep + ".." + Sep + "Deck";
            Assert.IsTrue(AssetLibraryPolicy.IsWithinRoot(Root, internalPath),
                "Traversal that normalizes back inside the root stays allowed.");

            string siblingViaTraversal = Root + Sep + ".." + Sep + "oPE-AssetRoot2" + Sep + "file.png";
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, siblingViaTraversal),
                "Traversal into a prefix-sharing sibling is rejected.");
        }

        [TestMethod]
        public void ContainmentIsCaseInsensitiveAndRejectsBlankOrInvalidInput()
        {
            Assert.IsTrue(AssetLibraryPolicy.IsWithinRoot(Root, Root.ToUpperInvariant()),
                "The containment check is case-insensitive (Windows paths).");
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, null));
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(Root, "   "));
            Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(null, Root));
        }

        [TestMethod]
        public void FolderNamesAreSanitizedToSafeUniqueishValues()
        {
            Assert.AreEqual("unnamed", AssetLibraryPolicy.MakeValidFolderName(null));
            Assert.AreEqual("unnamed", AssetLibraryPolicy.MakeValidFolderName("   "));
            Assert.AreEqual("My_Folder", AssetLibraryPolicy.MakeValidFolderName(" My Folder "));
            Assert.AreEqual("abcd", AssetLibraryPolicy.MakeValidFolderName("a:b*c\u002Fd"),
                "Invalid characters are removed, not replaced.");

            string longName = new string('x', 150);
            Assert.AreEqual(100, AssetLibraryPolicy.MakeValidFolderName(longName).Length,
                "Names are capped at 100 characters.");
        }

        [TestMethod]
        public void NewFolderDestinationsPreserveNestedParentsAndOrdinarySanitization()
        {
            string parent = Path.Combine(Root, "Slides", "Client");
            Assert.AreEqual(Path.Combine(parent, "My_Folder"),
                AssetLibraryPolicy.ResolveNewFolderPath(Root, parent, " My Folder "));
            Assert.AreEqual(Path.Combine(parent, "abcd"),
                AssetLibraryPolicy.ResolveNewFolderPath(Root, parent, "ab*c?d"));
            Assert.AreEqual(Path.Combine(parent, "unnamed"),
                AssetLibraryPolicy.ResolveNewFolderPath(Root, parent, null));
        }

        [DataTestMethod]
        [DataRow(".")]
        [DataRow("..")]
        [DataRow("...")]
        [DataRow(" .. ")]
        [DataRow("name.")]
        [DataRow(@"..\outside")]
        [DataRow("../outside")]
        [DataRow(@"nested\child")]
        [DataRow(@"C:\outside")]
        [DataRow(@"C:outside")]
        [DataRow(@"\outside")]
        [DataRow(@"\\server\share\outside")]
        [DataRow(@"\\?\C:\outside")]
        public void NewFolderDestinationsRejectPathsAndWindowsDotAliases(string name)
        {
            Assert.ThrowsException<ArgumentException>(() =>
                AssetLibraryPolicy.ResolveNewFolderPath(Root, Path.Combine(Root, "Slides"), name));
        }

        [TestMethod]
        public void NewFolderDestinationsRejectOutsideAndPrefixSiblingParents()
        {
            foreach (var parent in new[] { Path.GetTempPath(), Root + "2", Path.Combine(Root, "..", "outside") })
                Assert.ThrowsException<UnauthorizedAccessException>(() =>
                    AssetLibraryPolicy.ResolveNewFolderPath(Root, parent, "Child"));
        }

        [TestMethod]
        public void ValidatedDestinationCreatesOnlyTheRequestedChild()
        {
            string workspace = Path.Combine(Path.GetTempPath(), "oPE-folder-test-" + Guid.NewGuid().ToString("N"));
            string library = Path.Combine(workspace, "library");
            string parent = Path.Combine(library, "Slides", "Client");
            Directory.CreateDirectory(parent);
            try
            {
                // The old Favorites combine allowed this target outside the library.
                string escaped = Path.GetFullPath(Path.Combine(library, @"..\outside"));
                Assert.IsFalse(AssetLibraryPolicy.IsWithinRoot(library, escaped));
                Assert.ThrowsException<ArgumentException>(() =>
                    AssetLibraryPolicy.ResolveNewFolderPath(library, library, @"..\outside"));
                Assert.IsFalse(Directory.Exists(escaped));

                string destination = AssetLibraryPolicy.ResolveNewFolderPath(library, parent, "Client Notes");
                Directory.CreateDirectory(destination);
                Assert.IsTrue(Directory.Exists(Path.Combine(parent, "Client_Notes")));
                CollectionAssert.AreEqual(new[] { destination }, Directory.GetDirectories(parent));
                Assert.IsFalse(Directory.Exists(escaped));
            }
            finally
            {
                // workspace is an exact GUID child of the OS temporary directory.
                Directory.Delete(workspace, true);
            }
        }
    }
}
