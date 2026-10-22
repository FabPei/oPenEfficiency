using System;

namespace oPenEfficiency.DocumentAutomation
{
    public static class OneBasedCollectionTraversal
    {
        public static void Forward(int count, Action<int> visit)
        {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (visit == null) throw new ArgumentNullException(nameof(visit));
            for (int index = 1; index <= count; index++) visit(index);
        }

        public static void Reverse(int count, Action<int> visit)
        {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (visit == null) throw new ArgumentNullException(nameof(visit));
            for (int index = count; index >= 1; index--) visit(index);
        }

        public static void VisitLeaves<T>(T root, Func<T, int> childCount, Func<T, int, T> childAt, Action<T> visitLeaf)
        {
            if (childCount == null) throw new ArgumentNullException(nameof(childCount));
            if (childAt == null) throw new ArgumentNullException(nameof(childAt));
            if (visitLeaf == null) throw new ArgumentNullException(nameof(visitLeaf));
            int count = childCount(root);
            if (count == 0) { visitLeaf(root); return; }
            Forward(count, index => VisitLeaves(childAt(root, index), childCount, childAt, visitLeaf));
        }
    }
}
