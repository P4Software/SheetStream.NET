using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Operations;

namespace SheetStream.Analyzers
{
    /// <summary>
    /// Decides whether one operation is guaranteed to run before another inside a single method body, using
    /// lexical order plus the operation tree instead of a full control-flow graph. The test is deliberately
    /// conservative: when a branch, loop, lambda or similar sits between the first operation and the point where
    /// the two paths meet, the answer is "not sure", and the analyzer stays silent. False negatives are
    /// acceptable; false positives are not.
    /// </summary>
    internal static class OperationOrder
    {
        /// <summary>
        /// True when <paramref name="first"/> has certainly executed by the time <paramref name="second"/> runs:
        /// it appears earlier in the source, and nothing between it and the common ancestor of the two can skip
        /// or defer it. <paramref name="second"/> itself may sit inside a branch or a lambda; if that code ever
        /// runs, <paramref name="first"/> already ran.
        /// </summary>
        public static bool RunsBefore(IOperation first, IOperation second)
        {
            if (first.Syntax.SyntaxTree != second.Syntax.SyntaxTree) return false;
            if (first.Syntax.SpanStart >= second.Syntax.SpanStart) return false;

            var secondAncestors = new HashSet<IOperation>();
            for (var a = second.Parent; a != null; a = a.Parent) secondAncestors.Add(a);

            for (var a = first.Parent; a != null; a = a.Parent)
            {
                // The common ancestor decides the final verdict: a block (or an expression that evaluates its
                // children left to right) keeps the lexical order; an if/loop/try/switch may run only one child.
                if (secondAncestors.Contains(a)) return !MayNotRunChildren(a);
                if (MayNotRunChildren(a)) return false;
            }
            return false; // Different operation blocks (field initialisers vs. body): no ordering claim.
        }

        /// <summary>
        /// Constructs whose children do not simply run in order: branches, loops (zero iterations), exception
        /// handling, deferred bodies (lambdas, local functions) and short-circuiting expressions.
        /// </summary>
        private static bool MayNotRunChildren(IOperation operation)
        {
            switch (operation)
            {
                case IConditionalOperation _:
                case ILoopOperation _:
                case ISwitchOperation _:
                case ISwitchExpressionOperation _:
                case ITryOperation _:
                case ICatchClauseOperation _:
                case IAnonymousFunctionOperation _:
                case ILocalFunctionOperation _:
                case IConditionalAccessOperation _:
                case ICoalesceOperation _:
                case ICoalesceAssignmentOperation _:
                case ILabeledOperation _:
                    return true;
                case IBinaryOperation binary:
                    return binary.OperatorKind == BinaryOperatorKind.ConditionalAnd || binary.OperatorKind == BinaryOperatorKind.ConditionalOr;
                default:
                    return false;
            }
        }
    }
}
