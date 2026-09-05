using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services.CustomActions;

namespace OutlookAI.Services.TextEditing
{
    /// <summary>
    /// Selects custom actions that are explicitly scoped to transforming the
    /// current Word-editor selection. Surface, context source and output are
    /// all checked so editor commands cannot leak into assistant menus.
    /// </summary>
    public static class TextActionCatalog
    {
        public const string Surface = "editor_selection";
        public const string Source = "selected_text";
        public const string Output = "replace_selection";

        public static bool IsTextAction(CustomActionDefinition action)
        {
            return action != null
                && action.Context != null
                && EqualsNormalized(action.Surface, Surface)
                && EqualsNormalized(action.Context.Source, Source)
                && EqualsNormalized(action.Output, Output);
        }

        /// <summary>
        /// Returns enabled, applicable text actions in group-order followed by
        /// their original order within each group. Returned actions are clones
        /// so callers cannot mutate the persisted catalog accidentally.
        /// </summary>
        public static IReadOnlyList<CustomActionDefinition> GetActions(
            CustomActionFile catalog,
            string itemType,
            string direction)
        {
            var applicability = new CustomActionApplicabilityContext
            {
                ItemType = itemType,
                Direction = direction
            };

            return (catalog?.Groups ?? new CustomActionGroup[0])
                .Where(group => group != null)
                .OrderBy(group => group.Order)
                .SelectMany(group => group.Actions ?? new CustomActionDefinition[0])
                .Where(action => IsTextAction(action)
                    && !action.Disabled
                    && CustomActionApplicability.IsApplicable(action, applicability))
                .Select(action => action.Clone())
                .ToArray();
        }

        private static bool EqualsNormalized(string value, string expected)
        {
            return string.Equals(
                (value ?? "").Trim(),
                expected,
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
