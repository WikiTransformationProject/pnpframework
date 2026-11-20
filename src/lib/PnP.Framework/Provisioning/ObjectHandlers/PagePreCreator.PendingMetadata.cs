#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    public static partial class PagePreCreator
    {
        // written by LLM, 2026-11-14
        private sealed class PageCreationInfo
        {
            internal (string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist, int? itemId, long? createdUtcMs) Result { get; }
            internal IReadOnlyCollection<string> AvailableFields { get; }

            internal PageCreationInfo(
                (string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist, int? itemId, long? createdUtcMs) result,
                IEnumerable<string>? availableFields = null)
            {
                Result = result;
                AvailableFields = (availableFields ?? Array.Empty<string>()).Where(name =>
                    name == WikiTraccsPageCreationFields.ItemIdFieldName
                    || name == WikiTraccsPageCreationFields.CreatedUtcMsFieldName).ToArray();
            }
        }

        // written by LLM, 2026-11-14
        private static async Task<(string? pageUrl, List<Action<TokenParser, Web>> parserActions, bool? didExist, int? itemId, long? createdUtcMs)>
            CompletePreCreationAsync(Task<PageCreationInfo> firstTask, bool wasScheduled,
                Func<Task<PageCreationInfo>> readCurrentPage, BaseClientSidePage page,
                Action<PageCreationInfo> rememberCreation)
        {
            PageCreationInfo? first = null;
            try
            {
                first = await firstTask.ConfigureAwait(false);
            }
            catch when (wasScheduled)
            {
                // A failed earlier request has no receipt. Let this request try again.
            }
            var current = wasScheduled ? await readCurrentPage().ConfigureAwait(false) : first!;
            if (current.Result.createdUtcMs.HasValue)
            {
                rememberCreation(current);
            }
            var receipt = current.Result.createdUtcMs.HasValue ? current : first;
            if (!current.Result.itemId.HasValue || receipt?.Result.itemId != current.Result.itemId
                || !string.Equals(receipt?.Result.pageUrl, current.Result.pageUrl, StringComparison.OrdinalIgnoreCase))
            {
                receipt = null;
            }
            if (receipt?.Result.createdUtcMs is long createdUtcMs)
            {
                WikiTraccsPageCreationFields.FillTemplate(page.FieldValues, current.Result.itemId!.Value,
                    createdUtcMs, receipt.AvailableFields);
            }
            // For a repeated request, mark as "didExist".
            return (current.Result.pageUrl, current.Result.parserActions,
                wasScheduled ? true : current.Result.didExist, current.Result.itemId, receipt?.Result.createdUtcMs);
        }
    }
}
